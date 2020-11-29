use com::{
    runtime::{init_apartment, ApartmentType},
    sys::{HRESULT, IID},
};
use log::info;
use std::{
    collections::BTreeMap,
    ffi::c_void,
    os::raw::c_long,
    ptr::NonNull,
    sync::{
        mpsc::RecvTimeoutError::{Disconnected, Timeout},
        Arc, Mutex,
    },
    thread,
    thread::JoinHandle,
    time::Duration,
};
use winapi::{
    shared::{minwindef::BOOL, winerror::S_OK},
    um::oaidl::{SAFEARRAY, VARIANT},
};

// The CLSID of this RTD server. This GUID needs to be different for
// every RTD application.
pub const CLSID_DOG_CLASS: IID = IID {
    data1: 0xf99a1daa,
    data2: 0xdab5,
    data3: 0xfac1,
    data4: [0x8f, 0x6a, 0x83, 0xdc, 0x88, 0x98, 0x0a, 0x64],
};

const PROG_ID: &str = "Haka.PFY";

#[no_mangle]
extern "system" fn DllMain(
    hinstance: *mut c_void,
    fdw_reason: u32,
    _reserved: *mut c_void,
) -> BOOL {
    const DLL_PROCESS_ATTACH: u32 = 1;
    if fdw_reason == DLL_PROCESS_ATTACH {
        // this function specifically says it can be called from DllMain
        win_dbg_logger::rust_win_dbg_logger_init_info();
        info!("muppet loaded.");
    }

    // TODO: do this a different way
    artydee::make_body(|| Box::new(MuppetDataFeed::default()));

    artydee::dll_main(hinstance, fdw_reason, _reserved)
}

#[derive(Default)]
struct MuppetDataFeed {
    //
    cat_data: Arc<Mutex<CatData>>,
    // foo
    cat_guts: Arc<Mutex<CatGuts>>,
}

pub struct CatData {
    /// the shutdown channel
    shutdown: Option<std::sync::mpsc::Sender<()>>,

    /// cat_loop handle
    cat_loop_joinhandle: Option<std::thread::JoinHandle<()>>,
}

pub struct CatGuts {
    // update_event: *const IRTDUpdateEvent,
    update_event: Option<NonNull<NonNull<<artydee::IRTDUpdateEvent as com::Interface>::VTable>>>, // (callback_object.as_ref().as_ref().PutHeartbeatInterval)(callback_object, 1000 );

    // live topics
    topics: BTreeMap<c_long, Vec<String>>,
}

impl CatGuts {
    // callback
    fn update_notify(&self) {
        info!("calling notify");
        let callback_object = self.update_event.unwrap();
        unsafe {
            (callback_object.as_ref().as_ref().UpdateNotify)(callback_object);
        }
    }

    fn connect_data(
        &mut self,
        topic_id: c_long,
        strings: &[&str],
        get_new_values: bool,
    ) -> Result<Option<VARIANT>, HRESULT> {
        info!(
            "cat_guts connect_data: topic_id={} strings=? get_new_values={}",
            topic_id, get_new_values
        );

        let fields = strings
            .iter()
            .map(|s| s.to_string())
            .collect::<Vec<String>>();
        self.topics.insert(topic_id, fields);

        Ok(None)
    }
    unsafe fn refresh_data(
        &self,
        /*[in,out]*/ topic_count: *mut c_long,
        /*[out,retval]*/ parray_out: *mut *mut SAFEARRAY,
    ) -> HRESULT {
        info!("cat_guts refresh_data");

        // make up some data and return it for every topic
        let now = chrono::Local::now().format(" %a %b %e %T %Y");

        let updated_topics = self.topics.iter().map(|(topic, v)| {
            let mut data = v.join(",");
            data = data + &now.to_string();
            let data = artydee::variant::make_bstr(data);
            (*topic, data)
        });

        let sa = match artydee::topic_updates_to_safearray(updated_topics) {
            Ok(sa) => sa,
            Err(hr) => return hr,
        };

        *parray_out = sa;
        *topic_count = self.topics.len() as c_long; //yolo
        S_OK
    }
    fn disconnect_data(&mut self, topic_id: c_long) -> HRESULT {
        self.topics.remove(&topic_id);
        S_OK
    }
}

unsafe impl Send for CatGuts {}

impl Default for CatGuts {
    fn default() -> Self {
        Self {
            update_event: None,
            topics: BTreeMap::new(),
        }
    }
}
impl Default for CatData {
    fn default() -> Self {
        Self {
            shutdown: None,
            cat_loop_joinhandle: None,
        }
    }
}

fn cat_loop(newarc: Arc<Mutex<CatGuts>>, ctrl_chan: std::sync::mpsc::Receiver<()>) {
    info!("starting the worker thread");
    init_apartment(ApartmentType::Multithreaded).unwrap();

    let timeout = Duration::from_millis(1000);
    // wait for updates to data and add relevant changes to the dirty list
    loop {
        match ctrl_chan.recv_timeout(timeout) {
            Ok(()) => {
                // nothing is supposed to send on this channel - the close down
                // signal is just dropping the transmitter
            }
            Err(Timeout) => {
                let cat_guts = newarc.lock().unwrap();
                cat_guts.update_notify();
            }
            Err(Disconnected) => {
                break;
            }
        }
    }
    info!("the worker thread has ended");
}

impl artydee::RtdServer for MuppetDataFeed {
    unsafe fn server_start(
        &self,
        /*[in]*/
        callback_object: NonNull<NonNull<<artydee::IRTDUpdateEvent as com::Interface>::VTable>>,
    ) -> Result<bool, HRESULT> {
        info!("in muppet's server_start!");

        (callback_object.as_ref().as_ref().PutHeartbeatInterval)(callback_object, 30000);
        info!("got here");
        let mut cat_data = self.cat_data.lock().unwrap();
        let newarc = Arc::clone(&self.cat_guts);

        // // TODO: can this fail? if not, why not?
        let mut cat_guts = self.cat_guts.lock().unwrap();

        // // store the callback to excel
        cat_guts.update_event = Some(callback_object);
        let (tx, rx) = std::sync::mpsc::channel::<()>();
        drop(cat_guts);

        cat_data.shutdown = Some(tx);

        // store the JoinHandle so that the thread can be waited on during shutdown
        let joinhandle = thread::spawn(move || cat_loop(newarc, rx));
        cat_data.cat_loop_joinhandle = Some(joinhandle);

        Ok(true)
    }

    fn connect_data(
        &self,
        topic_id: c_long,
        strings: &[&str],
        get_new_values: bool,
    ) -> Result<Option<VARIANT>, com::sys::HRESULT> {
        // info!(
        //     "connect_data: topic_id={:?} strings={} get_new_values={:x}",
        //     topic_id, strings, get_new_values
        // );
        let mut cat_guts = self.cat_guts.lock().unwrap();
        cat_guts.connect_data(topic_id, strings, get_new_values)
    }

    unsafe fn refresh_data(
        &self,
        /*[in,out]*/ topic_count: *mut c_long,
        /*[out,retval]*/ parray_out: *mut *mut winapi::um::oaidl::SAFEARRAY,
    ) -> com::sys::HRESULT {
        info!("refresh_data: topic_count={}", *topic_count);
        let cat_guts = self.cat_guts.lock().unwrap();
        cat_guts.refresh_data(topic_count, parray_out)
    }

    fn disconnect_data(&self, topic_id: c_long) -> com::sys::HRESULT {
        info!("disconnect_data: topic_id={}", topic_id);
        let mut cat_guts = self.cat_guts.lock().unwrap();
        cat_guts.disconnect_data(topic_id)
    }

    fn heartbeat(&self) -> Result<bool, HRESULT> {
        info!("heartbeat");
        Ok(true)
    }

    fn server_terminate(&self) -> com::sys::HRESULT {
        info!("server_terminate");
        let mut cat_data = self.cat_data.lock().unwrap();

        // drop our end of the shutdown notification channel
        cat_data.shutdown = None;

        // wait on the thread
        cat_data.cat_loop_joinhandle.take().map(JoinHandle::join);

        S_OK
    }
}

#[no_mangle]
unsafe extern "stdcall" fn DllGetClassObject(
    class_id: *const ::com::sys::CLSID,
    iid: *const ::com::sys::IID,
    result: *mut *mut c_void,
) -> ::com::sys::HRESULT {
    //artydee::dll_get_class_object(class_id, iid, result)
    assert!(
        !class_id.is_null(),
        "class id passed to DllGetClassObject should never be null"
    );

    let class_id = &*class_id;
    if class_id == &CLSID_DOG_CLASS {
        let instance =
            <artydee::BritishShortHairCat as ::com::production::Class>::Factory::allocate();
        instance.QueryInterface(&*iid, result)
    } else {
        ::com::sys::CLASS_E_CLASSNOTAVAILABLE
    }
}

#[no_mangle]
extern "stdcall" fn DllRegisterServer() -> ::com::sys::HRESULT {
    info!("DllRegisterServer");
    artydee::dll_register_server(&mut artydee::get_relevant_registry_keys(
        &PROG_ID,
        &CLSID_DOG_CLASS,
    ))
}

#[no_mangle]
extern "stdcall" fn DllUnregisterServer() -> ::com::sys::HRESULT {
    info!("DllUnregisterServer");
    artydee::dll_unregister_server(&mut artydee::get_relevant_registry_keys(
        &PROG_ID,
        &CLSID_DOG_CLASS,
    ))
}

#[no_mangle]
extern "stdcall" fn DllCanUnloadNow() -> ::com::sys::HRESULT {
    info!("DllCanUnloadNow()");

    // if there have been any calls to LockServer (which is not declared
    // in this dll) or any objects are live this dll cannot be unloaded

    use com::sys::S_FALSE;
    S_FALSE
}
