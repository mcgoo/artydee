use com::{
    runtime::{init_apartment, ApartmentType},
    sys::{FAILED, HRESULT, IID, S_OK},
};
use log::{info, trace};
use std::{
    collections::BTreeMap, ffi::c_void, ptr::null_mut, ptr::NonNull, sync::Arc, sync::Mutex,
    thread, time::Duration,
};
use winapi::shared::wtypesbase::LPOLESTR;
use winapi::{
    ctypes::c_long,
    shared::{
        guiddef::REFIID,
        minwindef::{UINT, WORD},
        ntdef::{LCID, LONG, ULONG},
        winerror::{E_FAIL, E_NOTIMPL, E_POINTER},
        wtypes::VT_VARIANT,
        wtypes::{VARIANT_BOOL, VARTYPE},
    },
    um::{
        oaidl::{ITypeInfo, DISPID, DISPPARAMS, EXCEPINFO, SAFEARRAY, SAFEARRAYBOUND, VARIANT},
        oleauto::{SafeArrayAccessData, SafeArrayUnaccessData},
    },
};

mod dll;
mod interfaces;
mod registry;
pub mod variant;

use interfaces::*;

pub type IRtdServer = interfaces::IRtdServer;

pub use dll::{dll_can_unload_now, /*dll_get_class_object,*/ dll_main};
pub use registry::{dll_register_server, dll_unregister_server, get_relevant_registry_keys};

macro_rules! htry {
    ($expr:expr) => {
        let hr = $expr;
        //use crate::FAILED;
        if FAILED(hr) {
            return Err(hr);
        }
    };
}

// The CLSID of this RTD server. This GUID needs to be different for
// every RTD application.
// pub const CLSID_CAT_CLASS: IID = IID {
//     data1: 0x0aea1daa,
//     data2: 0xdab5,
//     data3: 0xfac1,
//     data4: [0x8f, 0x6a, 0x83, 0xdc, 0x88, 0x98, 0x0a, 0x64],
// };

// The CLSID for the IRtdServer interface Excel uses to call us. Being an Excel
// interface, this GUID never changes.
pub const IID_IRTDSERVER: IID = IID {
    data1: 0xEC0E6191,
    data2: 0xDB51,
    data3: 0x11D3,
    data4: [0x8F, 0x3E, 0x00, 0xC0, 0x4F, 0x36, 0x51, 0xB8],
};

pub const PROG_ID: &str = "Haka.PFX";

// the module handle for this dll
static mut _HMODULE: *mut c_void = null_mut();

// the typeinfo for IID_IRTDServer
static mut _ITYPEINFO: *mut ITypeInfo = null_mut();
//static mut foo: Option<ITypeInfo> = None; // this works

/// Get the ITypeInfo for IID_IRTDServer
fn get_itypeinfo(hinstance: *mut c_void) -> *mut ITypeInfo {
    use std::ffi::OsStr;
    use std::os::windows::prelude::*;
    use winapi::um::oaidl::ITypeLib;
    use winapi::um::oleauto::{LoadTypeLibEx, REGKIND_NONE};

    let mut modulename = unsafe { ::com::production::registration::get_dll_file_path(hinstance) };
    modulename += "\\1";
    let wmodulename = OsStr::new(&modulename)
        .encode_wide()
        .chain(Some(0).into_iter())
        .collect::<Vec<u16>>();
    let mut ptlib: *mut ITypeLib = null_mut(); // LPTYPELIB ptlib = NULL;
    unsafe {
        let hr = LoadTypeLibEx(wmodulename.as_ptr(), REGKIND_NONE, &mut ptlib);
        if FAILED(hr) {
            info!(
                "Could not load typelib from {}: HRESULT={:x}",
                modulename, hr
            );
            return null_mut();
        };
    }

    let mut ptinfo: *mut ITypeInfo = null_mut();

    unsafe {
        // Get type information for interface of the object.
        let hr = (*ptlib).GetTypeInfoOfGuid(
            &IID_IRTDSERVER as *const com::sys::GUID as *const winapi::shared::guiddef::GUID,
            &mut ptinfo,
        );
        if FAILED(hr) {
            info!("GetTypeInfoOfGuid() failed: HRESULT={:x}", hr);
            (*ptlib).Release();
            return null_mut();
        }
    }

    ptinfo
}

com::class! {
    // cat_data is the
    pub class BritishShortHairCat: IRtdServer(IDispatch) {
        cat_data: Arc<Mutex<CatData>>,
        // foo
        cat_guts: Arc<Mutex<CatGuts>>,

        body: Arc<Mutex<Option<Box<dyn RtdServer>>>>
    }

    impl IRtdServer for BritishShortHairCat {
        unsafe fn server_start(
            &self,
            /*[in]*/ callback_object: NonNull<NonNull<<IRTDUpdateEvent as com::Interface>::VTable>>,
            /*[out,retval]*/
            pf_res: *mut c_long,
        ) -> HRESULT {

            info!("server_start");
            (callback_object.as_ref().as_ref().PutHeartbeatInterval)(callback_object, 30000 );
            info!("got here");
            let mut cat_data = self.cat_data.lock().unwrap();
            let newarc = Arc::clone(&self.cat_guts);

            // // TODO: can this fail? if not, why not?
            let mut cat_guts = self.cat_guts.lock().unwrap();

            // // store the callback to excel
            cat_guts.update_event = Some(callback_object);
            let (tx,rx)=std::sync::mpsc::channel::<()>();
            drop(cat_guts);

            cat_data.shutdown = Some(tx);

            // TODO: store the JoinHandle so that the thread can be waited on
            // during shutdown
            let joinhandle = thread::spawn(move || {
                cat_loop(newarc, rx)
            });
            cat_data.cat_loop_joinhandle = Some(joinhandle);

            *pf_res=1; // success
            S_OK
        }
        unsafe fn connect_data(
            &self,
            /*[in]*/ topic_id: c_long,
            /*[in]*/ strings: *mut *mut SAFEARRAY,
            /*[in,out]*/ get_new_values: *mut VARIANT_BOOL,
            /*[out,retval]*/ pvar_out: *mut VARIANT,
        ) -> HRESULT {
            info!("connect_data: topic_id={} strings=? get_new_values={:x}",topic_id,*get_new_values);
            let mut cat_guts = self.cat_guts.lock().unwrap();
            cat_guts.connect_data(topic_id, strings, get_new_values, pvar_out)

        }
        unsafe fn refresh_data(
            &self,
            /*[in,out]*/ topic_count: *mut c_long,
            /*[out,retval]*/ parray_out: *mut *mut SAFEARRAY,
        ) -> HRESULT {
            info!("refresh_data: topic_count={}", *topic_count);
            let cat_guts = self.cat_guts.lock().unwrap();
            cat_guts.refresh_data(topic_count, parray_out)
        }

        unsafe fn disconnect_data(&self, /*[in]*/ topic_id: c_long) -> HRESULT {
            info!("disconnect_data: topic_id={}",topic_id);
            let mut cat_guts = self.cat_guts.lock().unwrap();
            cat_guts.disconnect_data(topic_id)
        }
        unsafe fn heartbeat(&self, /*[out,retval]*/ pf_res: *mut c_long) -> HRESULT {
            info!("heartbeat");
            *pf_res = 1;
            S_OK
        }
        unsafe fn server_terminate(&self) -> HRESULT {
            info!("server_terminate");
            let mut cat_data = self.cat_data.lock().unwrap();

            // drop our end of the shutdown notification channel
            cat_data.shutdown = None;

            // wait on the thread
            cat_data.cat_loop_joinhandle.take().map(std::thread::JoinHandle::join);

            S_OK
        }
    }

    impl IDispatch for BritishShortHairCat {
        unsafe fn get_type_info_count(&self, pctinfo: *mut UINT) -> HRESULT {
            trace!("get_type_info_count");
            *pctinfo = 0;
            S_OK
        }
        unsafe fn get_type_info(
            &self,
            _itinfo: UINT,
            _lcid: LCID,
            _pptinfo: *mut *mut ITypeInfo,
        ) -> HRESULT {
            trace!("get_type_info");
            E_NOTIMPL
        }
        unsafe fn get_ids_of_names(
            &self,
            _riid: REFIID,
            rgszNames: *mut LPOLESTR,
            cNames: UINT,
            _lcid: LCID,
            rgdispid: *mut DISPID,
        ) -> HRESULT {
            trace!("get_ids_of_names");

            //_ITYPEINFO = get_itypeinfo(_HMODULE);

            if _ITYPEINFO != null_mut() {
               (*_ITYPEINFO).GetIDsOfNames(rgszNames, cNames, rgdispid)
            } else {
                info!("get_ids_of_names running without the typelib which is never going to work");
                E_NOTIMPL
            }
        }
        unsafe fn invoke(
            &self,
            dispidMember: DISPID,
            _riid: REFIID,
            _lcid: LCID,
            wFlags: WORD,
            pdispparams: *mut DISPPARAMS,
            pvarResult: *mut VARIANT,
            pexcepinfo: *mut EXCEPINFO,
            puArgErr: *mut UINT,
        ) -> HRESULT {
            trace!("invoke");
            if _ITYPEINFO != null_mut() {
                (*_ITYPEINFO).Invoke(self as *const BritishShortHairCat as *mut _, dispidMember, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr)
            } else {
                info!("invoke running without the typelib which is never going to work");
                E_NOTIMPL
            }
        }
    }
}

impl BritishShortHairCat {
    pub fn set_something(&mut self, body: Box<dyn RtdServer>) {
        let mut opt = self.body.lock().unwrap();
        opt.replace(body);
    }
}

pub struct CatData {
    /// the shutdown channel
    shutdown: Option<std::sync::mpsc::Sender<()>>,

    /// cat_loop handle
    cat_loop_joinhandle: Option<std::thread::JoinHandle<()>>,
}

pub struct CatGuts {
    // update_event: *const IRTDUpdateEvent,
    update_event: Option<NonNull<NonNull<<IRTDUpdateEvent as com::Interface>::VTable>>>, // (callback_object.as_ref().as_ref().PutHeartbeatInterval)(callback_object, 1000 );

    // live topics
    topics: BTreeMap<c_long, Vec<String>>,
}

pub trait RtdServer {
    // callback object
    // fn update_notify(&self);

    // unsafe fn connect_data(
    //     &mut self,
    //     /*[in]*/ topic_id: c_long,
    //     /*[in]*/ _strings: *mut *mut SAFEARRAY,
    //     /*[in,out]*/ get_new_values: *mut VARIANT_BOOL,
    //     /*[out,retval]*/ _pvar_out: *mut VARIANT,
    // ) -> HRESULT;
    // unsafe fn refresh_data(
    //     &self,
    //     /*[in,out]*/ topic_count: *mut c_long,
    //     /*[out,retval]*/ parray_out: *mut *mut SAFEARRAY,
    // ) -> HRESULT;
    // unsafe fn disconnect_data(&mut self, /*[in]*/ topic_id: c_long) -> HRESULT;
}

impl CatGuts {
    // the IRTDUpdateEvent should have a lifetime of this CatGuts
    // and it should Addref and Release the pointer, or better,
    // store it in something that will do that automatically.
    // fn update_event(&self) -> &IRTDUpdateEvent {
    //     unsafe { &self.update_event.unwrap().as_ref() }
    // }

    // callback
    fn update_notify(&self) {
        info!("calling notify");
        let callback_object = self.update_event.unwrap();
        unsafe {
            (callback_object.as_ref().as_ref().UpdateNotify)(callback_object);
        }
    }

    unsafe fn connect_data(
        &mut self,
        /*[in]*/ topic_id: c_long,
        /*[in]*/ _strings: *mut *mut SAFEARRAY,
        /*[in,out]*/ get_new_values: *mut VARIANT_BOOL,
        /*[out,retval]*/ _pvar_out: *mut VARIANT,
    ) -> HRESULT {
        info!(
            "cat_guts connect_data: topic_id={} strings=? get_new_values={:x}",
            topic_id, *get_new_values
        );

        let mut sa = **_strings;
        let fields = decode_1d_safearray_of_variants_containing_strings(&mut sa);
        let fields = match fields {
            Ok(vs) => vs,
            Err(hr) => return hr,
        };

        self.topics.insert(topic_id, fields);
        S_OK
    }
    unsafe fn refresh_data(
        &self,
        /*[in,out]*/ topic_count: *mut c_long,
        /*[out,retval]*/ parray_out: *mut *mut SAFEARRAY,
    ) -> HRESULT {
        info!("cat_guts refresh_data");

        // make up some data and return it for every topic
        let now = chrono::Local::now().format(" %a %b %e %T %Y");

        let updated_topics = self
            .topics
            .iter()
            .map(|(topic, v)| {
                let mut data = v.join(",");
                data = data + &now.to_string();
                let data = variant::make_bstr(data);
                (*topic, data)
            })
            .collect::<Vec<_>>();

        let sa = match topic_updates_to_safearray(&updated_topics) {
            Ok(sa) => sa,
            Err(hr) => return hr,
        };

        *parray_out = sa;
        *topic_count = self.topics.len() as c_long; //yolo
        S_OK
    }
    unsafe fn disconnect_data(&mut self, /*[in]*/ topic_id: c_long) -> HRESULT {
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

    let timeout = Duration::from_millis(5000);
    // wait for updates to data and add relevant changes to the dirty list
    loop {
        match ctrl_chan.recv_timeout(timeout) {
            Ok(()) => {
                // nothing is supposed to send on this channel - the close down
                // signal is just dropping the transmitter
            }
            Err(std::sync::mpsc::RecvTimeoutError::Timeout) => {
                let cat_guts = newarc.lock().unwrap();
                cat_guts.update_notify();
            }
            Err(std::sync::mpsc::RecvTimeoutError::Disconnected) => {
                break;
            }
        }
    }
    info!("the worker thread has ended");
}

fn decode_1d_safearray_of_variants_containing_strings(
    sa: *mut SAFEARRAY,
) -> Result<Vec<String>, HRESULT> {
    unsafe {
        if sa == null_mut() {
            return Err(E_POINTER);
        }

        if (*sa).cDims != 1 {
            info!("got the wrong number of dims: {}", (*sa).cDims);
            return Err(E_FAIL);
        }

        let mut pvararr: *mut VARIANT = null_mut();
        let hr = SafeArrayAccessData(sa, &mut pvararr as *mut *mut VARIANT as *mut *mut _);
        if FAILED(hr) {
            info!("SafeArrayAccessData failed: hr={:x}", hr);

            return Err(hr);
        }

        let mut res = vec![];
        // not UB in this case since we are actually only using one dimension
        // the isize cast could wrap
        for i in 0..(*sa).rgsabound[0].cElements as isize {
            let x = (*(pvararr.offset(i))).n1.n2().n3.bstrVal();
            let s = widestring::U16CStr::from_ptr_str(*x);
            res.push(s.to_string_lossy());
        }
        if FAILED(SafeArrayUnaccessData(sa)) {
            info!("SafeArrayUnaccessData failed");
        }

        Ok(res)
    }
}

/// convert a Vec of (topic, new value of topic) into a safearray for excel
///
/// This does not check that the variants are of the copyable types. This
/// should be fine for types that are accepted by Excel.
fn topic_updates_to_safearray(data: &Vec<(c_long, VARIANT)>) -> Result<*mut SAFEARRAY, HRESULT> {
    let mut bounds = [
        SAFEARRAYBOUND {
            cElements: 2,
            lLbound: 0,
        },
        SAFEARRAYBOUND {
            cElements: data.len() as ULONG,
            lLbound: 0,
        },
    ];
    unsafe {
        let sa = SafeArrayCreate(VT_VARIANT as VARTYPE, 2, bounds.as_mut_ptr());

        let mut sa_idx: LONG = 0;
        for topic in data {
            let mut index1: [LONG; 2] = [0, sa_idx];
            let mut v1 = variant::make_i32(topic.0);
            htry!(SafeArrayPutElement(
                sa,
                index1.as_mut_ptr(),
                &mut v1 as *mut VARIANT as *mut _
            ));

            let mut index2: [LONG; 2] = [1, sa_idx];
            // This is not a safe copy. VT_ARRAY etc might not work,
            // not that Excel handles it anyway
            let mut v2 = topic.1;
            htry!(SafeArrayPutElement(
                sa,
                index2.as_mut_ptr(),
                &mut v2 as *mut VARIANT as *mut _
            ));

            sa_idx += 1;
        }

        Ok(sa)
    }
}

extern "system" {
    pub fn SafeArrayCreate(
        vt: VARTYPE,
        cDims: UINT,
        rgsabound: *mut SAFEARRAYBOUND,
    ) -> *mut SAFEARRAY;
    pub fn SafeArrayPtrOfIndex(
        psa: *mut SAFEARRAY,
        rgIndices: *mut LONG,
        ppvData: *mut *mut c_void,
    ) -> HRESULT;
    pub fn SafeArrayPutElement(
        psa: *mut SAFEARRAY,
        rgIndices: *mut LONG,
        pv: *mut c_void,
    ) -> HRESULT;
}

#[cfg(test)]
mod tests {
    use crate::variant::*;
    use crate::*;
    use winapi::shared::wtypes::{VT_BSTR, VT_I4, VT_R8};

    // get the variant at [r1,r2] from the array
    pub fn from_2d_safearray<'a>(
        sa: &'a *mut SAFEARRAY,
        r1: isize,
        r2: isize,
    ) -> Result<&'a VARIANT, HRESULT> {
        let mut indices: [LONG; 2] = [r1 as LONG, r2 as LONG];

        let mut v: *mut VARIANT = std::ptr::null_mut();

        unsafe {
            htry!(SafeArrayPtrOfIndex(
                *sa,
                indices.as_mut_ptr(),
                &mut v as *mut *mut VARIANT as *mut *mut _
            ));
            Ok(&*v)
        }
    }

    #[test]
    fn safearray_of_strings_to_vec_of_string() {
        use crate::decode_1d_safearray_of_variants_containing_strings;
        use oaidl::{SafeArrayExt, VariantWrapper, Variants};

        let v: Vec<Box<dyn VariantWrapper>> = vec![
            Box::new(Variants::String("One".to_string())),
            Box::new(Variants::String("Another".to_string())),
        ];
        let sa = v.into_iter().into_safearray().unwrap();

        let data = decode_1d_safearray_of_variants_containing_strings(sa.as_ptr()).unwrap();
        assert_eq!(data, vec!["One".to_string(), "Another".to_string()]);
    }

    #[test]
    fn create_safearray_for_topic_responses() {
        let v1 = crate::variant::make_f64(99.9);
        let v2 = crate::variant::make_bstr("beaster");

        let data = vec![(1, v1), (7, v2)];

        let sa = topic_updates_to_safearray(&data).unwrap();
        unsafe {
            let v = from_2d_safearray(&sa, 0, 0).unwrap();
            assert_eq!(v.n1.n2().vt, VT_I4 as VARTYPE);
            assert_eq!(*v.n1.n2().n3.lVal(), 1);

            let v = from_2d_safearray(&sa, 1, 0).unwrap();
            assert_eq!(v.n1.n2().vt, VT_R8 as VARTYPE);
            assert_eq!(*v.n1.n2().n3.dblVal(), 99.9); // yolo

            let v = from_2d_safearray(&sa, 0, 1).unwrap();
            assert_eq!(v.n1.n2().vt, VT_I4 as VARTYPE);
            assert_eq!(*v.n1.n2().n3.lVal(), 7);

            let v = from_2d_safearray(&sa, 1, 1).unwrap();
            assert_eq!(v.n1.n2().vt, VT_BSTR as VARTYPE);
            assert_eq!(variant_bstr_to_string(v).unwrap(), "beaster");
        }
    }
}
