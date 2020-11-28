use com::sys::IID;
use log::info;
use std::ffi::c_void;
use std::os::raw::c_long;
use std::ptr::NonNull;
use winapi::shared::minwindef::BOOL;

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
    artydee::make_body(|| Box::new(MuppetDataFeed {}));

    artydee::dll_main(hinstance, fdw_reason, _reserved)
}

struct MuppetDataFeed {
    //
}

impl artydee::RtdServer for MuppetDataFeed {
    unsafe fn server_start(
        &self, // /*[in]*/ callback_object: IRTDUpdateEvent,
        /*[in]*/
        callback_object: NonNull<NonNull<<artydee::IRTDUpdateEvent as com::Interface>::VTable>>,
        /*[out,retval]*/
        pf_res: *mut c_long,
    ) -> com::sys::HRESULT {
        info!("in muppet's server_start!");
        winapi::shared::winerror::S_OK
    }

    unsafe fn connect_data(
        &self,
        /*[in]*/ topic_id: winapi::ctypes::c_long,
        /*[in]*/ strings: *mut *mut winapi::um::oaidl::SAFEARRAY,
        /*[in,out]*/ get_new_values: *mut winapi::shared::wtypes::VARIANT_BOOL,
        /*[out,retval]*/ pvar_out: *mut winapi::um::oaidl::VARIANT,
    ) -> com::sys::HRESULT {
        todo!()
    }

    unsafe fn refresh_data(
        &self,
        /*[in,out]*/ topic_count: *mut winapi::ctypes::c_long,
        /*[out,retval]*/ parray_out: *mut *mut winapi::um::oaidl::SAFEARRAY,
    ) -> com::sys::HRESULT {
        todo!()
    }

    unsafe fn disconnect_data(
        &self,
        /*[in]*/ topic_id: winapi::ctypes::c_long,
    ) -> com::sys::HRESULT {
        todo!()
    }

    unsafe fn heartbeat(
        &self,
        /*[out,retval]*/ pf_res: *mut winapi::ctypes::c_long,
    ) -> com::sys::HRESULT {
        todo!()
    }

    unsafe fn server_terminate(&self) -> com::sys::HRESULT {
        todo!()
    }
}

#[no_mangle]
unsafe extern "stdcall" fn DllGetClassObject(
    class_id: *const ::com::sys::CLSID,
    iid: *const ::com::sys::IID,
    result: *mut *mut c_void,
) -> ::com::sys::HRESULT {
    // artydee::dll_get_class_object(class_id, iid, result)
    assert!(
        !class_id.is_null(),
        "class id passed to DllGetClassObject should never be null"
    );

    let class_id = &*class_id;
    if class_id == &CLSID_DOG_CLASS {
        let instance/*: com::production::ClassAllocation<artydee::BritishShortHairCatClassFactory>*/ =
            <artydee::BritishShortHairCat as ::com::production::Class>::Factory::allocate();

        //let body = Box::new(MuppetDataFeed {});

        // (**instance).set_something(body);
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

#[cfg(test)]
mod tests {
    #[test]
    fn it_works() {
        assert_eq!(2 + 2, 4);
    }
}
