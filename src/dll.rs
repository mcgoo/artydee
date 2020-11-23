use log::info;
use winapi::shared::minwindef::{BOOL, TRUE};

use crate::*;

#[no_mangle]
extern "system" fn DllMain(
    hinstance: *mut ::std::ffi::c_void,
    fdw_reason: u32,
    _reserved: *mut ::std::ffi::c_void,
) -> BOOL {
    const DLL_PROCESS_ATTACH: u32 = 1;
    if fdw_reason == DLL_PROCESS_ATTACH {
        // this function specifically says it can be called from DllMain
        win_dbg_logger::rust_win_dbg_logger_init_info();
        info!("loaded.");
        unsafe {
            _HMODULE = hinstance;
            _ITYPEINFO = get_itypeinfo(hinstance);
        }
    }
    TRUE
}

#[no_mangle]
unsafe extern "stdcall" fn DllGetClassObject(
    class_id: *const ::com::sys::CLSID,
    iid: *const ::com::sys::IID,
    result: *mut *mut ::std::ffi::c_void,
) -> ::com::sys::HRESULT {
    assert!(
        !class_id.is_null(),
        "class id passed to DllGetClassObject should never be null"
    );

    let class_id = &*class_id;
    if class_id == &CLSID_CAT_CLASS {
        let instance = <BritishShortHairCat as ::com::production::Class>::Factory::allocate();
        instance.QueryInterface(&*iid, result)
    } else {
        ::com::sys::CLASS_E_CLASSNOTAVAILABLE
    }
}

#[no_mangle]
extern "stdcall" fn DllRegisterServer() -> ::com::sys::HRESULT {
    info!("DllRegisterServer");
    dll_register_server(&mut get_relevant_registry_keys())
}

#[no_mangle]
extern "stdcall" fn DllUnregisterServer() -> ::com::sys::HRESULT {
    info!("DllUnregisterServer");
    dll_unregister_server(&mut get_relevant_registry_keys())
}

#[no_mangle]
extern "stdcall" fn DllCanUnloadNow() -> ::com::sys::HRESULT {
    info!("DllCanUnloadNow()");

    // if there have been any calls to LockServer (which is not declared
    // in this dll) or any objects are live this dll cannot be unloaded

    use com::sys::S_FALSE;
    S_FALSE
}
