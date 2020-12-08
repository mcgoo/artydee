use com::{
    runtime::{init_apartment, ApartmentType},
    sys::{FAILED, HRESULT, IID, S_OK},
};
use log::{info, trace};
use std::{
    collections::BTreeMap, ffi::c_void, ptr::null_mut, ptr::NonNull, sync::Arc, sync::Mutex,
    thread, time::Duration,
};
use winapi::shared::minwindef::{BOOL, TRUE};
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

// The CLSID of this RTD server. This GUID needs to be different for
// every RTD application.
pub const CLSID_DOG_CLASS: IID = IID {
    data1: 0xf99a1daa,
    data2: 0xdab5,
    data3: 0xfac1,
    data4: [0x8f, 0x6a, 0x83, 0xdc, 0x88, 0x98, 0x0a, 0x64],
};

const PROG_ID: &str = "Haka.PFY";

//use crate::*;

#[no_mangle]
extern "system" fn DllMain(
    hinstance: *mut ::std::ffi::c_void,
    fdw_reason: u32,
    _reserved: *mut ::std::ffi::c_void,
) -> BOOL {
    const DLL_PROCESS_ATTACH: u32 = 1;
    if fdw_reason == DLL_PROCESS_ATTACH {
        //     // this function specifically says it can be called from DllMain
        win_dbg_logger::rust_win_dbg_logger_init_info();
        //     info!("loaded.");
        //     unsafe {
        //         _HMODULE = hinstance;
        //         _ITYPEINFO = get_itypeinfo(hinstance);
        //     }
    }
    artydee::dll_main(hinstance, fdw_reason, _reserved)

    // const DLL_PROCESS_ATTACH: u32 = 1;
    // if fdw_reason == DLL_PROCESS_ATTACH {
    //     // this function specifically says it can be called from DllMain
    //     win_dbg_logger::rust_win_dbg_logger_init_info();
    //     info!("loaded.");
    //     unsafe {
    //         _HMODULE = hinstance;
    //         _ITYPEINFO = get_itypeinfo(hinstance);
    //     }
    // }
    // TRUE
}

#[no_mangle]
unsafe extern "stdcall" fn DllGetClassObject(
    class_id: *const ::com::sys::CLSID,
    iid: *const ::com::sys::IID,
    result: *mut *mut ::std::ffi::c_void,
) -> ::com::sys::HRESULT {
    // artydee::dll_get_class_object(class_id, iid, result)
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

#[cfg(test)]
mod tests {
    #[test]
    fn it_works() {
        assert_eq!(2 + 2, 4);
    }
}
