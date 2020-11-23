// lightly hacked on registry manipulation code copied in
// from the com crate. Altered to write to HKCU instead of
// HKLM/HKCR

use com::sys::{
    RegCloseKey, RegCreateKeyExA, RegDeleteKeyA, RegSetValueExA, ERROR_SUCCESS, FAILED, GUID, HKEY,
    HRESULT, LSTATUS, SELFREG_E_CLASS, S_OK,
};
use log::debug;
use std::convert::TryInto;
use std::ffi::{c_void, CString};
use std::str;

/// get the keys to set in the registry
pub(crate) fn get_relevant_registry_keys() -> Vec<RegistryKeyInfo> {
    use super::{CLSID_CAT_CLASS, _HMODULE};

    let file_path = unsafe { ::com::production::registration::get_dll_file_path(_HMODULE) };
    let guid = guid_to_string(&CLSID_CAT_CLASS);
    let key = HKEY_CURRENT_USER_ROOT;
    let prefix = r"Software\Classes\";
    vec![
        RegistryKeyInfo::new(
            key,
            &format!("{}{}", prefix, "Haka.PFX"),
            "",
            "hakafeed.hakafeed",
        ),
        RegistryKeyInfo::new(key, &format!("{}{}", prefix, "Haka.PFX\\CLSID"), "", &guid),
        RegistryKeyInfo::new(key, &format!("{}CLSID\\{}", prefix, &guid), "", "Haka.PFX"),
        RegistryKeyInfo::new(
            key,
            &format!("{}CLSID\\{}\\InprocServer32", prefix, &guid),
            "",
            &file_path,
        ),
    ]
}

/// Register the supplied keys with the registry
#[doc(hidden)]
#[inline]
pub fn dll_register_server(relevant_keys: &mut Vec<RegistryKeyInfo>) -> HRESULT {
    debug!("dll_register_server: {:?}", relevant_keys);
    println!("dll_register_server: {:?}", relevant_keys);

    let hr = register_keys(relevant_keys);
    if FAILED(hr) {
        dll_unregister_server(relevant_keys);
    }

    hr
}

/// Unregister the supplied keys with the registry
#[doc(hidden)]
#[inline]
pub fn dll_unregister_server(relevant_keys: &mut Vec<RegistryKeyInfo>) -> HRESULT {
    relevant_keys.reverse();
    unregister_keys(relevant_keys)
}

#[doc(hidden)]
#[derive(Debug)]
pub struct RegistryKeyInfo {
    key: HKEY,
    key_path: CString,
    key_value_name: CString,
    key_value_data: CString,
}

#[doc(hidden)]
impl RegistryKeyInfo {
    pub fn new(
        key: HKEY,
        key_path: &str,
        key_value_name: &str,
        key_value_data: &str,
    ) -> RegistryKeyInfo {
        RegistryKeyInfo {
            key,
            key_path: CString::new(key_path).unwrap(),
            key_value_name: CString::new(key_value_name).unwrap(),
            key_value_data: CString::new(key_value_data).unwrap(),
        }
    }
}

#[doc(hidden)]
pub fn register_keys(registry_keys_to_add: &Vec<RegistryKeyInfo>) -> HRESULT {
    for key_info in registry_keys_to_add.iter() {
        let result = add_class_key(&key_info);
        if result as u32 != ERROR_SUCCESS {
            return SELFREG_E_CLASS;
        }
    }

    S_OK
}

#[doc(hidden)]
pub fn unregister_keys(registry_keys_to_remove: &Vec<RegistryKeyInfo>) -> HRESULT {
    let mut hr = S_OK;
    for key_info in registry_keys_to_remove.iter() {
        let result = remove_class_key(&key_info);
        if result as u32 != ERROR_SUCCESS {
            hr = SELFREG_E_CLASS;
        }
    }

    hr
}

//const HKEY_CLASSES_ROOT: HKEY = 0x8000_0000 as HKEY;
pub const HKEY_CURRENT_USER_ROOT: HKEY = 0x8000_0001 as HKEY;
const KEY_ALL_ACCESS: u32 = 0x000F_003F;
const REG_OPTION_NON_VOLATILE: u32 = 0x00000000;
fn create_class_key(key_info: &RegistryKeyInfo) -> Result<HKEY, LSTATUS> {
    let mut hk_result = std::ptr::null_mut::<c_void>();
    let lp_class = std::ptr::null_mut::<u8>();
    let lp_security_attributes = std::ptr::null_mut::<c_void>();
    let lpdw_disposition = std::ptr::null_mut::<u32>();
    let result = unsafe {
        RegCreateKeyExA(
            key_info.key,
            key_info.key_path.as_ptr(),
            0,
            lp_class,
            REG_OPTION_NON_VOLATILE,
            KEY_ALL_ACCESS,
            lp_security_attributes,
            &mut hk_result as *mut _,
            lpdw_disposition,
        )
    };
    if result as u32 != ERROR_SUCCESS {
        return Err(result);
    }

    Ok(hk_result)
}

const REG_SZ: u32 = 1;
fn set_class_key(key_handle: HKEY, key_info: &RegistryKeyInfo) -> Result<HKEY, LSTATUS> {
    let key_value_name_ptr = match key_info.key_value_name.as_bytes().len() {
        0 => std::ptr::null(),
        _ => key_info.key_value_name.as_ptr(),
    };

    let result = unsafe {
        RegSetValueExA(
            key_handle,
            //   key_info.key_value_name.as_ptr(),
            key_value_name_ptr,
            0,
            REG_SZ,
            key_info.key_value_data.as_ptr() as *const u8,
            key_info
                .key_value_data
                .to_bytes_with_nul()
                .len()
                .try_into()
                .unwrap(),
        )
    };

    debug!(
        "res={} key:{:x} key_path:{:?} val_name:{:?}({:x}) val_data:{:?}",
        result,
        key_info.key as u32,
        key_info.key_path,
        key_info.key_value_name,
        key_value_name_ptr as u64,
        key_info.key_value_data
    );
    if result as u32 != ERROR_SUCCESS {
        return Err(result);
    }

    Ok(key_handle)
}

fn add_class_key(key_info: &RegistryKeyInfo) -> LSTATUS {
    let key_handle = match create_class_key(key_info) {
        Ok(key_handle) => key_handle,
        Err(e) => return e,
    };

    let key_handle = match set_class_key(key_handle, key_info) {
        Ok(key_handle) => key_handle,
        Err(e) => return e,
    };

    unsafe { RegCloseKey(key_handle) }
}

fn remove_class_key(key_info: &RegistryKeyInfo) -> LSTATUS {
    unsafe { RegDeleteKeyA(key_info.key, key_info.key_path.as_ptr()) }
}

fn guid_to_string(guid: &GUID) -> String {
    format!(
        "{{{:08X}-{:04X}-{:04X}-{:02X}{:02X}-{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}}}",
        guid.data1,
        guid.data2,
        guid.data3,
        guid.data4[0],
        guid.data4[1],
        guid.data4[2],
        guid.data4[3],
        guid.data4[4],
        guid.data4[5],
        guid.data4[6],
        guid.data4[7],
    )
}

#[test]
fn guid_to_string_format() {
    assert_eq!(
        "{08EA1DAA-DAB5-FAC1-8F6A-83DC88980A64}",
        guid_to_string(&com::IID {
            data1: 0x08ea1daa,
            data2: 0xdab5,
            data3: 0xfac1,
            data4: [0x8f, 0x6a, 0x83, 0xdc, 0x88, 0x98, 0x0a, 0x64],
        })
    );
}
