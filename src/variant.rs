use com::sys::HRESULT;
use std::mem::MaybeUninit;
use winapi::shared::winerror::E_FAIL;
use winapi::shared::wtypes::{VARTYPE, VT_ARRAY, VT_BSTR, VT_I4, VT_R8};
use winapi::um::oaidl::SAFEARRAY;
use winapi::um::oleauto::{SysAllocString, VariantInit};
use winapi::{ctypes::c_long, um::oaidl::VARIANT};

pub fn make_empty() -> VARIANT {
    unsafe {
        let mut v1 = MaybeUninit::uninit();
        VariantInit(v1.as_mut_ptr());
        v1.assume_init()
    }
}

pub fn make_i32(d: c_long) -> VARIANT {
    let mut v1 = make_empty();
    unsafe {
        v1.n1.n2_mut().vt = VT_I4 as VARTYPE;
        *v1.n1.n2_mut().n3.lVal_mut() = d;
    }
    v1
}
pub fn make_f64(d: f64) -> VARIANT {
    let mut v1 = make_empty();
    unsafe {
        v1.n1.n2_mut().vt = VT_R8 as VARTYPE;
        *v1.n1.n2_mut().n3.dblVal_mut() = d;
    }
    v1
}
pub fn make_bstr<S: AsRef<str>>(d: S) -> VARIANT {
    let mut v1 = make_empty();
    unsafe {
        v1.n1.n2_mut().vt = VT_BSTR as VARTYPE;
        let wide = widestring::U16CString::from_str(d).unwrap();
        let bstr = SysAllocString(wide.as_ptr());
        *v1.n1.n2_mut().n3.bstrVal_mut() = bstr;
    }
    v1
}

/// wrap a safearray in a variant
///
/// (I don't know of any situation where this is usable since Excel RTDs
/// can't return an array.)
pub fn make_array(psa: *mut SAFEARRAY) -> VARIANT {
    let mut v1 = make_empty();
    unsafe {
        v1.n1.n2_mut().vt = VT_ARRAY as VARTYPE;
        *v1.n1.n2_mut().n3.parray_mut() = psa;
    }
    v1
}

// convert a variant containing a bstr to a rust string
pub fn variant_bstr_to_string(v: &VARIANT) -> Result<String, HRESULT> {
    unsafe {
        if v.n1.n2().vt != VT_BSTR as VARTYPE {
            return Err(E_FAIL);
        }
        let x = v.n1.n2().n3.bstrVal();
        let s = widestring::U16CStr::from_ptr_str(*x);
        Ok(s.to_string_lossy())
    }
}
