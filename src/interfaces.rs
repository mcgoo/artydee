use com::interfaces::iunknown::IUnknown;
use com::sys::HRESULT;
use winapi::ctypes::c_long;
use winapi::shared::guiddef::REFIID;
use winapi::shared::{
    minwindef::{UINT, WORD},
    ntdef::LCID,
    wtypes::VARIANT_BOOL,
    wtypesbase::LPOLESTR,
};
use winapi::um::oaidl::{ITypeInfo, DISPID, DISPPARAMS, EXCEPINFO, SAFEARRAY, VARIANT};

com::interfaces! {

    #[uuid("00020400-0000-0000-C000-000000000046")]
    pub unsafe interface IDispatch: IUnknown {
        unsafe fn GetTypeInfoCount(&self, pctinfo: *mut UINT) -> HRESULT;

        unsafe fn get_type_info(
            &self,
            i_tinfo: UINT,
            lcid: LCID,
            pp_tinfo: *mut *mut ITypeInfo,
        ) -> HRESULT;
        unsafe fn get_ids_of_names(
            &self,
            riid: REFIID,
            rgsz_names: *mut LPOLESTR,
            c_names: UINT,
            lcid: LCID,
            rg_disp_id: *mut DISPID,
        ) -> HRESULT;
        unsafe fn invoke(
            &self,
            disp_id_member: DISPID,
            riid: REFIID,
            lcid: LCID,
            w_flags: WORD,
            p_disp_params: *mut DISPPARAMS,
            p_var_result: *mut VARIANT,
            p_excep_info: *mut EXCEPINFO,
            pu_arg_err: *mut UINT,
        ) -> HRESULT;
    }

    #[uuid("ec0e6191-db51-11d3-8f3e-00c04f3651b8")]
    pub unsafe interface IRtdServer: IDispatch {
        //
        // Raw methods provided by interface
        //

        pub unsafe fn server_start(
            &self,
            /*[in]*/ callback_object: IRTDUpdateEvent,
            /*[out,retval]*/
            pfres: *mut c_long,
        ) -> HRESULT;
        pub unsafe fn connect_data(
            &self,
            /*[in]*/ topic_id: c_long,
            /*[in]*/ strings: *mut *mut SAFEARRAY,
            /*[in,out]*/ get_new_values:*mut VARIANT_BOOL,
            /*[out,retval]*/ pvar_out: *mut VARIANT,
        ) -> HRESULT;
        pub unsafe fn refresh_data(
            &self,
            /*[in,out]*/ topic_count: *mut c_long,
            /*[out,retval]*/ parray_out: *mut *mut SAFEARRAY,
        ) -> HRESULT;
        pub unsafe fn disconnect_data(&self, /*[in]*/ topic_id: c_long) -> HRESULT;
        pub unsafe fn heartbeat(&self, /*[out,retval]*/ pf_res: *mut c_long) -> HRESULT;
        pub unsafe fn server_terminate(&self) -> HRESULT;
    }

    #[uuid("a43788c1-d91b-11d3-8f39-00c04f3651b8")]
    pub unsafe interface IRTDUpdateEvent: IDispatch {
        unsafe fn update_notify(&self) -> HRESULT;
        unsafe fn get_heartbeat_interval(
            &self,
            /*[out,retval]*/ _pl_ret_val: *mut c_long,
        ) -> HRESULT;
        unsafe fn put_heartbeat_interval(&self, /*[in]*/ _pl_ret_val: c_long) -> HRESULT; // lRetVal?
        unsafe fn disconnect(&self) -> HRESULT;
    }
}
