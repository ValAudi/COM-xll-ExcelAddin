use std::{ffi::c_void, ptr};
use windows::{Win32::System::{Com::*, Ole::*}, core::*};

const EXCEL_APP: GUID = GUID::from_values(0x00024500, 0x0000, 0x0000, [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46]);

pub fn ole_active_object() -> Result<IDispatch> {

    // GetActiveObject from Ole Database registration

    let punknw: *mut Option<IUnknown> = Box::into_raw(Box::new(unsafe { std::mem::zeroed() }));
    let raw_ptr: *mut c_void = ptr::null_mut();
    let active_object_punknw = unsafe { GetActiveObject(&EXCEL_APP, raw_ptr, punknw) };
    if active_object_punknw.is_ok() {
        let active_object_interface: IUnknown = unsafe { Box::from_raw(punknw).unwrap() };
        let excel_dispatch: IDispatch = active_object_interface.cast()?;
        return Ok(excel_dispatch);   
    } else {
        let error_message: Error = active_object_punknw.unwrap_err();
        return Err(error_message);
    }
}
    

