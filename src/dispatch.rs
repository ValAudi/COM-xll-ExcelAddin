use windows::{Win32::System::{Com::*, Variant::*}, core::*};
use crate::variant::*;

const NULL: GUID = GUID::from_values(0x00000000, 0x0000, 0x0000, [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]);

pub fn get_dispatch_interface(dispatch_interface: IDispatch, dispid: i32) -> Result<IDispatch> {            
        
    // preliminary variables for IDispacth interface initialization 

    let error_code: *mut u32 = Box::into_raw(Box::new(0 as u32));
    let exception_info: *mut EXCEPINFO = Box::into_raw(Box::new(EXCEPINFO::default()));
    let result_variant: *mut VARIANT = Box::into_raw(Box::new(VARIANT::default()));
    let args: *const DISPPARAMS = Box::into_raw(Box::new(DISPPARAMS::default()));
    
    let method_results = unsafe { 
        dispatch_interface.Invoke 
        ( 
            dispid, 
            &NULL,
            0,
            DISPATCH_PROPERTYGET, 
            args, 
            Some(result_variant), 
            Some(exception_info), 
            Some(error_code)
        ) 
    };

    if method_results.is_ok() {
        let result_box: Box<VARIANT> = unsafe { Box::from_raw(result_variant) };
        let dispatch_interface: IDispatch = unsafe { result_box.Anonymous.Anonymous.Anonymous.pdispVal.clone() }.take().unwrap(); 
        return Ok(dispatch_interface);
    } else {
        let error_message: Error = method_results.unwrap_err();
        return Err(error_message);
    }
}

pub fn get_range_interface(dispatch_interface: IDispatch, dispid: i32) -> Result<IDispatch> {            
        
    // preliminary variables for IDispacth interface initialization 

    let error_code: *mut u32 = Box::into_raw(Box::new(0 as u32));
    let exception_info: *mut EXCEPINFO = Box::into_raw(Box::new(EXCEPINFO::default()));
    let result_variant: *mut VARIANT = Box::into_raw(Box::new(VARIANT::default()));
    let variant1 = variant_initialize(None,  VT_BSTR, VariantType::VT_BSTR(BSTR::from("D2")));
    let variant2 = variant_initialize(None, VT_BSTR, VariantType::VT_BSTR(BSTR::from("E3")));
    let mut rgargs: [VARIANT;2] = [variant1, variant2];
    let prgars: *mut VARIANT = rgargs.as_mut_ptr();
    let mut params = DISPPARAMS::default();
    params.rgvarg = prgars;
    params.cArgs = 2;
    let args: *const DISPPARAMS = Box::into_raw(Box::new(params));
    
    let method_results = unsafe { 
        dispatch_interface.Invoke 
        ( 
            dispid, 
            &NULL,
            0,
            DISPATCH_PROPERTYGET, 
            args, 
            Some(result_variant), 
            Some(exception_info), 
            Some(error_code)
        ) 
    };

    if method_results.is_ok() {
        let result_box: Box<VARIANT> = unsafe { Box::from_raw(result_variant) };
        let dispatch_interface: IDispatch = unsafe { result_box.Anonymous.Anonymous.Anonymous.pdispVal.clone() }.take().unwrap(); 
        return Ok(dispatch_interface);
    } else {
        let error_message: Error = method_results.unwrap_err();
        return Err(error_message);
    }
}
