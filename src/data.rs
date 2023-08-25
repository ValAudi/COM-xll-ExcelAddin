use windows::{Win32::System::{Com::*, Variant::*}, core::*};
use crate::variant::*;

pub const CELL_VALUES_ID: i32 = 6; // Get value of cells
const NULL: GUID = GUID::from_values(0x00000000, 0x0000, 0x0000, [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]);

pub fn get_cell_data(dispatch_interface: IDispatch, dispid: i32) -> Result<()> {            
        
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
        println!("{:#?}", unsafe { result_box.Anonymous.Anonymous.vt} );
        // let value_array = unsafe { result_box.Anonymous.Anonymous.Anonymous.bstrVal.clone()}.to_string(); 
        // let value_array = unsafe { result_box.Anonymous.Anonymous.Anonymous.dblVal}; 
        // println!("{:#?}", value_array );
        return Ok(());
    } else {
        let result_box: Box<u32> = unsafe { Box::from_raw(error_code) };
        println!("{}", result_box);
        let error_message: Error = method_results.unwrap_err();
        return Err(error_message);
    }
}

pub fn set_cell_value(dispatch_interface: IDispatch, dispid: i32) -> Result<()> {            
        
    // preliminary variables for IDispacth interface initialization 

    let error_code: *mut u32 = Box::into_raw(Box::new(0 as u32));
    let exception_info: *mut EXCEPINFO = Box::into_raw(Box::new(EXCEPINFO::default()));
    let result_variant: *mut VARIANT = Box::into_raw(Box::new(VARIANT::default()));
    
    let variant1 = variant_initialize(None,  VT_BSTR, VariantType::VT_BSTR(BSTR::from("Changed Again")));
    let mut rgargs: [VARIANT;1] = [variant1];
    let prgars = rgargs.as_mut_ptr();
    let named_param = Box::into_raw(Box::new(-3 as i32));
    let mut params = DISPPARAMS::default();
    params.rgvarg = prgars;
    params.rgdispidNamedArgs = named_param;
    params.cArgs = 1;
    params.cNamedArgs = 1;
    let args: *const DISPPARAMS = Box::into_raw(Box::new(params));
    
    let method_results = unsafe { 
        dispatch_interface.Invoke 
        ( 
            dispid, 
            &NULL,
            0,
            DISPATCH_PROPERTYPUT, 
            args, 
            Some(result_variant), 
            Some(exception_info), 
            Some(error_code)
        ) 
    };

    if method_results.is_ok() {
        let result_box: Box<VARIANT> = unsafe { Box::from_raw(result_variant) };
        println!("{:#?}", unsafe { result_box.Anonymous.Anonymous.vt} );
        // let value_array = unsafe { result_box.Anonymous.Anonymous.Anonymous.bstrVal.clone()}.to_string(); 
        return Ok(());
    } else {
        println!("{:#?}", unsafe {Box::from_raw(exception_info)} );
        let error_message: Error = method_results.unwrap_err();
        return Err(error_message);
    }
}