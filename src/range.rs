use std::{slice, ffi::c_void};
use windows::{Win32::System::{Com::*, Ole::*, Variant::*}, core::*};
use crate::variant::{init_str_variant, init_array_variant};

pub const CELL_RANGE_ID: i32 = 197; // Get cells within a certain range
pub const RANGE_VALUES_ID: i32 = 6; // Get value of cells
const NULL: GUID = GUID::from_values(0x00000000, 0x0000, 0x0000, [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]);

pub fn get_range_array(dispatch_interface: IDispatch, dispid: i32) -> Result<*mut SAFEARRAY> {            
        
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
        let value_array = unsafe { result_box.Anonymous.Anonymous.Anonymous.parray }; 
        return Ok(value_array);
    } else {
        let error_message: Error = method_results.unwrap_err();
        return Err(error_message);
    }
}

pub fn get_range_data(array: *mut SAFEARRAY) {

    // Get range data from safe array
    let const_array = array.cast_const();
    let range_data = unsafe { SafeArrayLock(const_array) };
    if range_data.is_ok() {
        let ptr_data = unsafe { *const_array }.pvData as *mut VARIANT;
        let full_data = unsafe { slice::from_raw_parts_mut(ptr_data, 4) }; 
        for i in 0..4 {
            println!("{:#?}", unsafe { (full_data[i]).Anonymous.Anonymous.Anonymous.bstrVal.clone()}.to_string() );
        }
    } else {
        let error_message: Error = range_data.unwrap_err();
        println!("{}", error_message.to_string());
    }
}  

pub fn set_range_data() -> VARIANT {

    let variant1 = init_str_variant(None,  8, BSTR::from("Changed"));

    // Creating a safe array Bound
    let mut sab: SAFEARRAYBOUND = SAFEARRAYBOUND::default();
    sab.lLbound = 0;
    sab.cElements = 2;

    let mut sab1: SAFEARRAYBOUND = SAFEARRAYBOUND::default();
    sab1.lLbound = 0;
    sab1.cElements = 2;

    let bounds = [sab, sab1];
    // let bounds = [sab];
    let rgsabound = bounds.as_ptr();

    // Creating a safearray using the OLE method SafeArrayCreate
    let safe_array = unsafe { SafeArrayCreate(VARENUM(12), 2, rgsabound) };
    println!("{:#?}", unsafe {*safe_array}); 
    
    let index: [i32; 2] = [0, 0];
    let rgindices = index.as_ptr();
    let res = unsafe {SafeArrayPutElement(safe_array, rgindices, Box::into_raw(Box::new(variant1)) as *const c_void)};
    if res.is_ok() {
        println!("{:#?}", safe_array);
    } else {
        let error_message: Error = res.unwrap_err();
        println!("{}", error_message.to_string());
    } 
    println!("{:#?}", unsafe {*safe_array}); 

    let pointer: *mut *mut c_void = std::ptr::null_mut();
    let safearray: *const SAFEARRAY = safe_array.cast_const();
    let r = unsafe { SafeArrayAccessData(safearray, pointer)};
    if r.is_ok() {
        let t = unsafe {*pointer} as *mut VARIANT;
        println!("{:#?}", unsafe{(*t).Anonymous.Anonymous.vt});
    } else {
        let error_message: Error = r.unwrap_err();
        println!("{}", error_message.to_string());
    }

    // Build a variant from the safe-array
    let var_safe_array = init_array_variant(None, 8204, safe_array);
    var_safe_array
}  

pub fn set_range_array(dispatch_interface: IDispatch, dispid: i32, variant:VARIANT) -> Result<()> {            
        
    // preliminary variables for IDispacth interface initialization 

    let error_code: *mut u32 = Box::into_raw(Box::new(0 as u32));
    let exception_info: *mut EXCEPINFO = Box::into_raw(Box::new(EXCEPINFO::default()));
    let result_variant: *mut VARIANT = Box::into_raw(Box::new(VARIANT::default()));
    let mut rgargs: [VARIANT;1] = [variant];
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
        let _result_box: Box<VARIANT> = unsafe { Box::from_raw(result_variant) };
        return Ok(());
    } else {
        println!("{:#?}", unsafe {Box::from_raw(exception_info)} );
        let error_message: Error = method_results.unwrap_err();
        return Err(error_message);
    }
}