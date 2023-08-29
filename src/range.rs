use std::{slice, ffi::c_void};
use windows::{Win32::System::{Com::*, Ole::*, Variant::*}, core::*};
use crate::variant::*;

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

    let _variant1 = variant_initialize(None,  VT_BSTR, VariantType::VT_BSTR(BSTR::from("Changed")));

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
    
    // Single Element Insert
    // let index: [i32; 2] = [0, 0];
    // let rgindices = index.as_ptr();
    // let res = unsafe {SafeArrayPutElement(safe_array, rgindices, Box::into_raw(Box::new(variant1)) as *const c_void)};
    // if res.is_ok() {
    //     println!("{:#?}", safe_array);
    // } else {
    //     let error_message: Error = res.unwrap_err();
    //     println!("{}", error_message.to_string());
    // } 

    // Multi-element insert
    let mut empty_variant: VARIANT = unsafe {std::mem::zeroed()};
    let pointer: *mut *mut c_void = &mut empty_variant as *mut _ as *mut *mut c_void;
    let safearray: *mut SAFEARRAY = safe_array;
    let r = unsafe { SafeArrayAccessData(safearray, pointer)};
    if r.is_ok() {
        let mut vec_variant: Vec<VARIANT> = Vec::new();
        vec_variant.push(variant_initialize(None,  VT_BSTR, VariantType::VT_BSTR(BSTR::from("Changed2"))));
        vec_variant.push(variant_initialize(None,  VT_BSTR, VariantType::VT_BSTR(BSTR::from("Changed3"))));
        vec_variant.push(variant_initialize(None,  VT_BSTR, VariantType::VT_BSTR(BSTR::from("Changed4"))));
        vec_variant.push(variant_initialize(None,  VT_BSTR, VariantType::VT_BSTR(BSTR::from("Changed5"))));

        unsafe {
            for i in 0..vec_variant.len() {
                let curr_variant = vec_variant[i].clone();
                *pointer.offset(i  as isize) =  Box::into_raw(Box::new(curr_variant)) as *mut _ as *mut c_void;             
            }          
        }
        let t = unsafe {std::ptr::read(*pointer.offset(2) as *mut VARIANT)};
        println!("{:#?}", unsafe{t.Anonymous.Anonymous.Anonymous.bstrVal.clone()});
        let s = unsafe {SafeArrayUnaccessData(safearray)};
        if s.is_ok() {
            // println!("Unacess data");
        } else {
            println!("Entered Error block");
            let error_message: Error = s.unwrap_err();
            println!("{}", error_message.to_string());
        }
    } else {
        let error_message: Error = r.unwrap_err();
        println!("{}", error_message.to_string());
    }
    println!("Got here!!!!");
    println!("--------------------------------------------------------------");


    // Build a variant from the single element insert safe_array and return
    // let var_safe_array = variant_initialize(None, VARENUM(8204), VariantType::VT_ARRAY(safe_array));
    // var_safe_array

    // Build an element from a multi-element insert safearray and return
    let var_safearray = variant_initialize(None, VARENUM(8204), VariantType::VT_ARRAY(safearray));
    // println!("{:#?}", unsafe {var_safearray.Anonymous.Anonymous.vt});
    // println!("{:#?}", unsafe {var_safearray.Anonymous.Anonymous.Anonymous.parray});
    println!("--------------------------------------------------------------");
    let variant_safearray = unsafe {var_safearray.Anonymous.Anonymous.Anonymous.parray};
    let const_array = variant_safearray.cast_const();
    let range_data = unsafe { SafeArrayLock(const_array) };
    if range_data.is_ok() {
        println!("{:#?}", unsafe {*const_array}.cDims);
        println!("{:#?}", unsafe {*const_array}.cLocks);
        println!("{:#?}", unsafe {*const_array}.cbElements);
        println!("{:#?}", unsafe {*const_array}.fFeatures);
        let ptr_data = unsafe { *const_array }.pvData as *mut VARIANT;
        let full_data = unsafe { slice::from_raw_parts_mut(ptr_data, 4) }; 
        for i in 0..4 {
            println!("{:#?}", unsafe { (full_data[i]).Anonymous.Anonymous.Anonymous.bstrVal.clone()}.to_string() );
        }
    } else {
        println!("Came here!");
        let error_message: Error = range_data.unwrap_err();
        println!("{}", error_message.to_string());
    }
    var_safearray
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