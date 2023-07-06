pub mod excel_automation {

    use std::{ffi::c_void, ptr, mem::ManuallyDrop};
    use windows::{Win32::{System::{Com::*, Ole::*}, Foundation::DECIMAL}, core::*};

    const EXCEL_APP: GUID = GUID::from_values(0x00024500, 0x0000, 0x0000, [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46]);
    const NULL: GUID = GUID::from_values(0x00000000, 0x0000, 0x0000, [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]);
    const ACTIVE_WORKBOOK_ID: i32 = 308; // Method to get an active workbook interface
    const ACTIVE_WORKSHEET_ID: i32 = 307; // Method to get an interface to active worksheet  
    const CELL_RANGE_ID: i32 = 197; // Get cells within a certain range
    const CELL_VALUES_ID: i32 = 6; // Get value of cells

    pub fn insert_value() -> Result<()>
    {
        // Initialize a COM 
        unsafe { 
            CoInitializeEx(None, COINIT_APARTMENTTHREADED) 
        }?;

        let app_dispatch = ole_active_object()?;
        let workbook_dispacth = invoke_dispatch(app_dispatch, ACTIVE_WORKBOOK_ID)?;
        let worksheet_dispatch = invoke_dispatch(workbook_dispacth, ACTIVE_WORKSHEET_ID)?;
        let range_dispatch = invoke_specific_range_dispatch(worksheet_dispatch, CELL_RANGE_ID)?;
       let _set_value = invoke_set_value_dispatch(range_dispatch, CELL_VALUES_ID)?;
        unsafe {
            CoUninitialize()
        };
        Ok(())
    }

    pub fn invoke_set_value_dispatch(dispatch_interface: IDispatch, dispid: i32) -> Result<()> {            
        
        // preliminary variables for IDispacth interface initialization 

        let error_code: *mut u32 = Box::into_raw(Box::new(0 as u32));
        let exception_info: *mut EXCEPINFO = Box::into_raw(Box::new(EXCEPINFO::default()));
        let result_variant: *mut VARIANT = Box::into_raw(Box::new(VARIANT::default()));
        
        let variant1 = variant_initialize(None,  8, BSTR::from("Changed"));
        let mut rgargs: [VARIANT;1] = [variant1];
        let prgars = rgargs.as_mut_ptr();
        let mut params = DISPPARAMS::default();
        params.rgvarg = prgars;
        params.cArgs = 1;
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
            println!("Got here!");
            let result_box: Box<u32> = unsafe { Box::from_raw(error_code) };
            println!("{}", result_box);
            let error_message: Error = method_results.unwrap_err();
            return Err(error_message);
        }
    }

    pub fn variant_initialize(dec_val: Option<DECIMAL>, varenum: u16, value: BSTR ) -> VARIANT {
        let mut variant_make = unsafe { VariantInit() };
        let outer_variant_union = &mut variant_make.Anonymous;
        if let Some(dec_val) = dec_val {
            outer_variant_union.decVal = dec_val;
        } 
        let struct_inner = unsafe { &mut outer_variant_union.Anonymous };
        struct_inner.vt = VARENUM(varenum);
        struct_inner.Anonymous = VARIANT_0_0_0{ bstrVal: ManuallyDrop::new(value)};
        return variant_make;
    }

    pub fn invoke_specific_range_dispatch(dispatch_interface: IDispatch, dispid: i32) -> Result<IDispatch> {            
        
        // preliminary variables for IDispacth interface initialization 

        let error_code: *mut u32 = Box::into_raw(Box::new(0 as u32));
        let exception_info: *mut EXCEPINFO = Box::into_raw(Box::new(EXCEPINFO::default()));
        let result_variant: *mut VARIANT = Box::into_raw(Box::new(VARIANT::default()));
        let variant1 = variant_initialize(None,  8, BSTR::from("A2"));
        let variant2 = variant_initialize(None, 8, BSTR::from("A2"));
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

    pub fn invoke_dispatch(dispatch_interface: IDispatch, dispid: i32) -> Result<IDispatch> {            
        
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
}