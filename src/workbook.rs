use windows::{Win32::System::Com::*, core::*};

pub const ACTIVE_WORKBOOK_ID: i32 = 308; // Method to get an active workbook interface
const NULL: GUID = GUID::from_values(0x00000000, 0x0000, 0x0000, [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]);


// Refine This method


pub fn get_sheetname(dispatch_interface: IDispatch, dispid: i32) -> Result<String> {            
        
    // preliminary variables for IDispacth interface initialization 

    let error_code: *mut u32 = Box::into_raw(Box::new(0 as u32));
    let exception_info: *mut EXCEPINFO = Box::into_raw(Box::new(EXCEPINFO::default()));
    let result_variant: *mut VARIANT = Box::into_raw(Box::new(VARIANT::default()));
    let args: *const DISPPARAMS = Box::into_raw(Box::new(DISPPARAMS::default()));
    
    let method_results = unsafe { dispatch_interface.Invoke 
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
        let dispatch_interface: String  = unsafe { result_box.Anonymous.Anonymous.Anonymous.bstrVal.clone() }.to_string(); 
        return Ok(dispatch_interface);
    } else {
        let error_message: Error = method_results.unwrap_err();
        return Err(error_message);
    }
}