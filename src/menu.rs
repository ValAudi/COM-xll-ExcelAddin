use windows::{core::*, Win32::System::{Variant::*, Com::*}};
use crate::variant::*;

pub const COMMAND_BARS_ID: i32 = 1439; // Method to get an interface to command bars 
pub const COMMAND_BARS_ADD: i32 = 1610809346; // Method Adds a command Bar to Command Bars
pub const COMMAND_BARS_COUNT: i32 = 1610809347; // Method to get count of command bars
pub const COMMAND_BARS_ITEM: i32 = 0; // Method to get/set command bar item interface  
pub const COMMAND_BAR_NAME: i32 = 1610874894; // Method to get/set command bar item name 
pub const COMMAND_BAR_TYPE: i32 = 1610874909; // Sets/gets command Bar Type to the newly created command Bar
pub const COMMAND_BAR_CAPTION: i32 = 1610874883; // Method to get/set command bar caption
pub const COMMAND_BAR_DESCRIPTION: i32 = 1610874888; // Method to get/set command bar description text
pub const COMMAND_BAR_TOOLTIP: i32 = 1610874917; // Method to get/set command bar tooltip 
pub const COMMAND_BAR_TAG: i32 = 1610874915; // Method to get/set command bar tag
pub const COMMAND_BAR_POSITION: i32 = 1610874899; // Method to get/set command bar position
pub const COMMAND_BAR_VISIBILITY: i32 = 1610874910; // Method Adds a command Bar Type to the newly created command Bar
pub const COMMAND_BAR_CONTROL: i32 = 1610874887; // Method to get command bar control interface
pub const COMMAND_BAR_CONTROLS: i32 = 1610874883; // Method to get command bar controls interface 
pub const COMMAND_BARS_CONTROLS_ADD: i32 = 1610809344; // Method Adds a command Bar to Command Bars
pub const COMMAND_BAR_CONTROLS_COUNT: i32 = 1610809345; // Method to get count of command bar controls 
pub const COMMAND_BAR_POPUP_CONTROLS: i32 = 1610940417; // Method to get command bar controls interface 
pub const COMMAND_BAR_POPUP_VISIBILITY: i32 = 1610874921; // Method to delete command bar 
pub const CONTROL_INTERFACE_ID: i32 = 1610874885; // Method to get control dispatch interface
pub const COMMAND_BAR_ONACTION: i32 = 1610874906; // Method to get/set command bar on action function callback 
pub const COMMAND_BAR_BUTTON_PARAMETER: i32 = 1610874909; // Method to get/set command bar button parameter 
pub const COMMAND_BAR_BUTTON_TYPE: i32 = 1610874920; // Method to get/set command bar button type
pub const COMMAND_BAR_BUTTON_HYPERLINK_TYPE: i32 = 1610940428; // Method to get/set command bar button hyperlink 
pub const COMMAND_BAR_BUTTON_STYLE: i32 = 1610940426; // Method to get/set command bar button style 
pub const COMMAND_BAR_BUTTON_STATE: i32 = 1610940424; // Method to get/set command bar button state 
pub const COMMAND_BAR_DELETE: i32 = 1610874884; // Method to delete command bar 

const NULL: GUID = GUID::from_values(0x00000000, 0x0000, 0x0000, [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]);

pub fn get_int_values(dispatch_interface: &IDispatch, dispid: i32) -> Result<i32> {            
        
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
        let count: i32  = unsafe { result_box.Anonymous.Anonymous.Anonymous.intVal}; 
        return Ok(count);
    } else {
        let error_message: Error = method_results.unwrap_err();
        return Err(error_message);
    }
}

pub fn get_text_values(dispatch_interface: &IDispatch, dispid: i32) -> Result<String> {            
        
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
        let name: String  = unsafe { result_box.Anonymous.Anonymous.Anonymous.bstrVal.clone() }.to_string();  
        return Ok(name);
    } else {
        let error_message: Error = method_results.unwrap_err();
        return Err(error_message);
    }
}

pub fn get_visibility(dispatch_interface: &IDispatch, dispid: i32) -> Result<bool> {            
        
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
        let visbility: bool  = unsafe { result_box.Anonymous.Anonymous.Anonymous.boolVal }.into();  
        return Ok(visbility);
    } else {
        let error_message: Error = method_results.unwrap_err();
        return Err(error_message);
    }
}

pub fn set_visibility(dispatch_interface: &IDispatch, dispid: i32) -> Result<()> {            
        
    // preliminary variables for IDispacth interface initialization 

    let error_code: *mut u32 = Box::into_raw(Box::new(0 as u32));
    let exception_info: *mut EXCEPINFO = Box::into_raw(Box::new(EXCEPINFO::default()));
    let result_variant: *mut VARIANT = Box::into_raw(Box::new(VARIANT::default()));
    let visibility = variant_initialize(None,  VT_BOOL, VariantType::VT_BOOL(true.into()));
    let mut rgargs: [VARIANT;1] = [visibility];
    let prgars: *mut VARIANT = rgargs.as_mut_ptr();
    let named_param = Box::into_raw(Box::new(-3 as i32));
    let mut params = DISPPARAMS::default();
    params.rgvarg = prgars;
    params.rgdispidNamedArgs = named_param;
    params.cArgs = 1;
    params.cNamedArgs = 1;
    let args: *const DISPPARAMS = Box::into_raw(Box::new(params));
    
    let method_results = unsafe { dispatch_interface.Invoke 
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
        return Ok(());
    } else {
        let error_message: Error = method_results.unwrap_err();
        return Err(error_message);
    }
}

pub fn get_command_bars_interface(dispatch_interface: &IDispatch, dispid: i32) -> Result<IDispatch> {            
        
    // preliminary variables for IDispacth interface initialization 

    let error_code: *mut u32 = Box::into_raw(Box::new(0 as u32));
    let exception_info: *mut EXCEPINFO = Box::into_raw(Box::new(EXCEPINFO::default()));
    let result_variant: *mut VARIANT = Box::into_raw(Box::new(VARIANT::default()));

    let menu_name = variant_initialize(None,  VT_BSTR, VariantType::VT_BSTR(BSTR::from("National Accounts")));
    let menu_position = variant_initialize(None,  VT_INT, VariantType::VT_INT(1));
    let menu_bar = variant_initialize(None,  VT_BOOL, VariantType::VT_BOOL(true.into()));
    let temporary = variant_initialize(None,  VT_BOOL, VariantType::VT_BOOL(true.into()));

    let mut rgargs: [VARIANT;4] = [menu_name, menu_position, menu_bar, temporary];
    let prgars: *mut VARIANT = rgargs.as_mut_ptr();
    let named_param = Box::into_raw(Box::new(0 as i32));
    let mut params = DISPPARAMS::default();
    params.rgdispidNamedArgs = named_param;
    params.rgvarg = prgars;
    params.cArgs = 4;
    params.cNamedArgs = 0;
    let args: *const DISPPARAMS = Box::into_raw(Box::new(params));
    
    let method_results = unsafe { 
        dispatch_interface.Invoke 
        ( 
            dispid, 
            &NULL,
            0,
            DISPATCH_METHOD, 
            args, 
            Some(result_variant), 
            Some(exception_info), 
            Some(error_code)
        ) 
    };

    if method_results.is_ok() {
        println!("Got here");
        let result_box: Box<VARIANT> = unsafe { Box::from_raw(result_variant) };
        let dispatch_interface: IDispatch = unsafe { result_box.Anonymous.Anonymous.Anonymous.pdispVal.clone() }.take().unwrap(); 
        return Ok(dispatch_interface);
    } else {
        let error_message: Error = method_results.unwrap_err();
        return Err(error_message);
    }
}

pub fn get_command_bar_controls_interface(dispatch_interface: &IDispatch, dispid: i32) -> Result<IDispatch> {            
        
    // preliminary variables for IDispacth interface initialization 

    let error_code: *mut u32 = Box::into_raw(Box::new(0 as u32));
    let exception_info: *mut EXCEPINFO = Box::into_raw(Box::new(EXCEPINFO::default()));
    let result_variant: *mut VARIANT = Box::into_raw(Box::new(VARIANT::default()));
    let control_type = variant_initialize(None,  VT_INT, VariantType::VT_INT(1)); // MSOControlType BUtton
    let control_id = variant_initialize(None,  VT_NULL, VariantType::VT_NULL());
    let control_parameter = variant_initialize(None,  VT_NULL, VariantType::VT_NULL());
    let control_before = variant_initialize(None,  VT_NULL, VariantType::VT_NULL());
    let control_temporary = variant_initialize(None,  VT_BOOL, VariantType::VT_BOOL(true.into()));
    let mut rgargs: [VARIANT;5] = [control_type, control_id, control_parameter, control_before, control_temporary];
    let prgars: *mut VARIANT = rgargs.as_mut_ptr();
    let mut params = DISPPARAMS::default();
    params.rgvarg = prgars;
    params.cArgs = 5;
    let args: *const DISPPARAMS = Box::into_raw(Box::new(params));
    
    let method_results = unsafe { 
        dispatch_interface.Invoke 
        ( 
            dispid, 
            &NULL,
            0,
            DISPATCH_METHOD, 
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

pub fn command_bar_delete(dispatch_interface: &IDispatch, dispid: i32) -> Result<()> {            
        
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
            DISPATCH_METHOD, 
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
        let error_message: Error = method_results.unwrap_err();
        return Err(error_message);
    }
}