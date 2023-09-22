// use std::{fs::File, ffi::c_void};
// use std::io::Read;
// use windows::Win32::System::LibraryLoader::GetModuleHandleA;
use windows::{core::*, Win32::System::{Variant::*, Com::*}};

use crate::variant::*;

pub const COMMAND_BARS_ID: i32 = 1439; // Method to get an interface to command bars  
pub const COMMAND_BARS_COUNT: i32 = 1610809347; // Method to get count of command bars  
pub const COMMAND_BARS_ADD: i32 = 1610809346; // Method Adds a command Bar to Command Bars 
pub const COMMAND_BAR_TYPE: i32 = 1610874909; // Method Adds a command Bar Type to the newly created command Bar
pub const COMMAND_BAR_VISIBILITY: i32 = 1610874910; // Method Adds a command Bar Type to the newly created command Bar
pub const COMMAND_BARS_ITEM: i32 = 0; // Method to get command bar item interface  
pub const COMMAND_BAR_NAME: i32 = 1610874894; // Method to get command bar item name  4294962293
pub const COMMAND_BAR_CONTROL: i32 = 1610874887; // Method to get command bar control interface
pub const COMMAND_BAR_CONTROLS: i32 = 1610874883; // Method to get command bar controls interface
pub const CONTROL_INTERFACE_ID: i32 = 1610874885; // Method to get control dispatch interface 
pub const COMMAND_BAR_POSITION: i32 = 1610874899; // Method to get command bar position 
pub const UIRIBBONXML: &str = "excelRibbon.xml";
const _IRIBBONEXTENSIBILITY: GUID = GUID::from_values(0x000C0396, 0x0000, 0x0000, [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46]);
const NULL: GUID = GUID::from_values(0x00000000, 0x0000, 0x0000, [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]);

pub fn get_position(dispatch_interface: &IDispatch, dispid: i32) -> Result<i32> {            
        
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

pub fn get_count(dispatch_interface: &IDispatch, dispid: i32) -> Result<i32> {            
        
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

pub fn get_name(dispatch_interface: &IDispatch, dispid: i32) -> Result<String> {            
        
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

// fn read_ribbon_customizations() -> windows::core::Result<IStream> {
//     let stream = unsafe { CreateStreamOnHGlobal(None, true) }?;

//     // Load file into in-memory then into IStream
//     let mut buffer = Vec::new();
//     if let Ok(mut file) = File::open(UIRIBBONXML){
//         if let Ok(file_size) = file.read_to_end(&mut buffer){
//             let data: *const c_void = buffer.as_ptr() as *const c_void;
//             let _stream_res = unsafe { stream.Write(data, file_size as u32, None) };
//         } else {
//             println!("Error Creating buffer from file data and size!");
//         };
//     } else {
//         println!("Error reading UI Customization file!");
//     };      

//     Ok(stream)
// }

// pub fn load_customized_ribbon() -> Result<()> {
//     println!("----step 0");
//     let res_ribbon_framework_interface: Result<IUIFramework> = unsafe { CoCreateInstance(&UIRibbonFramework, None, CLSCTX_ALL)}; 
//     if res_ribbon_framework_interface.is_ok() {
//         println!("----step 1");
//         let ribbon_framework_interface: IUIFramework = res_ribbon_framework_interface.expect("unable to unwrap the UI framework interface");
//         let excel_hinstance = unsafe { GetModuleHandleA(None) };
//         if excel_hinstance.is_ok() {
//             println!("----step 2");
//             let excel_module: HMODULE = excel_hinstance.unwrap();
//             unsafe { ribbon_framework_interface.LoadUI(excel_module, resourcename) };
//             // let it = _intermed.
//         } else {
//             let error_message = excel_hinstance.err();
//             println!("{:#?}", error_message);
//         }
//         // let pointer: *mut *const c_void = Box::into_raw(Box::new(std::ptr::null()));
//         // let ri: HRESULT = unsafe { intermediate.query(&_IRIBBONEXTENSIBILITY, pointer) };
//         // if ri.is_ok() {
//         //     println!("----step 2");
//         //     // let _intermed = ri.unwrap();
//         //     // let it = _intermed.
//         // } else {
//         //     let error_message = ri.0;
//         //     println!("{:#?}", error_message);
//         // }
//     } else {
//         let error_message = res_ribbon_framework_interface.unwrap_err();
//         println!("{:#?}", error_message.to_string());
//     }

    
//     // println!("----step 1");
//     // let iribbon_dispatch: IDispatch = ribbon_ext_interface.cast()?;
//     // println!("----step 2");
//     // let iribbon_dispatch_typeinfo: ITypeInfo = unsafe { iribbon_dispatch.GetTypeInfo(0, 0)? };
//     // println!("----step 3");
//     // let funcs = unsafe { iribbon_dispatch_typeinfo.GetTypeAttr() }?;
//     // println!("----step 4");
//     // let f = unsafe { *funcs };
//     // println!("{:#?}", f.cFuncs);
//     // if let Ok(custom_stream_data) = read_ribbon_customizations(){

//     // } else {

//     // }
//     Ok(())
// }
