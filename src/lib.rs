use com::INationalAccounts;
use windows::{Win32::{System::{Com::*, Variant::*}, Foundation::*, UI::WindowsAndMessaging::MessageBoxA}, core::{Result, IUnknown, GUID, HRESULT, PCSTR, s, ComInterface}};
// use xlcall::LPXLOPER;
// use xlexcel4::*;
use crate::variant::*;

pub mod automation;
pub mod rot;
pub mod workbook;
pub mod dispatch;
pub mod variant;
pub mod worksheet;
pub mod range;
pub mod data;
pub mod ribbon;
pub mod menu;
pub mod com;
pub mod registry;
pub mod typelib;
pub mod xlcall;
pub mod xlexcel4;
pub mod xlvariant;
pub mod xllregister;
pub mod regkeys;

// #[no_mangle]
// #[allow(non_snake_case, unused_variables)]
// extern "stdcall" fn xlAutoOpen() -> i32{
//     // Function implementation goes here
//     // You can return an integer as in the original signature.
//     // If this function is a placeholder, you can just return 0.\
//     let calc = Command::new("calc.exe").spawn().unwrap();
//     xllregister::reg_xll_functions();
//     1
// }

// #[no_mangle]
// #[allow(non_snake_case, unused_variables)]
// extern "stdcall" fn xlAutoRegister(Variant::from_str("Xlladdin").as_mut_xloper(): xlcall::LPXLOPER) -> xlcall::XLOPER {
//     let res = set_sheetname();
//     if res.is_ok() {
//         return 0;
//     } else {
//         return 1;    
//     }
// }

// #[no_mangle]
// #[allow(non_snake_case, unused_variables)]
// extern "stdcall" fn ChangeSheetName() {
//     let res = set_sheetname();
// }

// #[no_mangle]
// #[allow(non_snake_case, unused_variables)]
// extern "stdcall" fn GetSumValentine(param1: XLOPER, param2: XLOPER) -> XLOPER {
//     let res =  unsafe {param1.val.w + param2.val.w};
//     let ret = xlvariant::Variant::from_int(res).as_mut_xloper();
//     unsafe {*ret}
// }

#[no_mangle]
#[allow(non_snake_case, unused_variables)]
extern "system" fn DllMain(
    dll_module: HINSTANCE,
    call_reason: u32,
    reserved: *const u32)
    -> BOOL
{
    const DLL_PROCESS_ATTACH: u32 = 1;
    const DLL_PROCESS_DETACH: u32 = 0;

    match call_reason {
        DLL_PROCESS_ATTACH => (), // Any functioin can go on here that sets up things
        DLL_PROCESS_DETACH => (),
        _ => ()
    }
    TRUE
}

#[no_mangle]
#[allow(non_snake_case, unused_variables, dead_code)]
extern "stdcall" fn DllGetClassObject(rclsid: *const GUID, riid: *const GUID, ppv: *mut *mut std::ffi::c_void) -> HRESULT {
    let cls_guid = GUID::from_values(0x3B68712C, 0xB6BD, 0x4E60, [0x8D, 0x75, 0x0D, 0xD2, 0x90, 0xB1, 0x52, 0x29]);
    let cls_iid = GUID::from_values(0xAB7742F6, 0x83AF, 0x427C, [0xA0, 0xC4, 0x9B, 0xF4, 0xD1, 0xB7, 0xCB, 0xE8]);
    if unsafe { *rclsid } == cls_guid {
        let res: Result<IUnknown>  = unsafe { CoCreateInstance(rclsid, None, CLSCTX_ALL) };
        if res.is_ok() {
            let iunknwn = res.unwrap();
            let INatComObject = INationalAccounts(iunknwn);
            // INatComObject.CreateInstance(None, Box::into_raw(Box::new(cls_iid)), ppv);
            let query_res = unsafe { INatComObject.query(Box::into_raw(Box::new(cls_iid)), ppv) };
            if query_res == HRESULT(0) {
                return S_OK;
            } else {
                return CLASS_E_CLASSNOTAVAILABLE; 
            }
            // if let Ok(ins) = create_inst {
            //     return S_OK;
            // } else {
            //     return CLASS_E_CLASSNOTAVAILABLE;
            // }
        } else {
            return CLASS_E_CLASSNOTAVAILABLE;
        }    
    }  else {
        return CLASS_E_CLASSNOTAVAILABLE;
    }  
}

#[no_mangle]
#[allow(non_snake_case, unused_variables, dead_code)]
extern "stdcall" fn DllRegisterServer() -> HRESULT {
    // let res = automation::excel_automation::register_com_interfaces(); 
    // if res.is_ok() {
    //     unsafe {
    //         MessageBoxA(HWND(0),
    //         PCSTR(String::from("National Accounts Library Successfully Installed!\0").as_ptr()),
    //         s!("National Accounts Library"),
    //         Default::default());
    //     };
    //     return S_OK;
    // } else {
    //     unsafe {
    //         MessageBoxA(HWND(0),
    //         PCSTR(std::format!("National Accounts Library Registration Encountered an Error: {} \0", res.unwrap_err().to_string()).as_ptr()),
    //         s!("National Accounts Library"),
    //         Default::default());
    //     };
    //     return S_FALSE;
    // }
    S_OK
}

#[no_mangle]
#[allow(non_snake_case, unused_variables, dead_code)]
extern "stdcall" fn DllUnregisterServer() -> HRESULT {
    let res = automation::excel_automation::register_com_interfaces(); 
    if res.is_ok() {
        unsafe {
            MessageBoxA(HWND(0),
            PCSTR(String::from("National Accounts Library Successfully Uninstalled!\0").as_ptr()),
            s!("National Accounts Library"),
            Default::default());
        };
        return S_OK;
    } else {
        unsafe {
            MessageBoxA(HWND(0),
            PCSTR(std::format!("National Accounts Library Encountered an Error during Uninstallation: {} \0", res.unwrap_err().to_string()).as_ptr()),
            s!("National Accounts Library"),
            Default::default());
        };
        return S_FALSE;
    }
}

#[no_mangle]
#[allow(non_snake_case, unused_variables, dead_code)]
extern "stdcall" fn DllCanUnloadNow() -> HRESULT {
    S_OK
}

pub fn add(left: usize, right: usize) -> usize {
    left + right
}

pub fn registry_work() -> windows::core::Result<()>{
    let _y = automation::excel_automation::register_com_interfaces()?;
    Ok(())
}

pub fn test_ribbon_ui() -> windows::core::Result<()> {
    let _t = automation::excel_automation::modify_ui_ribbon()?;
    Ok(())
}

pub fn test_command_bar() -> Result<()> {
    let _t = automation::excel_automation::modify_ui_menu()?;
    Ok(())
}

pub fn test_variant() {
    let t = variant_initialize(None, VT_I2, VariantType::VT_I2(2));
    println!("{:#?}", unsafe { t.Anonymous.Anonymous.vt } );
    
    println!("{:#?}", unsafe { t.Anonymous.Anonymous.Anonymous.iVal} );
}

pub fn get_sheetname() -> Result<()> {
    let _t = automation::excel_automation::retrieve_worksheet_name()?;
    Ok(())
}

pub fn set_sheetname() -> Result<()> {
    let _t = automation::excel_automation::set_worksheet_name()?;
    Ok(())
}

pub fn insert() -> Result<()> {
    let _t = automation::excel_automation::insert_value()?;
    Ok(())
}

pub fn retrieve() -> Result<()> {
    let _t = automation::excel_automation::retrieve_value()?;
    Ok(())
}

pub fn retrieve_array() -> Result<()> {
    let _t = automation::excel_automation::retrieve_range_values()?;
    Ok(())
}

pub fn set_values_array() -> Result<()> {
    let _t = automation::excel_automation::set_range_values()?;
    Ok(())
}

#[cfg(test)]
mod tests {

    use super::*;

    #[test]
    fn it_works() {
        let result = add(2, 2);
        assert_eq!(result, 4);
    }

    #[test]
    fn registry_test()  -> Result<()> {
        let _s = registry_work()?;   
        Ok(())
    }

    #[test]
    fn ribbon_ui_test()  -> Result<()> {
        let _ts = test_ribbon_ui()?;   
        Ok(())
    }

    #[test]
    fn menu_bar_test()  -> Result<()> {
        let _ts = test_command_bar()?;   
        Ok(())
    }

    #[test]
    fn works() {
        test_variant();   
    }

    #[test]
    fn test_com_set_array_values() -> Result<()> {
        let result = set_values_array()?;
        assert_eq!(result, ());
        Ok(())
    }

    #[test]
    fn test_com_get_array_values() -> Result<()> {
        let result = retrieve_array()?;
        assert_eq!(result, ());
        Ok(())
    }

    #[test]
    fn test_com_get_worksheet_name() -> Result<()> {
        let result = get_sheetname()?;
        assert_eq!(result, ());
        Ok(())
    }
    
    #[test]
    fn test_com_set_worksheet_name() -> Result<()> {
        let result = set_sheetname()?;
        assert_eq!(result, ());
        Ok(())
    }

    #[test]
    fn test_com_get_value() -> Result<()> {
        let result = retrieve()?;
        assert_eq!(result, ());
        Ok(())
    }

    #[test]
    fn test_com_set_value() -> Result<()> {
        let result = insert()?;
        assert_eq!(result, ());
        Ok(())
    }

}
