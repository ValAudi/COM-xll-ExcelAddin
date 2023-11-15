use windows::core::*;
use windows::Win32::System::Registry::*;
use crate::typelib::*;

pub fn create_registry_entry<T: RegistryConfigs>(reg_info: &T) -> windows::core::Result<()> {
    let subkey_path = format!("{{{:#?}}}\\{{{:#?}}}", reg_info.get_subkey(), unsafe {*reg_info.get_iid()});
    let lp_subkey = convert_to_pcwstr(&subkey_path.as_str());
    
    let key: *mut HKEY = Box::into_raw(Box::new(unsafe { std::mem::zeroed() }));
    let lp_dwdisposition: Option<*mut REG_CREATE_KEY_DISPOSITION> = Some(Box::into_raw(Box::new(unsafe { std::mem::zeroed() })));
    let _create_key = unsafe { 
        RegCreateKeyExW(
            HKEY_CURRENT_USER, 
            lp_subkey, 
            0, 
            None, 
            REG_OPTION_NON_VOLATILE, 
            KEY_ALL_ACCESS, 
            None, 
            key, 
            lp_dwdisposition
        )
    }?;
    let _ = unsafe {RegCloseKey(*key)};
    Ok(())
}

pub fn delete_reg_key<T: RegistryConfigs> (reg_info: &T) -> windows::core::Result<()> {
    let _delete_key = unsafe {
        RegDeleteKeyExW(
            reg_info.get_hkey(), 
            convert_to_pcwstr(reg_info.get_subkey().as_str()), 
            KEY_WOW64_64KEY.0, 
            0
        )
    }?;
    Ok(())
}

pub fn open_registry_entry() -> windows::core::Result<()> {
    let clsid = GUID::from_values(0x00000514, 0x0000, 0x0010, [0x80, 0x00, 0x00, 0xAA, 0x00, 0x6D, 0x2E, 0xA4]);
    let subkey_path = format!("CLSID\\{{{:#?}}}", clsid);
    let lp_subkey = convert_to_pcwstr(subkey_path.as_str());

    let key: *mut HKEY = Box::into_raw(Box::new(unsafe { std::mem::zeroed() }));
    let _create_key = unsafe { 
        RegOpenKeyExW(
            HKEY_CLASSES_ROOT, 
            lp_subkey,  
            0,
            KEY_QUERY_VALUE,
            key
        )
    }?;
    let _ = unsafe {RegCloseKey(*key)};
    Ok(())
}

pub fn set_registry_key_value<T: RegistryConfigs> (reg_info: &T) -> windows::core::Result<()> {
    // cbData depends on this value for Example REG_SZ is null terminated by a single 0 hence cbData = 1
    // whereas for REG_MULTI_SZ is double null terminated hence cbData = 2
    let _att = unsafe {
        RegSetKeyValueW(
            reg_info.get_hkey(), 
            convert_to_pcwstr(reg_info.get_subkey().as_str()), 
            reg_info.get_reg_conf()[0].value_name.unwrap(),
            reg_info.get_reg_conf()[0].value_type.unwrap().0, 
            reg_info.get_reg_conf()[0].value_data, 
            reg_info.get_reg_conf()[0].cb_data,
        )
    }?;
    Ok(())    
}