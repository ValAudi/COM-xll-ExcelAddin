use windows::core::*;
use windows::Win32::System::Registry::*;
use crate::typelib::*;

pub fn create_registry_entry<T: RegistryConfigs>(reg_info: &T) -> windows::core::Result<()> {
    let lp_sub = format!("{{{:#?}}}\\{{{:#?}}}", reg_info.get_subkey(), unsafe {*reg_info.get_iid()});
    let _ = create_reg_entry(reg_info, lp_sub)?;
    if let Some(subkeys) = reg_info.get_extra_subkey() {
        for j in 0..subkeys.len() {
            let lp_subkeys = format!("{{{:#?}}}\\{{{:#?}}}\\{{{:#?}}}", reg_info.get_subkey(), unsafe {*reg_info.get_iid()}, subkeys[j]);
            let _ = create_reg_entry(reg_info, lp_subkeys)?;
        }
    }
    Ok(())
}

fn create_reg_entry<T: RegistryConfigs>(registry_info: &T, subkey: String) -> windows::core::Result<()> {
    let lp_subkey = convert_to_pcwstr(&subkey.as_str());
    let key: *mut HKEY = Box::into_raw(Box::new(unsafe { std::mem::zeroed() }));
    let lp_dwdisposition: Option<*mut REG_CREATE_KEY_DISPOSITION> = Some(Box::into_raw(Box::new(unsafe { std::mem::zeroed() })));
    let _create_key = unsafe { 
        RegCreateKeyExW(
            registry_info.get_hkey(), 
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
    for i in 0..reg_info.get_reg_conf().len() {
        let _att = unsafe {
            RegSetKeyValueW(
                reg_info.get_hkey(), 
                convert_to_pcwstr(format!("{{{:#?}}}\\{{{:#?}}}", reg_info.get_subkey(), reg_info.get_reg_conf()[i].subkey).as_str()), 
                reg_info.get_reg_conf()[i].value_name.unwrap(), // CHeck how it works for default values
                reg_info.get_reg_conf()[i].value_type.unwrap().0, 
                reg_info.get_reg_conf()[i].value_data, 
                reg_info.get_reg_conf()[i].cb_data,
            )
        }?;
    }
    Ok(())    
}