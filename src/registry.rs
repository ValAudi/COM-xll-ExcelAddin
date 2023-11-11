use windows::core::*;
use windows::Win32::System::Registry::*;

pub fn create_registry_entry() -> windows::core::Result<()> {
    let clsid = GUID::new()?;
    let subkey_path = format!("SOFTWARE\\Classes\\CLSID\\{{{:#?}}}", clsid);
    println!("{}", subkey_path);
    let mut bytes_str: Vec<u8> = subkey_path.as_bytes().into_iter().map(|char| *char as u8).collect();
    bytes_str.push(0);
    let lp_subkey = PCSTR::from_raw(bytes_str.as_ptr() as *const u8);
    let lpc_subkey = lp_subkey.clone();

    let key: *mut HKEY = Box::into_raw(Box::new(unsafe { std::mem::zeroed() }));
    let lp_dwdisposition: Option<*mut REG_CREATE_KEY_DISPOSITION> = Some(Box::into_raw(Box::new(unsafe { std::mem::zeroed() })));
    let create_key = unsafe { 
        RegCreateKeyExA(
            HKEY_LOCAL_MACHINE, 
            lp_subkey, 
            0, 
            None, 
            REG_OPTION_NON_VOLATILE, 
            KEY_ALL_ACCESS, 
            None, 
            key, 
            lp_dwdisposition
        )
    };
    if create_key.is_ok() {
        println!("Succcessfully created a Registry Key");
        let _ = unsafe {RegCloseKey(*key)};
    } else {
        let err = create_key.unwrap_err().message().to_string();
        println!("Failed to create a Registry Key: {}", err);
    }

    let delete_key = unsafe {
        RegDeleteKeyExA(
            HKEY_LOCAL_MACHINE, 
            lpc_subkey, 
            KEY_WOW64_64KEY.0, 
            0
        )
    };
    if delete_key.is_ok() {
        println!("Succcessfully deleted a Registry Key");
    } else {
        let err = delete_key.unwrap_err().message().to_string();
        println!("Failed to delete the Registry Key: {}", err);
    }

    Ok(())
}

pub fn open_registry_entry() -> windows::core::Result<()> {
    let clsid = GUID::from_values(0x00000514, 0x0000, 0x0010, [0x80, 0x00, 0x00, 0xAA, 0x00, 0x6D, 0x2E, 0xA4]);
    let subkey_path = format!("CLSID\\{{{:#?}}}", clsid);
    let mut bytes_str: Vec<u8> = subkey_path.as_bytes().into_iter().map(|char| *char as u8).collect();
    bytes_str.push(0);
    let lp_subkey = PCSTR::from_raw(bytes_str.as_ptr());

    let key: *mut HKEY = Box::into_raw(Box::new(unsafe { std::mem::zeroed() }));
    let create_key = unsafe { 
        RegOpenKeyExA(HKEY_CLASSES_ROOT, 
                        lp_subkey,  
                        0,
                        KEY_QUERY_VALUE,
                        key)
    };
    if create_key.is_ok() {
        println!("Succcessfully Opened a Registry Key");
        let _ = unsafe {RegCloseKey(*key)};
    } else {
        let err = create_key.unwrap_err().message().to_string();
        println!("Failed to open a Registry Key: {}", err);
    }
    Ok(())
}