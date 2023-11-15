use std::ffi::c_void;
use windows::Win32::System::Registry::REG_SZ;
use crate::{typelib::*, registry::*};

pub fn app_configuration() -> windows::core::Result<()> {
    // I have three interfaces to register. 
    //      1: is the function interface
    //      2: Is the COM CLass interface. Here is where the InprocServer Value will also be set
    //      3: The type library interface.
    // Immediately after these registrations with the registry the typelibrary has to be built and the Information saved
    let mut tlb_data = TypeLibDef::new()?;
    // -------------------------------------------------------------------------------------------------------------------
    // COM Class Registry configurations
    tlb_data.coclass.extra_subkeys = Some(vec![
        String::from("InprocServer32"), // Version
        String::from("ProgID"), //LCID
        String::from("Version"),
        ]
    );
    tlb_data.coclass.reg_conf.push(
        RegConfigs::add(
            String::from("InprocServer32"), 
            None, 
            Some(REG_SZ), 
            Some(Box::into_raw(Box::new(convert_to_pcwstr("C:\\System32\\NationalAccount\\nationalaccounts.dll"))) as *const c_void), 
            0
        )
    );
    tlb_data.coclass.reg_conf.push(
        RegConfigs::add(
            String::from("ProgID"), 
            None, 
            Some(REG_SZ), 
            Some(Box::into_raw(Box::new(convert_to_pcwstr("NationalAccountCOM+"))) as *const c_void), 
            0
        )
    );
    tlb_data.coclass.reg_conf.push(
        RegConfigs::add(
            String::from("Version"), 
            None, 
            Some(REG_SZ), 
            Some(Box::into_raw(Box::new("1.0")) as *const c_void), 
            0
        )
    );  

    let _ = create_registry_entry(&tlb_data.coclass)?;
    let _ = set_registry_key_value(&tlb_data.coclass)?; 

    // ---------------------------------------------------------------------------------------------------------- 
    // Interface Registry configurations
    tlb_data.interface.reg_conf.push(
        RegConfigs::add(
            String::from("ProxyStubClsid32"), 
            None, // Needs some value or unwrap will panic
            Some(REG_SZ), 
            Some(tlb_data.coclass.iid as *const c_void), 
            0 // Change accordingly
        )
    );
    tlb_data.interface.reg_conf.push( 
        RegConfigs::add(
            String::from("TypeLib"), 
            None, // Needs some value or unwrap will panic
            Some(REG_SZ), 
            Some(tlb_data.type_library.iid as *const c_void), 
            0 // Change accordingly
        )
    );

    let _ = create_registry_entry(&tlb_data.interface)?;
    let _ = set_registry_key_value(&tlb_data.interface)?;
    
    //-------------------------------------------------------------------------------------------------------------------- 

    // Type Library Registry configurations --- NOT SURE THEY ARE NEEDED
    tlb_data.type_library.extra_subkeys = Some(vec![
        String::from("1.0"), // Version
        String::from("1.0\\0"), //LCID
        String::from("1.0\\FLAGS"),
        String::from("1.0\\HELPDIR"),
        String::from("1.0\\0\\win32"), // platform
        ]
    );
    tlb_data.type_library.reg_conf.push( 
        RegConfigs::add(
            String::from("1.0"), 
            None, // Needs some value or unwrap will panic
            Some(REG_SZ), 
            Some(Box::into_raw(Box::new(convert_to_pcwstr("National Accounts Automation"))) as *const c_void), // String won't cut it
            0 // Change accordingly
        )
    );
    tlb_data.type_library.reg_conf.push( 
        RegConfigs::add(
            String::from("1.0\\HELPDIR"), 
            None, // Needs some value or unwrap will panic
            Some(REG_SZ), 
            Some(Box::into_raw(Box::new(convert_to_pcwstr("C:\\System32\\"))) as *const c_void), 
            0 // Change accordingly
        )
    );
    tlb_data.type_library.reg_conf.push( 
        RegConfigs::add(
            String::from("1.0\\FLAGS"), 
            None, // Needs some value or unwrap will panic
            Some(REG_SZ), 
            Some(Box::into_raw(Box::new(0 as i32)) as *const c_void), 
            0 // Change accordingly
        )
    );    
    tlb_data.type_library.reg_conf.push( 
        RegConfigs::add(
            String::from("1.0\\0\\win32"), 
            None, // Needs some value or unwrap will panic
            Some(REG_SZ), 
            Some(Box::into_raw(Box::new(convert_to_pcwstr("C:\\System32\\NationalAccounts.tlb"))) as *const c_void), 
            0 // Change accordingly
        )
    );

    let _ = create_registry_entry(&tlb_data.type_library)?;
    let _ = set_registry_key_value(&tlb_data.type_library)?;

    //-------------------------------------------------------------------------------------------------------------------- 

    Ok(())

}