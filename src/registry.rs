use std::ffi::CString;
use windows::Win32::System::Com::*;
use windows::Win32::System::Ole::*;
use windows::Win32::System::SystemServices::LANG_NEUTRAL;
use windows::core::*;
use windows::Win32::System::Registry::*;

const CLASS_IID: GUID = GUID::from_values(0x00000000, 0x0000, 0x0000, [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46]);
const TYPELIBRARY_GUID: GUID = GUID::from_values(0x00000000, 0x0000, 0x0000, [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46]);
const TYPELIBRARY_IID:  GUID = GUID::from_values(0x10000000, 0x0000, 0x0000, [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46]);
const COCLASS_IID: GUID = GUID::from_values(0x00000000, 0x0000, 0x0000, [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46]);
const IUNKNOWN_GUID: GUID = GUID::from_values(0x00000000, 0x0000, 0x0000, [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46]);

pub fn create_registry_entry() -> windows::core::Result<()> {
    let subkey_path = format!("CLSID\\{{{:#?}}}", CLASS_IID);
    println!("{}", subkey_path);
    let intermed = CString::new(subkey_path).expect("Unable to create a C style String");
    let bytes_str: Vec<u8> = intermed.as_bytes().into_iter().map(|char| *char as u8).collect();
    println!("{:#?}", bytes_str);
    let lp_subkey = PCSTR::from_raw(intermed.as_ptr() as *const u8);

    let key: *mut HKEY = Box::into_raw(Box::new(unsafe { std::mem::zeroed() }));
    let lp_dwdisposition: Option<*mut REG_CREATE_KEY_DISPOSITION> = Some(Box::into_raw(Box::new(unsafe { std::mem::zeroed() })));
    
    let create_key = unsafe { 
        RegCreateKeyExA(HKEY_CLASSES_ROOT, 
                        lp_subkey, 
                        0, 
                        None, 
                        REG_OPTION_NON_VOLATILE, 
                        KEY_SET_VALUE, 
                        None, 
                        key, 
                        lp_dwdisposition)
    };
    if create_key.is_ok() {
        println!("Succcessfully created a Registry Key");
        let _ = unsafe {RegCloseKey(*key)};
    } else {
        let err = create_key.unwrap_err().message().to_string();
        println!("Failed to create a Registry Key: {}", err);
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

pub fn create_typelibrary () -> windows::core::Result<()> {

    let (icreate_typelib_int, icreate_typeinfo) = initializing_typelib
    (
        "C:\\mylib.tlb",
        "MyAddinLibrary", 
        "My Test Addin Library Description",
        "MyaddinInt"
    )?;

    let _  = configuring_coclass
    (
        icreate_typelib_int, 
        TYPELIBRARY_IID, 
        "MyaddinCoClass"
    )?;

    let _ = linking_to_iunknown(icreate_typeinfo)?;
    
    Ok(())
}

fn convert_to_pcwstr(string: &str) -> PCWSTR {
    let mut string_vec: Vec<u16> = string.as_bytes().into_iter().map(|char| *char as u16).collect();
    string_vec.push(0);
    let p_string = string_vec.as_ptr();
    let pcw_string: PCWSTR = PCWSTR::from_raw(p_string);
    pcw_string
}

fn initializing_typelib
(
    typelib_path: &str, 
    typelib_name: &str, 
    typelib_desc: &str, 
    typelib_int_name: &str
) -> windows::core::Result<(ICreateTypeLib2, ICreateTypeInfo)> {
    // initializing a type library
    let sz_file: PCWSTR = convert_to_pcwstr(typelib_path);
    let sz_name: PCWSTR = convert_to_pcwstr(typelib_name);
    let sz_doc: PCWSTR = convert_to_pcwstr(typelib_desc);
    let icreate_typelib: ICreateTypeLib2 = unsafe { CreateTypeLib2(SYS_WIN32, sz_file) }?;
    
    unsafe { 
        icreate_typelib.SetName(sz_name)?; 
        icreate_typelib.SetVersion(1, 0)?;
        icreate_typelib.SetDocString(sz_doc)?;
        icreate_typelib.SetLcid(LANG_NEUTRAL)?;
        icreate_typelib.SetGuid(Box::into_raw(Box::new(TYPELIBRARY_GUID)))?;
    }

    // Configuring Interface for type library GUID
    let in_sz_name = convert_to_pcwstr(typelib_int_name);
    let icreate_typeinfo: ICreateTypeInfo = unsafe { icreate_typelib.CreateTypeInfo(in_sz_name, TKIND_INTERFACE) }?;
    unsafe {
        icreate_typeinfo.SetGuid(Box::into_raw(Box::new(TYPELIBRARY_IID)))?;
        icreate_typeinfo.SetTypeFlags(256)?;
    }
    Ok((icreate_typelib, icreate_typeinfo))
}

fn configuring_coclass
(
    typelib_interface: ICreateTypeLib2, 
    typelib_iid: GUID,
    coclass_name: &str
) -> windows::core::Result<()> {
    // Configuring COCLASS Interface GUID
    let in_sz_name = convert_to_pcwstr(coclass_name);
    let icreate_typeinfo = unsafe { typelib_interface.CreateTypeInfo(in_sz_name, TKIND_COCLASS) }?;
    unsafe {
        icreate_typeinfo.SetGuid(Box::into_raw(Box::new(COCLASS_IID)))?;
        icreate_typeinfo.SetTypeFlags(2)?;
    }

    // Associate the COCLASS with interface above it and setting Implementation type flags
    let t_lib: ITypeLib = typelib_interface.cast()?;
    let t_info = unsafe { t_lib.GetTypeInfoOfGuid(Box::into_raw(Box::new(typelib_iid))) }?;

    let hreftype: *const u32 = Box::into_raw(Box::new(unsafe { std::mem::zeroed() }));
    let t = unsafe { icreate_typeinfo.AddRefTypeInfo(&t_info, hreftype) };
    
    // Because I am not sure if hreftype is working 
    if t.is_ok() {
        println!("Succcessfully Got HrefType");
        println!("HREFTYPE Value: {}", unsafe{ *hreftype });
    } else {
        let err = t.unwrap_err().message().to_string();
        println!("Failed to open a Registry Key: {}", err);
    }

    unsafe { 
        icreate_typeinfo.AddImplType(0, *hreftype)?; 
        icreate_typeinfo.SetImplTypeFlags(0, IMPLTYPEFLAG_FDEFAULT)?;
    };
    Ok(())
}

fn linking_to_iunknown(typelib_interface_icreate_typeinfo: ICreateTypeInfo) -> windows::core::Result<()> {
    // Associating our interface with the IUknown interface
    let guid_std_ole = GUID::from_values(0x00020430, 0x0000, 0x0000, [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46]);
    let rguid_std_ole: *const GUID = Box::into_raw(Box::new(guid_std_ole));
    println!("1");
    let std_ole_typelib = unsafe { LoadRegTypeLib(rguid_std_ole, STDOLE_MAJORVERNUM as u16, STDOLE_MINORVERNUM as u16, STDOLE_LCID)? };
    println!("2");
    let iunknwn_type_info = unsafe { std_ole_typelib.GetTypeInfoOfGuid(Box::into_raw(Box::new(IUNKNOWN_GUID)))? };
    
    // Associating our interface with the IUknown interface
    let h_reftype: *const u32 = Box::into_raw(Box::new(unsafe { std::mem::zeroed() }));
    let y = unsafe { typelib_interface_icreate_typeinfo.AddRefTypeInfo(&iunknwn_type_info, h_reftype) };
    
    // Because I am not sure if hreftype is working 
    if y.is_ok() {
        println!("Succcessfully Got HrefType for IUnknown");
        println!("HREFTYPE Value: {}", unsafe{ *h_reftype });
    } else {
        let err = y.unwrap_err().message().to_string();
        println!("Failed Got HrefType for IUnknown: {}", err);
    }

    unsafe { 
        typelib_interface_icreate_typeinfo.AddImplType(0, *h_reftype)?; 
    };
    Ok(())
}

