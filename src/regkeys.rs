use std::ffi::c_void;
use windows::{Win32::System::Registry::*, core::*};
use crate::{typelib::*, registry::*};

// Holds the Type Library Information
#[derive(Debug)]
pub struct TypeLibInfo {
    pub path: PCWSTR,
    pub name: PCWSTR,
    pub desc: PCWSTR,
    pub iid: *const GUID,
    pub hkey: HKEY,
    pub subkey: String,
    pub extra_subkeys: Option<Vec<String>>,
    pub reg_conf: Vec<RegConfigs>,
}

impl TypeLibInfo {
    pub fn new(path: &str, name: &str, desc: &str) -> windows::core::Result<TypeLibInfo> {
        let instance = TypeLibInfo {
            path: convert_to_pcwstr(path),
            name: convert_to_pcwstr(name),
            desc: convert_to_pcwstr(desc),
            iid: Box::into_raw(Box::new(GUID::new()?)),
            hkey: HKEY_CLASSES_ROOT,
            subkey: String::from("TypeLib\\"), 
            extra_subkeys: Some(Vec::new()),
            reg_conf: Vec::new(),
        };
        Ok(instance)
    }
}

// Holds the function interface Information
#[derive(Debug)]
pub struct FInterface {
    pub name: PCWSTR,
    pub iid: *const GUID,
    pub hkey: HKEY,
    pub subkey: String,
    pub extra_subkeys: Option<Vec<String>>,
    pub reg_conf: Vec<RegConfigs>,
}

impl FInterface {
    pub fn new(name: &str) -> windows::core::Result<FInterface> {
        let instance = FInterface {
            name: convert_to_pcwstr(name),
            iid: Box::into_raw(Box::new(GUID::new()?)),
            hkey: HKEY_CLASSES_ROOT,
            subkey: String::from("Interfaces\\"),
            extra_subkeys: None,
            reg_conf: Vec::new(),
        };
        Ok(instance)
    }
}
// Holds the COM Class Object information
#[derive(Debug)]
pub struct CoClassInt {
    pub name: PCWSTR,
    pub iid: *const GUID,
    pub hkey: HKEY,
    pub subkey: String,
    pub extra_subkeys: Option<Vec<String>>,
    pub reg_conf: Vec<RegConfigs> 
}

impl CoClassInt {
    pub fn new(name: &str) -> windows::core::Result<CoClassInt> {
        let instance = CoClassInt {
            name: convert_to_pcwstr(name),
            iid: Box::into_raw(Box::new(GUID::new()?)),
            hkey: HKEY_CLASSES_ROOT, 
            subkey: String::from("CLSID"),
            extra_subkeys: None,
            reg_conf: Vec::new(),
        };
        Ok(instance)
    }
}
#[derive(Debug)]
pub struct RegConfigs {
    pub subkey: String,
    pub value_name: Option<PCWSTR>,
    pub value_type: Option<REG_VALUE_TYPE>,
    pub value_data: Option<*const c_void>,
    pub cb_data: u32
}

impl RegConfigs {
    pub fn add(
        subkey: String,
        value_name: Option<PCWSTR>, 
        value_type: Option<REG_VALUE_TYPE>,
        value_data: Option<*const c_void>,
        cb_data: u32,
     ) -> RegConfigs {
        let instance = RegConfigs {
            subkey: subkey,
            value_name: value_name,
            value_type: value_type,
            value_data: value_data,
            cb_data: cb_data
        };
        instance
    }
}

pub trait RegistryConfigs {
    fn get_name(&self) -> PCWSTR;
    fn get_iid(&self) -> *const GUID;
    fn get_subkey(&self) -> String;
    fn get_hkey(&self) -> HKEY;
    fn get_reg_conf(&self) -> &Vec<RegConfigs>;
    fn get_extra_subkey(&self) -> &Option<Vec<String>>;
}

impl RegistryConfigs for TypeLibInfo {
    fn get_name(&self) -> PCWSTR { self.name }
    fn get_iid(&self) -> *const GUID { self.iid }
    fn get_subkey(&self) -> String { self.subkey.clone() }
    fn get_hkey(&self) -> HKEY { self.hkey}
    fn get_reg_conf(&self) -> &Vec<RegConfigs> { &self.reg_conf }
    fn get_extra_subkey(&self) -> &Option<Vec<String>> { &self.extra_subkeys }
}

impl RegistryConfigs for CoClassInt {
    fn get_name(&self) -> PCWSTR { self.name }
    fn get_subkey(&self) -> String { self.subkey.clone() }
    fn get_iid(&self) -> *const GUID { self.iid }
    fn get_hkey(&self) -> HKEY { self.hkey}
    fn get_reg_conf(&self) -> &Vec<RegConfigs> { &self.reg_conf }
    fn get_extra_subkey(&self) -> &Option<Vec<String>> { &self.extra_subkeys }
}

impl RegistryConfigs for FInterface {
    fn get_name(&self) -> PCWSTR { self.name }
    fn get_subkey(&self) -> String { self.subkey.clone() }
    fn get_iid(&self) -> *const GUID { self.iid }
    fn get_hkey(&self) -> HKEY { self.hkey}
    fn get_reg_conf(&self) -> &Vec<RegConfigs> { &self.reg_conf }
    fn get_extra_subkey(&self) -> &Option<Vec<String>> { &self.extra_subkeys }
}
#[derive(Debug)]
pub struct TypeLibDef {
    pub type_library: TypeLibInfo,
    pub interface: FInterface,
    pub coclass: CoClassInt,
}

impl TypeLibDef {
    pub fn new() ->  windows::core::Result<TypeLibDef> {
        let instance = TypeLibDef {
            type_library: TypeLibInfo::new("C:\\System32\\NationalAccounts.tlb", "National Accounts", "National Accounts Type Library")?,
            interface: FInterface::new("INationalAccounts")?,
            coclass: CoClassInt::new("NationalAccountsCoClass")?,
        };
        Ok(instance)
    }
}

pub fn registry_all() -> windows::core::Result<()> {
    let mut tlb_data = TypeLibDef::new()?;
    // -------------------------------------------------------------------------------------------------------------------
    // COM Class Registry configurations
    tlb_data.coclass.extra_subkeys = Some(vec![
        String::from("InprocServer32"), // Version
        String::from("ProgID"), //LCID
        String::from("TypeLib"),
        String::from("Version")
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
            String::from("TypeLib"), 
            None, 
            Some(REG_SZ), 
            Some(tlb_data.type_library.iid as *const c_void), 
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
            Some(tlb_data.coclass.iid as *const c_void), // Encapsulate the GUID in culy brackets
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
    // Build and register the type library
    let _ = create_typelibrary(&tlb_data)?;
    Ok(())

}