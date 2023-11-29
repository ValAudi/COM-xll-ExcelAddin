use std::ffi::c_void;
use windows::{Win32::System::Registry::*, core::*};
use crate::{typelib::*, registry::*, typelib::convert_to_pcwstr};

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
            hkey: HKEY_LOCAL_MACHINE,
            subkey: String::from("SOFTWARE\\Classes\\TypeLib"), 
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
            hkey: HKEY_LOCAL_MACHINE,
            subkey: String::from("SOFTWARE\\Classes\\Interface"),
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
            hkey: HKEY_LOCAL_MACHINE, 
            subkey: String::from("SOFTWARE\\Classes\\CLSID"),
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
            type_library: TypeLibInfo::new("C:\\NationalAccounts\\ntlacc.tlb", "National Accounts", "National Accounts Type Library")?,
            interface: FInterface::new("INationalAccounts")?,
            coclass: CoClassInt::new("NationalAccounts")?,
        };
        Ok(instance)
    }
}

pub fn registry_all() -> windows::core::Result<()> {
    let mut tlb_data = TypeLibDef::new()?;
    // -------------------------------------------------------------------------------------------------------------------
    // COM Class Registry configurations     todo!("Create an APPID key for the COM CLASS");
    tlb_data.coclass.extra_subkeys = Some(vec![
        String::from("InprocServer32"), 
        String::from("ProgID"),
        String::from("TypeLib"),
        String::from("Version")
        ]
    );
    tlb_data.coclass.reg_conf.push(
        RegConfigs::add(
            String::from("InprocServer32"), 
            None, 
            Some(REG_SZ), 
            Some(HSTRING::from("C:\\NationalAccounts\\ntlacc.dll\0").as_wide().as_ptr() as *const c_void), 
            (HSTRING::from("C:\\NationalAccount\\ntlacc.dll\0").len() * 2) as u32
        )
    );
    tlb_data.coclass.reg_conf.push(
        RegConfigs::add(
            String::from("ProgID"), 
            None, 
            Some(REG_SZ), 
            Some(HSTRING::from("NationalAccountCOM+\0").as_wide().as_ptr() as *const c_void), 
            (HSTRING::from("NationalAccountCOM+\0").len() * 2) as u32
        )
    );
    tlb_data.coclass.reg_conf.push(
        RegConfigs::add(
            String::from("TypeLib"), 
            None, 
            Some(REG_SZ), 
            Some(HSTRING::from(format!("{{{:#?}}}\0", unsafe{*tlb_data.type_library.iid})).as_wide().as_ptr() as *const c_void), 
            (HSTRING::from(format!("{{{:#?}}}\0", unsafe{*tlb_data.type_library.iid})).len() * 2) as u32
        )
    ); 
    tlb_data.coclass.reg_conf.push(
        RegConfigs::add(
            String::from("Version"), 
            None, 
            Some(REG_SZ), 
            Some(HSTRING::from("1.0\0").as_wide().as_ptr() as *const c_void),
            (HSTRING::from("1.0\0").len() * 2) as u32
        )
    );  

    let _ = create_registry_entry(&tlb_data.coclass)?;
    let _ = set_registry_key_value(&tlb_data.coclass)?; 

    // ---------------------------------------------------------------------------------------------------------- 
    // Interface Registry configurations. It is automatically created
    // tlb_data.interface.reg_conf.push( 
    //     RegConfigs::add(
    //         String::from("AppID"), // I need AppID for the COCLASS
    //         None, 
    //         Some(REG_SZ), 
    //         Some(HSTRING::from(format!("{{{:#?}}}\0", unsafe{*tlb_data.type_library.iid})).as_wide().as_ptr() as *const c_void), 
    //         (HSTRING::from(format!("{{{:#?}}}\0", unsafe{*tlb_data.type_library.iid})).len() * 2) as u32
    //     )
    // );

    // let _ = create_registry_entry(&tlb_data.interface)?;
    // let _ = set_registry_key_value(&tlb_data.interface)?;

    //-------------------------------------------------------------------------------------------------------------------- 
    // Build and register the type library
    let _ = create_typelibrary(&tlb_data)?;
    Ok(())

}