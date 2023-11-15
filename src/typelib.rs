use std::ffi::c_void;
use windows::{core::*, Win32::System::{Ole::*, Com::*, SystemServices::LANG_NEUTRAL, Variant::*, Registry::*}};

const IUNKNOWN: GUID = GUID::from_values(0x00000000, 0x0000, 0x0000, [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46]);
const STD_OLE_GUID: GUID = GUID::from_values(0x00020430, 0x0000, 0x0000, [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46]);

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

pub fn create_typelibrary(tl_data: TypeLibDef) -> windows::core::Result<()> {
    // Create a Type Library Configurator
    let tlb_conf = unsafe { CreateTypeLib2(SYS_WIN32, tl_data.type_library.path) }?;
    unsafe { 
        tlb_conf.SetName(tl_data.type_library.name)?; 
        tlb_conf.SetVersion(1, 0)?;
        tlb_conf.SetDocString(tl_data.type_library.desc)?;
        tlb_conf.SetLcid(LANG_NEUTRAL)?;
        tlb_conf.SetGuid(tl_data.type_library.iid)?;
    }
    // Get the Interace and CoClass Type Info configurators from the TypeLibrary
    let iid_typeinfo = create_type_info(&tlb_conf, &tl_data.interface, TKIND_INTERFACE, 256)?;
    let coclass_typeinfo = create_type_info(&tlb_conf, &tl_data.coclass, TKIND_COCLASS, 2)?;

    // Associate the COCLASS interface with the TypeLib interface
    let t_info = unsafe { tlb_conf.cast::<ITypeLib>()?.GetTypeInfoOfGuid(tl_data.interface.iid) }?;
    let _void = interface_association(&t_info, &coclass_typeinfo, Some(IMPLTYPEFLAG_FDEFAULT))?;

    // Associate the TypeLib interface with the IUNKNOWN interface
    let std_ole_typelib = unsafe { LoadRegTypeLib(Box::into_raw(Box::new(STD_OLE_GUID)), 2, 0, 0)? };
    let unknwn = unsafe { std_ole_typelib.GetTypeInfoOfGuid(Box::into_raw(Box::new(IUNKNOWN)))? };
    let _ = interface_association(&unknwn, &iid_typeinfo, None)?;

    // Build the function description and saving the changes
    let _ = function_description(&iid_typeinfo)?;
    let _ = unsafe { tlb_conf.SaveAllChanges() }?;

    // Regsiter the Type Library
    let _ = unsafe { 
        RegisterTypeLib(
            &tlb_conf.cast::<ITypeLib>()?, 
            convert_to_pcwstr("C:\\System32\\NationalAccounts.tlb"), 
            convert_to_pcwstr("")
        ) 
    }?;

    Ok(())
}
#[allow(dead_code)]
pub fn function_description (create_typeinfo: &ICreateTypeInfo) -> windows::core::Result<()> {

    let mut type_desc = TYPEDESC::default();
    type_desc.vt = VT_INT;

    let elem_desc = ELEMDESC::default();
    let mut elem_desc_array: [ELEMDESC;3] = [elem_desc, elem_desc, elem_desc];
    let mut counter = 0;
    while counter < 3 {
        elem_desc_array[counter].tdesc.vt = VT_INT;
        elem_desc_array[counter].tdesc.Anonymous.lptdesc = Box::into_raw(Box::new(type_desc));
        if counter == 2{
            elem_desc_array[counter].Anonymous.paramdesc.wParamFlags = PARAMFLAG_FRETVAL|PARAMFLAG_FOUT;
        }
        counter += 1;
    } 

    let mut func_desc: FUNCDESC = FUNCDESC::default();
    func_desc.memid = 1;
    func_desc.lprgscode = std::ptr::null_mut();
    func_desc.lprgelemdescParam = elem_desc_array.as_ptr() as *mut ELEMDESC;
    func_desc.funckind = FUNC_PUREVIRTUAL;
    func_desc.invkind = INVOKE_FUNC;
    func_desc.callconv = CC_STDCALL;
    func_desc.cParams = 3;
    func_desc.cParamsOpt = 0;
    func_desc.oVft = 0;
    func_desc.cScodes = 0;
    func_desc.elemdescFunc.tdesc.vt = VT_HRESULT;
    func_desc.elemdescFunc.tdesc.Anonymous.lptdesc = Box::into_raw(Box::new(TYPEDESC::default()));
    func_desc.wFuncFlags = FUNCFLAG_FDEFAULTCOLLELEM;

    let p_func_desc: *const FUNCDESC = Box::into_raw(Box::new(func_desc));
    let param_names:[PCWSTR;3] = [
        convert_to_pcwstr("SUM"), 
        convert_to_pcwstr("Value 1"), 
        convert_to_pcwstr("Value 2")
    ];
    let _func_desc_res = unsafe { 
        let _ = create_typeinfo.AddFuncDesc(0, p_func_desc); 
        let _ = create_typeinfo.SetFuncAndParamNames(1, &param_names);
    };

    // Assigning Offsets
    let _ = unsafe { create_typeinfo.LayOut() }?;
    Ok(())
}

fn create_type_info<T: RegistryConfigs>(tlb_conf: &ICreateTypeLib2, confs: &T, int_kind: TYPEKIND, flag: u32) -> windows::core::Result<ICreateTypeInfo>  {
    let type_info = unsafe { tlb_conf.CreateTypeInfo(confs.get_name(), int_kind) }?;
    unsafe {
        let _res = type_info.SetGuid(confs.get_iid());
        let _ts_res = type_info.SetTypeFlags(flag)?;
    }
    Ok(type_info)
}

fn interface_association(typeinfo: &ITypeInfo, createtypeinfo: &ICreateTypeInfo, flags: Option<IMPLTYPEFLAGS>) -> windows::core::Result<()> {
    let h_reftype: *const u32 = Box::into_raw(Box::new(unsafe { std::mem::zeroed() }));
    unsafe { 
        createtypeinfo.AddRefTypeInfo(typeinfo, h_reftype)?;
        createtypeinfo.AddImplType(0, *h_reftype)?; 
        if flags.is_some() {
            createtypeinfo.SetImplTypeFlags(0, flags.unwrap())?;  
        }
    };
    Ok(())
}

pub fn convert_to_pcwstr (string: &str) -> PCWSTR {
    let mut string_vec: Vec<u16> = string.as_bytes().into_iter().map(|char| *char as u16).collect();
    string_vec.push(0);
    let p_string = string_vec.as_ptr();
    let pcw_string: PCWSTR = PCWSTR::from_raw(p_string);
    pcw_string
}