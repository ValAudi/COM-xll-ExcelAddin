use windows::{core::*, Win32::System::{Ole::*, Com::*, SystemServices::LANG_NEUTRAL, Variant::*}};

const IUNKNOWN: GUID = GUID::from_values(0x00000000, 0x0000, 0x0000, [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46]);
const STD_OLE_GUID: GUID = GUID::from_values(0x00020430, 0x0000, 0x0000, [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46]);

pub struct TypeLibInfo {
    pub path: PCWSTR,
    pub name: PCWSTR,
    pub desc: PCWSTR,
    pub clsid: *const GUID,
}

impl TypeLibInfo {
    pub fn new(path: &str, name: &str, desc: &str) -> windows::core::Result<TypeLibInfo> {
        let instance = TypeLibInfo {
            path: convert_to_pcwstr(path),
            name: convert_to_pcwstr(name),
            desc: convert_to_pcwstr(desc),
            clsid: Box::into_raw(Box::new(GUID::new()?)),
        };
        Ok(instance)
    }
}

pub struct DerivedInt {
    pub name: PCWSTR,
    pub iid: *const GUID,
}

impl DerivedInt {
    pub fn new(name: &str) -> windows::core::Result<DerivedInt> {
        let instance = DerivedInt {
            name: convert_to_pcwstr(name),
            iid: Box::into_raw(Box::new(GUID::new()?)),
        };
        Ok(instance)
    }
}

pub struct TypeLibDef {
    type_library: TypeLibInfo,
    interface: DerivedInt,
    coclass: DerivedInt,
}

impl TypeLibDef {
    pub fn new() ->  windows::core::Result<TypeLibDef> {
        let instance = TypeLibDef {
            type_library: TypeLibInfo::new("C:\\mylib.tlb", "MyAddinLibrary", "My Addin Library Description")?,
            interface: DerivedInt::new("MyAddinInterface")?,
            coclass: DerivedInt::new("MyAddinCoClass")?,
        };
        Ok(instance)
    }
}

pub fn create_typelibrary(tl_data: TypeLibDef) -> windows::core::Result<ICreateTypeInfo> {
    // Create a Type Library Configurator
    let tlb_conf = unsafe { CreateTypeLib2(SYS_WIN32, tl_data.type_library.path) }?;
    unsafe { 
        tlb_conf.SetName(tl_data.type_library.name)?; 
        tlb_conf.SetVersion(1, 0)?;
        tlb_conf.SetDocString(tl_data.type_library.desc)?;
        tlb_conf.SetLcid(LANG_NEUTRAL)?;
        tlb_conf.SetGuid(tl_data.type_library.clsid)?;
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
    let _void = interface_association(&unknwn, &iid_typeinfo, None)?;

    Ok(iid_typeinfo)
}

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

    Ok(())
}

fn create_type_info(tlb_conf: &ICreateTypeLib2, confs: &DerivedInt, int_kind: TYPEKIND, flag: u32) -> windows::core::Result<ICreateTypeInfo> {
    let type_info = unsafe { tlb_conf.CreateTypeInfo(confs.name, int_kind) }?;
    unsafe {
        let _res = type_info.SetGuid(confs.iid);
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

fn convert_to_pcwstr (string: &str) -> PCWSTR {
    let mut string_vec: Vec<u16> = string.as_bytes().into_iter().map(|char| *char as u16).collect();
    string_vec.push(0);
    let p_string = string_vec.as_ptr();
    let pcw_string: PCWSTR = PCWSTR::from_raw(p_string);
    pcw_string
}