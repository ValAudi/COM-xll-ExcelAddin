use windows::{core::*, Win32::System::{Ole::*, Com::*, SystemServices::LANG_NEUTRAL, Variant::*}};
use crate::regkeys::*;

const IUNKNOWN: GUID = GUID::from_values(0x00000000, 0x0000, 0x0000, [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46]);
const STD_OLE_GUID: GUID = GUID::from_values(0x00020430, 0x0000, 0x0000, [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46]);

pub fn create_typelibrary(tl_data: &TypeLibDef) -> windows::core::Result<()> {
    // Create a Type Library Configurator
    println!("Inside Type Library: {:#?}", unsafe {*tl_data.type_library.iid});
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

    println!("Inside Type Library Right before the function description: {:#?}", unsafe {*tl_data.type_library.iid});
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
    println!("Registered Type Library of guid: {:#?}", unsafe {*tl_data.type_library.iid});
    Ok(())
}
#[allow(dead_code)]
pub fn function_description (create_typeinfo: &ICreateTypeInfo) -> windows::core::Result<()> {
    println!("Inside the function description method");
    let mut type_desc = TYPEDESC::default();
    type_desc.vt = VT_INT;

    let elem_desc = ELEMDESC::default(); // ELEMDESC is Element description or Variable description. Each variable must be described.
    let mut elem_desc_array: [ELEMDESC;2] = [elem_desc, elem_desc];
    
    // First/Input and Second/Return Variable
    elem_desc_array[0].tdesc.vt = VT_INT;
    elem_desc_array[0].tdesc.Anonymous.lptdesc = Box::into_raw(Box::new(type_desc));
    elem_desc_array[0].Anonymous.paramdesc.wParamFlags = PARAMFLAG_FIN; // The Input value can be the input identifier number for a command. Have one handler method

    elem_desc_array[1].tdesc.vt = VT_INT;
    elem_desc_array[1].tdesc.Anonymous.lptdesc = Box::into_raw(Box::new(type_desc));
    elem_desc_array[1].Anonymous.paramdesc.wParamFlags = PARAMFLAG_FRETVAL|PARAMFLAG_FOUT;

    let mut func_desc: FUNCDESC = FUNCDESC::default();
    func_desc.memid = 1;
    func_desc.lprgscode = std::ptr::null_mut();
    func_desc.lprgelemdescParam = elem_desc_array.as_ptr() as *mut ELEMDESC;
    func_desc.funckind = FUNC_PUREVIRTUAL;
    func_desc.invkind = INVOKE_FUNC;
    func_desc.callconv = CC_STDCALL;
    func_desc.cParams = 2;
    func_desc.cParamsOpt = 0;
    func_desc.oVft = 0;
    func_desc.cScodes = 0;
    func_desc.elemdescFunc.tdesc.vt = VT_INT;
    func_desc.elemdescFunc.tdesc.Anonymous.lptdesc = Box::into_raw(Box::new(TYPEDESC::default()));
    func_desc.wFuncFlags = FUNCFLAG_FDEFAULTCOLLELEM;

    let p_func_desc: *const FUNCDESC = Box::into_raw(Box::new(func_desc));
    let param_names:[PCWSTR;3] = [
        convert_to_pcwstr("OPERATION"), 
        convert_to_pcwstr("Input Code"), 
        convert_to_pcwstr("Return Code")
    ];
    let _func_desc_res = unsafe { 
        let _ = create_typeinfo.AddFuncDesc(0, p_func_desc)?; 
        let _ = create_typeinfo.SetFuncAndParamNames(0, &param_names)?;
    };

    // Assigning Offsets
    let _ = unsafe { create_typeinfo.LayOut() }?;
    println!("TypeInfo Layout created");
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
    let pcw_string: PCWSTR = PCWSTR::from_raw(string_vec.as_ptr());
    pcw_string
}

pub fn unregsiter_typelib(tl_data: &TypeLibDef) -> windows::core::Result<()> { 
    let res = unsafe { UnRegisterTypeLib(tl_data.type_library.iid, 1, 0, 0, SYS_WIN32) };
    if res.is_ok() {
        let _res_un = res.unwrap();
    } else {
        let error_message: Error = res.unwrap_err();
        println!("{}", error_message.to_string());
    }
    Ok(())
}