use std::{fs::File, ffi::c_void};
use std::io::Read;
use windows::{core::*, Win32::{UI::Ribbon::*, System::Com::{*, StructuredStorage::*}}};

pub const UIRIBBONXML: &str = "excelRibbon.xml";
const _IRIBBONEXTENSIBILITY: GUID = GUID::from_values(0x000C0396, 0x0000, 0x0000, [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46]);

fn read_ribbon_customizations() -> windows::core::Result<IStream> {
    let stream = unsafe { CreateStreamOnHGlobal(None, true) }?;

    // Load file into in-memory then into IStream
    let mut buffer = Vec::new();
    if let Ok(mut file) = File::open(UIRIBBONXML){
        if let Ok(file_size) = file.read_to_end(&mut buffer){
            let data: *const c_void = buffer.as_ptr() as *const c_void;
            let _stream_res = unsafe { stream.Write(data, file_size as u32, None) };
        } else {
            println!("Error Creating buffer from file data and size!");
        };
    } else {
        println!("Error reading UI Customization file!");
    };      

    Ok(stream)
}

pub fn load_customized_ribbon() -> Result<()> {
    println!("----step 0");
    let ribbon_interface: Result<IUnknown> = unsafe { CoCreateInstance(&UIRibbonFramework, None, CLSCTX_ALL)}; 
    if ribbon_interface.is_ok() {
        println!("----step 1");
        let intermediate = ribbon_interface.unwrap();
        // let pointer: *mut *const c_void = Box::into_raw(Box::new(std::ptr::null()));
        let ri: HRESULT = unsafe { intermediate.query(&_IRIBBONEXTENSIBILITY, pointer) };
        if ri.is_ok() {
            println!("----step 2");
            // let _intermed = ri.unwrap();
            // let it = _intermed.
        } else {
            let error_message = ri.0;
            println!("{:#?}", error_message);
        }
    } else {
        let error_message = ribbon_interface.unwrap_err();
        println!("{:#?}", error_message.to_string());
    }

    
    // println!("----step 1");
    // let iribbon_dispatch: IDispatch = ribbon_ext_interface.cast()?;
    // println!("----step 2");
    // let iribbon_dispatch_typeinfo: ITypeInfo = unsafe { iribbon_dispatch.GetTypeInfo(0, 0)? };
    // println!("----step 3");
    // let funcs = unsafe { iribbon_dispatch_typeinfo.GetTypeAttr() }?;
    // println!("----step 4");
    // let f = unsafe { *funcs };
    // println!("{:#?}", f.cFuncs);
    // if let Ok(custom_stream_data) = read_ribbon_customizations(){

    // } else {

    // }
    Ok(())
}
