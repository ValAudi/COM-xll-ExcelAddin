use std::{fs::File, ffi::c_void};
use std::io::Read;
use windows::Win32::Foundation::HMODULE;
use windows::Win32::System::Com::{IStream, CoCreateInstance, CLSCTX_ALL};
use windows::Win32::System::Com::StructuredStorage::CreateStreamOnHGlobal;
use windows::Win32::System::LibraryLoader::GetModuleHandleA;
use windows::Win32::UI::Ribbon::{UIRibbonFramework, IUIFramework};
// use windows::core::GUID;
use windows::core::{Result, PCSTR};

pub const UIRIBBONXML: &str = "excelRibbon.xml";
// const _IRIBBONEXTENSIBILITY: GUID = GUID::from_values(0x000C0396, 0x0000, 0x0000, [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46]);
// const NULL: GUID = GUID::from_values(0x00000000, 0x0000, 0x0000, [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]);

pub fn read_ribbon_customizations() -> windows::core::Result<IStream> {
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

pub fn load_customized_ribbon() ->Result<()> {
    let _ribbon_framework_interface: IUIFramework = unsafe { CoCreateInstance(&UIRibbonFramework, None, CLSCTX_ALL)}?; 
    println!("----step 1");
    let excel_module_name = String::from("C://Program Files//Microsoft Office//root//Office16//EXCEL.EXE");
    let excel_name_pcstr = PCSTR::from_raw(excel_module_name.as_ptr() as *const u8);
    let _excel_hinstance: HMODULE = unsafe { GetModuleHandleA(excel_name_pcstr) }?;
    println!("----step 2");

   
    
    // unsafe { ribbon_framework_interface.LoadUI(excel_hinstance, resourcename) };
    // let it = _intermed.
    // let pointer: *mut *const c_void = Box::into_raw(Box::new(std::ptr::null()));
    // let ri: HRESULT = unsafe { intermediate.query(&_IRIBBONEXTENSIBILITY, pointer) };
    // if ri.is_ok() {
    //     println!("----step 2");
    //     // let _intermed = ri.unwrap();
    //     // let it = _intermed.
    // } else {
    //     let error_message = ri.0;
    //     println!("{:#?}", error_message);
    // }

    Ok(())
}
