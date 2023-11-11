use libloading::Library;
use windows::Win32::UI::WindowsAndMessaging::MessageBoxA;
use windows::Win32::Foundation::HWND;
use windows::core::{PCSTR, s};
use crate::xlcall::*;
use crate::xlvariant::Variant;

const XLCALL32DLL: &str = "C:\\Program Files\\Microsoft Office\\root\\Office16\\XLCALL32.DLL";
// pub static DLLNAME: LPXLOPER = Variant::from_str("Xlladdin").as_mut_xloper();
// pub fn get_dll_name() -> LPXLOPER {
//     let dll_var = Variant::from_str("Xlladdin").as_mut_xloper();
//     dll_var
// }
type EXCEL4V = extern "stdcall" fn(
    xlfn: ::std::os::raw::c_int, 
    xloper4res: LPXLOPER,
    count: ::std::os::raw::c_int,
    rgpxloper4: *const LPXLOPER
) -> ::std::os::raw::c_int;

type EXCEL4 = extern "cdecl" fn (
        xlfn: ::std::os::raw::c_int,
        operRes: LPXLOPER,
        count: ::std::os::raw::c_int,
        ...
    ) -> ::std::os::raw::c_int;

pub fn excel4v(   
    xlfn: ::std::os::raw::c_int, 
    xloper4res: LPXLOPER,
    count: ::std::os::raw::c_int,
    rgpxloper4: *const LPXLOPER
) -> Result<::std::os::raw::c_int, Box<dyn std::error::Error>> {
    let lib = unsafe { Library::new(XLCALL32DLL) }?;
    let excel_4v: libloading::Symbol<EXCEL4V> = unsafe { lib.get(b"Excel4v") }?;
    let ans = excel_4v(xlfn, xloper4res, count, rgpxloper4);
    let _f = lib.close();
    Ok(ans)
}

pub fn excel4(   
    xlfn: ::std::os::raw::c_int,
    oper_res: LPXLOPER,
    count: ::std::os::raw::c_int
) -> Result<::std::os::raw::c_int, Box<dyn std::error::Error>> {
    let lib = unsafe { Library::new(XLCALL32DLL) }?;
    let excel_4: libloading::Symbol<EXCEL4> = unsafe { lib.get(b"Excel4") }?;
    let ans = excel_4(xlfn, oper_res, count);
    let _f = lib.close();
    Ok(ans)
}

/// Call into Excel, passing a function number as defined in xlcall and a slice
/// of Variant, and returning a Variant. To find out the number and type of
/// parameters and the expected result, please consult the Excel SDK documentation.
/// 
/// Note that this is a slightly inefficient call, in that it allocates a vector
/// of pointers. For example, if you have a single argument, it is faster to invoke
/// the single arg version.
pub fn excel4_wrapper(xlfn: u32, opers: &mut [Variant]) -> Result<xloper, Box<dyn std::error::Error>> {
    let mut result = Variant::new();
    let mut args: Vec<LPXLOPER> = Vec::with_capacity(opers.len());    
    for oper in opers.iter_mut() {
        args.push(oper.as_mut_xloper());
    }
    let opers = &mut args;
    let operation_result = result.as_mut_xloper();
    let t = excel4v(xlfn as i32, operation_result, opers.len() as i32, opers.as_ptr() as *const *mut xloper)?;
    unsafe {
        MessageBoxA(HWND(0),
        PCSTR(std::format!("Got a result of : {}\0", t).as_ptr()),
        s!("xlcall32.dll"),
        Default::default());
    };
    Ok(unsafe { *operation_result })
}

