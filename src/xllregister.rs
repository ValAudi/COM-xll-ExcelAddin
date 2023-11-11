use windows::Win32::UI::WindowsAndMessaging::MessageBoxA;
use windows::core::PCSTR;
use windows::Win32::Foundation::HWND;
use windows::core::s;

use crate::xlvariant::*;
use crate::xlcall::*;
use crate::xlexcel4::*;

#[allow(dead_code)]
struct ExcelFunction {
    dll_name: Variant,
    procedure_name: Variant,
    type_text: Variant,
    fn_name: Variant,
    arg_text: Variant,
    macro_type: Variant,
    category: Variant,
    shortcut: Variant,
    help_text: Variant,
    fn_help: Variant,
    arg_help: Variant,
}

pub fn reg_single_function() {

    let mut result = Variant::new();
    let operation_result = result.as_mut_xloper();
    let get_dllname = excel4(xlfGetName as i32, operation_result, 0);
    if get_dllname.is_ok() {
        let dans = get_dllname.unwrap();
        unsafe {
            MessageBoxA(HWND(0),
            PCSTR(std::format!("Got an xllcall result: {} \0", dans).as_ptr()),
            s!("Get XLL Name: XLL CALL RESULT"),
            Default::default());
        };
        unsafe {
            MessageBoxA(HWND(0),
            PCSTR(std::format!("Got an xloper of type: {} \0", (*operation_result).xltype).as_ptr()),
            s!("Get XLL Name: XLOPER TYPE"),
            Default::default());
        };
    }

    let func_desc: ExcelFunction = ExcelFunction {    
        dll_name: Variant::from_str("Xlladdin"), 
        procedure_name: Variant::from_float(1.0), 
        type_text: Variant::from_str(""), 
        fn_name: Variant::from_str("Valentine's Sum"), 
        arg_text: Variant::from_str("Param 1, Param 2"), 
        macro_type: Variant::from_float(1.0),
        category: Variant::from_float(10.0),  
        shortcut: Variant::from_str("P"),
        help_text: Variant::from_str("Sum of two Numbers "),
        fn_help: Variant::from_str("Test command"),
        arg_help: Variant::from_str("Returns sum"),
    };

    let mut func_data_vec = vec![
        func_desc.dll_name,
        func_desc.procedure_name,
        func_desc.type_text,
        func_desc.fn_name,
        func_desc.arg_text,
        func_desc.macro_type,
        func_desc.category,
        func_desc.shortcut,
        func_desc.help_text,
        func_desc.fn_help,
        func_desc.arg_help
    ];

    let reg_res = excel4_wrapper(xlfRegister, func_data_vec.as_mut_slice());
    if reg_res.is_ok() {
        println!("Change Sheet Function Registered as a Command!");
        let ans = reg_res.unwrap();
        unsafe {
            MessageBoxA(HWND(0),
            PCSTR(std::format!("Got an xloper of type: {}: {}\0", ans.xltype, ans.val.err).as_ptr()),
            s!("Register XLL"),
            Default::default());
        };
    } else {
        println!("Failed to register Command!");
    }
}

pub fn reg_xll_functions() {
    let mut result = Variant::new();
    let operation_result = result.as_mut_xloper();
    let mut dll_name = Variant::from_str("c://Users//Vampete//Documents//xlladdin.xll");
    let mut args: Vec<LPXLOPER> = vec![dll_name.as_mut_xloper()];
    let opers = &mut args;
    let get_dllname = excel4v(xlfRegister as i32, operation_result, 1,opers.as_ptr());
    if get_dllname.is_ok() {
        let dans = get_dllname.unwrap();
        // let _ans_clone = (unsafe { *operation_result }).clone();
        // let ans_str = unsafe { (*operation_result).val.str };
        // let ans = unsafe { CStr::from_ptr(ans_str) }.to_str();
        // let rust_ans = String::from(ans.unwrap());
        unsafe {
            MessageBoxA(HWND(0),
            PCSTR(std::format!("Got an xllcall result: {} \0", dans).as_ptr()),
            s!("Get XLL Name: XLL CALL RESULT"),
            Default::default());
        };
        unsafe {
            MessageBoxA(HWND(0),
            PCSTR(std::format!("Got an xloper of type: {} \0", (*operation_result).xltype).as_ptr()),
            // PCSTR(std::format!("Got an xloper of type: {} : {}\0", (*operation_result).xltype, rust_ans ).as_ptr()),
            s!("Get XLL Name: XLOPER TYPE"),
            Default::default());
        };
    }
}