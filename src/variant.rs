use std::mem::ManuallyDrop;
use windows::{Win32::{System::{Com::*, Variant::*}, Foundation::{DECIMAL, VARIANT_BOOL}}, core::*};

// An Enum to hold the various dataTypes that can be passed into a variant
#[allow(non_camel_case_types)]
pub enum VariantType {
    VT_I1(u8),
    VT_UI1(u8),
    VT_I2(i16),
    VT_UI2(u16),
    VT_I4(i32),
    VT_UI4(u32),
    VT_INT(i32),
    VT_UINT(u32),
    VT_I8(i64),
    VT_UI8(u64),
    VT_R4(f32),
    VT_R8(f64),
    VT_DATE(f64),
    VT_CY(CY),
    VT_BSTR(BSTR),
    VT_DISPATCH(IDispatch),
    VT_BOOL(VARIANT_BOOL),
    VT_VARIANT(VARIANT),
    VT_UNKNOWN(IUnknown),
    VT_DECIMAL(DECIMAL),
    VT_RECORD(VARIANT_0_0_0_0),
    VT_ARRAY(*mut SAFEARRAY),
    VT_ERROR(i32),
}

pub fn variant_initialize(dec_val: Option<DECIMAL>, varenum: VARENUM, value: VariantType) -> VARIANT {
    
    //initialize variant with decimal value if one is intended
    let mut variant_make = unsafe { VariantInit() };
    let outer_variant_union = &mut variant_make.Anonymous;
    if let Some(dec_val) = dec_val {
        outer_variant_union.decVal = dec_val;
    } 
    
    // Specify the varenum and get data type to match on
    let struct_inner = unsafe { &mut outer_variant_union.Anonymous };
    struct_inner.vt = varenum;

    // Matching on Variant type 
    match value {
        VariantType::VT_I2(value) => {
            // Short or iVal or Varenum 2
            struct_inner.Anonymous = VARIANT_0_0_0{ iVal: value };
            return variant_make;
        },
        VariantType::VT_I4(value) => {
            // UShort or uiVal or Varenum 18
            struct_inner.Anonymous = VARIANT_0_0_0{ lVal: value };
            return variant_make;
        },
        VariantType::VT_R4(value) => {
            // Also known as INT or intVal or Varenum 22 or VT_INT
            struct_inner.Anonymous = VARIANT_0_0_0{ fltVal: value };
            return variant_make;
        },
        VariantType::VT_R8(value) => {
            // Also known as UINT or uintVal or Varenum 23 or VT_UINT
            struct_inner.Anonymous = VARIANT_0_0_0{ dblVal: value};
            return variant_make;
        },
        VariantType::VT_CY(value) => {
            // Also known as LONG or lVal or Varenum 3 or VT_I4
            struct_inner.Anonymous = VARIANT_0_0_0{ cyVal: value };
            return variant_make;
        },
        VariantType::VT_DATE(value) => {
            // Also known as ULONG or lVal or Varenum 19 or VT_UI4
            struct_inner.Anonymous = VARIANT_0_0_0{ date: value};
            return variant_make;
        },
        VariantType::VT_BSTR(value) => {
            // Also known as FLOAT or fltVal or Varenum 4 or VT_R4
            struct_inner.Anonymous = VARIANT_0_0_0{ bstrVal: ManuallyDrop::new(value) };
            return variant_make;
        },
        VariantType::VT_DISPATCH(value) => {
            // Also known as DOUBLE or dblVal or Varenum 5 or VT_R8
            struct_inner.Anonymous = VARIANT_0_0_0{ pdispVal: ManuallyDrop::new(Some(value))};
            return variant_make;
        },
        VariantType::VT_ERROR(value) => {
            // Also known as BSTR or bstrVal or Varenum 8 or VT_BSTR
            struct_inner.Anonymous = VARIANT_0_0_0{ scode: value };
            return variant_make;
        },
        VariantType::VT_BOOL(value) => {
            // Also known as UINT or uintVal or Varenum 23 or VT_UINT
            struct_inner.Anonymous = VARIANT_0_0_0{ boolVal: value };
            return variant_make;
        },
        VariantType::VT_VARIANT(value) => {
            // Also known as UINT or uintVal or Varenum 23 or VT_UINT
            struct_inner.Anonymous = VARIANT_0_0_0{ pvarVal: Box::into_raw(Box::new(value)) };
            return variant_make;
        },
        VariantType::VT_UNKNOWN(value) => {
            // Also known as UINT or uintVal or Varenum 23 or VT_UINT
            struct_inner.Anonymous = VARIANT_0_0_0{ punkVal: ManuallyDrop::new(Some(value)) };
            return variant_make;
        },
        VariantType::VT_DECIMAL(value) => {
            // Also known as UINT or uintVal or Varenum 23 or VT_UINT
            struct_inner.Anonymous = VARIANT_0_0_0{ pdecVal: Box::into_raw(Box::new(value)) };
            return variant_make;
        },
        VariantType::VT_I1(value) => {
            // Also known as UINT or uintVal or Varenum 23 or VT_UINT
            struct_inner.Anonymous = VARIANT_0_0_0{ cVal: value };
            return variant_make;
        },
        VariantType::VT_UI1(value) => {
            // Also known as UINT or uintVal or Varenum 23 or VT_UINT
            struct_inner.Anonymous = VARIANT_0_0_0{ bVal: value };
            return variant_make;
        },
        VariantType::VT_UI2(value) => {
            // Also known as UINT or uintVal or Varenum 23 or VT_UINT
            struct_inner.Anonymous = VARIANT_0_0_0{ uiVal: value };
            return variant_make;
        },
        VariantType::VT_UI4(value) => {
            // Also known as UINT or uintVal or Varenum 23 or VT_UINT
            struct_inner.Anonymous = VARIANT_0_0_0{ ulVal: value };
            return variant_make;
        },
        VariantType::VT_I8(value) => {
            // Also known as UINT or uintVal or Varenum 23 or VT_UINT
            struct_inner.Anonymous = VARIANT_0_0_0{ llVal: value };
            return variant_make;
        },
        VariantType::VT_UI8(value) => {
            // Also known as UINT or uintVal or Varenum 23 or VT_UINT
            struct_inner.Anonymous = VARIANT_0_0_0{ ullVal: value };
            return variant_make;
        },
        VariantType::VT_INT(value) => {
            // Also known as UINT or uintVal or Varenum 23 or VT_UINT
            struct_inner.Anonymous = VARIANT_0_0_0{ intVal: value };
            return variant_make;
        },
        VariantType::VT_UINT(value) => {
            // Also known as UINT or uintVal or Varenum 23 or VT_UINT
            struct_inner.Anonymous = VARIANT_0_0_0{ uintVal: value };
            return variant_make;
        },
        VariantType::VT_ARRAY(value) => {
            // Also known as UINT or uintVal or Varenum 23 or VT_UINT
            struct_inner.Anonymous = VARIANT_0_0_0{ parray: value };
            return variant_make;
        },
        _ => {
            println!("Unknown data Type. Returning empty VARIANT");
            return variant_make;
        }
    }   
}