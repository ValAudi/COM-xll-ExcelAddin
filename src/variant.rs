use std::mem::ManuallyDrop;
use windows::{Win32::{System::{Com::*, Variant::*}, Foundation::DECIMAL}, core::*};

pub fn init_str_variant(dec_val: Option<DECIMAL>, varenum: u16, value: BSTR ) -> VARIANT {
    let mut variant_make = unsafe { VariantInit() };
    let outer_variant_union = &mut variant_make.Anonymous;
    if let Some(dec_val) = dec_val {
        outer_variant_union.decVal = dec_val;
    } 
    let struct_inner = unsafe { &mut outer_variant_union.Anonymous };
    struct_inner.vt = VARENUM(varenum);
    struct_inner.Anonymous = VARIANT_0_0_0{ bstrVal: ManuallyDrop::new(value)};
    return variant_make;
}

pub fn init_array_variant(dec_val: Option<DECIMAL>, varenum: u16, value: *mut SAFEARRAY ) -> VARIANT {
    let mut variant_make = unsafe { VariantInit() };
    let outer_variant_union = &mut variant_make.Anonymous;
    if let Some(dec_val) = dec_val {
        outer_variant_union.decVal = dec_val;
    } 
    let struct_inner = unsafe { &mut outer_variant_union.Anonymous };
    struct_inner.vt = VARENUM(varenum);
    struct_inner.Anonymous = VARIANT_0_0_0{ parray: value};
    return variant_make;
}