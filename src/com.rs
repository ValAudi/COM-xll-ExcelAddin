use std::ffi::c_void;
use std::fmt::{Debug, Formatter};
use windows::Win32::Foundation::{S_OK, BOOL};
use windows::Win32::System::Com::IClassFactory_Impl;
use windows::core::*;

macro_rules! interface_hierarchy {
    ($child:ty, $parent:ty) => {
        impl CanInto<$parent> for $child {}
    };
    ($child:ty, $first:ty, $($rest:ty),+) => {
        $crate::imp::interface_hierarchy!($child, $first);
        $crate::imp::interface_hierarchy!($child, $($rest),+);
    };
}

#[repr(transparent)]
pub struct INationalAccounts(pub IUnknown);
impl INationalAccounts {
    pub unsafe fn operation() -> HRESULT {
        S_OK
    }
}
interface_hierarchy!(INationalAccounts, IUnknown);
impl PartialEq for INationalAccounts {
    fn eq(&self, other: &Self) -> bool {
        self.0 == other.0
    }
}

impl Eq for INationalAccounts {}
impl Debug for INationalAccounts {
    fn fmt(&self, f: &mut Formatter<'_>) -> core::fmt::Result {
        f.debug_tuple("INationalAccounts").field(&self.0).finish()
    }
}
unsafe impl Interface for INationalAccounts {
    type Vtable = INationalAccounts_Vtbl;
}
impl Clone for INationalAccounts {
    fn clone(&self) -> Self {
        Self(self.0.clone())
    }
}
unsafe impl ComInterface for INationalAccounts {
    const IID: GUID = GUID::from_values(0x3B68712C, 0xB6BD, 0x4E60, [0x8D, 0x75, 0x0D, 0xD2, 0x90, 0xB1, 0x52, 0x29]);
}
#[allow(non_snake_case)]
impl IClassFactory_Impl for INationalAccounts {
    fn CreateInstance(&self, _: Option<&IUnknown>, _: *const GUID, _: *mut *mut c_void) 
        -> std::result::Result<(), windows::core::Error> { 
            todo!()
    }
    fn LockServer(&self, _: BOOL) -> std::result::Result<(), windows::core::Error> { todo!() }
}

#[repr(C)]
#[allow(non_camel_case_types)]
pub struct INationalAccounts_Vtbl {
    pub base__: IUnknown_Vtbl,
    pub operation: unsafe extern "system" fn(this: *mut ::core::ffi::c_void) -> HRESULT,
}