use windows::{core::{IUnknown, IUnknown_Vtbl, Result, ComInterface, GUID, Interface, HRESULT}, Win32::{System::Com::IClassFactory_Impl, Foundation::{BOOL, S_OK}}};
use core::fmt::{Debug, Formatter};
use std::ffi::c_void;
use windows::core::*;

#[repr(transparent)]
pub struct INationalAccounts(IUnknown);
impl INationalAccounts {
    pub unsafe fn register(&self) -> Result<()> {

        Ok(())
    }
}
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
#[allow(non_snake_case)]
impl IUnknownImpl for INationalAccounts {
    type Impl = u32;
    fn get_impl(&self) -> &Self::Impl {
    
        &0
    }
    /// The classic `QueryInterface` method from COM.
    ///
    /// # Safety
    ///
    /// This function is safe to call as long as the interface pointer is non-null and valid for writes
    /// of an interface pointer.
    unsafe fn QueryInterface(&self, iid: *const GUID, interface: *mut *mut std::ffi::c_void) -> HRESULT {

        S_OK
    }
    /// Increments the reference count of the interface
    fn AddRef(&self) -> u32 {
        1
    }
    /// Decrements the reference count causing the interface's memory to be freed when the count is 0
    ///
    /// # Safety
    ///
    /// This function should only be called when the interfacer pointer is no longer used as calling `Release`
    /// on a non-aliased interface pointer and then using that interface pointer may result in use after free.
    unsafe fn Release(&self) -> u32 {
        1
    }
}

#[repr(C)]
#[allow(non_camel_case_types)]
pub struct INationalAccounts_Vtbl {
    pub base__: IUnknown_Vtbl,
    pub register: unsafe extern "system" fn(this: *mut ::core::ffi::c_void) -> HRESULT,
}
