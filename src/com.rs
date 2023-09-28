use windows::{core::{IUnknown, IUnknown_Vtbl, Result, ComInterface, GUID, Interface, HRESULT}, Win32::{System::Com::IClassFactory_Impl, Foundation::BOOL, UI::Ribbon::{IUIApplication_Impl, UI_VIEWTYPE, UI_COMMANDTYPE, IUICommandHandler, UI_VIEWVERB}}};
use core::fmt::{Debug, Formatter};
use std::ffi::c_void;

#[repr(transparent)]
pub struct IExcelComAddin(IUnknown);
impl IExcelComAddin {
    pub unsafe fn register(&self) -> Result<()> {

        Ok(())
    }
    pub unsafe fn install(&self) -> Result<()> {

        Ok(())
    }
    pub unsafe fn deregister(&self) -> Result<()> {

        Ok(())
    }
    pub unsafe fn uninstall(&self) -> Result<()> {

        Ok(())
    }
}

impl PartialEq for IExcelComAddin {
    fn eq(&self, other: &Self) -> bool {
        self.0 == other.0
    }
}
impl Eq for IExcelComAddin {}
impl Debug for IExcelComAddin {
    fn fmt(&self, f: &mut Formatter<'_>) -> core::fmt::Result {
        f.debug_tuple("IExcelComAddin").field(&self.0).finish()
    }
}
unsafe impl Interface for IExcelComAddin {
    type Vtable = IExcelComAddin_Vtbl;
}
impl Clone for IExcelComAddin {
    fn clone(&self) -> Self {
        Self(self.0.clone())
    }
}
unsafe impl ComInterface for IExcelComAddin {
    const IID: GUID = GUID::from_u128(0x75ae0a2d_dc03_4c9f_8883_069660d0beb6);
}
#[allow(non_snake_case)]
impl IClassFactory_Impl for IExcelComAddin {
    fn CreateInstance(&self, _: Option<&IUnknown>, _: *const GUID, _: *mut *mut c_void) 
        -> std::result::Result<(), windows::core::Error> { 
            todo!() 
    }
    fn LockServer(&self, _: BOOL) -> std::result::Result<(), windows::core::Error> { todo!() }
}
#[allow(non_snake_case)]
impl IUIApplication_Impl for IExcelComAddin {
    fn OnViewChanged(&self, _: u32, _: UI_VIEWTYPE, _: Option<&IUnknown>, _: UI_VIEWVERB, _: i32) 
        -> std::result::Result<(), windows::core::Error> { 
            todo!() 
    }
    fn OnCreateUICommand(&self, _: u32, _: UI_COMMANDTYPE) 
        -> std::result::Result<IUICommandHandler, windows::core::Error> { 
            todo!() 
    }
    fn OnDestroyUICommand(&self, _: u32, _: UI_COMMANDTYPE, _: Option<&IUICommandHandler>) 
        -> std::result::Result<(), windows::core::Error> { 
            todo!() 
    }
}

#[repr(C)]
#[allow(non_camel_case_types)]
pub struct IExcelComAddin_Vtbl {
    pub base__: IUnknown_Vtbl,
    pub register: unsafe extern "system" fn(this: *mut ::core::ffi::c_void) -> HRESULT,
    pub install: unsafe extern "system" fn(this: *mut ::core::ffi::c_void) -> HRESULT,
    pub deregister: unsafe extern "system" fn(this: *mut ::core::ffi::c_void) -> HRESULT,
    pub uninstall: unsafe extern "system" fn(this: *mut ::core::ffi::c_void) -> HRESULT,
}
