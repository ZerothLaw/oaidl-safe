use std::ptr::null_mut;

use winapi::um::oleauto::{SysAllocStringLen, SysFreeString, SysStringLen};
use widestring::U16String;

use errors::BStringError;
use ptr::Ptr;

// pub type wchar_t = u16;
// pub type WCHAR = wchar_t;
// pub type OLECHAR = WCHAR;
// pub type BSTR = *mut OLECHAR;

//This is how C/Rust look at it, but the memory returned by SysX methods is a bit different
pub type BSTR = *mut u16; 

pub trait BStringExt {
    fn allocate_bstr(&mut self) -> Result<Ptr<u16>, BStringError>;
    fn allocate_managed_bstr(&mut self) -> Result<DroppableBString, BStringError>;
    fn deallocate_bstr(bstr: Ptr<u16>);
    fn from_bstr(bstr: *mut u16) -> U16String;
    fn from_pbstr(bstr: Ptr<u16>) -> U16String;
    fn from_boxed_bstr(bstr: Box<u16>) -> U16String;
}

impl BStringExt for U16String {
    fn allocate_bstr(&mut self) -> Result<Ptr<u16>, BStringError> {
        let sz = self.len();
        let cln = self.clone();
        let rw = cln.as_ptr();
        let bstr: BSTR = unsafe {SysAllocStringLen(rw, sz as u32)};
        match Ptr::with_checked(bstr) {
            Some(pbstr) => Ok(pbstr), 
            None => Err(BStringError::AllocateFailed{len: sz})
        }
    }

    fn allocate_managed_bstr(&mut self) -> Result<DroppableBString, BStringError> {
        Ok(DroppableBString{ inner: Some(self.allocate_bstr()?) })
    }

    fn deallocate_bstr(bstr: Ptr<u16>) {
        let bstr: BSTR = bstr.as_ptr();
        unsafe { SysFreeString(bstr) }
    }

    fn from_bstr(bstr: *mut u16) -> U16String {
        assert!(!bstr.is_null());
        let sz = unsafe {SysStringLen(bstr)};
        unsafe {U16String::from_ptr(bstr, sz as usize)}
    }

    fn from_pbstr(bstr: Ptr<u16>) -> U16String {
        U16String::from_bstr(bstr.as_ptr())
    }

    fn from_boxed_bstr(bstr: Box<u16>) -> U16String {
        U16String::from_bstr(Box::into_raw(bstr))
    }
}

/// Struct that holds pointer to Sys* allocated memory. 
/// It will automatically free the memory via the Sys* 
/// functions unless it has been consumed. 
pub struct DroppableBString {
    inner: Option<Ptr<u16>>
}

impl DroppableBString {
    /// consume() -> *mut u16 returns the contained data
    /// while also setting a flag that the data has been
    /// consumed. It is your responsibility to manage the 
    /// memory yourself. Most uses of BSTR in FFI will
    /// free the memory for you. 
    #[allow(dead_code)]
    pub fn consume(&mut self) -> *mut u16 {
        let ret = match self.inner {
            Some(ptr) => ptr.as_ptr(), 
            None => null_mut()
        };
        self.inner = None;
        ret
    }
}

impl Drop for DroppableBString {
    fn drop(&mut self) {
        match self.inner {
            Some(ptr) => {
                unsafe { SysFreeString(ptr.as_ptr())}
            }, 
            None => {}
        }
    }
}