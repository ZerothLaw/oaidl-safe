use std::ptr::null_mut;

use winapi::um::oleauto::{SysAllocStringLen, SysFreeString, SysStringLen};
use winapi::shared::wtypes::BSTR;
pub(crate) use widestring::U16String;

use super::errors::{BStringError, FromSafeArrayError, FromVariantError, IntoSafeArrayError, IntoSafeArrElemError, IntoVariantError};
use super::ptr::Ptr;
use super::types::TryConvert;

/// This trait is implemented on `U16String` to enable the convenient and safe conversion of
/// It utilizes the Sys* functions to manage the allocated memory. 
/// Generally you will want to use [`allocate_managed_bstr`] because it provides a
/// type that will automatically free the BSTR when dropped. 
/// 
/// For FFI, you **cannot** use a straight up `*mut u16` when an interface calls for a 
/// BSTR. The reason being is that at the four bytes before where the BSTR pointer points to, 
/// there is a length prefix. In addition, the memory will be freed by the same allocator used in 
/// `SysAllocString`, which can cause UB if you didn't allocate the memory that way. **Any** other
/// allocation method will cause UB and crashes. 
/// 
/// ## Example
/// ```
/// extern crate oaidl;
/// extern crate widestring;
/// 
/// use oaidl::{BStringError, BStringExt};
/// use widestring::U16String;
/// 
/// fn main() -> Result<(), BStringError> {
///     let mut ustr = U16String::from_str("testing abc1267 ?Ťũřǐꝥꞔ");
///     // Automagically dropped once you leave scope. 
///     let bstr = ustr.allocate_managed_bstr()?;
/// 
///     //Unless you call .consume() on it
///     // bstr.consume(); <-- THIS WILL LEAK if you don't take care.
///     Ok(())
/// }
/// ```
/// 
/// [`allocate_managed_bstr`]: #tymethod.allocate_managed_bstr
/// [`DroppableBString`]: struct.DroppableBString.html
pub trait BStringExt {
    /// Allocates a [`Ptr<u16>`] (aka a `*mut u16` aka a BSTR)
    fn allocate_bstr(&mut self) -> Result<Ptr<u16>, BStringError>;

    /// Allocates a [`DroppableBString`] container - automatically frees the memory properly if dropped.
    fn allocate_managed_bstr(&mut self) -> Result<DroppableBString, BStringError>;

    /// Manually and correct free the memory allocated via Sys* methods
    fn deallocate_bstr(bstr: Ptr<u16>);
    
    /// Convenience method for conversion to a good intermediary type
    fn from_bstr(bstr: *mut u16) -> U16String;
    
    /// Convenience method for conversion to a good intermediary type
    
    fn from_pbstr(bstr: Ptr<u16>) -> U16String;
    
    /// Convenience method for conversion to a good intermediary type
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
/// 
/// ## Safety
/// 
/// This wraps up a pointer to Sys* allocated memory and 
/// will automatically clean up that memory correctly
/// unless the memory has been leaked by `consume()`.
/// 
/// One would use the `.consume()` method when sending the 
/// pointer through FFI.
/// 
/// If you don't manually free the memory yourself (correctly)
/// or send it to an FFI function that will do so, then it 
/// *will* be leaked memory. 
/// 
/// If you have a memory leak and you're using this type, 
/// then check your use of consume. 
pub struct DroppableBString {
    inner: Option<Ptr<u16>>
}

impl DroppableBString {
    /// `consume()` -> `*mut u16` returns the contained data
    /// while also setting a flag that the data has been
    /// consumed. It is your responsibility to manage the 
    /// memory yourself. Most uses of BSTR in FFI will
    /// free the memory for you. 
    /// 
    /// This method is very unsafe to use unless you know
    /// how to handle it correctly, hence the `unsafe` marker. 
    #[allow(dead_code)]
    pub unsafe fn consume(&mut self) -> *mut u16 {
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

impl TryConvert<String, BStringError> for U16String {
    fn try_convert(s: String) -> Result<Self, BStringError> {
        Ok(U16String::from_str(&s))
    }
}

impl<'s> TryConvert<&'s str, BStringError> for U16String {
    fn try_convert(s: &str) -> Result<Self, BStringError> {
        Ok(U16String::from_str(s))
    }
}

impl TryConvert<U16String, BStringError> for String {
    fn try_convert(u: U16String) -> Result<Self, BStringError> {
        Ok(u.to_string_lossy())
    }
}

impl TryConvert<U16String, IntoVariantError> for BSTR {
    fn try_convert(u: U16String) -> Result<Self, IntoVariantError> {
        Ok(u.clone().allocate_bstr()?.as_ptr())
    }
}

impl TryConvert<BSTR, FromVariantError> for U16String {
    fn try_convert(p: BSTR) -> Result<Self, FromVariantError> {
        Ok(U16String::from_bstr(p))
    }
}

impl TryConvert<U16String, IntoSafeArrayError> for BSTR {
    fn try_convert(u: U16String) -> Result<Self, IntoSafeArrayError> {
        match u.clone().allocate_bstr() {
            Ok(ptr) => Ok(ptr.as_ptr()), 
            Err(bse) => Err(IntoSafeArrayError::from_element_err(IntoSafeArrElemError::from(bse), 0))
        }
    }
}

impl TryConvert<BSTR, FromSafeArrayError> for U16String {
    fn try_convert(p: BSTR) -> Result<Self, FromSafeArrayError> {
        Ok(U16String::from_bstr(p))
    }
}

impl TryConvert<U16String, IntoSafeArrElemError> for BSTR {
    fn try_convert(u: U16String) -> Result<Self,IntoSafeArrElemError> {
         match u.clone().allocate_bstr() {
            Ok(ptr) => Ok(ptr.as_ptr()), 
            Err(bse) => Err(IntoSafeArrElemError::from(bse))
        }
    } 
}