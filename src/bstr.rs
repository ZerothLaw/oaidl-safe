use std;
use std::mem;
use winapi::um::oleauto::{SysAllocString, SysAllocStringLen, SysFreeString, SysStringLen, SysStringByteLen, SysAllocStringByteLen};

use ptr::Ptr;

// pub type wchar_t = u16;
// pub type WCHAR = wchar_t;
// pub type OLECHAR = WCHAR;
// pub type BSTR = *mut OLECHAR;

//This is how C/Rust look at it, but the memory returned by SysX methods is a bit different
pub type BSTR = *mut u16; 

#[derive(Debug, Default, Clone, Eq, Hash, Ord, PartialEq, PartialOrd)]
pub struct BString {
    inner: Vec<u16>
}

#[derive(Debug, Eq, Hash, Ord, PartialEq, PartialOrd)]
pub struct BStr {
    inner: [u16]
}

pub struct RawBStr {
    inner: Ptr<u16>
}

impl Drop for RawBStr {
    fn drop(&mut self) {
        BString::deallocate_raw(self.inner)
    }
}

impl BString {
    pub fn new() -> Self {
        Self{ inner: vec![]}
    }

    pub fn from_vec(raw: impl Into<Vec<u16>>) -> Self {
        Self { inner: raw.into() }
    }

    pub unsafe fn from_ptr(p: *const u16, len: usize) -> Self {
        if len == 0 {
            return Self::new();
        }
        assert!(!p.is_null());
        let slice = std::slice::from_raw_parts(p, len);
        Self::from_vec(slice)
    }

    pub fn with_capacity(capacity: usize) -> Self {
        Self {
            inner: Vec::with_capacity(capacity)
        }
    }

    pub fn capacity(&self) -> usize {
        self.inner.capacity()
    }

    pub fn clear(&mut self) {
        self.inner.clear()
    }

    pub fn reserve(&mut self, additional: usize) {
        self.inner.reserve(additional)
    }

    pub fn into_vec(self) -> Vec<u16> {
        self.inner
    }

    pub fn as_bstr(&self) -> &BStr {
        self
    }

    pub fn push(&mut self, s: impl AsRef<BString>) {
        self.inner.extend_from_slice(&s.as_ref().inner)
    }

    pub fn push_slice(&mut self, s: impl AsRef<[u16]>) {
        self.inner.extend_from_slice(&s.as_ref())
    }

    pub fn shrink_to_fit(&mut self) {
        self.inner.shrink_to_fit();
    }

    pub fn into_boxed_bstr(self) -> Box<BStr> {
        let rw = Box::into_raw(self.inner.into_boxed_slice()) as *mut BStr;
        unsafe {Box::from_raw(rw)}
    }

    pub fn allocate_raw(&mut self) -> Result<Ptr<u16>, ()> {
        let sz = self.len();
        let rw = self.inner.as_mut_ptr();
        let bstr: BSTR = unsafe {SysAllocStringLen(rw, sz as u32)};
        match Ptr::with_checked(bstr) {
            Some(p) => Ok(p), 
            None => Err(())
        }
    }

    pub fn deallocate_raw(p: Ptr<u16>) {
        let bstr: BSTR = p.as_ptr();
        unsafe { SysFreeString(bstr) }
    }
}

impl BStr {
    pub fn new<S: AsRef<Self> + ?Sized>(s: &S) -> &Self {
        s.as_ref()
    }

    pub unsafe fn from_ptr<'a>(p: *const u16, len: usize) -> &'a Self {
        assert!(!p.is_null());
        mem::transmute(std::slice::from_raw_parts(p, len))
    }

    pub fn from_slice(slice: &[u16]) -> &Self {
        unsafe { mem::transmute(slice) }
    }

    pub fn to_bstring(&self) -> BString {
        BString::from_vec(&self.inner)
    }

    pub fn as_slice(&self) -> &[u16] {
        &self.inner
    }

    pub fn len(&self) -> usize {
        self.inner.len()
    }

    pub fn is_empty(&self) -> bool {
        self.inner.is_empty()
    }

    pub fn into_bstring(self: Box<Self>) -> BString {
        let boxed = unsafe {Box::from_raw(Box::into_raw(self) as *mut [u16])};
        BString { inner: boxed.into_vec() }
    }
}

impl std::ops::Index<std::ops::RangeFull> for BString {
    type Output = BStr;

    #[inline]
    fn index(&self, _index: std::ops::RangeFull) -> &BStr {
        BStr::from_slice(&self.inner)
    }
}

impl std::ops::Deref for BString {
    type Target = BStr;
    #[inline]
    fn deref(&self) -> &BStr {
        &self[..]
    }
}

