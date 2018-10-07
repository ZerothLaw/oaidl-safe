use std::fmt;
use std::ptr::NonNull;

#[derive(Debug, Eq, Hash, PartialOrd, PartialEq)]
pub struct Ptr<T> {
    inner: NonNull<T>
}

impl<T: Copy> Copy for Ptr<T> {}
impl<T: Clone> Clone for Ptr<T> {
    fn clone(&self) -> Self {
        Ptr {inner: self.inner.clone()}
    }
}

#[allow(dead_code)]
impl<T> Ptr<T> {
    pub fn new(p: NonNull<T>) -> Ptr<T> {
        Ptr { inner: p }
    }

    pub fn with_checked(p: *mut T) -> Option<Ptr<T>> {
        match NonNull::new(p) {
            Some(p) => Some(Ptr::new(p)),
            None => None
        }
    }

    pub fn as_ptr(&self) -> *mut T {
        self.inner.as_ptr()
    }

    pub unsafe fn as_ref(&self) -> &T {
        self.inner.as_ref()
    }

    pub fn cast<U>(self) -> Ptr<U> {
        Ptr::new(self.inner.cast())
    }
}

impl<T> fmt::Pointer for Ptr<T> {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "{:p}", self.inner)
    }
}

impl<T> From<NonNull<T>> for Ptr<T> {
    fn from(nn: NonNull<T>) -> Self {
        Ptr::new(nn)
    }
}

impl<T> Into<NonNull<T>> for Ptr<T> {
    fn into(self) -> NonNull<T> {
        self.inner
    }
}

impl<T> AsRef<T> for Ptr<T> {
    fn as_ref(&self) -> &T {
        unsafe {self.as_ref()}
    }
}