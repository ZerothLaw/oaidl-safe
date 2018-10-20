use std::fmt;
use std::ptr::NonNull;

/// Convenience type for holding value of `*mut T`
/// Mostly just a projection of [`NonNull<T>`] functionality
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

impl<T> Ptr<T> {
    /// Wraps a valid [`NonNull<T>`] 
    /// [`NonNull<T>`]: https://doc.rust-lang.org/nightly/core/ptr/struct.NonNull.html
    pub fn new(p: NonNull<T>) -> Ptr<T> {
        Ptr { inner: p }
    }

    /// Checks a `*mut T` for null and wraps it up for easier handling.
    pub fn with_checked(p: *mut T) -> Option<Ptr<T>> {
        match NonNull::new(p) {
            Some(p) => Some(Ptr::new(p)),
            None => None
        }
    }

    /// Get inner ptr
    pub fn as_ptr(&self) -> *mut T {
        self.inner.as_ptr()
    }

    /// Get inner reference
    /// ## Safety
    /// 
    /// The underlying `.as_ref` call is unsafe so this is unsafe as well, 
    /// in order to propagate the unsafety invariant forward.
    /// 
    /// The lifetime of the provided reference is tied to self. 
    /// 
    /// If you need an unbound lifetime, use `&*my_ptr.as_ptr()` instead.
    pub unsafe fn as_ref(&self) -> &T {
        self.inner.as_ref()
    }

    /// Cast a `Ptr<T>` to `Ptr<U>`
    pub fn cast<U>(self) -> Ptr<U> {
        Ptr::new(self.inner.cast())
    }
}

impl<T> fmt::Pointer for Ptr<T> {
    /// Formats [`Ptr<T>`] as a pointer value (ie, hexadecimal)
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "{:p}", self.inner)
    }
}

impl<T> From<NonNull<T>> for Ptr<T> {
    /// Free, zero-allocation conversion from [`NonNull<T>`] 
    fn from(nn: NonNull<T>) -> Self {
        Ptr::new(nn)
    }
}

impl<T> Into<NonNull<T>> for Ptr<T> {
    /// Need to use Into because of orphan rules
    fn into(self) -> NonNull<T> {
        self.inner
    }
}