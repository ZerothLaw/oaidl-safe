use std::fmt;
use std::marker::PhantomData;
use std::ptr::{drop_in_place, NonNull};

///Trait for objects that clean up pointed to memory
pub trait PtrDestructor<T> {
    /// Code that cleans up a pointer
    fn drop(ptr: NonNull<T>);
}

/// Default destructor that does not free the memory. 
#[derive(Copy, Clone, Debug, Eq, Hash, PartialEq, PartialOrd)]
pub struct DefaultDestructor;
impl<T> PtrDestructor<T> for DefaultDestructor 
{
    /// This impl *will* leak the `*mut T`. 
    fn drop(_ptr: NonNull<T>) {}
}

/// Automatically frees the memory pointed to by `Ptr<T>`. 
/// Uses `std::ptr::drop_in_place` to drop T generically. 
/// This isn't desirable in all situations however.  
#[derive(Copy, Clone, Debug, Eq, Hash, PartialEq, PartialOrd)]
pub struct DropDestructor;
impl<T> PtrDestructor<T> for DropDestructor 
{
    /// 
    fn drop(ptr: NonNull<T>) {
        let p: *mut T = ptr.as_ptr();
        drop(ptr);
        unsafe {drop_in_place(p)};
    }
}

/// Convenience type for holding value of `*mut T`.
/// Mostly just a projection of [`NonNull<T>`] functionality.
/// 
/// ## Invariants
/// 
/// `*mut T` is guaranteed to be NonNull. 
/// 
/// ## Memory Safety
/// 
/// The default destructor will not drop the *mut T. Therefore, 
/// when the `Ptr<T>` is dropped, the memory will be leaked. 
/// 
/// Use `DropDestructor` if you want to drop the *mut T in the standard way.  
#[derive(Debug, Eq, Hash, PartialOrd, PartialEq)]
pub struct Ptr<T, D = DefaultDestructor> 
where
    D: PtrDestructor<T>
{
    inner: Option<NonNull<T>>, 
    _marker: PhantomData<D>
}

impl<T: Clone> Clone for Ptr<T> {
    fn clone(&self) -> Self {
        Ptr {inner: self.inner.clone(), _marker: PhantomData}
    }
}

impl<T, D> Drop for Ptr<T, D> 
where
    D: PtrDestructor<T> 
{
    fn drop(&mut self){
        if let Some(nn) = self.inner.take() {
            D::drop(nn)
        }
    }
}

impl<T, D> Ptr<T, D> 
where
    D: PtrDestructor<T>
{
    /// Wraps a valid [`NonNull<T>`] 
    /// [`NonNull<T>`]: https://doc.rust-lang.org/nightly/core/ptr/struct.NonNull.html
    pub fn new(p: NonNull<T>) -> Ptr<T, D> {
        Ptr { inner: Some(p), _marker: PhantomData }
    }

    /// Checks a `*mut T` for null and wraps it up for easier handling.
    pub fn with_checked(p: *mut T) -> Option<Ptr<T, D>> {
        match NonNull::new(p) {
            Some(p) => Some(Ptr::new(p)),
            None => None
        }
    }

    /// Get inner ptr
    pub fn as_ptr(&self) -> *mut T {
        if let Some(ref nn) = &self.inner {
            nn.as_ptr()
        } else {
            //safe because only time this is null is right before destruction (ie, a cast<U,Q>)
            // or drop
            unreachable!()
        }
    }

    /// Get inner reference. 
    /// 
    /// ## Safety
    /// 
    /// The underlying `.as_ref` call is unsafe so this is unsafe as well, 
    /// in order to propagate the unsafety invariant forward.
    /// 
    /// The lifetime of the provided reference is tied to self. 
    /// 
    /// If you need an unbound lifetime, use `&*my_ptr.as_ptr()` instead.
    pub unsafe fn as_ref(&self) -> &T {
        if let Some(ref nn) = &self.inner {
            nn.as_ref()
        } else {
            //safe because only time this is null is right before destruction (ie, a cast<U,Q>)
            // or drop
            unreachable!()
        }
    }

    /// Get inner mutable reference. 
    /// 
    /// ## Safety
    /// 
    /// The underlying `.as_ref` call is unsafe so this is unsafe as well, 
    /// in order to propagate the unsafety invariant forward.
    /// 
    /// The lifetime of the provided reference is tied to self. 
    /// 
    /// If you need an unbound lifetime, use `&mut *my_ptr.as_ptr()` instead.
    pub unsafe fn as_mut(&mut self) -> &mut T {
        if let Some(ref mut nn) = &mut self.inner {
            nn.as_mut()
        } else {
            //safe because only time this is null is right before destruction (ie, a cast<U,Q>)
            // or drop
            unreachable!()
        }
    }

    /// Cast a `Ptr<T, D>` to `Ptr<U, Q>`. 
    /// Whatever the `D` is here, the `T` will not be dropped. 
    pub fn cast<U, Q>(mut self) -> Ptr<U, Q> 
    where
        Q: PtrDestructor<U>
    {
        // We take here so as to ensure the T isn't dropped
        if let Some(nn) = self.inner.take() {
            Ptr::new(nn.cast())
        } else {
            //safe because only time this is null is right before destruction (ie, a cast<U,Q>)
            // or drop
            unreachable!()
        }
    }
}

impl<T> fmt::Pointer for Ptr<T> {
    /// Formats [`Ptr<T>`] as a pointer value (ie, hexadecimal)
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        if let Some(ref nn) = &self.inner {
            write!(f, "{:p}", nn)
        } else {
            write!(f, "0x{}", 0)
        }
    }
}

impl<T, D> From<NonNull<T>> for Ptr<T, D> 
where
    D: PtrDestructor<T>
{
    /// Free, zero-allocation conversion from [`NonNull<T>`] 
    fn from(nn: NonNull<T>) -> Self {
        Ptr::new(nn)
    }
}

impl<T> Into<NonNull<T>> for Ptr<T> {
    /// Need to use Into because of orphan rules
    fn into(mut self) -> NonNull<T> {
        // Use .take to ensure that the `*mut T` doesn't get dropped. 
        if let Some(nn) = self.inner.take() {
            nn
        } else {
            unreachable!()
        }
    }
}