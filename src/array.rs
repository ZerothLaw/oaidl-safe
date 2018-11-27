use std::marker::PhantomData;
use std::mem;
use std::ptr::{drop_in_place, null_mut, NonNull};

use super::bstr::U16String;

use winapi::ctypes::{c_long, c_void};
use winapi::shared::minwindef::{UINT, ULONG};
use winapi::shared::ntdef::HRESULT;
use winapi::shared::wtypes::{
    BSTR,
    CY,
    DATE,
    DECIMAL,
    VARIANT_BOOL,
    VARTYPE,
    VT_BOOL,
    //VT_RECORD,
    VT_BSTR,
    VT_CY,
    VT_DATE,
    VT_DECIMAL,
    VT_DISPATCH,
    VT_ERROR,
    VT_I1,
    VT_I2,
    VT_I4,
    VT_INT,
    VT_R4,
    VT_R8,
    VT_UI1,
    VT_UI2,
    VT_UI4,
    VT_UINT,
    VT_UNKNOWN,
    VT_VARIANT,
};
use winapi::shared::wtypesbase::SCODE;

use winapi::um::oaidl::{
    IDispatch, LPSAFEARRAY, LPSAFEARRAYBOUND, SAFEARRAY, SAFEARRAYBOUND, VARIANT,
};
use winapi::um::unknwnbase::IUnknown;

use super::errors::{
    ElementError, FromSafeArrElemError, FromSafeArrayError, FromVariantError, IntoSafeArrElemError,
    IntoSafeArrayError, IntoVariantError, SafeArrayError,
};
use super::ptr::{DefaultDestructor, Ptr, PtrDestructor};
use super::types::{Currency, Date, DecWrapper, Int, SCode, TryConvert, UInt, VariantBool};
use super::variant::{Variant, VariantDestructor, VariantExt, VariantWrapper, Variants};
macro_rules! check_and_throw {
    ($hr:ident, $success:expr, $fail:expr) => {
        match $hr {
            0 => $success,
            _ => $fail,
        }
    };
}

// Handles dropping zeroed memory (technically initialized, but can't be dropped.)
struct EmptyMemoryDestructor<T> {
    pub(crate) inner: *mut T,
    _marker: PhantomData<T>,
}

///
impl<T> EmptyMemoryDestructor<T> {
    /// Takes a pointer to work around the borrow checker (can't take &mut on a & borrow.)
    fn new(t: *mut T) -> EmptyMemoryDestructor<T> {
        EmptyMemoryDestructor {
            inner: t,
            _marker: PhantomData,
        }
    }
}

impl<T> Drop for EmptyMemoryDestructor<T> {
    /// Checks if pointer is null, and if so, doesn't do anything.
    /// Otherwise, it drops the allocated memory in place.
    fn drop(&mut self) {
        if self.inner.is_null() {
            return;
        }
        unsafe {
            drop_in_place(self.inner);
        }
        self.inner = null_mut();
    }
}

/// Helper trait implemented for types that can be converted into a safe array.
///
/// Implemented for types:
///
/// * [`i8`], [`u8`], [`i16`], [`u16`], [`i32`], [`u32`]
/// * [`bool`], [`f32`], [`f64`]
/// * [`String`], [`Variant<T>`],
/// * [`Ptr<IUnknown>`], [`Ptr<IDispatch>`]
///  
/// [`Variant<T>`]: struct.Variant.html
/// [`Ptr<IUnknown>`]: struct.Ptr.html
/// [`Ptr<IDispatch>`]: struct.Ptr.html
/// [`i8`]: https://doc.rust-lang.org/std/i8/index.html
/// [`u8`]: https://doc.rust-lang.org/std/u8/index.html
/// [`f32`]: https://doc.rust-lang.org/std/f32/index.html
/// [`f64`]: https://doc.rust-lang.org/std/f64/index.html
/// [`i16`]: https://doc.rust-lang.org/std/i16/index.html
/// [`i32`]: https://doc.rust-lang.org/std/i32/index.html
/// [`i64`]: https://doc.rust-lang.org/std/i64/index.html
/// [`u16`]: https://doc.rust-lang.org/std/u16/index.html
/// [`u32`]: https://doc.rust-lang.org/std/u32/index.html
/// [`u64`]: https://doc.rust-lang.org/std/u64/index.html
/// [`String`]: https://doc.rust-lang.org/std/string/struct.String.html
/// [`bool`]: https://doc.rust-lang.org/std/primitive.bool.html
/// [`SCode`]: struct.SCode.html
/// [`Currency`]: struct.Currency.html
/// [`Date`]: struct.Date.html
/// [`Int`]: struct.Int.html
/// [`UInt`]: struct.UInt.html
///
/// ## Example usage
///
/// Generally, you shouldn't implement this on your types without great care. Therefore this
/// example only shows the basic interface, but not implementation details.
///
/// ```
/// extern crate oaidl;
/// extern crate winapi;
///
/// use std::vec::IntoIter;
/// use oaidl::{SafeArrayElement, SafeArrayExt, SafeArrayError};
///
/// fn main() -> Result<(), SafeArrayError> {
///     let v = vec![-3i16, -2, -1, 0, 1, 2, 3];
///     let arr = v.into_iter().into_safearray()?;
///     let out = IntoIter::<i16>::from_safearray(arr)?;    
///     println!("{:?}", out);
///     Ok(())
/// }
/// ```
pub trait SafeArrayElement
where
    Self: Sized,
{
    #[doc(hidden)]
    const SFTYPE: u32;

    #[doc(hidden)]
    type Element: TryConvert<Self, ElementError>;

    #[doc(hidden)]
    fn from_safearray(psa: *mut SAFEARRAY, ix: i32) -> Result<Self::Element, ElementError> {
        let mut def_val: Self::Element = unsafe { mem::zeroed() };
        let mut empty = EmptyMemoryDestructor::new(&mut def_val);
        #[allow(trivial_casts)]
        let hr = unsafe { SafeArrayGetElement(psa, &ix, &mut def_val as *mut _ as *mut c_void) };
        match hr {
            0 => {
                empty.inner = null_mut();
                Ok(def_val)
            }
            _ => Err(ElementError::from(FromSafeArrElemError::GetElementFailed {
                hr: hr,
            })),
        }
    }

    #[doc(hidden)]
    fn into_safearray(self, psa: *mut SAFEARRAY, ix: i32) -> Result<(), ElementError> {
        let mut slf = Self::Element::try_convert(self)?;
        #[allow(trivial_casts)]
        let hr = unsafe { SafeArrayPutElement(psa, &ix, &mut slf as *mut _ as *mut c_void) };
        match hr {
            0 => Ok(()),
            _ => Err(ElementError::from(IntoSafeArrElemError::PutElementFailed {
                hr: hr,
            })),
        }
    }
}

#[doc(hidden)]
pub trait SafeArrayPtrElement: SafeArrayElement {
    #[doc(hidden)]
    fn from_safearray(psa: *mut SAFEARRAY, ix: i32) -> Result<Self::Element, ElementError>;

    #[doc(hidden)]
    fn into_safearray(self, psa: *mut SAFEARRAY, ix: i32) -> Result<(), ElementError>;
}

impl<Array, VD> SafeArrayPtrElement for Array
where
    Array: SafeArrayElement<Element = Ptr<VARIANT, VD>>,
    Ptr<VARIANT, VD>: TryConvert<Array, ElementError>,
    VD: PtrDestructor<VARIANT>,
{
    fn from_safearray(psa: *mut SAFEARRAY, ix: i32) -> Result<Ptr<VARIANT, VD>, ElementError> {
        let mut def_val: VARIANT = unsafe { mem::zeroed() };
        let mut empty = EmptyMemoryDestructor::new(&mut def_val);
        #[allow(trivial_casts)]
        let hr = unsafe { SafeArrayGetElement(psa, &ix, &mut def_val as *mut _ as *mut c_void) };
        match hr {
            0 => {
                empty.inner = null_mut();
                Ok(Ptr::with_checked(Box::into_raw(Box::new(def_val))).unwrap())
            }
            _ => Err(ElementError::from(FromSafeArrElemError::GetElementFailed {
                hr: hr,
            })),
        }
    }

    fn into_safearray(self, psa: *mut SAFEARRAY, ix: i32) -> Result<(), ElementError> {
        let slf = <Ptr<VARIANT, VD> as TryConvert<Array, ElementError>>::try_convert(self)?;
        let hr = unsafe { SafeArrayPutElement(psa, &ix, slf.as_ptr() as *mut c_void) };
        match hr {
            0 => Ok(()),
            _ => Err(ElementError::from(IntoSafeArrElemError::PutElementFailed {
                hr: hr,
            })),
        }
    }
}

macro_rules! impl_safe_arr_elem {
    ($(#[$attrs:meta])* $t:ty => $element:ty, $vtype:ident) => {
        $(#[$attrs])*
        impl SafeArrayElement for $t {
            const SFTYPE: u32 = $vtype;
            type Element = $element;
        }
    };
    ($(#[$attrs:meta])* $t:ty, $vtype:ident) => {
        $(#[$attrs])*
        impl SafeArrayElement for $t {
            const SFTYPE: u32 = $vtype;
            type Element = $t;
        }
    };
}

impl_safe_arr_elem!(
    #[doc = "`SafeArrayElement` impl for `i16`. This allows it to be converted into SAFEARRAY with vt = `VT_I2`."]
    i16,
    VT_I2
);
impl_safe_arr_elem!(
    #[doc = "`SafeArrayElement` impl for `i32`. This allows it to be converted into SAFEARRAY with vt = `VT_I4`."]
    i32,
    VT_I4
);
impl_safe_arr_elem!(
    #[doc = "`SafeArrayElement` impl for `f32`. This allows it to be converted into SAFEARRAY with vt = `VT_R4`."]
    f32,
    VT_R4
);
impl_safe_arr_elem!(
    #[doc = "`SafeArrayElement` impl for `f64`. This allows it to be converted into SAFEARRAY with vt = `VT_R8`."]
    f64,
    VT_R8
);
impl_safe_arr_elem!(#[doc="`SafeArrayElement` impl for [`Currency`]:struct.Currency.html . This allows it to be converted into SAFEARRAY with vt = `VT_CY`."] Currency => CY, VT_CY);
impl_safe_arr_elem!(#[doc="`SafeArrayElement` impl for ['Date']: struct.Date.html. This allows it to be converted into SAFEARRAY with vt = `VT_DATE`."] Date => DATE, VT_DATE);

/// `SafeArrayElement` impl for [`U16String`]. This allows it to be converted into SAFEARRAY with vt = `VT_BSTR`.
/// This overrides the default implementation of `into_safearray` because `*mut *mut u16` is the incorrect
/// type to put in a SAFEARRAY.
///
/// [`U16String`]: https://docs.rs/widestring/0.4.0/widestring/type.U16String.html
impl SafeArrayElement for U16String {
    const SFTYPE: u32 = VT_BSTR;
    type Element = BSTR;

    fn into_safearray(self, psa: *mut SAFEARRAY, ix: i32) -> Result<(), ElementError> {
        let slf = <BSTR as TryConvert<U16String, ElementError>>::try_convert(self)?;
        let hr = unsafe { SafeArrayPutElement(psa, &ix, slf as *mut c_void) };
        match hr {
            0 => Ok(()),
            _ => Err(ElementError::from(IntoSafeArrElemError::PutElementFailed {
                hr: hr,
            })),
        }
    }
}
impl_safe_arr_elem!(
    #[doc="`SafeArrayElement` impl for ['SCode']. This allows it to be converted into SAFEARRAY with vt = `VT_ERROR`."]
    SCode => SCODE, 
    VT_ERROR
);
impl_safe_arr_elem!(
    #[doc="`SafeArrayElement` impl for ['VariantBool']. This allows it to be converted into SAFEARRAY with vt = `VT_BOOL`."]
    VariantBool => VARIANT_BOOL, 
    VT_BOOL
);

/// SafeArrayElement` impl for ['Variant<D,T>']. This allows it to be converted into SAFEARRAY with vt = `VT_VARIANT`.
/// This overrides the default impl of `from_safearray` and `into_safearray` because `*mut VARIANT` doesn't need
/// an additional indirection to be put into a `SAFEARRAY`.
impl<D, T, VD> SafeArrayElement for Variant<D, T, VD>
where
    D: VariantExt<T, VD>
        + TryConvert<T, FromVariantError>
        + super::variant::private::VariantAccess<VD, Field = T>,
    VD: PtrDestructor<VARIANT>,
    T: TryConvert<D, IntoVariantError>,
{
    const SFTYPE: u32 = VT_VARIANT;
    type Element = Ptr<VARIANT, VD>;

    #[doc(hidden)]
    fn from_safearray(psa: *mut SAFEARRAY, ix: i32) -> Result<Self::Element, ElementError> {
        <Self as SafeArrayPtrElement>::from_safearray(psa, ix)
    }

    #[doc(hidden)]
    fn into_safearray(self, psa: *mut SAFEARRAY, ix: i32) -> Result<(), ElementError> {
        <Self as SafeArrayPtrElement>::into_safearray(self, psa, ix)
    }
}

impl SafeArrayElement for Variants {
    const SFTYPE: u32 = VT_VARIANT;
    type Element = Ptr<VARIANT, DefaultDestructor>;

    #[doc(hidden)]
    fn from_safearray(psa: *mut SAFEARRAY, ix: i32) -> Result<Self::Element, ElementError> {
        <Self as SafeArrayPtrElement>::from_safearray(psa, ix)
    }

    #[doc(hidden)]
    fn into_safearray(self, psa: *mut SAFEARRAY, ix: i32) -> Result<(), ElementError> {
        <Self as SafeArrayPtrElement>::into_safearray(self, psa, ix)
    }
}
impl SafeArrayElement for Box<VariantWrapper> {
    const SFTYPE: u32 = VT_VARIANT;
    type Element = Ptr<VARIANT, VariantDestructor>;

    #[doc(hidden)]
    fn from_safearray(psa: *mut SAFEARRAY, ix: i32) -> Result<Self::Element, ElementError> {
        <Self as SafeArrayPtrElement>::from_safearray(psa, ix)
    }

    #[doc(hidden)]
    fn into_safearray(self, psa: *mut SAFEARRAY, ix: i32) -> Result<(), ElementError> {
        <Self as SafeArrayPtrElement>::into_safearray(self, psa, ix)
    }
}

impl_safe_arr_elem!(
    #[doc="`SafeArrayElement` impl for ['DecWrapper']. This allows it to be converted into SAFEARRAY with vt = `VT_DECIMAL`."]
    DecWrapper => DECIMAL, 
    VT_DECIMAL
);
//VT_RECORD
impl_safe_arr_elem!(
    #[doc = "`SafeArrayElement` impl for `i8`. This allows it to be converted into SAFEARRAY with vt = `VT_I1`."]
    i8,
    VT_I1
);
impl_safe_arr_elem!(
    #[doc = "`SafeArrayElement` impl for `u8`. This allows it to be converted into SAFEARRAY with vt = `VT_UI1`."]
    u8,
    VT_UI1
);
impl_safe_arr_elem!(
    #[doc = "`SafeArrayElement` impl for `u16`. This allows it to be converted into SAFEARRAY with vt = `VT_UI2`."]
    u16,
    VT_UI2
);
impl_safe_arr_elem!(
    #[doc = "`SafeArrayElement` impl for `u32`. This allows it to be converted into SAFEARRAY with vt = `VT_UI4`."]
    u32,
    VT_UI4
);
impl_safe_arr_elem!(
    #[doc="`SafeArrayElement` impl for [`Int`]. This allows it to be converted into SAFEARRAY with vt = `VT_INT`."]
    Int => i32, 
    VT_INT
);
impl_safe_arr_elem!(
    #[doc="`SafeArrayElement` impl for [`UInt`]. This allows it to be converted into SAFEARRAY with vt = `VT_INT`."]
    UInt => u32, 
    VT_UINT
);

/// SafeArrayElement` impl for ['Ptr<IUnknown>']. This allows it to be converted into SAFEARRAY with vt = `VT_UNKNOWN`.
/// This overrides the default impl of `from_safearray` and `into_safearray` because *mut IUnknown doesn't need
/// an additional indirection to be put into a SAFEARRAY.
impl SafeArrayElement for Ptr<IUnknown> {
    #[doc(hidden)]
    const SFTYPE: u32 = VT_UNKNOWN;

    #[doc(hidden)]
    type Element = *mut IUnknown;

    #[doc(hidden)]
    fn from_safearray(psa: *mut SAFEARRAY, ix: i32) -> Result<Self::Element, ElementError> {
        let mut def_val: IUnknown = unsafe { mem::zeroed() };
        let mut empty = EmptyMemoryDestructor::new(&mut def_val);
        #[allow(trivial_casts)]
        let hr = unsafe { SafeArrayGetElement(psa, &ix, &mut def_val as *mut _ as *mut c_void) };
        match hr {
            0 => {
                empty.inner = null_mut();
                Ok(&mut def_val)
            }
            _ => Err(ElementError::from(FromSafeArrElemError::GetElementFailed {
                hr: hr,
            })),
        }
    }

    #[doc(hidden)]
    fn into_safearray(self, psa: *mut SAFEARRAY, ix: i32) -> Result<(), ElementError> {
        let slf = self.as_ptr();
        let hr = unsafe { SafeArrayPutElement(psa, &ix, slf as *mut c_void) };
        match hr {
            0 => Ok(()),
            _ => Err(ElementError::from(IntoSafeArrElemError::PutElementFailed {
                hr: hr,
            })),
        }
    }
}

/// SafeArrayElement` impl for ['Ptr<IDispatch>']. This allows it to be converted into SAFEARRAY with vt = `VT_DISPATCH`.
/// This overrides the default impl of `from_safearray` and `into_safearray` because *mut IDispatch doesn't need
/// an additional indirection to be put into a SAFEARRAY.
impl SafeArrayElement for Ptr<IDispatch> {
    #[doc(hidden)]
    const SFTYPE: u32 = VT_DISPATCH;

    #[doc(hidden)]
    type Element = *mut IDispatch;

    #[doc(hidden)]
    fn from_safearray(psa: *mut SAFEARRAY, ix: i32) -> Result<Self::Element, ElementError> {
        let mut def_val: IDispatch = unsafe { mem::zeroed() };
        let mut empty = EmptyMemoryDestructor::new(&mut def_val);
        #[allow(trivial_casts)]
        let hr = unsafe { SafeArrayGetElement(psa, &ix, &mut def_val as *mut _ as *mut c_void) };
        match hr {
            0 => {
                empty.inner = null_mut();
                Ok(&mut def_val)
            }
            _ => Err(ElementError::from(FromSafeArrElemError::GetElementFailed {
                hr: hr,
            })),
        }
    }

    #[doc(hidden)]
    fn into_safearray(self, psa: *mut SAFEARRAY, ix: i32) -> Result<(), ElementError> {
        let slf = self.as_ptr();
        let hr = unsafe { SafeArrayPutElement(psa, &ix, slf as *mut c_void) };
        match hr {
            0 => Ok(()),
            _ => Err(ElementError::from(IntoSafeArrElemError::PutElementFailed {
                hr: hr,
            })),
        }
    }
}

impl TryConvert<Ptr<IUnknown>, ElementError> for *mut IUnknown {
    /// Unwraps the Ptr with [`as_ptr()`]
    fn try_convert(p: Ptr<IUnknown>) -> Result<Self, ElementError> {
        Ok(p.as_ptr())
    }
}

impl TryConvert<Ptr<IDispatch>, ElementError> for *mut IDispatch {
    /// Unwraps the Ptr with [`as_ptr()`]
    fn try_convert(p: Ptr<IDispatch>) -> Result<Self, ElementError> {
        Ok(p.as_ptr())
    }
}

/// Workhorse trait and main interface for converting to/from SAFEARRAY.
/// Default impl is on `ExactSizeIterator<Item=SafeArrayElement>`
pub trait SafeArrayExt<T: SafeArrayElement> {
    /// Use `t.into_safearray()` to convert a type into a SAFEARRAY
    fn into_safearray(&mut self) -> Result<Ptr<SAFEARRAY>, SafeArrayError>;

    /// Use `T::from_safearray(psa)` to convert a safearray pointer into the relevant T
    fn from_safearray(psa: Ptr<SAFEARRAY>) -> Result<Vec<T>, SafeArrayError>;
}

///
#[derive(Copy, Clone, Debug)]
pub struct SafeArrayDestructor;
impl PtrDestructor<SAFEARRAY> for SafeArrayDestructor {
    fn drop(ptr: NonNull<SAFEARRAY>) {
        unsafe { SafeArrayDestroy(ptr.as_ptr()) };
    }
}

/// Blanket implementation, requires that `TryConvert` is implement between `Itr::Item` and `Elem` where
/// `Elem` is the target type for conversion into/from the SAFEARRAY
impl<Itr, Elem> SafeArrayExt<Itr::Item> for Itr
where
    Itr: ExactSizeIterator,
    Itr::Item: SafeArrayElement<Element = Elem> + TryConvert<Elem, ElementError>,
    Elem: TryConvert<Itr::Item, ElementError>,
{
    fn into_safearray(&mut self) -> Result<Ptr<SAFEARRAY>, SafeArrayError> {
        let c_elements: ULONG = self.len() as u32;
        let vartype = Itr::Item::SFTYPE;
        let mut sab = SAFEARRAYBOUND {
            cElements: c_elements,
            lLbound: 0i32,
        };
        let psa = unsafe { SafeArrayCreate(vartype as u16, 1, &mut sab) };
        assert!(!psa.is_null());
        let psa: Ptr<SAFEARRAY, SafeArrayDestructor> = Ptr::with_checked(psa).unwrap();

        for (ix, mut elem) in self.enumerate() {
            match <Itr::Item as SafeArrayElement>::into_safearray(elem, psa.as_ptr(), ix as i32) {
                Ok(()) => continue,
                //Safe-ish to do because memory allocated will be freed.
                Err(e) => {
                    return Err(SafeArrayError::from(IntoSafeArrayError::from_element_err(
                        e, ix,
                    )))
                }
            }
        }

        // No point in clearing the allocated memory at this point.
        // This converts Ptr<SAFEARRAY, SAD> to Ptr<SAFEARRAY, DefaultDestructor>
        Ok(psa.cast())
    }

    fn from_safearray(psa: Ptr<SAFEARRAY>) -> Result<Vec<Itr::Item>, SafeArrayError> {
        let psa: Ptr<SAFEARRAY, SafeArrayDestructor> = psa.cast();
        let sa_dims = unsafe { SafeArrayGetDim(psa.as_ptr()) };
        assert!(sa_dims > 0); //Assert its not a dimensionless safe array
        let vt = unsafe {
            let mut vt: VARTYPE = 0;
            let hr = SafeArrayGetVartype(psa.as_ptr(), &mut vt);
            check_and_throw!(hr, {}, {
                return Err(SafeArrayError::from(
                    FromSafeArrayError::SafeArrayGetVartypeFailed { hr: hr },
                ));
            });
            vt
        };

        if vt as u32 != Itr::Item::SFTYPE {
            return Err(SafeArrayError::from(
                FromSafeArrayError::VarTypeDoesNotMatch {
                    expected: Itr::Item::SFTYPE,
                    found: vt as u32,
                },
            ));
        }

        if sa_dims == 1 {
            let (l_bound, r_bound) = unsafe {
                let mut l_bound: c_long = 0;
                let mut r_bound: c_long = 0;
                let hr = SafeArrayGetLBound(psa.as_ptr(), 1, &mut l_bound);
                check_and_throw!(hr, {}, {
                    return Err(SafeArrayError::from(
                        FromSafeArrayError::SafeArrayLBoundFailed { hr: hr },
                    ));
                });
                let hr = SafeArrayGetUBound(psa.as_ptr(), 1, &mut r_bound);
                check_and_throw!(hr, {}, {
                    return Err(SafeArrayError::from(
                        FromSafeArrayError::SafeArrayRBoundFailed { hr: hr },
                    ));
                });
                (l_bound, r_bound)
            };

            let mut vc: Vec<Itr::Item> = Vec::new();
            for ix in l_bound..=r_bound {
                match Itr::Item::from_safearray(psa.as_ptr(), ix) {
                    Ok(val) => {
                        let v = match Itr::Item::try_convert(val) {
                            Ok(v) => v,
                            Err(ex) => {
                                return Err(SafeArrayError::from(
                                    FromSafeArrayError::from_element_err(ex, ix as usize),
                                ))
                            }
                        };
                        vc.push(v);
                    }
                    Err(e) => {
                        return Err(SafeArrayError::from(FromSafeArrayError::from_element_err(
                            e,
                            ix as usize,
                        )))
                    }
                }
            }
            Ok(vc)
        } else {
            Err(SafeArrayError::from(
                FromSafeArrayError::SafeArrayDimsInvalid { sa_dims: sa_dims },
            ))
        }
    }
}

#[link(name = "OleAut32")]
extern "system" {
    fn SafeArrayCreate(vt: VARTYPE, cDims: UINT, rgsabound: LPSAFEARRAYBOUND) -> LPSAFEARRAY;
    fn SafeArrayDestroy(safe: LPSAFEARRAY) -> HRESULT;

    fn SafeArrayGetDim(psa: LPSAFEARRAY) -> UINT;

    fn SafeArrayGetElement(psa: LPSAFEARRAY, rgIndices: *const c_long, pv: *mut c_void) -> HRESULT;

    fn SafeArrayGetLBound(psa: LPSAFEARRAY, nDim: UINT, plLbound: *mut c_long) -> HRESULT;
    fn SafeArrayGetUBound(psa: LPSAFEARRAY, nDim: UINT, plUbound: *mut c_long) -> HRESULT;

    fn SafeArrayGetVartype(psa: LPSAFEARRAY, pvt: *mut VARTYPE) -> HRESULT;

    fn SafeArrayPutElement(psa: LPSAFEARRAY, rgIndices: *const c_long, pv: *mut c_void) -> HRESULT;
}

#[cfg(test)]
mod test {
    use super::*;
    use rust_decimal::Decimal;
    use std::vec::IntoIter;
    macro_rules! validate_safe_arr {
        ($t:ident, $vals:expr, $vt:expr) => {
            let v: Vec<$t> = $vals;

            let p = v.into_iter().into_safearray().unwrap();

            let r: Result<Vec<$t>, SafeArrayError> = IntoIter::<$t>::from_safearray(p);
            let r = r.unwrap();
            assert_eq!(r, $vals);
        };
    }
    #[test]
    fn test_i16() {
        validate_safe_arr!(i16, vec![0, 1, 2, 3, 4], VT_I2);
    }
    #[test]
    fn test_i32() {
        validate_safe_arr!(i32, vec![0, 1, 2, 3, 4], VT_I4);
    }
    #[test]
    fn test_f32() {
        validate_safe_arr!(f32, vec![0.0f32, -1.333f32, 2f32, 3f32, 4f32], VT_R4);
    }
    #[test]
    fn test_f64() {
        validate_safe_arr!(f64, vec![0.0f64, -1.333f64, 2f64, 3f64, 4f64], VT_R8);
    }
    #[test]
    fn test_cy() {
        validate_safe_arr!(Currency, vec![Currency::from(-1), Currency::from(2)], VT_CY);
    }
    #[test]
    fn test_date() {
        validate_safe_arr!(
            Date,
            vec![Date::from(0.01), Date::from(100.0 / 99.0)],
            VT_DATE
        );
    }

    #[test]
    fn test_str() {
        let v: Vec<U16String> = vec![
            U16String::from_str("validate"),
            U16String::from_str("test string"),
        ];
        let mut v = v.into_iter();
        let p = <IntoIter<U16String> as SafeArrayExt<U16String>>::into_safearray(&mut v).unwrap();

        let r: Result<Vec<U16String>, SafeArrayError> = IntoIter::from_safearray(p);

        let r = r.unwrap();
        assert_eq!(
            r,
            vec![
                U16String::from_str("validate"),
                U16String::from_str("test string")
            ]
        );
    }

    #[test]
    fn test_scode() {
        validate_safe_arr!(SCode, vec![SCode::from(100), SCode::from(10000)], VT_ERROR);
    }
    #[test]
    fn test_bool() {
        validate_safe_arr!(
            VariantBool,
            vec![
                VariantBool::from(true),
                VariantBool::from(false),
                VariantBool::from(true),
                VariantBool::from(true),
                VariantBool::from(false),
                VariantBool::from(false),
                VariantBool::from(true)
            ],
            VT_BOOL
        );
    }

    #[test]
    fn test_variant() {
        let v: Vec<Variant<u64, u64>> = vec![Variant::wrap(100u64)];

        let mut p = v.into_iter();
        let p = <IntoIter<Variant<u64, u64>> as SafeArrayExt<Variant<u64, u64>>>::into_safearray(
            &mut p,
        )
        .unwrap();

        let r: Result<Vec<Variant<u64, u64>>, SafeArrayError> = IntoIter::from_safearray(p);
        let r = r.unwrap();
        assert_eq!(r, vec![Variant::wrap(100u64)]);
    }

    #[test]
    fn variants() {
        let v: Vec<Variants> = vec![
            Variants::LongLong(100000),
            Variants::Long(10000),
            Variants::Char(126),
            Variants::Short(6747),
            Variants::Float(3.33),
            Variants::Double(-85.875509864564),
            Variants::Bool(true),
            Variants::Error(SCode::from(10)),
            Variants::Cy(Currency::from(7587)),
            Variants::Date(Date::from(786.67)),
            Variants::String(String::from("vi7r89748jh u o7o")),
            Variants::Byte(57),
            Variants::UShort(675),
            Variants::ULong(65769866),
            Variants::ULongLong(85784675378),
            Variants::Int(Int::from(56576)),
            Variants::UInt(UInt::from(86576)),
        ];

        let mut ii = v.clone().into_iter();
        let p = <IntoIter<Variants> as SafeArrayExt<Variants>>::into_safearray(&mut ii).unwrap();
        let r = IntoIter::from_safearray(p);
        assert_eq!(v, r.unwrap());
    }

    #[test]
    fn test_decimal() {
        validate_safe_arr!(
            DecWrapper,
            vec![
                DecWrapper::from(Decimal::new(2, 2)),
                DecWrapper::from(Decimal::new(3, 3))
            ],
            VE_DECIMAL
        );
    }
    #[test]
    fn test_i8() {
        validate_safe_arr!(i8, vec![-1, 0, 1, 2, 3, 4], VT_I1);
    }
    #[test]
    fn test_u8() {
        validate_safe_arr!(u8, vec![0, 1, 2, 3, 4], VT_UI1);
    }
    #[test]
    fn test_u16() {
        validate_safe_arr!(u16, vec![0, 1, 2, 3, 4], VT_UI2);
    }
    #[test]
    fn test_u32() {
        validate_safe_arr!(u32, vec![0, 1, 2, 3, 4], VT_UI4);
    }
}
