use std::marker::PhantomData;
use std::mem;
use std::ptr::{drop_in_place, null_mut};

use widestring::U16String;

use winapi::ctypes::{c_long, c_void};
use winapi::shared::minwindef::{UINT, ULONG,};
use winapi::shared::ntdef::HRESULT;
use winapi::shared::wtypes::{
    BSTR,
    CY, 
    DATE, 
    DECIMAL,  
    VARTYPE,
    VARIANT_BOOL,
    VT_BSTR,
    VT_BOOL,
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

use winapi::um::oaidl::{IDispatch, LPSAFEARRAY, LPSAFEARRAYBOUND, SAFEARRAY, SAFEARRAYBOUND, VARIANT};
use winapi::um::unknwnbase::IUnknown;

use super::errors::{
    FromSafeArrayError, 
    FromSafeArrElemError, 
    IntoSafeArrayError, 
    IntoSafeArrElemError,
};
use super::ptr::Ptr;
use super::types::{Currency, Date, DecWrapper, Int, SCode, TryConvert, UInt, VariantBool};
use super::variant::{Variant, VariantExt};
macro_rules! check_and_throw {
    ($hr:ident, $success:expr, $fail:expr) => {
        match $hr {
            0 => $success, 
            _ => $fail
        }
    };
}

// Handles dropping zeroed memory (technically initialized, but can't be dropped.)
struct EmptyMemoryDestructor<T> {
    pub inner: *mut T, 
    _marker: PhantomData<T>
}

impl<T> EmptyMemoryDestructor<T> {
    fn new(t: *mut T) -> EmptyMemoryDestructor<T> {
        EmptyMemoryDestructor{
            inner: t, 
            _marker: PhantomData
        }
    }
}

impl<T> Drop for EmptyMemoryDestructor<T> {
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
/// * `i8`, `u8`, `i16`, `u16`, `i32`, `u32`
/// * `bool`, `f32`, `f64`
/// * `String`, [`Variant<T>`], 
/// * [`Ptr<IUnknown>`], [`Ptr<IDispatch>`]
///  
/// [`Variant<T>`]: struct.Variant.html
/// [`Ptr<IUnknown>`]: struct.Ptr.html
/// [`Ptr<IDispatch>`]: struct.Ptr.html
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
    Self: Sized
{
    #[doc(hidden)]
    const SFTYPE: u32;

    #[doc(hidden)]
    type Element: TryConvert<Self, IntoSafeArrElemError>;

    #[doc(hidden)]
    fn from_safearray(psa: *mut SAFEARRAY, ix: i32) -> Result<Self::Element, FromSafeArrElemError> {
        let mut def_val: Self::Element = unsafe {mem::zeroed()};
        let mut empty = EmptyMemoryDestructor::new(&mut def_val);
        let hr = unsafe {SafeArrayGetElement(psa, &ix, &mut def_val as *mut _ as *mut c_void)};
        match hr {
            0 => {
                empty.inner = null_mut();
                Ok(def_val)
            }, 
            _ => {
                Err(FromSafeArrElemError::GetElementFailed{hr: hr})
            }
        }
    }

    #[doc(hidden)]
    fn into_safearray(self, psa: *mut SAFEARRAY, ix: i32) -> Result<(), IntoSafeArrElemError> {
        let mut slf = Self::Element::try_convert(self)?;
        let hr = unsafe { SafeArrayPutElement(psa, &ix, &mut slf as *mut _ as *mut c_void)};
        match hr {
            0 => Ok(()), 
            _ => Err(IntoSafeArrElemError::PutElementFailed{hr: hr})
        }
    }
}

macro_rules! impl_safe_arr_elem {
    ($t:ty => $element:ty, $vtype:ident) => {
        impl SafeArrayElement for $t {
            const SFTYPE: u32 = $vtype;
            type Element = $element;
        }
    };
    ($t:ty, $vtype:ident) => {
        impl SafeArrayElement for $t {
            const SFTYPE: u32 = $vtype;
            type Element = $t;
        }
    };
}

impl_safe_arr_elem!(i16, VT_I2);
impl_safe_arr_elem!(i32, VT_I4);
impl_safe_arr_elem!(f32, VT_R4);
impl_safe_arr_elem!(f64, VT_R8);
impl_safe_arr_elem!(Currency => CY, VT_CY);
impl_safe_arr_elem!(Date => DATE, VT_DATE);

// Need to override default implementation here because *mut *mut u16 is invalid.
impl SafeArrayElement for U16String {
    const SFTYPE: u32 = VT_BSTR;
    type Element = BSTR;

    fn into_safearray(self, psa: *mut SAFEARRAY, ix: i32) -> Result<(), IntoSafeArrElemError> {
        let slf = <BSTR as TryConvert<U16String, IntoSafeArrElemError>>::try_convert(self)?;
        let hr = unsafe { SafeArrayPutElement(psa, &ix, slf as *mut _ as *mut c_void)};
        match hr {
            0 => Ok(()), 
            _ => Err(IntoSafeArrElemError::PutElementFailed{hr: hr})
        }
    }
}
impl_safe_arr_elem!(SCode => SCODE, VT_ERROR);
impl_safe_arr_elem!(VariantBool => VARIANT_BOOL, VT_BOOL);

// Overriding default behavior here because its special.
impl<D, T> SafeArrayElement for Variant<D, T> 
where
    D: VariantExt<T>
{
    const SFTYPE: u32 = VT_VARIANT;
    type Element = *mut VARIANT;

    #[doc(hidden)]
    fn from_safearray(psa: *mut SAFEARRAY, ix: i32) -> Result<Self::Element, FromSafeArrElemError> {
        let mut def_val: VARIANT = unsafe {mem::zeroed()};
        let mut empty = EmptyMemoryDestructor::new(&mut def_val);
        let hr = unsafe {SafeArrayGetElement(psa, &ix, &mut def_val as *mut _ as *mut c_void)};
        match hr {
            0 => {
                empty.inner = null_mut();
                Ok(&mut def_val)
            }, 
            _ => {
                Err(FromSafeArrElemError::GetElementFailed{hr: hr})
            }
        }
    }

    #[doc(hidden)]
    fn into_safearray(self, psa: *mut SAFEARRAY, ix: i32) -> Result<(), IntoSafeArrElemError> {
        let slf = <*mut VARIANT as TryConvert<Variant<D, T>, IntoSafeArrElemError>>::try_convert(self)?;
        let hr = unsafe { SafeArrayPutElement(psa, &ix, slf as *mut _ as *mut c_void)};
        match hr {
            0 => Ok(()), 
            _ => Err(IntoSafeArrElemError::PutElementFailed{hr: hr})
        }
    }
}
impl_safe_arr_elem!(DecWrapper => DECIMAL, VT_DECIMAL);
//VT_RECORD
impl_safe_arr_elem!(i8, VT_I1);
impl_safe_arr_elem!(u8, VT_UI1);
impl_safe_arr_elem!(u16, VT_UI2);
impl_safe_arr_elem!(u32, VT_UI4);
impl_safe_arr_elem!(Int => i32, VT_INT);
impl_safe_arr_elem!(UInt => u32, VT_UINT);

impl TryConvert<Ptr<IUnknown>, IntoSafeArrElemError> for *mut IUnknown {
    fn try_convert(p: Ptr<IUnknown>) -> Result<Self, IntoSafeArrElemError> {
        Ok(p.as_ptr())
    }
}

impl SafeArrayElement for Ptr<IUnknown> {
    const SFTYPE: u32 = VT_UNKNOWN;
    type Element = *mut IUnknown;

    fn from_safearray(psa: *mut SAFEARRAY, ix: i32) -> Result<Self::Element, FromSafeArrElemError> {
        let mut def_val: IUnknown = unsafe {mem::zeroed()};
        let mut empty = EmptyMemoryDestructor::new(&mut def_val);
        let hr = unsafe {SafeArrayGetElement(psa, &ix, &mut def_val as *mut _ as *mut c_void)};
        match hr {
            0 => {
                empty.inner = null_mut();
                Ok(&mut def_val)
            }, 
            _ => {
                Err(FromSafeArrElemError::GetElementFailed{hr: hr})
            }
        }
    }

    fn into_safearray(self, psa: *mut SAFEARRAY, ix: i32) -> Result<(), IntoSafeArrElemError> {
        let slf = self.as_ptr();
        let hr = unsafe { SafeArrayPutElement(psa, &ix, slf as *mut _ as *mut c_void)};
        match hr {
            0 => Ok(()), 
            _ => Err(IntoSafeArrElemError::PutElementFailed{hr: hr})
        }
    }
}

impl TryConvert<Ptr<IDispatch>, IntoSafeArrElemError> for *mut IDispatch {
    fn try_convert(p: Ptr<IDispatch>) -> Result<Self, IntoSafeArrElemError> {
        Ok(p.as_ptr())
    }
}

impl SafeArrayElement for Ptr<IDispatch> {
    const SFTYPE: u32 = VT_DISPATCH;
    type Element = *mut IDispatch;

    fn from_safearray(psa: *mut SAFEARRAY, ix: i32) -> Result<Self::Element, FromSafeArrElemError> {
        let mut def_val: IDispatch = unsafe {mem::zeroed()};
        let mut empty = EmptyMemoryDestructor::new(&mut def_val);
        let hr = unsafe {SafeArrayGetElement(psa, &ix, &mut def_val as *mut _ as *mut c_void)};
        match hr {
            0 => {
                empty.inner = null_mut();
                Ok(&mut def_val)
            }, 
            _ => {
                Err(FromSafeArrElemError::GetElementFailed{hr: hr})
            }
        }
    }

    fn into_safearray(self, psa: *mut SAFEARRAY, ix: i32) -> Result<(), IntoSafeArrElemError> {
        let slf = self.as_ptr();
        let hr = unsafe { SafeArrayPutElement(psa, &ix, slf as *mut _ as *mut c_void)};
        match hr {
            0 => Ok(()), 
            _ => Err(IntoSafeArrElemError::PutElementFailed{hr: hr})
        }
    }
}

/// Workhorse trait and main interface for converting to/from SAFEARRAY. 
/// Default impl is on `ExactSizeIterator<Item=SafeArrayElement>` 
pub trait SafeArrayExt<T: SafeArrayElement> {
    /// Use `t.into_safearray()` to convert a type into a SAFEARRAY
    fn into_safearray(&mut self) -> Result<Ptr<SAFEARRAY>, IntoSafeArrayError>;
    
    /// Use `T::from_safearray(psa)` to convert a safearray pointer into the relevant T
    fn from_safearray(psa: Ptr<SAFEARRAY>) -> Result<Vec<T>, FromSafeArrayError>;
}

// Ensures that the SAFEARRAY memory that is allocated gets cleaned up, even during a panic.
struct SafeArrayDestructor {
    inner: *mut SAFEARRAY, 
    _marker: PhantomData<SAFEARRAY>
}

impl SafeArrayDestructor {
    fn new(p: *mut SAFEARRAY) -> SafeArrayDestructor {
        assert!(!p.is_null(), "SafeArrayDestructor initialized with null *mut SAFEARRAY pointer.");
        SafeArrayDestructor{
            inner: p, 
            _marker: PhantomData
        }
    }
}

impl Drop for SafeArrayDestructor {
    fn drop(&mut self)  {
        if self.inner.is_null(){
            return;
        }
        unsafe {
            SafeArrayDestroy(self.inner)
        };
        self.inner = null_mut();
    }
}

/// Blanket implementation, requires that `TryConvert` is implement between `Itr::Item` and `Elem` where 
/// `Elem` is the target type for conversion into/from the SAFEARRAY
impl<Itr, Elem> SafeArrayExt<Itr::Item> for Itr
where 
    Itr: ExactSizeIterator, 
    Itr::Item: SafeArrayElement<Element=Elem> + TryConvert<Elem, FromSafeArrayError>, 
    Elem: TryConvert<Itr::Item, IntoSafeArrayError> + TryConvert<Itr::Item, IntoSafeArrElemError> 
{
    fn into_safearray(&mut self) -> Result<Ptr<SAFEARRAY>, IntoSafeArrayError > {
        let c_elements: ULONG = self.len() as u32;
        let vartype = Itr::Item::SFTYPE;
        let mut sab = SAFEARRAYBOUND { cElements: c_elements, lLbound: 0i32};
        let psa = unsafe { SafeArrayCreate(vartype as u16, 1, &mut sab)};
        assert!(!psa.is_null());
        let mut sad = SafeArrayDestructor::new(psa);
        
        for (ix, mut elem) in self.enumerate() {
            match <Itr::Item as SafeArrayElement>::into_safearray(elem, psa, ix as i32) {
                Ok(()) => continue,
                //Safe-ish to do because memory allocated will be freed. 
                Err(e) => return Err(IntoSafeArrayError::from_element_err(e, ix))
            }
        }

        // No point in clearing the allocated memory at this point. 
        sad.inner = null_mut();

        Ok(Ptr::with_checked(psa).unwrap())
    }

    fn from_safearray(psa: Ptr<SAFEARRAY>) -> Result<Vec<Itr::Item>, FromSafeArrayError> {
        let psa = psa.as_ptr();
        //Stack sentinel to ensure safearray is released even if there is a panic or early return.
        let _sad = SafeArrayDestructor::new(psa);
        let sa_dims = unsafe { SafeArrayGetDim(psa) };
        assert!(sa_dims > 0); //Assert its not a dimensionless safe array
        let vt = unsafe {
            let mut vt: VARTYPE = 0;
            let hr = SafeArrayGetVartype(psa, &mut vt);
            check_and_throw!(hr, {}, {return Err(FromSafeArrayError::SafeArrayGetVartypeFailed{hr: hr})});
            vt
        };

        if vt as u32 != Itr::Item::SFTYPE {
            return Err(FromSafeArrayError::VarTypeDoesNotMatch{expected: Itr::Item::SFTYPE, found: vt as u32});
        }

        if sa_dims == 1 {
            let (l_bound, r_bound) = unsafe {
                let mut l_bound: c_long = 0;
                let mut r_bound: c_long = 0;
                let hr = SafeArrayGetLBound(psa, 1, &mut l_bound);
                check_and_throw!(hr, {}, {return Err(FromSafeArrayError::SafeArrayLBoundFailed{hr: hr})});
                let hr = SafeArrayGetUBound(psa, 1, &mut r_bound);
                check_and_throw!(hr, {}, {return Err(FromSafeArrayError::SafeArrayRBoundFailed{hr: hr})});
                (l_bound, r_bound)
            };

            let mut vc: Vec<Itr::Item> = Vec::new();
            for ix in l_bound..=r_bound {
                match Itr::Item::from_safearray(psa, ix) {
                    Ok(val) => vc.push(Itr::Item::try_convert(val)?), 
                    Err(e) => return Err(FromSafeArrayError::from_element_err(e, ix as usize))
                }
            }
            Ok(vc)
        } else {
            Err(FromSafeArrayError::SafeArrayDimsInvalid{sa_dims: sa_dims})
        }
    }
} 

#[allow(dead_code)]
#[link(name="OleAut32")]
extern "system" {
     fn SafeArrayCreate(vt: VARTYPE, cDims: UINT, rgsabound: LPSAFEARRAYBOUND) -> LPSAFEARRAY;
	 fn SafeArrayDestroy(safe: LPSAFEARRAY)->HRESULT;
    
     fn SafeArrayGetDim(psa: LPSAFEARRAY) -> UINT;
	
     fn SafeArrayGetElement(psa: LPSAFEARRAY, rgIndices: *const c_long, pv: *mut c_void) -> HRESULT;
     fn SafeArrayGetElemSize(psa: LPSAFEARRAY) -> UINT;
    
     fn SafeArrayGetLBound(psa: LPSAFEARRAY, nDim: UINT, plLbound: *mut c_long)->HRESULT;
     fn SafeArrayGetUBound(psa: LPSAFEARRAY, nDim: UINT, plUbound: *mut c_long)->HRESULT;
    
     fn SafeArrayGetVartype(psa: LPSAFEARRAY, pvt: *mut VARTYPE) -> HRESULT;

     fn SafeArrayLock(psa: LPSAFEARRAY) -> HRESULT;
	 fn SafeArrayUnlock(psa: LPSAFEARRAY) -> HRESULT;
    
     fn SafeArrayPutElement(psa: LPSAFEARRAY, rgIndices: *const c_long, pv: *mut c_void) -> HRESULT;
}

#[cfg(test)]
mod test {
    use super::*;
    use std::vec::IntoIter;
    use rust_decimal::Decimal;
    macro_rules! validate_safe_arr {
        ($t:ident, $vals:expr, $vt:expr) => {
            let v: Vec<$t> = $vals;

            let p = v.into_iter().into_safearray().unwrap();
            
            let r: Result<Vec<$t>, FromSafeArrayError> = IntoIter::<$t>::from_safearray(p);
            let r = r.unwrap();
            assert_eq!(r, $vals);
        };
    }
    #[test]
    fn test_i16() {
        validate_safe_arr!(i16, vec![0,1,2,3,4], VT_I2 );
    }
    #[test]
    fn test_i32() {
        validate_safe_arr!(i32, vec![0,1,2,3,4], VT_I4 );
    }
    #[test]
    fn test_f32() {
        validate_safe_arr!(f32, vec![0.0f32,-1.333f32,2f32,3f32,4f32], VT_R4 );
    }
    #[test]
    fn test_f64() {
        validate_safe_arr!(f64, vec![0.0f64,-1.333f64,2f64,3f64,4f64], VT_R8 );
    }
    #[test]
    fn test_cy() {
        validate_safe_arr!(Currency, vec![Currency::from(-1), Currency::from(2)], VT_CY );
    }
    #[test]
    fn test_date() {
        validate_safe_arr!(Date, vec![Date::from(0.01), Date::from(100.0/99.0)], VT_DATE );
    }

    #[test]
    fn test_str() {
        let v: Vec<U16String> = vec![U16String::from_str("validate"), U16String::from_str("test string")];
        let mut v = v.into_iter();
        let p = <IntoIter<U16String> as SafeArrayExt<U16String>>::into_safearray(&mut v).unwrap();

        let r: Result<Vec<U16String>, FromSafeArrayError> = IntoIter::from_safearray(p);

        let r = r.unwrap();
        assert_eq!(r, vec![U16String::from_str("validate"), U16String::from_str("test string")]);
    }

    #[test]
    fn test_scode() {
        validate_safe_arr!(SCode, vec![SCode::from(100), SCode::from(10000)], VT_ERROR );
    }
    #[test]
    fn test_bool() {
        validate_safe_arr!(VariantBool, vec![VariantBool::from(true), VariantBool::from(false), VariantBool::from(true), VariantBool::from(true), VariantBool::from(false), VariantBool::from(false), VariantBool::from(true)], VT_BOOL );
    }

    #[test]
    fn test_variant() {
        let v: Vec<Variant<u64, u64>> = vec![Variant::wrap(100u64), Variant::wrap(100u64), Variant::wrap(103u64)];

        let mut p = v.into_iter();
        let p = <IntoIter<Variant<u64, u64>> as SafeArrayExt<Variant<u64, u64>>>::into_safearray(&mut p).unwrap();
        
        let r: Result< Vec<Variant<u64, u64>>, FromSafeArrayError> = IntoIter::from_safearray(p);
        let r = r.unwrap();
        assert_eq!(r,  vec![Variant::wrap(100u64), Variant::wrap(100u64), Variant::wrap(103u64)]);
    }

    #[test]
    fn test_decimal() {
        validate_safe_arr!(DecWrapper, vec![DecWrapper::from(Decimal::new(2, 2)), DecWrapper::from(Decimal::new(3, 3))], VE_DECIMAL );
    }
    #[test]
    fn test_i8() {
        validate_safe_arr!(i8, vec![-1, 0,1,2,3,4], VT_I1 );
    }
    #[test]
    fn test_u8() {
        validate_safe_arr!(u8, vec![0,1,2,3,4], VT_UI1 );
    }
    // #[test]
    // fn test_u16() {
    //     validate_safe_arr!(u16, vec![0,1,2,3,4], VT_UI2 );
    // }
    // #[test]
    // fn test_u32() {
    //     validate_safe_arr!(u32, vec![0,1,2,3,4], VT_UI4 );
    // }

    // #[test]
    // fn test_send() {
    //     fn assert_send<T: Send>() {}
    //     assert_send::<Variant<i64>>();
    // }

    // #[test]
    // fn test_sync() {
    //     fn assert_sync<T: Sync>() {}
    //     assert_sync::<Variant<i64>>();
    // }
}