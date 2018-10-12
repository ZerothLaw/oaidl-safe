//! Variant conversions
//! 
//! This module contains the trait [`VariantExt`] and the types [`Variant`], [`VtEmpty`], [`VtNull`].
//! 
//! It implements [`VariantExt`] for many built in types to enable conversions to VARIANT.  
//! 
//! [`VariantExt`]: trait.VariantExt.html
//! [`Variant`]: struct.Variant.html
//! [`VtEmpty`]: struct.VtEmpty.html
//! [`VtNull`]: struct.VtNull.html

/*
/// 
/// Reference:
/// typedef struct tagVARIANT {
///     union {
///         struct {
///             VARTYPE vt;
///             WORD    wReserved1;
///             WORD    wReserved2;
///             WORD    wReserved3;
///             union {
///                 LONGLONG     llVal;
///                 LONG         lVal;
///                 BYTE         bVal;
///                 SHORT        iVal;
///                 FLOAT        fltVal;
///                 DOUBLE       dblVal;
///                 VARIANT_BOOL boolVal;
///                 SCODE        scode;
///                 CY           cyVal;
///                 DATE         date;
///                 BSTR         bstrVal;
///                 IUnknown     *punkVal;
///                 IDispatch    *pdispVal;
///                 SAFEARRAY    *parray;
///                 BYTE         *pbVal;
///                 SHORT        *piVal;
///                 LONG         *plVal;
///                 LONGLONG     *pllVal;
///                 FLOAT        *pfltVal;
///                 DOUBLE       *pdblVal;
///                 VARIANT_BOOL *pboolVal;
///                 SCODE        *pscode;
///                 CY           *pcyVal;
///                 DATE         *pdate;
///                 BSTR         *pbstrVal;
///                 IUnknown     **ppunkVal;
///                 IDispatch    **ppdispVal;
///                 SAFEARRAY    **pparray;
///                 VARIANT      *pvarVal;
///                 PVOID        byref;
///                 CHAR         cVal;
///                 USHORT       uiVal;
///                 ULONG        ulVal;
///                 ULONGLONG    ullVal;
///                 INT          intVal;
///                 UINT         uintVal;
///                 DECIMAL      *pdecVal;
///                 CHAR         *pcVal;
///                 USHORT       *puiVal;
///                 ULONG        *pulVal;
///                 ULONGLONG    *pullVal;
///                 INT          *pintVal;
///                 UINT         *puintVal;
///                 struct {
///                     PVOID       pvRecord;
///                     IRecordInfo *pRecInfo;
///                 } __VARIANT_NAME_4;
///             } __VARIANT_NAME_3;
///         } __VARIANT_NAME_2;
///         DECIMAL decVal;
///     } __VARIANT_NAME_1;
/// } VARIANT;*/
/*
* VARENUM usage key,
*
* * [V] - may appear in a VARIANT
* * [T] - may appear in a TYPEDESC
* * [P] - may appear in an OLE property set
* * [S] - may appear in a Safe Array
*
*
*  VT_EMPTY            [V]   [P]     nothing
*  VT_NULL             [V]   [P]     SQL style Null
*  VT_I2               [V][T][P][S]  2 byte signed int
*  VT_I4               [V][T][P][S]  4 byte signed int
*  VT_R4               [V][T][P][S]  4 byte real
*  VT_R8               [V][T][P][S]  8 byte real
*  VT_CY               [V][T][P][S]  currency
*  VT_DATE             [V][T][P][S]  date
*  VT_BSTR             [V][T][P][S]  OLE Automation string
*  VT_DISPATCH         [V][T]   [S]  IDispatch *
*  VT_ERROR            [V][T][P][S]  SCODE
*  VT_BOOL             [V][T][P][S]  True=-1, False=0
*  VT_VARIANT          [V][T][P][S]  VARIANT *
*  VT_UNKNOWN          [V][T]   [S]  IUnknown *
*  VT_DECIMAL          [V][T]   [S]  16 byte fixed point
*  VT_RECORD           [V]   [P][S]  user defined type
*  VT_I1               [V][T][P][s]  signed char
*  VT_UI1              [V][T][P][S]  unsigned char
*  VT_UI2              [V][T][P][S]  unsigned short
*  VT_UI4              [V][T][P][S]  ULONG
*  VT_I8                  [T][P]     signed 64-bit int
*  VT_UI8                 [T][P]     unsigned 64-bit int
*  VT_INT              [V][T][P][S]  signed machine int
*  VT_UINT             [V][T]   [S]  unsigned machine int
*  VT_INT_PTR             [T]        signed machine register size width
*  VT_UINT_PTR            [T]        unsigned machine register size width
*  VT_VOID                [T]        C style void
*  VT_HRESULT             [T]        Standard return type
*  VT_PTR                 [T]        pointer type
*  VT_SAFEARRAY           [T]        (use VT_ARRAY in VARIANT)
*  VT_CARRAY              [T]        C style array
*  VT_USERDEFINED         [T]        user defined type
*  VT_LPSTR               [T][P]     null terminated string
*  VT_LPWSTR              [T][P]     wide null terminated string
*  VT_FILETIME               [P]     FILETIME
*  VT_BLOB                   [P]     Length prefixed bytes
*  VT_STREAM                 [P]     Name of the stream follows
*  VT_STORAGE                [P]     Name of the storage follows
*  VT_STREAMED_OBJECT        [P]     Stream contains an object
*  VT_STORED_OBJECT          [P]     Storage contains an object
*  VT_VERSIONED_STREAM       [P]     Stream with a GUID version
*  VT_BLOB_OBJECT            [P]     Blob contains an object 
*  VT_CF                     [P]     Clipboard format
*  VT_CLSID                  [P]     A Class ID
*  VT_VECTOR                 [P]     simple counted array
*  VT_ARRAY            [V]           SAFEARRAY*
*  VT_BYREF            [V]           void* for local use
*  VT_BSTR_BLOB                      Reserved for system use
*/
use std::marker::PhantomData;
use std::mem;
use std::ptr::{NonNull, null_mut};

use rust_decimal::Decimal;

use widestring::U16String;

use winapi::ctypes::c_void;
use winapi::shared::wtypes::{
    BSTR, CY, DATE, DECIMAL,
    VARIANT_BOOL,
    VT_ARRAY, 
    VT_BSTR, 
    VT_BOOL,
    VT_BYREF, 
    VT_CY,
    VT_DATE,
    VT_DECIMAL, 
    VT_DISPATCH,
    VT_EMPTY, 
    VT_ERROR, 
    VT_I1, 
    VT_I2, 
    VT_I4,
    VT_I8,
    VT_INT,  
    VT_NULL,
    VT_R4, 
    VT_R8, 
    //VT_RECORD,
    VT_UI1,
    VT_UI2,
    VT_UI4,
    VT_UI8,  
    VT_UINT, 
    VT_UNKNOWN, 
    VT_VARIANT, 
};
use winapi::shared::wtypesbase::SCODE;
use winapi::um::oaidl::{IDispatch,  __tagVARIANT, SAFEARRAY, VARIANT, VARIANT_n3, VARIANT_n1};
use winapi::um::oleauto::VariantClear;
use winapi::um::unknwnbase::IUnknown;

use super::array::{SafeArrayElement, SafeArrayExt};
use super::bstr::BStringExt;
use super::errors::{IntoVariantError, FromVariantError};
use super::ptr::Ptr;
use super::types::{Date, DecWrapper, Currency, Int, SCode, UInt, VariantBool };

const VT_PUI1:      u32 = VT_BYREF | VT_UI1;
const VT_PI2:       u32 = VT_BYREF | VT_I2;
const VT_PI4:       u32 = VT_BYREF | VT_I4;
const VT_PI8:       u32 = VT_BYREF | VT_I8;
const VT_PUI8:      u32 = VT_BYREF | VT_UI8;
const VT_PR4:       u32 = VT_BYREF | VT_R4;
const VT_PR8:       u32 = VT_BYREF | VT_R8;
const VT_PBOOL:     u32 = VT_BYREF | VT_BOOL;
const VT_PERROR:    u32 = VT_BYREF | VT_ERROR;
const VT_PCY:       u32 = VT_BYREF | VT_CY;
const VT_PDATE:     u32 = VT_BYREF | VT_DATE;
const VT_PBSTR:     u32 = VT_BYREF | VT_BSTR;
const VT_PUNKNOWN:  u32 = VT_BYREF | VT_UNKNOWN;
const VT_PDISPATCH: u32 = VT_BYREF | VT_DISPATCH;
const VT_PARRAY:    u32 = VT_BYREF | VT_ARRAY;
const VT_PDECIMAL:  u32 = VT_BYREF | VT_DECIMAL;
const VT_PI1:       u32 = VT_BYREF | VT_I1;
const VT_PUI2:      u32 = VT_BYREF | VT_UI2;
const VT_PUI4:      u32 = VT_BYREF | VT_UI4;
const VT_PINT:      u32 = VT_BYREF | VT_INT;
const VT_PUINT:     u32 = VT_BYREF | VT_UINT;

mod private {
    use super::*;
    #[doc(hidden)]
    pub trait VariantAccess: Sized {
        #[doc(hidden)]
        const VTYPE: u32;
        type Field;

        #[doc(hidden)]
        fn from_variant(n1: &VARIANT_n1, n3: &VARIANT_n3) -> Self::Field;
        
        #[doc(hidden)]
        fn into_variant(inner: Self::Field, n1: &mut VARIANT_n1, n3: &mut VARIANT_n3);
    }

    macro_rules! impl_conversions {
        ($t:ty, $f:ty, $vtype:ident, $member:ident, $member_mut:ident) => {
            impl VariantAccess for $t {
                const VTYPE: u32 = $vtype;
                type Field = $f;
                fn from_variant(_n1: &VARIANT_n1, n3: &VARIANT_n3) -> Self::Field {
                    unsafe {*n3.$member()}
                }
                
                fn into_variant(inner: Self::Field, _n1: &mut VARIANT_n1, n3: &mut VARIANT_n3) {
                    unsafe {
                        let n_ptr = n3.$member_mut();
                        *n_ptr = inner
                    }
                }
            }

            impl<'s> VariantAccess for &'s $t {
                const VTYPE: u32 = $vtype;
                type Field = $f;
                fn from_variant(_n1: &VARIANT_n1, n3: &VARIANT_n3) -> Self::Field {
                    unsafe {*n3.$member()}
                }
                
                fn into_variant(inner: Self::Field, _n1: &mut VARIANT_n1, n3: &mut VARIANT_n3) {
                    unsafe {
                        let n_ptr = n3.$member_mut();
                        *n_ptr = inner
                    }
                }
            }

            impl<'s> VariantAccess for &'s mut $t {
                const VTYPE: u32 = $vtype;
                type Field = $f;
                fn from_variant(_n1: &VARIANT_n1, n3: &VARIANT_n3) -> Self::Field {
                    unsafe {*n3.$member()}
                }
                
                fn into_variant(inner: Self::Field, _n1: &mut VARIANT_n1, n3: &mut VARIANT_n3) {
                    unsafe {
                        let n_ptr = n3.$member_mut();
                        *n_ptr = inner
                    }
                }
            }
        };
    }

    impl_conversions!(i64, i64, VT_I8, llVal, llVal_mut);
    impl_conversions!(i32, i32, VT_I4, lVal, lVal_mut);
    impl_conversions!(u8, u8, VT_UI1,  bVal, bVal_mut);
    impl_conversions!(i16, i16, VT_I2, iVal, iVal_mut);
    impl_conversions!(f32, f32, VT_R4, fltVal, fltVal_mut);
    impl_conversions!(f64, f64, VT_R8, dblVal, dblVal_mut);
    impl_conversions!(VariantBool, VARIANT_BOOL, VT_BOOL, boolVal, boolVal_mut);
    impl_conversions!(bool, VARIANT_BOOL, VT_BOOL, boolVal, boolVal_mut);
    impl_conversions!(SCode, SCODE, VT_ERROR, scode, scode_mut);
    impl_conversions!(Currency, CY, VT_CY, cyVal, cyVal_mut);
    impl_conversions!(Date, DATE, VT_DATE, date, date_mut);
    impl<S> VariantAccess for Vec<S>
    where
        S: SafeArrayElement
    {
        const VTYPE: u32 = VT_ARRAY;
        type Field = *mut SAFEARRAY;
        fn from_variant(_n1: &VARIANT_n1, n3: &VARIANT_n3) -> Self::Field {
            unsafe {*n3.parray()}
        }
        
        fn into_variant(inner: Self::Field, _n1: &mut VARIANT_n1, n3: &mut VARIANT_n3){
            unsafe {
                let n_ptr = n3.parray_mut();
                *n_ptr = inner;
            }
        }
    }   

    impl<'s, S> VariantAccess for &'s [S] 
    where 
        S: SafeArrayElement
    {
        const VTYPE: u32 = VT_ARRAY;
        type Field = *mut SAFEARRAY;
        fn from_variant(_n1: &VARIANT_n1, n3: &VARIANT_n3) -> Self::Field {
            unsafe {*n3.parray()}
        }
        
        fn into_variant(inner: Self::Field, _n1: &mut VARIANT_n1, n3: &mut VARIANT_n3){
            unsafe {
                let n_ptr = n3.parray_mut();
                *n_ptr = inner;
            }
        }
    } 
    
    impl_conversions!(U16String, BSTR, VT_BSTR, bstrVal, bstrVal_mut);
    impl_conversions!(Ptr<IUnknown>, *mut IUnknown, VT_UNKNOWN, punkVal, punkVal_mut);
    impl_conversions!(Ptr<IDispatch>, *mut IDispatch, VT_DISPATCH,  pdispVal, pdispVal_mut);
    impl_conversions!(Box<u8>, *mut u8,  VT_PUI1, pbVal, pbVal_mut);
    impl_conversions!(Box<i16>, *mut i16, VT_PI2,  piVal, piVal_mut);
    impl_conversions!(Box<i32>, *mut i32, VT_PI4, plVal, plVal_mut);
    impl_conversions!(Box<i64>, *mut i64, VT_PI8, pllVal, pllVal_mut);
    impl_conversions!(Box<f32>, *mut f32, VT_PR4, pfltVal, pfltVal_mut);
    impl_conversions!(Box<f64>, *mut f64, VT_PR8, pdblVal, pdblVal_mut);
    impl_conversions!(Box<VariantBool>, *mut VARIANT_BOOL, VT_PBOOL, pboolVal, pboolVal_mut);
    impl_conversions!(Box<SCode>, *mut SCODE, VT_PERROR, pscode, pscode_mut);
    impl_conversions!(Box<Currency>, *mut CY, VT_PCY, pcyVal, pcyVal_mut);
    impl_conversions!(Box<Date>, *mut DATE, VT_PDATE, pdate, pdate_mut);
    impl_conversions!(Box<U16String>, *mut BSTR, VT_PBSTR, pbstrVal, pbstrVal_mut);
    impl_conversions!(Box<Ptr<IUnknown>>, *mut *mut IUnknown,  VT_PUNKNOWN, ppunkVal, ppunkVal_mut);
    impl_conversions!(Box<Ptr<IDispatch>>, *mut *mut IDispatch, VT_PDISPATCH, ppdispVal, ppdispVal_mut);
    impl<S> VariantAccess for Box<Vec<S>>
    where
        S: SafeArrayElement
    {
        const VTYPE: u32 = VT_PARRAY;
        type Field = *mut *mut SAFEARRAY;
        fn from_variant(_n1: &VARIANT_n1, n3: &VARIANT_n3) -> Self::Field {
            unsafe {*n3.pparray()}
        }
        
        fn into_variant(inner: Self::Field, _n1: &mut VARIANT_n1, n3: &mut VARIANT_n3){
            unsafe {
                let n_ptr = n3.pparray_mut();
                *n_ptr = inner;
            }
        }
    }

    impl<'s, S> VariantAccess for Box<&'s [S]> 
    where 
        S: SafeArrayElement
    {
        const VTYPE: u32 = VT_PARRAY;
        type Field = *mut *mut SAFEARRAY;
        fn from_variant(_n1: &VARIANT_n1, n3: &VARIANT_n3) -> Self::Field {
            unsafe {*n3.pparray()}
        }
        
        fn into_variant(inner: Self::Field, _n1: &mut VARIANT_n1, n3: &mut VARIANT_n3){
            unsafe {
                let n_ptr = n3.pparray_mut();
                *n_ptr = inner;
            }
        }
    } 
    //pvarVal - need to redo Variant<T> first.
    impl_conversions!(Ptr<c_void>, *mut c_void, VT_BYREF, byref, byref_mut);
    impl_conversions!(i8, i8, VT_I1, cVal, cVal_mut);
    impl_conversions!(u16, u16, VT_UI2, uiVal, uiVal_mut);
    impl_conversions!(u32, u32, VT_UI4, ulVal, ulVal_mut);
    impl_conversions!(u64, u64, VT_UI8, ullVal, ullVal_mut);
    impl_conversions!(Int, i32, VT_INT, intVal, intVal_mut);
    impl_conversions!(UInt, u32, VT_UINT, uintVal, uintVal_mut);
    impl_conversions!(Box<DecWrapper>, *mut DECIMAL, VT_PDECIMAL, pdecVal, pdecVal_mut);
    impl_conversions!(Box<i8>, *mut i8, VT_PI1, pcVal, pcVal_mut);
    impl_conversions!(Box<u16>, *mut u16, VT_PUI2,  puiVal, puiVal_mut);
    impl_conversions!(Box<u32>, *mut u32, VT_PUI4, pulVal, pulVal_mut);
    impl_conversions!(Box<u64>, *mut u64, VT_PUI8, pullVal, pullVal_mut);
    impl_conversions!(Box<Int>, *mut i32, VT_PINT, pintVal, pintVal_mut);
    impl_conversions!(Box<UInt>, *mut u32, VT_PUINT, puintVal, puintVal_mut);
    impl VariantAccess for DecWrapper {
        const VTYPE: u32 = VT_DECIMAL;
        type Field = DECIMAL;
        fn from_variant(n1: &VARIANT_n1, _n3: &VARIANT_n3) -> Self::Field {
            unsafe {*n1.decVal()}
        }
        
        fn into_variant(inner: Self::Field, n1: &mut VARIANT_n1, _n3: &mut VARIANT_n3) {
            unsafe {
                let n_ptr = n1.decVal_mut();
                *n_ptr = inner;
            }
        }
    }

    impl<'s> VariantAccess for &'s DecWrapper {
        const VTYPE: u32 = VT_DECIMAL;
        type Field = DECIMAL;
        fn from_variant(n1: &VARIANT_n1, _n3: &VARIANT_n3) -> Self::Field {
            unsafe {*n1.decVal()}
        }
        
        fn into_variant(inner: Self::Field, n1: &mut VARIANT_n1, _n3: &mut VARIANT_n3) {
            unsafe {
                let n_ptr = n1.decVal_mut();
                *n_ptr = inner;
            }
        }
    }

    impl<'s> VariantAccess for &'s mut DecWrapper {
        const VTYPE: u32 = VT_DECIMAL;
        type Field = DECIMAL;
        fn from_variant(n1: &VARIANT_n1, _n3: &VARIANT_n3) -> Self::Field {
            unsafe {*n1.decVal()}
        }
        
        fn into_variant(inner: Self::Field, n1: &mut VARIANT_n1, _n3: &mut VARIANT_n3) {
            unsafe {
                let n_ptr = n1.decVal_mut();
                *n_ptr = inner;
            }
        }
    }
}

/// Trait implemented to convert the type into a VARIANT
/// Do not implement this yourself without care. 
pub trait VariantExt: Sized { //Would like Clone, but *mut IDispatch and *mut IUnknown don't implement them
    /// VARTYPE constant value for the type
    const VARTYPE: u32;

    /// Call this associated function on a Ptr<VARIANT> to obtain a value T
    fn from_variant(var: Ptr<VARIANT>) -> Result<Self, FromVariantError>;  

    /// Convert a value of type T into a Ptr<VARIANT>
    fn into_variant(self) -> Result<Ptr<VARIANT>, IntoVariantError>;
}

/// Helper struct to wrap a VARIANT compatible type into a VT_VARIANT marked VARIANT
#[derive(Clone, Copy, Debug, Eq, Hash, Ord, PartialEq, PartialOrd)]
pub struct Variant<T: VariantExt>(T);

/// `Variant<T: VariantExt>` type to wrap an impl of `VariantExt` - creates a VARIANT of VT_VARIANT
/// which wraps an inner variant that points to T. 
impl<T: VariantExt> Variant<T> {

    /// default constructor
    pub fn new(t: T) -> Variant<T> {
        Variant(t)
    }

    /// Get access to the inner value and the Variant is consumed
    pub fn unwrap(self) -> T {
        self.0
    }

    /// Borrow reference to inner value
    pub fn borrow(&self) -> &T {
        &self.0
    }

    /// Borrow mutable reference to inner value
    pub fn borrow_mut(&mut self) -> &mut T {
        &mut self.0
    }

    /// Converts the `Variant<T>` into a `Ptr<VARIANT>`
    /// Returns `Result<Ptr<VARIANT>, IntoVariantError>`
    pub fn into_variant(self) -> Result<Ptr<VARIANT>, IntoVariantError> {
        #[allow(unused_mut)]
        let mut n3: VARIANT_n3 = unsafe {mem::zeroed()};
        let mut n1: VARIANT_n1 = unsafe {mem::zeroed()};
        let var = self.0.into_variant()?;
        let var = var.as_ptr();
        unsafe {
            let n_ptr = n3.pvarVal_mut();
            *n_ptr = var;
        };

        let tv = __tagVARIANT { vt: <Self as VariantExt>::VARTYPE as u16, 
                        wReserved1: 0, 
                        wReserved2: 0, 
                        wReserved3: 0, 
                        n3: n3};
        unsafe {
            let n_ptr = n1.n2_mut();
            *n_ptr = tv;
        };
        let var = Box::new(VARIANT{ n1: n1 });
        Ok(Ptr::with_checked(Box::into_raw(var)).unwrap())
    }

    /// Converts `Ptr<VARIANT>` into  `Variant<T>` 
    /// Returns `Result<Variant<T>>, FromVariantError>`
    pub fn from_variant(var: Ptr<VARIANT>) -> Result<Variant<T>, FromVariantError> {
        let var = var.as_ptr();
        let mut _var_d = VariantDestructor::new(var);

        let mut n1 = unsafe {(*var).n1};
        let n3 = unsafe { n1.n2_mut().n3 };
        let n_ptr = unsafe {
            let n_ptr = n3.pvarVal();
            *n_ptr
        };

        let pnn = match Ptr::with_checked(n_ptr) {
            Some(nn) => nn, 
            None => return Err(FromVariantError::VariantPtrNull) 
        };
    
        let t = T::from_variant(pnn).unwrap();
        Ok(Variant(t))
    }
}

impl<T: VariantExt> AsRef<T> for Variant<T> {
    fn as_ref(&self) -> &T {
        self.borrow()
    }
}

impl<T: VariantExt> AsMut<T> for Variant<T> {
    fn as_mut(&mut self) -> &mut T {
        self.borrow_mut()
    }
}

struct VariantDestructor {
    inner: *mut VARIANT, 
    _marker: PhantomData<VARIANT>
}

impl VariantDestructor {
    fn new(p: *mut VARIANT) -> VariantDestructor {
        VariantDestructor {
            inner: p, 
            _marker: PhantomData
        }
    }
}

impl Drop for VariantDestructor {
    fn drop(&mut self) {
        if self.inner.is_null() {
            return;
        }
        unsafe { VariantClear(self.inner)};
        unsafe { let _dtor = *self.inner;}
        self.inner = null_mut();
    }
}

macro_rules! variant_impl {
    (
        impl $(<$tn:ident : $tc:ident>)* VariantExt for $t:ty {
            VARTYPE = $vt:expr ;
            $n_name:ident, $un_n:ident, $un_n_mut:ident
            from => {$from:expr}
            into => {$into:expr}
        }
    ) => {
        impl $(<$tn: $tc>)* VariantExt for $t {
            const VARTYPE: u32 = $vt;
            fn from_variant(var: Ptr<VARIANT>) -> Result<Self, FromVariantError>{
                let var = var.as_ptr();
                let mut var_d = VariantDestructor::new(var);

                #[allow(unused_mut)]
                let mut n1 = unsafe {(*var).n1};
                let vt = unsafe{n1.n2()}.vt;
                if vt as u32 != Self::VARTYPE {
                    return Err(FromVariantError::VarTypeDoesNotMatch{expected: Self::VARTYPE, found: vt as u32})
                }
                let ret = variant_impl!(@read $n_name, $un_n, $from, n1);

                var_d.inner = null_mut();
                ret
            }

            fn into_variant(self) -> Result<Ptr<VARIANT>, IntoVariantError> {
                #[allow(unused_mut)]
                let mut n3: VARIANT_n3 = unsafe {mem::zeroed()};
                let mut n1: VARIANT_n1 = unsafe {mem::zeroed()};
                variant_impl!(@write $n_name, $un_n_mut, $into, n3, n1, self);
                let tv = __tagVARIANT { vt: <Self as VariantExt>::VARTYPE as u16, 
                                wReserved1: 0, 
                                wReserved2: 0, 
                                wReserved3: 0, 
                                n3: n3};
                unsafe {
                    let n_ptr = n1.n2_mut();
                    *n_ptr = tv;
                };
                let var = Box::new(VARIANT{ n1: n1 });
                Ok(Ptr::with_checked(Box::into_raw(var)).unwrap())
            }
        }
    };
    (@read n3, $un_n:ident, $from:expr, $n1:ident) => {
        {
            let n3 = unsafe { $n1.n2_mut().n3 };
            let ret = unsafe {
                let n_ptr = n3.$un_n();
                $from(n_ptr)
            };
            ret 
        }
    };
    (@read n1, $un_n:ident, $from:expr, $n1:ident) => {
        {
            let ret = unsafe {
                let n_ptr = $n1.$un_n();
                $from(n_ptr)
            };
            ret
        }
    };
    (@write n3, $un_n_mut:ident, $into:expr, $n3:ident, $n1:ident, $slf:expr) => {
        unsafe {
            let n_ptr = $n3.$un_n_mut();
            *n_ptr = $into($slf)?
        }
    };
    (@write n1, $un_n_mut:ident, $into:expr, $n3:ident, $n1:ident, $slf:expr) => {
        unsafe {
            let n_ptr = $n1.$un_n_mut();
            *n_ptr = $into($slf)?
        }
    };
}

variant_impl!{
    impl VariantExt for i64 {
        VARTYPE = VT_I8;
        n3, llVal, llVal_mut
        from => {|n_ptr: &i64| {Ok(*n_ptr)}}
        into => {|slf: i64| -> Result<_, IntoVariantError> {Ok(slf)}}
    }
}
variant_impl!{
    impl VariantExt for i32 {
        VARTYPE = VT_I4;
        n3, lVal, lVal_mut
        from => {|n_ptr: &i32| Ok(*n_ptr)}
        into => {|slf: i32| -> Result<_, IntoVariantError> {Ok(slf)}}
    }
}
variant_impl!{
    impl VariantExt for u8 {
        VARTYPE = VT_UI1;
        n3, bVal, bVal_mut
        from => {|n_ptr: &u8| Ok(*n_ptr)}
        into => {|slf: u8| -> Result<_, IntoVariantError> {Ok(slf)}}
    }
}
variant_impl!{
    impl VariantExt for i16 {
        VARTYPE = VT_I2;
        n3, iVal, iVal_mut
        from => {|n_ptr: &i16| Ok(*n_ptr)}
        into => {|slf: i16| -> Result<_, IntoVariantError> {Ok(slf)}}
    }
}
variant_impl!{
    impl VariantExt for f32 {
        VARTYPE = VT_R4;
        n3, fltVal, fltVal_mut
        from => {|n_ptr: &f32| Ok(*n_ptr)}
        into => {|slf: f32| -> Result<_, IntoVariantError> {Ok(slf)}}
    }
}
variant_impl!{
    impl VariantExt for f64 {
        VARTYPE = VT_R8;
        n3, dblVal, dblVal_mut
        from => {|n_ptr: &f64| Ok(*n_ptr)}
        into => {|slf: f64| -> Result<_, IntoVariantError> {Ok(slf)}}
    }
}
variant_impl!{
    impl VariantExt for bool {
        VARTYPE = VT_BOOL;
        n3, boolVal, boolVal_mut
        from => {|n_ptr: &VARIANT_BOOL| Ok(bool::from(VariantBool::from(*n_ptr)))}
        into => {|slf: bool| -> Result<_, IntoVariantError> {
            Ok(VARIANT_BOOL::from(VariantBool::from(slf)))
        }}
    }
}
variant_impl!{
    impl VariantExt for SCode {
        VARTYPE = VT_ERROR;
        n3, scode, scode_mut
        from => {|n_ptr: &SCODE| Ok(SCode::from(*n_ptr))}
        into => {|slf: SCode| -> Result<_, IntoVariantError> { 
            Ok(i32::from(slf))
        }}
    }
}
variant_impl!{
    impl VariantExt for Currency {
        VARTYPE = VT_CY;
        n3, cyVal, cyVal_mut
        from => {|n_ptr: &CY| Ok(Currency::from(*n_ptr))}
        into => {|slf: Currency| -> Result<_, IntoVariantError> {Ok(CY::from(slf))}}
    }
}
variant_impl!{
    impl VariantExt for Date {
        VARTYPE = VT_DATE;
        n3, date, date_mut
        from => {|n_ptr: &DATE| Ok(Date::from(*n_ptr))}
        into => {|slf: Date| -> Result<_, IntoVariantError> {Ok(DATE::from(slf))}}
    }
}
variant_impl!{
    impl VariantExt for String {
        VARTYPE = VT_BSTR;
        n3, bstrVal, bstrVal_mut
        from => {|n_ptr: &*mut u16| {
            let bstr = U16String::from_bstr(*n_ptr);
            Ok(bstr.to_string_lossy())
        }}
        into => {|slf: String|{
            let mut bstr = U16String::from_str(&slf);
            match bstr.allocate_bstr(){
                Ok(ptr) => Ok(ptr.as_ptr()), 
                Err(bse) => Err(IntoVariantError::from(bse))
            }
        }}
    }
}
variant_impl!{
    impl VariantExt for Ptr<IUnknown> {
        VARTYPE = VT_UNKNOWN;
        n3, punkVal, punkVal_mut
        from => {|n_ptr: &* mut IUnknown| Ok(Ptr::with_checked(*n_ptr).unwrap())}
        into => {|slf: Ptr<IUnknown>| -> Result<_, IntoVariantError> {Ok(slf.as_ptr())}}
    }
}
variant_impl!{
    impl VariantExt for Ptr<IDispatch> {
        VARTYPE = VT_DISPATCH;
        n3, pdispVal, pdispVal_mut
        from => {|n_ptr: &*mut IDispatch| Ok(Ptr::with_checked(*n_ptr).unwrap())}
        into => {|slf: Ptr<IDispatch>| -> Result<_, IntoVariantError> { Ok(slf.as_ptr()) }}
    }
}
variant_impl!{
    impl VariantExt for Box<u8> {
        VARTYPE = VT_PUI1;
        n3, pbVal, pbVal_mut
        from => {|n_ptr: &* mut u8| Ok(Box::new(**n_ptr))}
        into => {|slf: Box<u8>| -> Result<_, IntoVariantError> {
            Ok(Box::into_raw(slf))
        }}
    }
}
variant_impl!{
    impl VariantExt for Box<i16> {
        VARTYPE = VT_PI2;
        n3, piVal, piVal_mut
        from => {|n_ptr: &* mut i16| Ok(Box::new(**n_ptr))}
        into => {|slf: Box<i16>| -> Result<_, IntoVariantError> {
            Ok(Box::into_raw(slf))
        }}
    }
}
variant_impl!{
    impl VariantExt for Box<i32> {
        VARTYPE = VT_PI4;
        n3, plVal, plVal_mut
        from => {|n_ptr: &* mut i32| Ok(Box::new(**n_ptr))}
        into => {|slf: Box<i32>| -> Result<_, IntoVariantError> {
            Ok(Box::into_raw(slf))
        }}
    }
}
variant_impl!{
    impl VariantExt for Box<i64> {
        VARTYPE = VT_PI8;
        n3, pllVal, pllVal_mut
        from => {|n_ptr: &* mut i64| Ok(Box::new(**n_ptr))}
        into => {|slf: Box<i64>| -> Result<_, IntoVariantError> {
            Ok(Box::into_raw(slf))
        }}
    }
}
variant_impl!{
    impl VariantExt for Box<f32> {
        VARTYPE = VT_PR4;
        n3, pfltVal, pfltVal_mut
        from => {|n_ptr: &* mut f32| Ok(Box::new(**n_ptr))}
        into => {|slf: Box<f32>| -> Result<_, IntoVariantError> {
            Ok(Box::into_raw(slf))
        }}
    }
}
variant_impl!{
    impl VariantExt for Box<f64> {
        VARTYPE = VT_PR8;
        n3, pdblVal, pdblVal_mut
        from => {|n_ptr: &* mut f64| Ok(Box::new(**n_ptr))}
        into => {|slf: Box<f64>| -> Result<_, IntoVariantError> {
            Ok(Box::into_raw(slf))
        }}
    }
}
variant_impl!{
    impl VariantExt for Box<bool> {
        VARTYPE = VT_PBOOL;
        n3, pboolVal, pboolVal_mut
        from => {
            |n_ptr: &*mut VARIANT_BOOL| Ok(Box::new(bool::from(VariantBool::from(**n_ptr))))
        }
        into => {
            |slf: Box<bool>|-> Result<_, IntoVariantError> {
                Ok(Box::into_raw(Box::new(VARIANT_BOOL::from(VariantBool::from(*slf)))))
            }
        }
    }
}
variant_impl!{
    impl VariantExt for Box<SCode> {
        VARTYPE = VT_PERROR;
        n3, pscode, pscode_mut
        from => {|n_ptr: &*mut SCODE| Ok(Box::new(SCode::from(**n_ptr)))}
        into => {|slf: Box<SCode>| -> Result<_, IntoVariantError> {
            Ok(Box::into_raw(Box::new(i32::from(*slf))))
        }}
    }
}
variant_impl!{
    impl VariantExt for Box<Currency> {
        VARTYPE = VT_PCY;
        n3, pcyVal, pcyVal_mut
        from => { |n_ptr: &*mut CY| Ok(Box::new(Currency::from(**n_ptr))) }
        into => {
            |slf: Box<Currency>|-> Result<_, IntoVariantError>  {
                Ok(Box::into_raw(Box::new(CY::from(*slf))))
            }
        }
    }
}
variant_impl!{
    impl VariantExt for Box<Date> {
        VARTYPE = VT_PDATE;
        n3, pdate, pdate_mut
        from => { |n_ptr: &*mut f64| Ok(Box::new(Date::from(**n_ptr))) }
        into => {
            |slf: Box<Date>|-> Result<_, IntoVariantError>  {
                let bptr = Box::new(DATE::from(*slf));
                Ok(Box::into_raw(bptr))
            }
        }
    }
}
variant_impl!{
    impl VariantExt for Box<String> {
        VARTYPE = VT_PBSTR;
        n3, pbstrVal, pbstrVal_mut
        from => {|n_ptr: &*mut *mut u16| {
            let bstr = U16String::from_bstr(**n_ptr);
            Ok(Box::new(bstr.to_string_lossy()))
        }}
        into => {|slf: Box<String>| -> Result<_, IntoVariantError> {
            let mut bstr = U16String::from_str(&*slf);
            let bstr = Box::new(bstr.allocate_bstr().unwrap().as_ptr());
            Ok(Box::into_raw(bstr))
        }}
    }
}
variant_impl! {
    impl VariantExt for Box<Ptr<IUnknown>> {
        VARTYPE = VT_PUNKNOWN;
        n3, ppunkVal, ppunkVal_mut
        from => {
            |n_ptr: &*mut *mut IUnknown| {
                match NonNull::new((**n_ptr).clone()) {
                    Some(nn) => Ok(Box::new(Ptr::new(nn))), 
                    None => Err(FromVariantError::UnknownPtrNull)
                }
            }
        }
        into => {
            |slf: Box<Ptr<IUnknown>>| -> Result<_, IntoVariantError> {
                Ok(Box::into_raw(Box::new((*slf).as_ptr())))
            }
        }
    }
}
variant_impl! {
    impl VariantExt for Box<Ptr<IDispatch>> {
        VARTYPE = VT_PDISPATCH;
        n3, ppdispVal, ppdispVal_mut
        from => {
            |n_ptr: &*mut *mut IDispatch| {
                match Ptr::with_checked((**n_ptr).clone()) {
                    Some(nn) => Ok(Box::new(nn)), 
                    None => Err(FromVariantError::DispatchPtrNull)
                }
            }
        }
        into => {
            |slf: Box<Ptr<IDispatch>>| -> Result<_, IntoVariantError> {
                Ok(Box::into_raw(Box::new((*slf).as_ptr())))
            }
        }
    }
}
variant_impl!{
    impl<T: VariantExt> VariantExt for Variant<T> {
        VARTYPE = VT_VARIANT;
        n3, pvarVal, pvarVal_mut
        from => {|n_ptr: &*mut VARIANT| {
            let pnn = match Ptr::with_checked(*n_ptr) {
                Some(nn) => nn, 
                None => return Err(FromVariantError::VariantPtrNull) 
            };
            Variant::<T>::from_variant(pnn)
        }}
        into => {|slf: Variant<T>| -> Result<_, IntoVariantError> {
            let pvar = slf.into_variant().unwrap();
            Ok(pvar.as_ptr())
        }}
    }
}
variant_impl!{
    impl<T: SafeArrayElement> VariantExt for Vec<T>{
        VARTYPE = VT_ARRAY;
        n3, parray, parray_mut
        from => {
            |n_ptr: &*mut SAFEARRAY| {
                match ExactSizeIterator::<Item=T>::from_safearray(*n_ptr) {
                    Ok(sa) => Ok(sa), 
                    Err(fsae) => Err(FromVariantError::from(fsae))
                }
            }
        }
        into => {
            |slf: Vec<T>| -> Result<_, IntoVariantError> {
                match slf.into_iter().into_safearray() {
                    Ok(psa) => {
                        Ok(psa.as_ptr())
                    }, 
                    Err(isae) => {
                        Err(IntoVariantError::from(isae))
                    }
                }
            }
        }
    }
}
variant_impl!{
    impl VariantExt for Ptr<c_void> {
        VARTYPE = VT_BYREF;
        n3, byref, byref_mut
        from => {|n_ptr: &*mut c_void| {
            match Ptr::with_checked(*n_ptr) {
                Some(nn) => Ok(nn), 
                None => Err(FromVariantError::CVoidPtrNull)
            }
        }}
        into => {|slf: Ptr<c_void>| -> Result<_, IntoVariantError> {
            Ok(slf.as_ptr())
        }}
    }
}
variant_impl!{
    impl VariantExt for i8 {
        VARTYPE = VT_I1;
        n3, cVal, cVal_mut
        from => {|n_ptr: &i8|Ok(*n_ptr)}
        into => {|slf: i8| -> Result<_, IntoVariantError> {Ok(slf)}}
    }
}
variant_impl!{
    impl VariantExt for u16 {
        VARTYPE = VT_UI2;
        n3, uiVal, uiVal_mut
        from => {|n_ptr: &u16|Ok(*n_ptr)}
        into => {|slf: u16| -> Result<_, IntoVariantError> {Ok(slf)}}
    }
}
variant_impl!{
    impl VariantExt for u32 {
        VARTYPE = VT_UI4;
        n3, ulVal, ulVal_mut
        from => {|n_ptr: &u32|Ok(*n_ptr)}
        into => {|slf: u32| -> Result<_, IntoVariantError> {Ok(slf)}}
    }
}
variant_impl!{
    impl VariantExt for u64 {
        VARTYPE = VT_UI8;
        n3, ullVal, ullVal_mut
        from => {|n_ptr: &u64|Ok(*n_ptr)}
        into => {|slf: u64| -> Result<_, IntoVariantError> {Ok(slf)}}
    }
}
variant_impl!{
    impl VariantExt for Int {
        VARTYPE = VT_INT;
        n3, intVal, intVal_mut
        from => {|n_ptr: &i32| Ok(Int::from(*n_ptr))}
        into => {|slf: Int| -> Result<_, IntoVariantError> {Ok(i32::from(slf))}}
    }
}
variant_impl!{
    impl VariantExt for UInt {
        VARTYPE = VT_UINT;
        n3, uintVal, uintVal_mut
        from => {|n_ptr: &u32| Ok(UInt::from(*n_ptr))}
        into => {|slf: UInt| -> Result<_, IntoVariantError> { Ok(u32::from(slf))}}
    }
}
variant_impl!{
    impl VariantExt for Box<DecWrapper> {
        VARTYPE = VT_PDECIMAL;
        n3, pdecVal, pdecVal_mut
        from => {|n_ptr: &*mut DECIMAL|Ok(Box::new(DecWrapper::from(**n_ptr)))}
        into => {|slf: Box<DecWrapper>| -> Result<_, IntoVariantError> {
            Ok(Box::into_raw( Box::new(DECIMAL::from(*slf))))
        }}
    }
}
variant_impl!{
    impl VariantExt for Box<Decimal> {
        VARTYPE = VT_PDECIMAL;
        n3, pdecVal, pdecVal_mut
        from => {|n_ptr: &*mut DECIMAL|Ok(Box::new(Decimal::from(DecWrapper::from(**n_ptr))))}
        into => {|slf: Box<Decimal>| -> Result<_, IntoVariantError> {
            let bptr = Box::new(DECIMAL::from(DecWrapper::from(*slf)));
            Ok(Box::into_raw(bptr))
        }}
    }
}
variant_impl!{
    impl VariantExt for Box<i8> {
        VARTYPE = VT_PI1;
        n3, pcVal, pcVal_mut
        from => {|n_ptr: &*mut i8|Ok(Box::new(**n_ptr))}
        into => {|slf: Box<i8>| -> Result<_, IntoVariantError> {
            Ok(Box::into_raw(slf))
        }}
    }
}
variant_impl!{
    impl VariantExt for Box<u16> {
        VARTYPE = VT_PUI2;
        n3, puiVal, puiVal_mut
        from => {|n_ptr: &*mut u16|Ok(Box::new(**n_ptr))}
        into => {|slf: Box<u16>| -> Result<_, IntoVariantError> {
            Ok(Box::into_raw(slf))
        }}
    }
}
variant_impl!{
    impl VariantExt for Box<u32> {
        VARTYPE = VT_PUI4;
        n3, pulVal, pulVal_mut
        from => {|n_ptr: &*mut u32|Ok(Box::new(**n_ptr))}
        into => {|slf: Box<u32>| -> Result<_, IntoVariantError> {
            Ok(Box::into_raw(slf))
        }}
    }
}
variant_impl!{
    impl VariantExt for Box<u64> {
        VARTYPE = VT_PUI8;
        n3, pullVal, pullVal_mut
        from => {|n_ptr: &*mut u64|Ok(Box::new(**n_ptr))}
        into => {|slf: Box<u64>| -> Result<_, IntoVariantError> {
            Ok(Box::into_raw(slf))
        }}
    }
}
variant_impl!{
    impl VariantExt for Box<Int> {
        VARTYPE = VT_PINT;
        n3, pintVal, pintVal_mut
        from => {|n_ptr: &*mut i32| Ok(Box::new(Int::from(**n_ptr)))}
        into => {|slf: Box<Int>|-> Result<_, IntoVariantError> { 
            Ok(Box::into_raw(Box::new(i32::from(*slf))))
        }}
    }
}
variant_impl!{
    impl VariantExt for Box<UInt> {
        VARTYPE = VT_PUINT;
        n3, puintVal, puintVal_mut
        from => {|n_ptr: &*mut u32| Ok(Box::new(UInt::from(**n_ptr)))}
        into => {|slf: Box<UInt>| -> Result<_, IntoVariantError> {
            Ok(Box::into_raw(Box::new(u32::from(*slf))))
        }}
    }
}
variant_impl!{
    impl VariantExt for DecWrapper {
        VARTYPE = VT_DECIMAL;
        n1, decVal, decVal_mut
        from => {|n_ptr: &DECIMAL|Ok(DecWrapper::from(*n_ptr))}
        into => {|slf: DecWrapper| -> Result<_, IntoVariantError> {
            Ok(DECIMAL::from(slf))
        }}
    }
}
variant_impl!{
    impl VariantExt for Decimal {
        VARTYPE = VT_DECIMAL;
        n1, decVal, decVal_mut
        from => {|n_ptr: &DECIMAL| Ok(Decimal::from(DecWrapper::from(*n_ptr)))}
        into => {|slf: Decimal| -> Result<_, IntoVariantError> {
            Ok(DECIMAL::from(DecWrapper::from(slf)))
        }}
    }
}

/// Helper type for VT_EMPTY variants
#[derive(Clone, Copy, Debug)]
pub struct VtEmpty{}

/// Helper type for VT_NULL variants
#[derive(Clone, Copy, Debug)]
pub struct VtNull{}

impl VariantExt for VtEmpty {
    const VARTYPE: u32 = VT_EMPTY;
    fn into_variant(self) -> Result<Ptr<VARIANT>, IntoVariantError> {
        let n3: VARIANT_n3 = unsafe {mem::zeroed()};
        let mut n1: VARIANT_n1 = unsafe {mem::zeroed()};

        let tv = __tagVARIANT { vt: <Self as VariantExt>::VARTYPE as u16, 
                        wReserved1: 0, 
                        wReserved2: 0, 
                        wReserved3: 0, 
                        n3: n3};
        unsafe {
            let n_ptr = n1.n2_mut();
            *n_ptr = tv;
        };
        let var = Box::new(VARIANT{ n1: n1 });
        Ok(Ptr::with_checked(Box::into_raw(var)).unwrap())
    }
    fn from_variant(var: Ptr<VARIANT>) -> Result<Self, FromVariantError> {
        let _var_d = VariantDestructor::new(var.as_ptr());
        Ok(VtEmpty{})
    }
}

impl VariantExt for VtNull {
    const VARTYPE: u32 = VT_NULL;
    fn into_variant(self) -> Result<Ptr<VARIANT>, IntoVariantError> {
        let n3: VARIANT_n3 = unsafe {mem::zeroed()};
        let mut n1: VARIANT_n1 = unsafe {mem::zeroed()};

        let tv = __tagVARIANT { vt: <Self as VariantExt>::VARTYPE as u16, 
                        wReserved1: 0, 
                        wReserved2: 0, 
                        wReserved3: 0, 
                        n3: n3};
        unsafe {
            let n_ptr = n1.n2_mut();
            *n_ptr = tv;
        };
        let var = Box::new(VARIANT{ n1: n1 });
        Ok(Ptr::with_checked(Box::into_raw(var)).unwrap())
    }
    fn from_variant(var: Ptr<VARIANT>) -> Result<Self, FromVariantError> {
        let _var_d = VariantDestructor::new(var.as_ptr());
        Ok(VtNull{})
    }
}

#[cfg(test)]
mod test {
    use super::*;
    macro_rules! validate_variant {
        ($t:ident, $val:expr, $vt:expr) => {
            let v = $val;
            let var = match v.clone().into_variant() {
                Ok(var) => var, 
                Err(_) => panic!("Error")
            };
            assert!(!var.as_ptr().is_null());
            unsafe {
                let pvar = var.as_ptr();
                let n1 = (*pvar).n1;
                let tv: &__tagVARIANT = n1.n2();
                assert_eq!(tv.vt as u32, $vt);
            };
            let var = $t::from_variant(var);
            assert_eq!(v, var.unwrap());
            };
    }
    #[test]
    fn test_i64() {
        validate_variant!(i64, 1337i64, VT_I8);
    }
    #[test]
    fn test_i32() {
        validate_variant!(i32, 1337i32, VT_I4);
    }
    #[test]
    fn test_u8() {
        validate_variant!(u8, 137u8, VT_UI1);
    }
        #[test]
    fn test_i16() {
        validate_variant!(i16, 1337i16, VT_I2);
    }
    #[test]
    fn test_f32() {
        validate_variant!(f32, 1337.9f32, VT_R4);
    }
    #[test]
    fn test_f64() {
        validate_variant!(f64, 1337.9f64, VT_R8);
    }
    #[test]
    fn test_bool_t() {
        validate_variant!(bool, true, VT_BOOL);
    }

    #[test]
    fn test_bool_f() {
        validate_variant!(bool, false, VT_BOOL);
    }

    #[test]
    fn test_scode() {
        validate_variant!(SCode, SCode::from(137), VT_ERROR);
    }

    #[test]
    fn test_cy() {
        validate_variant!(Currency, Currency::from(137), VT_CY);
    }

    #[test]
    fn test_date() {
        validate_variant!(Date, Date::from(137.7), VT_DATE);
    }

    #[test]
    fn test_str() {
        validate_variant!(String, String::from("testing abc1267 ?Ťũřǐꝥꞔ"), VT_BSTR);
    }

    #[test]
    fn test_box_u8() {
        type Bu8 = Box<u8>;
        validate_variant!(Bu8, Box::new(139), VT_PUI1);
    }

    #[test]
    fn test_box_i16() {
        type Bi16 = Box<i16>;
        validate_variant!(Bi16, Box::new(139), VT_PI2);
    }
    #[test]
    fn test_box_i32() {
        type Bi32 = Box<i32>;
        validate_variant!(Bi32, Box::new(139), VT_PI4);
    }
    #[test]
    fn test_box_i64() {
        type Bi64 = Box<i64>;
        validate_variant!(Bi64, Box::new(139), VT_PI8);
    }
    #[test]
    fn test_box_f32() {
        type Bf32 = Box<f32>;
        validate_variant!(Bf32, Box::new(1337.9f32), VT_PR4);
    }
    #[test]
    fn test_box_f64() {
        type Bf64 = Box<f64>;
        validate_variant!(Bf64, Box::new(1337.9f64), VT_PR8);
    }
    #[test]
    fn test_box_bool() {
        type Bbool = Box<bool>;
        validate_variant!(Bbool, Box::new(true), VT_PBOOL);
    }
    #[test]
    fn test_box_scode() {
        type BSCode = Box<SCode>;
        validate_variant!(BSCode, Box::new(SCode::from(-50)), VT_PERROR);
    }
    #[test]
    fn test_box_cy() {
        type BCy = Box<Currency>;
        validate_variant!(BCy, Box::new(Currency::from(137)), VT_PCY);
    }
    #[test]
    fn test_box_date() {
        type BDate = Box<Date>;
        validate_variant!(BDate, Box::new(Date::from(-10.333f64)), VT_PDATE);
    }
    #[test]
    fn test_box_str() {
        type BStr = Box<String>;
        validate_variant!(BStr, Box::new(String::from("testing abc1267 ?Ťũřǐꝥꞔ")), VT_PBSTR);
    }
    #[test]
    fn test_variant() {
        let v = Variant::new(1000u64);
        let var = match v.into_variant() {
            Ok(var) => var, 
            Err(_) => panic!("Error")
        };
        assert!(!var.as_ptr().is_null());
        unsafe {
            let pvar = var.as_ptr();
            let n1 = (*pvar).n1;
            let tv: &__tagVARIANT = n1.n2();
            assert_eq!(tv.vt as u32, VT_VARIANT);
        };
        let var = Variant::<u64>::from_variant(var);
        assert_eq!(v, var.unwrap());
    }
    //test SafeArray<T>
    //Ptr<c_void>
    #[test]
    fn test_i8() {
        validate_variant!(i8, -119i8, VT_I1);
    }
    #[test]
    fn test_u16() {
        validate_variant!(u16, 119u16, VT_UI2);
    }
    #[test]
    fn test_u32() {
        validate_variant!(u32, 11976u32, VT_UI4);
    }
    #[test]
    fn test_u64() {
        validate_variant!(u64, 11976u64, VT_UI8);
    }
    #[test]
    fn test_box_i8() {
        type Bi8 = Box<i8>;
        validate_variant!(Bi8, Box::new(-119i8), VT_PI1);
    }
    #[test]
    fn test_box_u16() {
        type Bu16 = Box<u16>;
        validate_variant!(Bu16, Box::new(119u16), VT_PUI2);
    }
    #[test]
    fn test_box_u32() {
        type Bu32 = Box<u32>;
        validate_variant!(Bu32, Box::new(11976u32), VT_PUI4);
    }
    #[test]
    fn test_box_u64() {
        type Bu64 = Box<u64>;
        validate_variant!(Bu64, Box::new(11976u64), VT_PUI8);
    }
    #[test]
    fn test_send() {
        fn assert_send<T: Send>() {}
        assert_send::<Variant<i64>>();
    }
    #[test]
    fn test_sync() {
        fn assert_sync<T: Sync>() {}
        assert_sync::<Variant<i64>>();
    }
}