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
/// } VARIANT;
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
use std::mem;
use std::ptr::NonNull;

use rust_decimal::Decimal;

use winapi::ctypes::c_void;
use winapi::shared::wtypes::{
    CY, DATE, DECIMAL,
    VARIANT_BOOL,
    VT_ARRAY, 
    //VT_BSTR, 
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
use winapi::shared::wtypesbase::{SCODE};
use winapi::um::oaidl::{IDispatch,  __tagVARIANT, SAFEARRAY, VARIANT, VARIANT_n3, VARIANT_n1};
use winapi::um::unknwnbase::IUnknown;

use array::{SafeArray, SafeArrayElement};
use ptr::Ptr;
use types::{Date, DecWrapper, Currency, Int, SCode, UInt, VariantBool };

pub const VT_PUI1: u32 = VT_BYREF | VT_UI1;
pub const VT_PI2: u32 = VT_BYREF | VT_I2;
pub const VT_PI4: u32 = VT_BYREF | VT_I4;
pub const VT_PI8: u32 = VT_BYREF | VT_I8;
pub const VT_PUI8: u32 = VT_BYREF | VT_UI8;
pub const VT_PR4: u32 = VT_BYREF | VT_R4;
pub const VT_PR8: u32 = VT_BYREF | VT_R8;
pub const VT_PBOOL: u32 = VT_BYREF | VT_BOOL;
pub const VT_PERROR: u32 = VT_BYREF | VT_ERROR;
pub const VT_PCY: u32 = VT_BYREF | VT_CY;
pub const VT_PDATE: u32 = VT_BYREF | VT_DATE;
//pub const VT_PBSTR: u32 = VT_BYREF | VT_BSTR;
pub const VT_PUNKNOWN: u32 = VT_BYREF | VT_UNKNOWN;
pub const VT_PDISPATCH: u32 = VT_BYREF | VT_DISPATCH;
pub const VT_PDECIMAL: u32 = VT_BYREF | VT_DECIMAL;
pub const VT_PI1: u32 = VT_BYREF | VT_I1;
pub const VT_PUI2: u32 = VT_BYREF | VT_UI2;
pub const VT_PUI4: u32 = VT_BYREF | VT_UI4;
pub const VT_PINT: u32 = VT_BYREF | VT_INT;
pub const VT_PUINT: u32 = VT_BYREF | VT_UINT;

pub trait VariantType: Sized { //Would like Clone + Default, but IDispatch and IUnknown don't implement them
    const VARTYPE: u32;

    fn from_variant(var: Ptr<VARIANT>) -> Result<Self, ()>;  
    fn into_variant(&mut self) -> Result<Ptr<VARIANT>, ()>;
}

#[derive(Clone, Copy, Debug, Eq, Hash, Ord, PartialEq, PartialOrd)]
pub struct Variant<T: VariantType>(T);

#[allow(dead_code)]
impl<T: VariantType> Variant<T> {
    pub fn new(t: T) -> Variant<T> {
        Variant(t)
    }

    pub fn unwrap(self) -> T {
        self.0
    }

    pub fn borrow(&self) -> &T {
        &self.0
    }

    pub fn borrow_mut(&mut self) -> &mut T {
        &mut self.0
    }

    pub fn to_variant(&mut self) -> Result<Ptr<VARIANT>, ()> {
        self.0.into_variant()        
    }

    pub fn from_variant(var: Ptr<VARIANT>) -> Result<Variant<T>, ()> {
        let t = T::from_variant(var).unwrap();
        Ok(Variant(t))
    }
}

macro_rules! variant_impl {
    (
        impl $(<$tn:ident : $tc:ident>)* Variant for $t:ty {
            VARTYPE = $vt:expr;
            n3, $un_n:ident, $un_n_mut:ident
            from => {$from:expr}
            into => {$into:expr}
        }
    ) => {
        impl $(<$tn: $tc>)* VariantType for $t {
            const VARTYPE: u32 = $vt;
            fn from_variant(var: Ptr<VARIANT>) -> Result<Self, ()>{
                let var = unsafe {*var.as_ptr()};
                let mut n1 = var.n1;

                let n3 = unsafe {
                    n1.n2_mut().n3
                };
                unsafe {
                    let n_ptr = n3.$un_n();
                    $from(n_ptr)
                }
            }

            fn into_variant(&mut self) -> Result<Ptr<VARIANT>, ()> {
                let mut n3: VARIANT_n3 = unsafe {mem::zeroed()};
                let mut n1: VARIANT_n1 = unsafe {mem::zeroed()};
                unsafe {
                    let n_ptr = n3.$un_n_mut();
                    *n_ptr = $into(self)
                }
                let tv = __tagVARIANT { vt: <Self as VariantType>::VARTYPE as u16, 
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

    (
        impl $(<$tn:ident : $tc:ident>)* Variant for $t:ty {
            VARTYPE = $vt:expr;
            n1, $un_n:ident, $un_n_mut:ident
            from => {$from:expr}
            into => {$into:expr}
        }
    ) => {
        impl $(<$tn: $tc>)* VariantType for $t {
            const VARTYPE: u32 = $vt;
            fn from_variant(var: Ptr<VARIANT>) -> Result<Self, ()>{
                let var = unsafe {*var.as_ptr()};
                let n1 = var.n1;
                unsafe {
                    let n_ptr = n1.$un_n();
                    $from(n_ptr)
                }
            }

           fn into_variant(&mut self) -> Result<Ptr<VARIANT>, ()> {
                let n3: VARIANT_n3 = unsafe {mem::zeroed()};
                let mut n1: VARIANT_n1 = unsafe {mem::zeroed()};
                unsafe {
                    let n_ptr = n1.$un_n_mut();
                    *n_ptr = $into(self)
                }
                let tv = __tagVARIANT { vt: <Self as VariantType>::VARTYPE as u16, 
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
}

variant_impl!{
    impl Variant for i64 {
        VARTYPE = VT_I8;
        n3, llVal, llVal_mut
        from => {|n_ptr: &i64| {Ok(*n_ptr)}}
        into => {|slf: &mut i64| *slf}
    }
}
variant_impl!{
    impl Variant for i32 {
        VARTYPE = VT_I4;
        n3, lVal, lVal_mut
        from => {|n_ptr: &i32| Ok(*n_ptr)}
        into => {|slf: &mut i32| *slf}
    }
}
variant_impl!{
    impl Variant for u8 {
        VARTYPE = VT_UI1;
        n3, bVal, bVal_mut
        from => {|n_ptr: &u8| Ok(*n_ptr)}
        into => {|slf: &mut u8| *slf}
    }
}
variant_impl!{
    impl Variant for i16 {
        VARTYPE = VT_I2;
        n3, iVal, iVal_mut
        from => {|n_ptr: &i16| Ok(*n_ptr)}
        into => {|slf: &mut i16| *slf}
    }
}
variant_impl!{
    impl Variant for f32 {
        VARTYPE = VT_R4;
        n3, fltVal, fltVal_mut
        from => {|n_ptr: &f32| Ok(*n_ptr)}
        into => {|slf: &mut f32| *slf}
    }
}
variant_impl!{
    impl Variant for f64 {
        VARTYPE = VT_R8;
        n3, dblVal, dblVal_mut
        from => {|n_ptr: &f64| Ok(*n_ptr)}
        into => {|slf: &mut f64| *slf}
    }
}
variant_impl!{
    impl Variant for bool {
        VARTYPE = VT_BOOL;
        n3, boolVal, boolVal_mut
        from => {|n_ptr: &VARIANT_BOOL| Ok(bool::from(VariantBool::from(*n_ptr)))}
        into => {|slf: &mut bool| VARIANT_BOOL::from(*slf)}
    }
}
variant_impl!{
    impl Variant for SCode {
        VARTYPE = VT_ERROR;
        n3, scode, scode_mut
        from => {|n_ptr: &SCODE| Ok(SCode(*n_ptr))}
        into => {|slf: &mut SCode| slf.0}
    }
}
variant_impl!{
    impl Variant for Currency {
        VARTYPE = VT_CY;
        n3, cyVal, cyVal_mut
        from => {|n_ptr: &CY| Ok(Currency::from(*n_ptr))}
        into => {|slf: &mut Currency| CY::from(*slf)}
    }
}
variant_impl!{
    impl Variant for Date {
        VARTYPE = VT_DATE;
        n3, date, date_mut
        from => {|n_ptr: &DATE| Ok(Date::from(*n_ptr))}
        into => {|slf: &mut Date| DATE::from(*slf)}
    }
}
//BSTR
variant_impl!{
    impl Variant for Ptr<IUnknown> {
        VARTYPE = VT_UNKNOWN;
        n3, punkVal, punkVal_mut
        from => {|n_ptr: &* mut IUnknown| Ok(Ptr::with_checked(*n_ptr).unwrap())}
        into => {|slf: &mut Ptr<IUnknown>| (*slf).as_ptr()}
    }
}
variant_impl!{
    impl Variant for Ptr<IDispatch> {
        VARTYPE = VT_DISPATCH;
        n3, pdispVal, pdispVal_mut
        from => {|n_ptr: &* mut IDispatch| Ok(Ptr::with_checked(*n_ptr).unwrap())}
        into => {|slf: &mut Ptr<IDispatch>| (*slf).as_ptr()}
    }
}
variant_impl!{
    impl Variant for Box<u8> {
        VARTYPE = VT_PUI1;
        n3, pbVal, pbVal_mut
        from => {|n_ptr: &* mut u8| Ok(Box::new(**n_ptr))}
        into => {|slf: &mut Box<u8>| Box::into_raw((*slf).clone())}
    }
}
variant_impl!{
    impl Variant for Box<i16> {
        VARTYPE = VT_PI2;
        n3, piVal, piVal_mut
        from => {|n_ptr: &* mut i16| Ok(Box::new(**n_ptr))}
        into => {|slf: &mut Box<i16>| Box::into_raw((*slf).clone())}
    }
}
variant_impl!{
    impl Variant for Box<i32> {
        VARTYPE = VT_PI4;
        n3, plVal, plVal_mut
        from => {|n_ptr: &* mut i32| Ok(Box::new(**n_ptr))}
        into => {|slf: &mut Box<i32>| Box::into_raw((*slf).clone())}
    }
}
variant_impl!{
    impl Variant for Box<i64> {
        VARTYPE = VT_PI8;
        n3, pllVal, pllVal_mut
        from => {|n_ptr: &* mut i64| Ok(Box::new(**n_ptr))}
        into => {|slf: &mut Box<i64>| Box::into_raw((*slf).clone())}
    }
}
variant_impl!{
    impl Variant for Box<f32> {
        VARTYPE = VT_PR4;
        n3, pfltVal, pfltVal_mut
        from => {|n_ptr: &* mut f32| Ok(Box::new(**n_ptr))}
        into => {|slf: &mut Box<f32>| Box::into_raw((*slf).clone())}
    }
}
variant_impl!{
    impl Variant for Box<f64> {
        VARTYPE = VT_PR8;
        n3, pdblVal, pdblVal_mut
        from => {|n_ptr: &* mut f64| Ok(Box::new(**n_ptr))}
        into => {|slf: &mut Box<f64>| Box::into_raw((*slf).clone())}
    }
}
variant_impl!{
    impl Variant for Box<bool> {
        VARTYPE = VT_PBOOL;
        n3, pboolVal, pboolVal_mut
        from => {
            |n_ptr: &*mut VARIANT_BOOL| Ok(Box::new(bool::from(VariantBool::from(**n_ptr))))
        }
        into => {
            |slf: &mut Box<bool>| {
                let bptr = Box::new(VARIANT_BOOL::from(**slf));
                Box::into_raw(bptr)
            }
        }
    }
}
variant_impl!{
    impl Variant for Box<SCode> {
        VARTYPE = VT_PERROR;
        n3, pscode, pscode_mut
        from => {|n_ptr: &*mut SCODE| Ok(Box::new(SCode(**n_ptr)))}
        into => {|slf: &mut Box<SCode>| Box::into_raw(Box::new((*slf).0))}
    }
}
variant_impl!{
    impl Variant for Box<Currency> {
        VARTYPE = VT_PCY;
        n3, pcyVal, pcyVal_mut
        from => { |n_ptr: &*mut CY| Ok(Box::new(Currency::from(**n_ptr))) }
        into => {
            |slf: &mut Box<Currency>| {
                let bptr = Box::new(CY::from(**slf));
                Box::into_raw(bptr)
            }
        }
    }
}
variant_impl!{
    impl Variant for Box<Date> {
        VARTYPE = VT_PDATE;
        n3, pdate, pdate_mut
        from => { |n_ptr: &*mut f64| Ok(Box::new(Date(**n_ptr))) }
        into => {
            |slf: &mut Box<Date>| {
                let bptr = Box::new(DATE::from(**slf));
                Box::into_raw(bptr)
            }
        }
    }
}
variant_impl! {
    impl Variant for Box<Ptr<IUnknown>> {
        VARTYPE = VT_PUNKNOWN;
        n3, ppunkVal, ppunkVal_mut
        from => {
            |n_ptr: &*mut *mut IUnknown| {
                match NonNull::new((**n_ptr).clone()) {
                    Some(nn) => Ok(Box::new(Ptr::new(nn))), 
                    None => Err(())
                }
            }
        }
        into => {
            |slf: &mut Box<Ptr<IUnknown>>| {
                let bptr = Box::new((**slf).as_ptr());
                Box::into_raw(bptr)
            }
        }
    }
}
variant_impl! {
    impl Variant for Box<Ptr<IDispatch>> {
        VARTYPE = VT_PDISPATCH;
        n3, ppdispVal, ppdispVal_mut
        from => {
            |n_ptr: &*mut *mut IDispatch| {
                match NonNull::new((**n_ptr).clone()) {
                    Some(nn) => Ok(Box::new(Ptr::new(nn))), 
                    None => Err(())
                }
            }
        }
        into => {
            |slf: &mut Box<Ptr<IDispatch>>| {
                let bptr = Box::new((**slf).as_ptr());
                Box::into_raw(bptr)
            }
        }
    }
}
variant_impl!{
    impl<T: VariantType> Variant for Variant<T> {
        VARTYPE = VT_VARIANT;
        n3, pvarVal, pvarVal_mut
        from => {|n_ptr: &*mut VARIANT| {
            let pnn = match NonNull::new(*n_ptr) {
                Some(nn) => Ptr::new(nn), 
                None => return Err(()) 
            };
            Variant::<T>::from_variant(pnn)
        }}
        into => {|slf: &mut Variant<T>| {
            let pvar = slf.borrow_mut().into_variant().unwrap();
            pvar.as_ptr()
        }}
    }
}
variant_impl!{
    impl<T: SafeArrayElement> Variant for SafeArray<T>{
        VARTYPE = VT_ARRAY;
        n3, parray, parray_mut
        from => {
            |n_ptr: &*mut SAFEARRAY| {
                match SafeArray::<T>::from_safearray(*n_ptr) {
                    Ok(sa) => Ok(sa), 
                    Err(_) => Err(())
                }
            }
        }
        into => {
            |slf: &mut SafeArray<T>| {
                match slf.into_safearray() {
                    Ok(psa) => {
                        psa.as_ptr()
                    }, 
                    Err(_) => {
                        panic!("Could not convert SafeArray<{}> into Ptr<SAFEARRAY>.", stringify!(T));
                    }
                }
            }
        }
    }
}
variant_impl!{
    impl Variant for Ptr<c_void> {
        VARTYPE = VT_BYREF;
        n3, byref, byref_mut
        from => {|n_ptr: &*mut c_void| {
            match NonNull::new(*n_ptr) {
                Some(nn) => Ok(Ptr::new(nn)), 
                None => Err(())
            }
        }}
        into => {|slf: &mut Ptr<c_void>| slf.as_ptr()}
    }
}
variant_impl!{
    impl Variant for i8 {
        VARTYPE = VT_I1;
        n3, cVal, cVal_mut
        from => {|n_ptr: &i8|Ok(*n_ptr)}
        into => {|slf: &mut i8| *slf}
    }
}
variant_impl!{
    impl Variant for u16 {
        VARTYPE = VT_UI2;
        n3, uiVal, uiVal_mut
        from => {|n_ptr: &u16|Ok(*n_ptr)}
        into => {|slf: &mut u16| *slf}
    }
}
variant_impl!{
    impl Variant for u32 {
        VARTYPE = VT_UI4;
        n3, ulVal, ulVal_mut
        from => {|n_ptr: &u32|Ok(*n_ptr)}
        into => {|slf: &mut u32| *slf}
    }
}
variant_impl!{
    impl Variant for u64 {
        VARTYPE = VT_UI8;
        n3, ullVal, ullVal_mut
        from => {|n_ptr: &u64|Ok(*n_ptr)}
        into => {|slf: &mut u64| *slf}
    }
}
variant_impl!{
    impl Variant for Int {
        VARTYPE = VT_INT;
        n3, intVal, intVal_mut
        from => {|n_ptr: &i32| Ok(Int(*n_ptr))}
        into => {|slf: &mut Int| slf.0}
    }
}
variant_impl!{
    impl Variant for UInt {
        VARTYPE = VT_UINT;
        n3, uintVal, uintVal_mut
        from => {|n_ptr: &u32| Ok(UInt(*n_ptr))}
        into => {|slf: &mut UInt| slf.0}
    }
}
variant_impl!{
    impl Variant for Box<DecWrapper> {
        VARTYPE = VT_PDECIMAL;
        n3, pdecVal, pdecVal_mut
        from => {|n_ptr: &*mut DECIMAL|Ok(Box::new(DecWrapper::from(**n_ptr)))}
        into => {|slf: &mut Box<DecWrapper>| {
            let bptr = Box::new(DECIMAL::from(**slf));
            Box::into_raw(bptr)
        }}
    }
}
variant_impl!{
    impl Variant for Box<Decimal> {
        VARTYPE = VT_PDECIMAL;
        n3, pdecVal, pdecVal_mut
        from => {|n_ptr: &*mut DECIMAL|Ok(Box::new(Decimal::from(DecWrapper::from(**n_ptr))))}
        into => {|slf: &mut Box<Decimal>| {
            let bptr = Box::new(DECIMAL::from(DecWrapper::from(**slf)));
            Box::into_raw(bptr)
        }}
    }
}
variant_impl!{
    impl Variant for Box<i8> {
        VARTYPE = VT_PI1;
        n3, pcVal, pcVal_mut
        from => {|n_ptr: &*mut i8|Ok(Box::new(**n_ptr))}
        into => {|slf: &mut Box<i8>| {
            Box::into_raw((*slf).clone())
        }}
    }
}
variant_impl!{
    impl Variant for Box<u16> {
        VARTYPE = VT_PUI2;
        n3, puiVal, puiVal_mut
        from => {|n_ptr: &*mut u16|Ok(Box::new(**n_ptr))}
        into => {|slf: &mut Box<u16>| {
            Box::into_raw((*slf).clone())
        }}
    }
}
variant_impl!{
    impl Variant for Box<u32> {
        VARTYPE = VT_PUI4;
        n3, pulVal, pulVal_mut
        from => {|n_ptr: &*mut u32|Ok(Box::new(**n_ptr))}
        into => {|slf: &mut Box<u32>| {
            Box::into_raw((*slf).clone())
        }}
    }
}
variant_impl!{
    impl Variant for Box<u64> {
        VARTYPE = VT_PUI8;
        n3, pullVal, pullVal_mut
        from => {|n_ptr: &*mut u64|Ok(Box::new(**n_ptr))}
        into => {|slf: &mut Box<u64>| {
            Box::into_raw((*slf).clone())
        }}
    }
}
variant_impl!{
    impl Variant for Box<Int> {
        VARTYPE = VT_PINT;
        n3, pintVal, pintVal_mut
        from => {|n_ptr: &*mut i32| Ok(Box::new(Int(**n_ptr)))}
        into => {|slf: &mut Box<Int>| Box::into_raw(Box::new((**slf).0))}
    }
}
variant_impl!{
    impl Variant for Box<UInt> {
        VARTYPE = VT_PUINT;
        n3, puintVal, puintVal_mut
        from => {|n_ptr: &*mut u32| Ok(Box::new(UInt(**n_ptr)))}
        into => {|slf: &mut Box<UInt>| Box::into_raw(Box::new((**slf).0))}
    }
}
variant_impl!{
    impl Variant for DecWrapper {
        VARTYPE = VT_DECIMAL;
        n1, decVal, decVal_mut
        from => {|n_ptr: &DECIMAL|Ok(DecWrapper::from(*n_ptr))}
        into => {|slf: &mut DecWrapper| {
            DECIMAL::from(*slf)
        }}
    }
}
variant_impl!{
    impl Variant for Decimal {
        VARTYPE = VT_DECIMAL;
        n1, decVal, decVal_mut
        from => {|n_ptr: &DECIMAL| Ok(Decimal::from(DecWrapper::from(*n_ptr)))}
        into => {|slf: &mut Decimal| DECIMAL::from(DecWrapper::from(*slf))}
    }
}

pub struct VtEmpty{}
pub struct VtNull{}

impl VariantType for VtEmpty {
    const VARTYPE: u32 = VT_EMPTY;
    fn into_variant(&mut self) -> Result<Ptr<VARIANT>, ()> {
        //TODO: fix this
        Err(())
    }
    fn from_variant(_var: Ptr<VARIANT>) -> Result<Self, ()> {
        Ok(VtEmpty{})
    }
}

impl VariantType for VtNull {
    const VARTYPE: u32 = VT_NULL;
    fn into_variant(&mut self) -> Result<Ptr<VARIANT>, ()> {
        Err(())
    }
    fn from_variant(_var: Ptr<VARIANT>) -> Result<Self, ()> {
        Ok(VtNull{})
    }
}

#[cfg(test)]
mod test {
    use super::*;
    
    #[test]
    fn test_variant() {
        let mut value = 100u8;
        let r = value.into_variant().unwrap();
        let new_v = u8::from_variant(r).unwrap();
        println!("{:?}", new_v);
        assert_eq!(new_v, 100u8 );
    }
}