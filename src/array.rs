use std::marker::PhantomData;
use std::mem;
use std::ptr::null_mut;

use rust_decimal::Decimal;

// use widestring::U16String;

#[cfg(windows)]
use winapi::ctypes::{c_long, c_void};

#[cfg(windows)]
use winapi::shared::minwindef::{UINT, ULONG,};

#[cfg(windows)]
use winapi::shared::ntdef::HRESULT;

#[cfg(windows)]
use winapi::shared::wtypes::{
    // BSTR,
    CY, 
    DATE, 
    DECIMAL,  
    VARTYPE,
    VARIANT_BOOL,
    // VT_BSTR, 
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
    //VT_RECORD,
    VT_UI1,
    VT_UI2,
    VT_UI4,
    VT_UINT,
    VT_UNKNOWN, 
    VT_VARIANT,   
};

use winapi::um::oaidl::{IDispatch, LPSAFEARRAY, LPSAFEARRAYBOUND, SAFEARRAY, SAFEARRAYBOUND, VARIANT};
use winapi::um::unknwnbase::IUnknown;

// use bstr::BStringExt;
use super::errors::{
    FromSafeArrayError, 
    FromSafeArrElemError, 
    IntoSafeArrayError, 
    IntoSafeArrElemError,
};
use super::ptr::Ptr;
use super::types::{Currency, Date, DecWrapper, Int, SCode, UInt, VariantBool};
use super::variant::{Variant, VariantExt};

/// Helper trait implemented for types that can be converted into a safe array. 
/// Generally, don't implement this yourself without care.
/// Implemented for types:
///     * i8, u8, i16, u16, i32, u32
///     * bool, f32, f64
///     * String, Variant<T>, 
///     * Ptr<IUnknown>, Ptr<IDispatch>
/// 
pub trait SafeArrayElement: Sized {
    /// This is the VT value used to create the SAFEARRAY
    const SFTYPE: u32;

    /// puts a type into the safearray at the specified index (default impls use SafeArrayPutElement)
    fn into_safearray(self, psa: *mut SAFEARRAY, ix: i32) -> Result<(), IntoSafeArrElemError>;
    /// gets a type from the safearray at the specified index (default impls use SafeArrayGetElement)
    fn from_safearray(psa: *mut SAFEARRAY, ix: i32) -> Result<Self, FromSafeArrElemError>;
}

/// Workhorse trait and main interface for converting to/from SAFEARRAY
/// Default impl is on `Vec<T: SafeArrayElement>` 
pub trait SafeArrayExt<T: SafeArrayElement> {
    /// Use `t.into_safearray()` to convert a type into a SAFEARRAY
    fn into_safearray(&mut self) -> Result<Ptr<SAFEARRAY>, IntoSafeArrayError>;
    
    /// Use `T::from_safearray(psa)` to convert a safearray pointer into the relevant T
    fn from_safearray(psa: *mut SAFEARRAY) -> Result<Vec<T>, FromSafeArrayError>;
}

macro_rules! check_and_throw {
    ($hr:ident, $success:expr, $fail:expr) => {
        match $hr {
            0 => $success, 
            _ => $fail
        }
    };
}

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

impl<I> SafeArrayExt<I::Item> for I 
where I: ExactSizeIterator + ?Sized, 
      I::Item: SafeArrayElement
{
    fn into_safearray(&mut self) -> Result<Ptr<SAFEARRAY>, IntoSafeArrayError > {
        let c_elements: ULONG = self.len() as u32;
        let vartype = I::Item::SFTYPE;
        let mut sab = SAFEARRAYBOUND { cElements: c_elements, lLbound: 0i32};
        let psa = unsafe { SafeArrayCreate(vartype as u16, 1, &mut sab)};
        assert!(!psa.is_null());
        let mut sad = SafeArrayDestructor::new(psa);

        for (ix, mut elem) in self.enumerate() {
            match elem.into_safearray(psa, ix as i32) {
                Ok(()) => continue, 
                Err(e) => return Err(IntoSafeArrayError::from_element_err(e, ix))
            }
        }
        sad.inner = null_mut();

        Ok(Ptr::with_checked(psa).unwrap())
    }

    fn from_safearray(psa: *mut SAFEARRAY) -> Result<Vec<I::Item>, FromSafeArrayError> {
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

        if vt as u32 != I::Item::SFTYPE {
            return Err(FromSafeArrayError::VarTypeDoesNotMatch{expected: I::Item::SFTYPE, found: vt as u32});
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

            let mut vc: Vec<I::Item> = Vec::new();
            for ix in l_bound..=r_bound {
                match I::Item::from_safearray(psa, ix) {
                    Ok(val) => vc.push(val), 
                    Err(e) => return Err(FromSafeArrayError::from_element_err(e, ix as usize))
                }
            }
            Ok(vc)
        } else {
            Err(FromSafeArrayError::SafeArrayDimsInvalid{sa_dims: sa_dims})
        }
    }
} 

// impl<'a, T: SafeArrayElement> SafeArrayExt<T> for IterMut<'a, T> {

// }

macro_rules! safe_arr_impl {
    (
        impl $(< $tn:ident : $tc:ident >)* SafeArrayElement for $t:ty {
            SFTYPE = $vt:expr;
            ptr
            def  => {$def:expr}
            from => {$from:expr}
            into => {$into:expr}
        }
    ) => {
        impl $(<$tn:$tc>)* SafeArrayElement for $t {
            const SFTYPE: u32 = $vt;
             fn from_safearray(psa: *mut SAFEARRAY, ix: i32) -> Result<Self, FromSafeArrElemError> {
                let val = $def;
                let hr = unsafe {SafeArrayGetElement(psa, &ix, val as *mut _ as *mut c_void)};
                check_and_throw!(hr, $from(val), {return Err(FromSafeArrElemError::GetElementFailed{hr: hr})})
            }
            
            fn into_safearray(self, psa: *mut SAFEARRAY, ix: i32) -> Result<(), IntoSafeArrElemError> {
                let slf = $into(self)?;
                let hr = unsafe {SafeArrayPutElement(psa, &ix, slf as *mut _ as *mut c_void)};
                check_and_throw!(hr, {return Ok(())}, {Err(IntoSafeArrElemError::PutElementFailed{hr: hr})})
            }
        }
    };
    (
        impl $(< $tn:ident : $tc:ident >)* SafeArrayElement for $t:ty {
            SFTYPE = $vt:expr;
            def  => {$def:expr}
            from => {$from:expr}
            into => {$into:expr}
        }
    ) => {
        impl $(<$tn:$tc>)* SafeArrayElement for $t {
            const SFTYPE: u32 = $vt;
             fn from_safearray(psa: *mut SAFEARRAY, ix: i32) -> Result<Self, FromSafeArrElemError> {
                let mut val = $def;
                let hr = unsafe {SafeArrayGetElement(psa, &ix, &mut val as *mut _ as *mut c_void)};
                check_and_throw!(hr, $from(val), {return Err(FromSafeArrElemError::GetElementFailed{hr: hr})})
            }
            
            fn into_safearray(self, psa: *mut SAFEARRAY, ix: i32) -> Result<(), IntoSafeArrElemError> {
                let mut slf = $into(self)?;
                let hr = unsafe {SafeArrayPutElement(psa, &ix, &mut slf as *mut _ as *mut c_void)};
                check_and_throw!(hr, {return Ok(())}, {Err(IntoSafeArrElemError::PutElementFailed{hr: hr})})
            }
        }
    };
}

safe_arr_impl!{impl SafeArrayElement for i16 {
    SFTYPE = VT_I2;
    def => { 0i16 }
    from => {|i| Ok(i)}
    into => { |slf: i16| -> Result<_, IntoSafeArrElemError> {Ok(slf)} }
}}
safe_arr_impl!{impl SafeArrayElement for i32 {
    SFTYPE = VT_I4;
    def => { 0i32 }
    from => {|i| Ok(i)}
    into => { |slf: _| -> Result<_, IntoSafeArrElemError> { Ok(slf) }}
}}
safe_arr_impl!{impl SafeArrayElement for f32 {
    SFTYPE = VT_R4;
    def => { 0.0f32 }
    from => {|i| Ok(i)}
    into => { |slf: _| -> Result<_, IntoSafeArrElemError> { Ok(slf) }}
}}
safe_arr_impl!{impl SafeArrayElement for f64 { 
    SFTYPE = VT_R8; 
    def => { 0.0f64 }
    from => {|i| Ok(i)}
    into => { |slf: _| -> Result<_, IntoSafeArrElemError> { Ok(slf) }}
}}
safe_arr_impl!{impl SafeArrayElement for Currency{
    SFTYPE = VT_CY; 
    def => { CY{int64: 0} }
    from => { |cy| Ok(Currency::from(cy)) }
    into => {|slf: Currency| -> Result<_, IntoSafeArrElemError> {Ok(CY::from(slf))}}
}}
safe_arr_impl!{impl SafeArrayElement for Date{
    SFTYPE = VT_DATE; 
    def =>  { 0f64 }
    from => { |dt| Ok(Date::from(dt)) } 
    into => { |slf: Date| -> Result<_, IntoSafeArrElemError> {Ok(DATE::from(slf)) }}
}}
// Need to wrap the string in a variant because its not working otherwise. 
safe_arr_impl!{impl SafeArrayElement for String {
    SFTYPE = VT_VARIANT;
    ptr
    def => {{
        let mut var: VARIANT = unsafe {mem::zeroed()};
        &mut var as *mut VARIANT
    }}
    from => {|pvar| {
        let pnn = match Ptr::with_checked(pvar) {
            Some(nn) => nn, 
            None => return Err(FromSafeArrElemError::VariantPtrNull)
        };
        match Variant::<String>::from_variant(pnn) {
            Ok(var) => Ok(var.unwrap()), 
            Err(_) => return Err(FromSafeArrElemError::FromVariantFailed)
        }
    }}
    into => {|slf: String|{
        let slf = Variant::new(slf);
        match slf.into_variant() {
            Ok(slf) => {
                let mut s = slf.as_ptr();
                Ok(s)
            }, 
            Err(ive) => Err(IntoSafeArrElemError::from(ive))
        }
    }}
}}
safe_arr_impl!{impl SafeArrayElement for Ptr<IDispatch>{
    SFTYPE = VT_DISPATCH; 
    ptr
    def => {{
        let mut var: IDispatch = unsafe {mem::zeroed()};
        &mut var as *mut IDispatch
    }}
    from => { |ptr: *mut IDispatch| {
        match Ptr::with_checked(ptr) {
            Some(pnn) => Ok(pnn), 
            None => Err(FromSafeArrElemError::DispatchPtrNull)
        }
    }}
    into => { |slf: Ptr<IDispatch>| -> Result<*mut IDispatch, IntoSafeArrElemError> {Ok(slf.as_ptr()) }}
}}
safe_arr_impl!{impl SafeArrayElement for SCode {
    SFTYPE = VT_ERROR;
    def => {0}
    from => {|sc| Ok(SCode::from(sc))}
    into => { |slf: SCode| -> Result<_, IntoSafeArrElemError> {Ok(i32::from(slf)) }}
}}
safe_arr_impl!{impl SafeArrayElement for bool {
    SFTYPE = VT_BOOL; 
    def => {0}
    from => {|vb| Ok(bool::from(VariantBool::from(vb)))} 
    into => {
        |slf: bool| -> Result<_, IntoSafeArrElemError> { Ok(VARIANT_BOOL::from(VariantBool::from(slf)))}
    }
}}
safe_arr_impl!{impl <T: VariantExt> SafeArrayElement for Variant<T> {
    SFTYPE = VT_VARIANT;
    ptr
    def => {{
        let mut var: VARIANT = unsafe {mem::zeroed()};
        &mut var as *mut VARIANT
    }}
    from => {|pvar| {
        let pnn = match Ptr::with_checked(pvar) {
            Some(nn) => nn, 
            None => return Err(FromSafeArrElemError::VariantPtrNull)
        };
        match Variant::<T>::from_variant(pnn) {
            Ok(var) => Ok(var), 
            Err(_) => Err(FromSafeArrElemError::FromVariantFailed)
        }
    }}
    into => {|slf: Variant<T>| -> Result<*mut VARIANT, IntoSafeArrElemError>{
        match slf.into_variant() {
            Ok(slf) => {
                let mut s = slf.as_ptr();
                Ok(s)
            }, 
            Err(ive) => Err(IntoSafeArrElemError::from(ive))
        }
    }}
}}
safe_arr_impl!{impl SafeArrayElement for Ptr<IUnknown> {
    SFTYPE = VT_UNKNOWN; 
    ptr
    def => {{
        let mut var: IUnknown = unsafe {mem::zeroed()};
        &mut var as *mut IUnknown
    }}
    from => {
        |ptr| {
            match Ptr::with_checked(ptr) {
                Some(ptr) => Ok(ptr), 
                None => Err(FromSafeArrElemError::UnknownPtrNull)
            }
        }
    }
    into => {
        |slf: Ptr<IUnknown>| -> Result<*mut IUnknown, IntoSafeArrElemError> {
            Ok(slf.as_ptr())
        }
    }
}}
safe_arr_impl!{impl SafeArrayElement for Decimal {
    SFTYPE = VT_DECIMAL; 
    def => {DECIMAL::from(DecWrapper::from(Decimal::new(0, 0)))}
    from => {|dec| Ok(Decimal::from(DecWrapper::from(dec)))}
    into => {
        |slf: Decimal| -> Result<_, IntoSafeArrElemError> {Ok(DECIMAL::from(DecWrapper::from(slf)))}
    }
}}
safe_arr_impl!{impl SafeArrayElement for DecWrapper { 
    SFTYPE = VT_DECIMAL; 
    def => {DECIMAL::from(DecWrapper::from(Decimal::new(0, 0)))}
    from => {|dec|Ok(DecWrapper::from(dec))} 
    into => { |slf: DecWrapper| -> Result<_, IntoSafeArrElemError> { Ok(DECIMAL::from(slf)) }}
}}
//VT_RECORD
safe_arr_impl!{impl SafeArrayElement for i8 {
    SFTYPE = VT_I1;
    def => { 0i8 }
    from => {|i| Ok(i)}
    into => { |slf: _| -> Result<_, IntoSafeArrElemError> { Ok(slf) }}
}}
safe_arr_impl!{impl SafeArrayElement for u8 {
    SFTYPE = VT_UI1;
    def => { 0u8}
    from => {|i| Ok(i)}
    into => { |slf: _| -> Result<_, IntoSafeArrElemError> { Ok(slf) }}
}}
safe_arr_impl!{impl SafeArrayElement for u16 {
    SFTYPE = VT_UI2;
    def => { 0u16 }
    from => {|i| Ok(i)}
    into => { |slf: _| -> Result<_, IntoSafeArrElemError> { Ok(slf) }}
}}
safe_arr_impl!{impl SafeArrayElement for u32 {
    SFTYPE = VT_UI4;
    def => { 0u32 }
    from => {|i| Ok(i)}
    into => { |slf: _| -> Result<_, IntoSafeArrElemError> { Ok(slf) }}
}}
safe_arr_impl!{impl SafeArrayElement for Int {
    SFTYPE = VT_INT;
    def => { 0i32 }
    from => {|i| Ok(Int::from(i))}
    into => { |slf: Int| -> Result<_, IntoSafeArrElemError> {Ok(i32::from(slf)) }}
}}
safe_arr_impl!{impl SafeArrayElement for UInt {
    SFTYPE = VT_UINT;
    def => { 0u32 }
    from => {|i| Ok(UInt::from(i))}
    into => { |slf: UInt| -> Result<_, IntoSafeArrElemError> {Ok(u32::from(slf)) }}
}}

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
    macro_rules! validate_safe_arr {
        ($t:ident, $vals:expr, $vt:expr) => {
            let v: Vec<$t> = $vals;

            let p = v.into_iter().into_safearray().unwrap();
            
            let r = ExactSizeIterator::<Item=$t>::from_safearray(p.as_ptr());
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
        let v: Vec<String> = vec![String::from("validate"), String::from("test string")];

        let p = v.into_iter().into_safearray().unwrap();

        let r = ExactSizeIterator::<Item=String>::from_safearray(p.as_ptr());

        let r = r.unwrap();
        assert_eq!(r, vec![String::from("validate"), String::from("test string")]);
    }

    #[test]
    fn test_scode() {
        validate_safe_arr!(SCode, vec![SCode::from(100), SCode::from(10000)], VT_ERROR );
    }
    #[test]
    fn test_bool() {
        validate_safe_arr!(bool, vec![true, false, true, true, false, false, true], VT_BOOL );
    }

    #[test]
    fn test_variant() {
        let v: Vec<Variant<u64>> = vec![Variant::new(100u64), Variant::new(100u64), Variant::new(103u64)];

        let p = v.into_iter().into_safearray().unwrap();
        
        let r = ExactSizeIterator::<Item=Variant<u64>>::from_safearray(p.as_ptr());
        let r = r.unwrap();
        assert_eq!(r,  vec![Variant::new(100u64), Variant::new(100u64), Variant::new(103u64)]);
    }

    #[test]
    fn test_decimal() {
        validate_safe_arr!(Decimal, vec![Decimal::new(2, 2), Decimal::new(3, 3)], VE_DECIMAL );
    }
    #[test]
    fn test_i8() {
        validate_safe_arr!(i8, vec![-1, 0,1,2,3,4], VT_I1 );
    }
    #[test]
    fn test_u8() {
        validate_safe_arr!(u8, vec![0,1,2,3,4], VT_UI1 );
    }
    #[test]
    fn test_u16() {
        validate_safe_arr!(u16, vec![0,1,2,3,4], VT_UI2 );
    }
    #[test]
    fn test_u32() {
        validate_safe_arr!(u32, vec![0,1,2,3,4], VT_UI4 );
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