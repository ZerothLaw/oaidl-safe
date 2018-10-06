use std::marker::PhantomData;
use std::mem;
use std::ptr::{NonNull, null_mut};

use rust_decimal::Decimal;

use widestring::U16String;

use winapi::ctypes::{c_long, c_void};

use winapi::shared::minwindef::{UINT, ULONG,};
use winapi::shared::ntdef::HRESULT;
use winapi::shared::wtypes::{
    DATE, 
    DECIMAL, 
    CY,  
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
    //VT_RECORD,
    VT_UI1,
    VT_UI2,
    VT_UI4,
    VT_UINT,
    VT_UNKNOWN, 
    VT_VARIANT,   
};
use winapi::shared::wtypesbase::{SCODE};

use winapi::um::oaidl::{IDispatch, LPSAFEARRAY, LPSAFEARRAYBOUND, SAFEARRAY, SAFEARRAYBOUND, VARIANT};
use winapi::um::unknwnbase::IUnknown;

use bstr::BStringExt;
use errors::{ElementError, FromSafeArrayError, FromSafeArrElemError, IntoSafeArrayError, IntoSafeArrElemError, SafeArrayError};
use ptr::Ptr;
use types::{Currency, Date, DecWrapper, Int, SCode, UInt, VariantBool};
use variant::{Variant, VariantExt};

pub trait SafeArrayElement: Sized {
    const SFTYPE: u32;

    fn into_safearray(&mut self, psa: *mut SAFEARRAY, ix: i32) -> Result<(), IntoSafeArrElemError>;
    fn from_safearray(psa: *mut SAFEARRAY, ix: i32) -> Result<Self, FromSafeArrElemError>;
}

pub trait SafeArrayExt<T: SafeArrayElement> {
    fn into_safearray(&mut self) -> Result<Ptr<SAFEARRAY>, IntoSafeArrayError>;
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

impl<T: SafeArrayElement> SafeArrayExt<T> for Vec<T> {
    fn into_safearray(&mut self) -> Result<Ptr<SAFEARRAY>, IntoSafeArrayError > {
        let c_elements: ULONG = self.len() as u32;
        let vartype = T::SFTYPE;
        let mut sab = SAFEARRAYBOUND { cElements: c_elements, lLbound: 0i32};
        let psa = unsafe { SafeArrayCreate(vartype as u16, 1, &mut sab)};
        assert!(!psa.is_null());
        let mut sad = SafeArrayDestructor::new(psa);

        for (ix, mut elem) in self.iter_mut().enumerate() {
            match elem.into_safearray(psa, ix as i32) {
                Ok(()) => continue, 
                Err(e) => return Err(IntoSafeArrayError::from_element_err(e, ix as usize))
            }
        }
        sad.inner = null_mut();

        Ok(Ptr::with_checked(psa).unwrap())
    }

    fn from_safearray(psa: *mut SAFEARRAY) -> Result<Vec<T>, FromSafeArrayError> {
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

        if vt as u32 != T::SFTYPE {
            return Err(FromSafeArrayError::VarTypeDoesNotMatch{expected: T::SFTYPE, found: vt as u32});
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

            let mut vc: Vec<T> = Vec::new();
            for ix in l_bound..=r_bound {
                match T::from_safearray(psa, ix) {
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
//TODO: rewrite this and associated macro calls to extra repetitive code
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
            
            fn into_safearray(&mut self, psa: *mut SAFEARRAY, ix: i32) -> Result<(), IntoSafeArrElemError> {
                let slf = $into(self);
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
            
            fn into_safearray(&mut self, psa: *mut SAFEARRAY, ix: i32) -> Result<(), IntoSafeArrElemError> {
                let mut slf = $into(self);
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
    into => { |slf: &mut _| *slf }
}}
safe_arr_impl!{impl SafeArrayElement for i32 {
    SFTYPE = VT_I4;
    def => { 0i32 }
    from => {|i| Ok(i)}
    into => { |slf: &mut _| *slf }
}}
safe_arr_impl!{impl SafeArrayElement for f32 {
    SFTYPE = VT_R4;
    def => { 0.0f32 }
    from => {|i| Ok(i)}
    into => { |slf: &mut _| *slf }
}}
safe_arr_impl!{impl SafeArrayElement for f64 { 
    SFTYPE = VT_R8; 
    def => { 0.0f64 }
    from => {|i| Ok(i)}
    into => { |slf: &mut _| *slf }
}}
safe_arr_impl!{impl SafeArrayElement for Currency{
    SFTYPE = VT_CY; 
    def => { CY{int64: 0} }
    from => { |cy| Ok(Currency::from(cy)) }
    into => {|slf: &mut Currency| CY::from(*slf)}
}}
safe_arr_impl!{impl SafeArrayElement for Date{
    SFTYPE = VT_DATE; 
    def =>  { 0f64 as DATE }
    from => { |dt| Ok(Date::from(dt)) } 
    into => { |slf: &mut Date| {DATE::from(*slf)} }
}}
// safe_arr_impl!(String => VT_BSTR);
safe_arr_impl!{impl SafeArrayElement for String {
    SFTYPE = VT_BSTR;
    def => {null_mut()}
    from => {|bstr| Ok(U16String::from_bstr(bstr).to_string_lossy())}
    into => {|slf: &mut String|{
        let bstr = U16String::from_str(slf);
        let bstr = match bstr.allocate_bstr() {
            Ok(bstr) => bstr, 
            Err(()) => panic!("Nope"),
        };
        bstr.as_ptr()
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
    into => { |slf: &mut Ptr<IDispatch>| (*slf).as_ptr() }
}}
safe_arr_impl!{impl SafeArrayElement for SCode {
    SFTYPE = VT_ERROR;
    def => {0 as SCODE}
    from => {|sc| Ok(SCode(sc))}
    into => { |slf: &mut SCode| slf.0 }
}}
safe_arr_impl!{impl SafeArrayElement for bool {
    SFTYPE = VT_BOOL; 
    def => {0 as VARIANT_BOOL}
    from => {|vb| Ok(bool::from(VariantBool::from(vb)))} 
    into => {
        |slf: &mut bool|  VARIANT_BOOL::from(VariantBool::from(*slf))
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
    into => {|slf: &mut Variant<T>|{
        match slf.into_variant() {
            Ok(slf) => {
                let mut s = slf.as_ptr();
                s
            }, 
            Err(()) => panic!("Could not alloc variant")
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
        |slf: &mut Ptr<IUnknown>| {
            (*slf).as_ptr()
        }
    }
}}
safe_arr_impl!{impl SafeArrayElement for Decimal {
    SFTYPE = VT_DECIMAL; 
    def => {DECIMAL::from(DecWrapper::from(Decimal::new(0, 0)))}
    from => {|dec| Ok(Decimal::from(DecWrapper::from(dec)))}
    into => {
        |slf: &mut Decimal| DECIMAL::from(DecWrapper::from(*slf))
    }
}}
safe_arr_impl!{impl SafeArrayElement for DecWrapper { 
    SFTYPE = VT_DECIMAL; 
    def => {DECIMAL::from(DecWrapper::from(Decimal::new(0, 0)))}
    from => {|dec|Ok(DecWrapper::from(dec))} 
    into => { |slf: &mut DecWrapper| DECIMAL::from(slf) }
}}
//VT_RECORD
safe_arr_impl!{impl SafeArrayElement for i8 {
    SFTYPE = VT_I1;
    def => { 0i8 }
    from => {|i| Ok(i)}
    into => { |slf: &mut _| *slf }
}}
safe_arr_impl!{impl SafeArrayElement for u8 {
    SFTYPE = VT_UI1;
    def => { 0u8}
    from => {|i| Ok(i)}
    into => { |slf: &mut _| *slf }
}}
safe_arr_impl!{impl SafeArrayElement for u16 {
    SFTYPE = VT_UI2;
    def => { 0u16 }
    from => {|i| Ok(i)}
    into => { |slf: &mut _| *slf }
}}
safe_arr_impl!{impl SafeArrayElement for u32 {
    SFTYPE = VT_UI4;
    def => { 0u32 }
    from => {|i| Ok(i)}
    into => { |slf: &mut _| *slf }
}}
safe_arr_impl!{impl SafeArrayElement for Int {
    SFTYPE = VT_INT;
    def => { 0i32 }
    from => {|i| Ok(Int(i))}
    into => { |slf: &mut Int| slf.0 }
}}
safe_arr_impl!{impl SafeArrayElement for UInt {
    SFTYPE = VT_UINT;
    def => { 0u32 }
    from => {|i| Ok(UInt(i))}
    into => { |slf: &mut UInt| slf.0 }
}}

#[allow(dead_code)]
#[link(name="OleAut32")]
extern "system" {
    pub fn SafeArrayCreate(vt: VARTYPE, cDims: UINT, rgsabound: LPSAFEARRAYBOUND) -> LPSAFEARRAY;
	pub fn SafeArrayDestroy(safe: LPSAFEARRAY)->HRESULT;
    
    pub fn SafeArrayGetDim(psa: LPSAFEARRAY) -> UINT;
	
    pub fn SafeArrayGetElement(psa: LPSAFEARRAY, rgIndices: *const c_long, pv: *mut c_void) -> HRESULT;
    pub fn SafeArrayGetElemSize(psa: LPSAFEARRAY) -> UINT;
    
    pub fn SafeArrayGetLBound(psa: LPSAFEARRAY, nDim: UINT, plLbound: *mut c_long)->HRESULT;
    pub fn SafeArrayGetUBound(psa: LPSAFEARRAY, nDim: UINT, plUbound: *mut c_long)->HRESULT;
    
    pub fn SafeArrayGetVartype(psa: LPSAFEARRAY, pvt: *mut VARTYPE) -> HRESULT;

    pub fn SafeArrayLock(psa: LPSAFEARRAY) -> HRESULT;
	pub fn SafeArrayUnlock(psa: LPSAFEARRAY) -> HRESULT;
    
    pub fn SafeArrayPutElement(psa: LPSAFEARRAY, rgIndices: *const c_long, pv: *mut c_void) -> HRESULT;
}

#[cfg(test)]
mod test {
    use super::*;
    macro_rules! validate_safe_arr {
        ($t:ident, $vals:expr, $vt:expr) => {
            let mut v: Vec<$t> = $vals;

            let p = v.into_safearray().unwrap();
            
            let r = Vec::<$t>::from_safearray(p.as_ptr());
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
    fn test_cy() {
        validate_safe_arr!(Currency, vec![Currency(-1), Currency(2)], VT_CY );
    }
    #[test]
    fn test_date() {
        validate_safe_arr!(Date, vec![Date(0.01), Date(100.0/99.0)], VT_DATE );
    }

    #[test]
    fn test_variant() {
        let mut v: Vec<Variant<u64>> = vec![Variant::new(100u64), Variant::new(100u64), Variant::new(103u64)];

        let p = v.into_safearray().unwrap();
        
        let r = Vec::<Variant<u64>>::from_safearray(p.as_ptr());
        let r = r.unwrap();
        assert_eq!(r,  vec![Variant::new(100u64), Variant::new(100u64), Variant::new(103u64)]);
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