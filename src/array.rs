use std::marker::PhantomData;
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

use winapi::um::oaidl::{IDispatch, LPSAFEARRAY, LPSAFEARRAYBOUND, VARIANT, SAFEARRAY, SAFEARRAYBOUND};
use winapi::um::unknwnbase::IUnknown;

use bstr::BStringExt;
use ptr::Ptr;
use types::{Currency, Date, DecWrapper, Int, SCode, UInt, VariantBool};
use variant::{Variant, VariantExt};

pub trait SafeArrayElement: Sized {
    const VARTYPE: u32;

    fn into_safearray(&mut self, psa: *mut SAFEARRAY, ix: i32) -> Result<(), i32>;
    fn from_safearray(psa: *mut SAFEARRAY, ix: i32) -> Result<Self, i32>;
}

pub trait SafeArrayExt<T: SafeArrayElement> {
    fn into_safearray(&mut self) -> Result<Ptr<SAFEARRAY>, ()>;
    fn from_safearray(psa: *mut SAFEARRAY) -> Result<Vec<T>, i32>;
}


struct SafeArrayDestructor {
    inner: *mut SAFEARRAY, 
    _marker: PhantomData<SAFEARRAY>
}

impl SafeArrayDestructor {
    fn new(p: *mut SAFEARRAY) -> SafeArrayDestructor {
        assert!(!p.is_null());
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
    fn into_safearray(&mut self) -> Result<Ptr<SAFEARRAY>, ()> {
        let c_elements: ULONG = self.len() as u32;
        let vartype = T::VARTYPE;
        let mut sab = SAFEARRAYBOUND { cElements: c_elements, lLbound: 0i32};
        let psa = unsafe { SafeArrayCreate(vartype as u16, 1, &mut sab)};
        let mut sad = SafeArrayDestructor::new(psa);

        for (ix, mut elem) in self.iter_mut().enumerate() {
            match elem.into_safearray(psa, ix as i32) {
                Ok(()) => continue, 
                Err(_) => return Err(())
            }
        }
        sad.inner = null_mut();

        Ok(Ptr::with_checked(psa).unwrap())
    }

    fn from_safearray(psa: *mut SAFEARRAY) -> Result<Vec<T>, i32> {
        //Stack sentinel to ensure safearray is released even if there is a panic or early return.
        let _sad = SafeArrayDestructor::new(psa);
        let sa_dims = unsafe { SafeArrayGetDim(psa) };
        assert!(sa_dims > 0); //Assert its not a dimensionless safe array
        let vt = unsafe {
            let mut vt: VARTYPE = 0;
            let _hr = SafeArrayGetVartype(psa, &mut vt);
            vt
        };
        assert!(vt as u32 == T::VARTYPE);

        if sa_dims == 1 {
            let (l_bound, r_bound) = unsafe {
                let mut l_bound: c_long = 0;
                let mut r_bound: c_long = 0;
                let _hr = SafeArrayGetLBound(psa, 1, &mut l_bound);
                let _hr = SafeArrayGetUBound(psa, 1, &mut r_bound);
                (l_bound, r_bound)
            };

            let mut vc: Vec<T> = Vec::new();
            for ix in l_bound..=r_bound {
                match T::from_safearray(psa, ix) {
                    Ok(val) => vc.push(val), 
                    Err(hr) => return Err(hr)
                }
            }
            Ok(vc)
        } else {
            panic!("Multiple dimension arrays not yet supported.")
        }
    }
} 

macro_rules! safe_arr_impl {
    (
        impl $(< $tn:ident : $tc:ident >)* SafeArrayElement for $t:ty {
            SFTYPE = $vt:expr;
            from => {$from:expr}
            into => {$into:expr}
        }
    ) => {
        impl $(<$tn:$tc>)* SafeArrayElement for $t {
            const VARTYPE: u32 = $vt;
             fn from_safearray(psa: *mut SAFEARRAY, ix: i32) -> Result<Self, i32> {
                let (hr, i) = unsafe {$from(psa, ix)};
                match hr {
                    0 => Ok(i), 
                    _ => Err(hr)
                }
            }
            
            fn into_safearray(&mut self, psa: *mut SAFEARRAY, ix: i32) -> Result<(), i32> {
                let hr = unsafe {$into(self, psa, ix)};
                match hr {
                    0 => Ok(()), 
                    _ => Err(hr)
                }
            }
        }
    };
}

safe_arr_impl!{impl SafeArrayElement for i16 {
    SFTYPE = VT_I2;
    from => {
        |psa, ix| {
            let mut i = 0i16;
            let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void);
            (hr, i)
        }
    }
    into => {
        |slf, psa, ix| {
            SafeArrayPutElement(psa, &ix, slf as *mut _ as *mut c_void)
        }
    }
}}
safe_arr_impl!{impl SafeArrayElement for i32 {
    SFTYPE = VT_I4;
    from => {
        |psa, ix|{
            let mut i = 0i32; 
            let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void); 
            (hr, i)
        }
    }
    into => {
        |slf, psa, ix| {
            SafeArrayPutElement(psa, &ix, slf as *mut _ as *mut c_void)
        }
    }
}}
safe_arr_impl!{impl SafeArrayElement for f32 {
    SFTYPE = VT_R4;
    from => {
        |psa, ix| {
            let mut f = 0f32;
            let hr = SafeArrayGetElement(psa, &ix, &mut f as *mut _ as *mut c_void);
            (hr, f)
        }
    }
    into => {
        |slf, psa, ix| {
            SafeArrayPutElement(psa, &ix, slf as *mut _ as *mut c_void)
        }
    }
}}
safe_arr_impl!{impl SafeArrayElement for f64 { 
    SFTYPE = VT_R8; 
    from => {
        |psa, ix| {
            let mut f = 0f64;
            let hr = SafeArrayGetElement(psa, &ix, &mut f as *mut _ as *mut c_void);
            (hr, f)
        }
    } 
    into => {
        |slf, psa, ix| {
            SafeArrayPutElement(psa, &ix, slf as *mut _ as *mut c_void)
        }
    }
}}
safe_arr_impl!{impl SafeArrayElement for Currency{
    SFTYPE = VT_CY; 
    from => {
        |psa, ix| {
            let mut cy = CY{int64: 0};
            let hr = SafeArrayGetElement(psa, &ix, &mut cy as *mut _ as *mut c_void);
            (hr, Currency::from(cy))
        }
    }
    into => {
        |slf: &mut Currency, psa, ix| {
            let mut slf = CY::from(*slf);
            SafeArrayPutElement(psa, &ix, &mut slf as *mut _ as *mut c_void)
        }
    }
}}
safe_arr_impl!{impl SafeArrayElement for Date{
    SFTYPE = VT_DATE; 
    from => {
        |psa, ix| {
            let mut dt: DATE = 0f64;
            let hr = SafeArrayGetElement(psa, &ix, &mut dt as *mut _ as *mut c_void);
            (hr, Date::from(dt))
        }
    } 
    into => {
        |slf: &mut Date, psa, ix| {
            let mut slf = DATE::from(*slf);
            SafeArrayPutElement(psa, &ix, &mut slf as *mut _ as *mut c_void)
        }
    }
}}
// safe_arr_impl!(String => VT_BSTR);
safe_arr_impl!{impl SafeArrayElement for String {
    SFTYPE = VT_BSTR;
    from => {|psa, ix|{
        let mut bstr: *mut u16 = null_mut();
        let hr = SafeArrayGetElement(psa, &ix, &mut bstr as *mut _ as *mut c_void );
        (hr, U16String::from_bstr(bstr).to_string_lossy())
    }}
    into => {|slf: &mut String, psa, ix|{
        let bstr = U16String::from_str(slf);
        let bstr = match bstr.allocate_bstr() {
            Ok(bstr) => bstr, 
            Err(()) => panic!("Nope"),
        };
        let mut bstr = bstr.as_ptr();
        SafeArrayPutElement(psa, &ix, &mut bstr as *mut _ as *mut c_void)
    }}
}}
safe_arr_impl!{impl SafeArrayElement for Ptr<IDispatch>{
    SFTYPE = VT_DISPATCH; 
    from => {
        |psa, ix| {
            let mut ptr: *mut IDispatch = null_mut();
            let hr = SafeArrayGetElement(psa, &ix, &mut ptr as *mut *mut _ as *mut c_void);
            (hr, Ptr::with_checked(ptr).unwrap())
        }
    }
    into => {
        |slf: &mut Ptr<IDispatch>, psa, ix| {
            let mut slf = (*slf).as_ptr();
            SafeArrayPutElement(psa, &ix, &mut slf as *mut *mut _ as *mut c_void)
        }
    }
}}
safe_arr_impl!{impl SafeArrayElement for SCode {
    SFTYPE = VT_ERROR;
    from => {
        |psa, ix| {
            let mut sc: SCODE = 0;
            let hr = SafeArrayGetElement(psa, &ix, &mut sc as *mut _ as *mut c_void);
            (hr, SCode(sc))
        }
    }
    into => {
        |slf: &mut SCode, psa, ix| {
            let mut slf = slf.0;
            SafeArrayPutElement(psa, &ix, &mut slf as *mut _ as *mut c_void)
        }
    }
}}
safe_arr_impl!{impl SafeArrayElement for bool {
    SFTYPE = VT_BOOL; 
    from => {
        |psa, ix| {
            let mut vb: VARIANT_BOOL = 0;
            let hr = SafeArrayGetElement(psa, &ix, &mut vb as *mut _ as *mut c_void);
            (hr, bool::from(VariantBool::from(vb)))
        }
    } 
    into => {
        |slf: &mut bool, psa, ix| {
            let mut slf = VARIANT_BOOL::from(VariantBool::from(*slf));
            SafeArrayPutElement(psa, &ix, &mut slf as *mut _ as *mut c_void)
        }
    }
}}
safe_arr_impl!{impl <T: VariantExt> SafeArrayElement for Variant<T> {
    SFTYPE = VT_VARIANT;
    from => {|psa, ix| {
        let mut pvar: *mut VARIANT = null_mut();
        let hr = SafeArrayGetElement(psa, &ix, &mut pvar as *mut _ as *mut c_void);
        let nn = match NonNull::new(pvar) {
            Some(nn) => nn, 
            None => panic!("Did not get a valid variant pointer")
        };
        let pnn = Ptr::new(nn);
        match Variant::<T>::from_variant(pnn) {
            Ok(var) => (hr, var), 
            Err(_) => panic!("Invalid variant pointer")
        }
    }}
    into => {|slf: &mut Variant<T>, psa, ix|{
        let mut slf = slf.to_variant();
        SafeArrayPutElement(psa, &ix, &mut slf as *mut _ as *mut c_void)
    }}
}}
safe_arr_impl!{impl SafeArrayElement for Ptr<IUnknown> {
    SFTYPE = VT_UNKNOWN; 
    from => {
        |psa, ix| {
            let mut ptr: *mut IUnknown = null_mut();
            let hr = SafeArrayGetElement(psa, &ix, &mut ptr as *mut *mut _ as *mut c_void);
            (hr, Ptr::with_checked(ptr).unwrap())
        }
    }
    into => {
        |slf: &mut Ptr<IUnknown>, psa, ix| {
            let mut slf = (*slf).as_ptr();
            SafeArrayPutElement(psa, &ix, &mut slf as *mut *mut _ as *mut c_void)
        }
    }
}}
safe_arr_impl!{impl SafeArrayElement for Decimal {
    SFTYPE = VT_DECIMAL; 
    from => {
        |psa, ix|{
            let mut dec: DECIMAL = DECIMAL::from(DecWrapper::from(Decimal::new(0, 0)));
            let hr = SafeArrayGetElement(psa, &ix, &mut dec as *mut _ as *mut c_void);
            (hr, Decimal::from(DecWrapper::from(dec)))
        }
    }
    into => {
        |slf: &mut Decimal, psa, ix| {
            let mut dec: DECIMAL =  DECIMAL::from(DecWrapper::from(*slf));
            SafeArrayPutElement(psa, &ix, &mut dec as *mut _ as *mut c_void)
        }
    }
}}
safe_arr_impl!{impl SafeArrayElement for DecWrapper { 
    SFTYPE = VT_DECIMAL; 
    from => {
        |psa, ix|{
            let mut dec: DECIMAL = DECIMAL::from(DecWrapper::from(Decimal::new(0, 0)));
            let hr = SafeArrayGetElement(psa, &ix, &mut dec as *mut _ as *mut c_void);
            (hr, DecWrapper::from(dec))
        }
    } 
    into => {
        |slf: &mut DecWrapper, psa, ix| {
            let mut dec: DECIMAL =  DECIMAL::from(slf);
            SafeArrayPutElement(psa, &ix, &mut dec as *mut _ as *mut c_void)
        }
    }
}}
//VT_RECORD
safe_arr_impl!{impl SafeArrayElement for i8 {
    SFTYPE = VT_I1;
    from => {
        |psa, ix|{
            let mut i = 0i8; 
            let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void); 
            (hr, i)
        }
    }
    into => {
        |slf, psa, ix| {
            SafeArrayPutElement(psa, &ix, slf as *mut _ as *mut c_void)
        }
    }
}}
safe_arr_impl!{impl SafeArrayElement for u8 {
    SFTYPE = VT_UI1;
    from => {
        |psa, ix|{
            let mut i = 0u8; 
            let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void); 
            (hr, i)
        }
    }
    into => {
        |slf, psa, ix| {
            SafeArrayPutElement(psa, &ix, slf as *mut _ as *mut c_void)
        }
    }
}}
safe_arr_impl!{impl SafeArrayElement for u16 {
    SFTYPE = VT_UI2;
    from => {
        |psa, ix|{
            let mut i = 0u16; 
            let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void); 
            (hr, i)
        }
    }
    into => {
        |slf, psa, ix| {
            SafeArrayPutElement(psa, &ix, slf as *mut _ as *mut c_void)
        }
    }
}}
safe_arr_impl!{impl SafeArrayElement for u32 {
    SFTYPE = VT_UI4;
    from => {
        |psa, ix|{
            let mut i = 0u32; 
            let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void); 
            (hr, i)
        }
    }
    into => {
        |slf, psa, ix| {
            SafeArrayPutElement(psa, &ix, slf as *mut _ as *mut c_void)
        }
    }
}}
safe_arr_impl!{impl SafeArrayElement for Int {
    SFTYPE = VT_INT;
    from => {
        |psa, ix|{
            let mut i = 0i32;
            let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void); 
            (hr, Int(i))
        }
    }
    into => {
        |slf: &mut Int, psa, ix| {
            SafeArrayPutElement(psa, &ix, &mut slf.0 as *mut _ as *mut c_void)
        }
    }
}}
safe_arr_impl!{impl SafeArrayElement for UInt {
    SFTYPE = VT_UINT;
    from => {
        |psa, ix|{
            let mut i = 0u32;
            let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void); 
            (hr, UInt(i))
        }
    }
    into => {
        |slf: &mut UInt, psa, ix| {
            SafeArrayPutElement(psa, &ix, &mut slf.0 as *mut _ as *mut c_void)
        }
    }
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
}