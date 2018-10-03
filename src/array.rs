use std::fmt::Display;
use std::ptr::null_mut;
use rust_decimal::Decimal;

use winapi::ctypes::{c_long, c_void};
use winapi::shared::minwindef::{UINT, ULONG,};
use winapi::shared::ntdef::HRESULT;
use winapi::shared::wtypes::{DATE, CY,  VARTYPE};
use winapi::um::oaidl::{IDispatch, SAFEARRAYBOUND, LPSAFEARRAYBOUND, SAFEARRAY};
use winapi::um::unknwnbase::IUnknown;


use ptr::Ptr;
use types::{Currency, Date, DecWrapper, VariantBool};
use variant::Variant;

use winapi::shared::wtypes::{
    VARIANT_BOOL,
    VT_BSTR, 
    VT_BOOL,
    VT_CY,
    VT_DATE,
    VT_DECIMAL, 
    VT_DISPATCH,
    VT_I1, 
    VT_I2, 
    VT_I4,
    VT_R4, 
    VT_R8, 
    //VT_RECORD,
    VT_UI1,
    VT_UI2,
    VT_UI4,
    VT_UNKNOWN, 
    VT_VARIANT,   
};



pub trait SafeArrayElement: Sized {
    const VARTYPE: u32;

    fn into_safearray(&mut self, psa: *mut SAFEARRAY, ix: i32) -> Result<(), i32>;
    fn from_safearray(psa: *mut SAFEARRAY, ix: i32) -> Result<Self, i32>;
}

pub struct SafeArr<T: SafeArrayElement> {
    array: Vec<T>
}

impl<T: SafeArrayElement> SafeArr<T> {
    pub fn new(arr: Vec<T>) -> SafeArr<T> {
        SafeArr { array: arr }
    }

    pub fn unwrap(self) -> Vec<T> {
        self.array
    }

    pub fn into_safearray(self) -> Result<Ptr<SAFEARRAY>, ()> {
        let c_elements: ULONG = self.array.len() as u32;
        let vartype = T::VARTYPE;
        let mut sab = SAFEARRAYBOUND { cElements: c_elements, lLbound: 0i32};
        let psa = unsafe { SafeArrayCreate(vartype as u16, 1, &mut sab)};

        for (ix, mut elem) in self.array.into_iter().enumerate() {
            match elem.into_safearray(psa, ix as i32) {
                Ok(()) => continue, 
                Err(_) => return Err(())
            }
        }

        Ok(Ptr::with_checked(psa).unwrap())
    }

    pub fn from_safearray(psa: *mut SAFEARRAY) -> Result<SafeArr<T>, i32> {
        let sa_dims = unsafe { SafeArrayGetDim(psa) };
        assert!(sa_dims > 0); //Assert its not a dimensionless sa
        let vt = unsafe {
            let mut vt: VARTYPE = 0;
            let _hr = SafeArrayGetVartype(psa, &mut vt);
            vt
        };
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
            Ok(SafeArr::new(vc))
        } else {
            panic!("Multiple dimension arrays not yet supported.")
        }
    }
}

macro_rules! safe_arr_impl {
    ($t:ty => $vt:expr, {$from:expr} <=> {$into:expr}) => {
        impl SafeArrayElement for $t {
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

safe_arr_impl!(i16 => VT_I2, 
    { 
        |psa, ix| {
            let mut i = 0i16;
            let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void);
            (hr, i)
        }
    } <=> {
        |slf, psa, ix| {
            SafeArrayGetElement(psa, &ix, slf as *mut _ as *mut c_void)
        }
    }
);
safe_arr_impl!(i32 => VT_I4, 
    {
        |psa, ix|{
            let mut i = 0i32; 
            let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void); 
            (hr, i)
        }
    } <=> {
        |slf, psa, ix| {
            SafeArrayPutElement(psa, &ix, slf as *mut _ as *mut c_void)
        }
    }
);
safe_arr_impl!(f32 => VT_R4, 
    {
        |psa, ix| {
            let mut f = 0f32;
            let hr = SafeArrayGetElement(psa, &ix, &mut f as *mut _ as *mut c_void);
            (hr, f)
        }
    } <=> {
        |slf, psa, ix| {
            SafeArrayPutElement(psa, &ix, slf as *mut _ as *mut c_void)
        }
    }
);
safe_arr_impl!(f64 => VT_R8, 
    {
        |psa, ix| {
            let mut f = 0f64;
            let hr = SafeArrayGetElement(psa, &ix, &mut f as *mut _ as *mut c_void);
            (hr, f)
        }
    } <=> {
        |slf, psa, ix| {
            SafeArrayPutElement(psa, &ix, slf as *mut _ as *mut c_void)
        }
    }
);
safe_arr_impl!(Currency => VT_CY, 
    {
        |psa, ix| {
            let mut cy = CY{int64: 0};
            let hr = SafeArrayGetElement(psa, &ix, &mut cy as *mut _ as *mut c_void);
            (hr, Currency::from(cy))
        }
    } <=> {
        |slf: &mut Currency, psa, ix| {
            let mut slf = CY::from(*slf);
            SafeArrayPutElement(psa, &ix, &mut slf as *mut _ as *mut c_void)
        }
    }
);
safe_arr_impl!(Date => VT_DATE, 
    {
        |psa, ix| {
            let mut dt: DATE = 0f64;
            let hr = SafeArrayGetElement(psa, &ix, &mut dt as *mut _ as *mut c_void);
            (hr, Date::from(dt))
        }
    } <=> {
        |slf: &mut Date, psa, ix| {
            let mut slf = DATE::from(*slf);
            SafeArrayPutElement(psa, &ix, &mut slf as *mut _ as *mut c_void)
        }
    }
);
// safe_arr_impl!(String => VT_BSTR);
safe_arr_impl!(Ptr<IDispatch> => VT_DISPATCH, 
    {
        |psa, ix| {
            let mut ptr: *mut IDispatch = null_mut();
            let hr = SafeArrayGetElement(psa, &ix, &mut ptr as *mut *mut _ as *mut c_void);
            (hr, Ptr::with_checked(ptr).unwrap())
        }
    } <=> {
        |slf: &mut Ptr<IDispatch>, psa, ix| {
            let mut slf = (*slf).as_ptr();
            SafeArrayPutElement(psa, &ix, &mut slf as *mut *mut _ as *mut c_void)
        }
    }
);
// //safe_arr_impl!(SCode => VT_ERROR);
safe_arr_impl!(bool => VT_BOOL, 
    {
        |psa, ix| {
            let mut vb: VARIANT_BOOL = 0;
            let hr = SafeArrayGetElement(psa, &ix, &mut vb as *mut _ as *mut c_void);
            (hr, bool::from(VariantBool::from(vb)))
        }
    } <=> {
        |slf: &mut bool, psa, ix| {
            let mut slf = VARIANT_BOOL::from(VariantBool::from(*slf));
            SafeArrayPutElement(psa, &ix, &mut slf as *mut _ as *mut c_void)
        }
    }
);
// safe_arr_impl!(Variant => VT_VARIANT);
safe_arr_impl!(Ptr<IUnknown> => VT_UNKNOWN, 
    {
        |psa, ix| {
            let mut ptr: *mut IUnknown = null_mut();
            let hr = SafeArrayGetElement(psa, &ix, &mut ptr as *mut *mut _ as *mut c_void);
            (hr, Ptr::with_checked(ptr).unwrap())
        }
    } <=> {
        |slf: &mut Ptr<IUnknown>, psa, ix| {
            let mut slf = (*slf).as_ptr();
            SafeArrayPutElement(psa, &ix, &mut slf as *mut *mut _ as *mut c_void)
        }
    }
);
// safe_arr_impl!(Decimal => VT_DECIMAL);
// safe_arr_impl!(DecWrapper => VT_DECIMAL);
safe_arr_impl!(i8 => VT_I1, 
    {
        |psa, ix|{
            let mut i = 0i8; 
            let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void); 
            (hr, i)
        }
    } <=> {
        |slf, psa, ix| {
            SafeArrayPutElement(psa, &ix, slf as *mut _ as *mut c_void)
        }
    }
);
safe_arr_impl!(u8 => VT_UI1, 
    {
        |psa, ix|{
            let mut i = 0u8; 
            let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void); 
            (hr, i)
        }
    } <=> {
        |slf, psa, ix| {
            SafeArrayPutElement(psa, &ix, slf as *mut _ as *mut c_void)
        }
    }
);
safe_arr_impl!(u16 => VT_UI2, 
    {
        |psa, ix|{
            let mut i = 0u16; 
            let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void); 
            (hr, i)
        }
    } <=> {
        |slf, psa, ix| {
            SafeArrayPutElement(psa, &ix, slf as *mut _ as *mut c_void)
        }
    }
);
safe_arr_impl!(u32 => VT_UI4, 
    {
        |psa, ix|{
            let mut i = 0u32; 
            let hr = SafeArrayGetElement(psa, &ix, &mut i as *mut _ as *mut c_void); 
            (hr, i)
        }
    } <=> {
        |slf, psa, ix| {
            SafeArrayPutElement(psa, &ix, slf as *mut _ as *mut c_void)
        }
    }
);
//safe_arr_impl!(Int => VT_INT);
//safe_arr_impl!(UInt => VT_UINT);

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

pub use winapi::um::oaidl::LPSAFEARRAY;

#[cfg(test)]
mod test {
    use super::*;
    #[test]
    fn test_create() {
        let v: Vec<i8> = vec![0,1,2,3,4];
        println!("{:?}", v);
        let mut sa = SafeArr::new(v);
        let p = sa.into_safearray().unwrap();
        println!("{:p}", p.as_ptr());

        let r = SafeArr::<i8>::from_safearray(p.as_ptr());
        println!("{:?}", r.unwrap().unwrap());
    }

}