

//SafeArrayElement 
//  into_safearray
//  from_safearray
#[derive(Debug, Fail)]
pub enum ElementError {
    #[fail(display = "{}", _0)]
    From(Box<FromSafeArrElemError>),
    #[fail(display = "{}", _0)]
    Into(Box<IntoSafeArrElemError>), 
}
#[derive(Copy, Clone, Debug, Fail)]
pub enum FromSafeArrElemError {
    #[fail(display = "SafeArrayGetElement failed with HRESULT=0x{:x}", hr)]
    GetElementFailed { hr: i32 },
    #[fail(display = "VARIANT pointer is null")]
    VariantPtrNull, 
    #[fail(display = "conversion from variant failed")]
    FromVariantFailed, 
    #[fail(display = "IUnknown pointer is null")]
    UnknownPtrNull,
    #[fail(display = "IDispatch pointer is null")]
    DispatchPtrNull,
}
#[derive(Debug, Fail)]
pub enum IntoSafeArrElemError {
    #[fail(display = "BSTR allocation failed for len: {}", len)]
    BStringAllocFailed{len: usize},
    #[fail(display = "VARIANT allocation failed for vartype: {}", vartype)]
    VariantAllocFailed{vartype: u32},
    #[fail(display = "SafeArrayPutElement failed with HRESULT = 0x{}", hr)]
    PutElementFailed { hr: i32 }, 
    #[fail(display = "IntoVariantError: {}", _0)]
    IntoVariantError(Box<IntoVariantError>),
}

impl From<FromSafeArrElemError> for ElementError {
    fn from(fsaee: FromSafeArrElemError) -> ElementError {
        ElementError::From(Box::new(fsaee))
    }
}

impl From<IntoSafeArrElemError> for ElementError {
    fn from(isaee: IntoSafeArrElemError) -> ElementError {
        ElementError::Into(Box::new(isaee))
    }
}

//SafeArrayExt
//  into_safearray
//  from_safearray
#[derive(Debug, Fail)]
pub enum SafeArrayError {
    #[fail(display = "{}", _0)]
    From(Box<FromSafeArrayError>),
    #[fail(display = "{}", _0)]
    Into(Box<IntoSafeArrayError>), 
}
#[derive(Debug, Fail)]
pub enum FromSafeArrayError{
    #[fail(display = "Safe array dimensions are invalid: {}", sa_dims)]
    SafeArrayDimsInvalid {sa_dims: u32},
    #[fail(display = "expected vartype was not found - expected: {} - found: {}", expected, found)]
    VarTypeDoesNotMatch {expected: u32, found: u32},
    #[fail(display = "SafeArrayGetLBound failed with HRESULT = 0x{}", hr)]
    SafeArrayLBoundFailed {hr: i32}, 
    #[fail(display = "SafeArrayGetRBound failed with HRESULT = 0x{}", hr)]
    SafeArrayRBoundFailed {hr: i32},
    #[fail(display = "SafeArrayGetVartype failed with HRESULT = 0x{}", hr)]
    SafeArrayGetVartypeFailed {hr: i32},
    #[fail(display = "element conversion failed at index {} with {}", index, element)]
    ElementConversionFailed {
        index: usize, 
        element: Box<ElementError>
    }
}
#[derive(Debug, Fail)]
pub enum IntoSafeArrayError {
    #[fail(display = "element conversion failed at index {} with {}", index, element)]
    ElementConversionFailed {
        index: usize, 
        element: Box<ElementError>
    },
    #[fail(display = "safe array creation failed")]
    SafeArrayCreateFailed,
}

impl From<FromSafeArrayError> for SafeArrayError {
    fn from(fsae: FromSafeArrayError) -> SafeArrayError {
        SafeArrayError::From(Box::new(fsae))
    }
}

impl From<IntoSafeArrayError> for SafeArrayError {
    fn from(isae: IntoSafeArrayError) -> SafeArrayError {
        SafeArrayError::Into(Box::new(isae))
    }
}

impl FromSafeArrayError {
    pub fn from_element_err<E: Into<ElementError>>(ee: E, index: usize) -> FromSafeArrayError {
        FromSafeArrayError::ElementConversionFailed{index: index, element: Box::new(ee.into())}
    }
}

impl IntoSafeArrayError {
    pub fn from_element_err<E: Into<ElementError>>(ee: E, index: usize) -> IntoSafeArrayError {
        IntoSafeArrayError::ElementConversionFailed{index: index, element: Box::new(ee.into())}
    }
}

//impl <T: VariantExt> SafeArrayElement for Variant<T> 
//can fail on invalid pointer

//BStringExt
//  allocate_bstr
//  allocate_managed_bstr
#[derive(Clone, Copy, Debug, Fail)]
pub enum BStringError {
    #[fail(display = "BSTR allocation failed for len: {}", len)]
    AllocateFailed {len: usize},    
}

impl From<BStringError> for IntoSafeArrElemError {
    fn from(bse: BStringError) -> IntoSafeArrElemError {
        match bse {
            BStringError::AllocateFailed{len} =>  IntoSafeArrElemError::BStringAllocFailed{len: len}
        }
    }
}

impl From<BStringError> for IntoVariantError {
    fn from(bse: BStringError) -> IntoVariantError {
        IntoVariantError::AllocBStrFailed(bse)
    }
}

//VariantExt
//  from_variant
//  into_variant

#[derive(Debug, Fail)]
pub enum FromVariantError {
    #[fail(display = "expected vartype was not found - expected: {} - found: {}", expected, found)]
    VarTypeDoesNotMatch { expected: u32, found: u32 },
    #[fail(display = "{}", _0)]
    AllocBStr(BStringError),
    #[fail(display = "IUnknown pointer is null")]
    UnknownPtrNull,
    #[fail(display = "IDispatch pointer is null")]
    DispatchPtrNull,
    #[fail(display = "VARIANT pointer is null")]
    VariantPtrNull,
    #[fail(display = "SAFEARRAY pointer is null")]
    ArrayPtrNull, 
    #[fail(display = "void pointer is null")]
    CVoidPtrNull,
    #[fail(display = "Safe array conversion failed: {}", _0)]
    SafeArrConvFailed(Box<SafeArrayError>),
}
#[derive(Debug, Fail)]
pub enum IntoVariantError {
    #[fail(display = "{}", _0)]
    AllocBStrFailed(BStringError),
    #[fail(display = "SafeArray conversion failed: {}", _0)]
    SafeArrConvFailed(Box<SafeArrayError>),
}

impl From<IntoVariantError> for IntoSafeArrElemError {
    fn from(ive: IntoVariantError) -> IntoSafeArrElemError {
        IntoSafeArrElemError::IntoVariantError(Box::new(ive))
    }
}

impl<I: Into<SafeArrayError>> From<I> for FromVariantError {
    fn from(i: I) -> FromVariantError {
        FromVariantError::SafeArrConvFailed(Box::new(i.into()))
    }
}

impl<I: Into<SafeArrayError>> From<I> for IntoVariantError {
    fn from(i: I) -> IntoVariantError {
        IntoVariantError::SafeArrConvFailed(Box::new(i.into()))
    }
}


//Variant
//  to_variant  
//  from_variant

//impl VariantExt for String
// calling unwrap without checking value

//impl VariantExt for Ptr<IUnknown>
//impl VariantExt for Ptr<IDispatch> 
//calling unwrap without checking value

//impl VariantExt for Box<String>
//calling unwrap without checking value

//impl<T: SafeArrayElement> VariantExt for Vec<T>
//panics and errs