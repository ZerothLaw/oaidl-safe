//SafeArrayElement 
//  into_safearray
//  from_safearray
#[derive(Debug)]
pub enum ElementError {
    From(FromSafeArrElemError),
    Into(IntoSafeArrElemError), 
}
#[derive(Debug)]
pub enum FromSafeArrElemError {
    GetElementFailed { hr: i32 },
    VariantPtrNull, 
    FromVariantFailed, 
    UnknownPtrNull,
    DispatchPtrNull,
}
#[derive(Debug)]
pub enum IntoSafeArrElemError {
    BStringAllocFailed{len: usize},
    VariantAllocFailed{vartype: u32},
    PutElementFailed { hr: i32 }
}

impl From<FromSafeArrElemError> for ElementError {
    fn from(fsaee: FromSafeArrElemError) -> ElementError {
        ElementError::From(fsaee)
    }
}

impl From<IntoSafeArrElemError> for ElementError {
    fn from(isaee: IntoSafeArrElemError) -> ElementError {
        ElementError::Into(isaee)
    }
}

//SafeArrayExt
//  into_safearray
//  from_safearray
#[derive(Debug)]
pub enum SafeArrayError {
    From(FromSafeArrayError),
    Into(IntoSafeArrayError), 
}
#[derive(Debug)]
pub enum FromSafeArrayError{
    SafeArrayDimsInvalid {sa_dims: u32},
    VarTypeDoesNotMatch {expected: u32, found: u32},
    SafeArrayLBoundFailed {hr: i32}, 
    SafeArrayRBoundFailed {hr: i32},
    SafeArrayGetVartypeFailed {hr: i32},
    ElementConversionFailed {
        index: usize, 
        element: ElementError
    }
}
#[derive(Debug)]
pub enum IntoSafeArrayError {
    ElementConversionFailed {
        index: usize, 
        element: ElementError
    },
    SafeArrayCreateFailed,
}

impl From<FromSafeArrayError> for SafeArrayError {
    fn from(fsae: FromSafeArrayError) -> SafeArrayError {
        SafeArrayError::From(fsae)
    }
}

impl From<IntoSafeArrayError> for SafeArrayError {
    fn from(isae: IntoSafeArrayError) -> SafeArrayError {
        SafeArrayError::Into(isae)
    }
}

impl FromSafeArrayError {
    pub fn from_element_err<E: Into<ElementError>>(ee: E, index: usize) -> FromSafeArrayError {
        FromSafeArrayError::ElementConversionFailed{index: index, element: ee.into()}
    }
}

impl IntoSafeArrayError {
    pub fn from_element_err<E: Into<ElementError>>(ee: E, index: usize) -> IntoSafeArrayError {
        IntoSafeArrayError::ElementConversionFailed{index: index, element: ee.into()}
    }
}

//impl <T: VariantExt> SafeArrayElement for Variant<T> 
//can fail on invalid pointer

//BStringExt
//  allocate_bstr
//  allocate_managed_bstr
#[derive(Debug)]
pub enum BStringError {
    AllocateFailed {len: usize},    
}

//VariantExt
//  from_variant
//  into_variant
#[derive(Debug)]
pub enum VariantError {
    From(FromVariantError), 
    Into(IntoVariantError),
}
#[derive(Debug)]
pub enum FromVariantError {
    VarTypeDoesNotMatch { expected: u32, found: u32 },
    AllocBStr(BStringError),
    UnknownPtrNull,
    DispatchPtrNull,
    VariantPtrNull,
    ArrayPtrNull, 
    CVoidPtrNull,
}
#[derive(Debug)]
pub enum IntoVariantError {
    AllocBStrFailed(BStringError),
    SafeArrConvFailed(SafeArrayError),
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