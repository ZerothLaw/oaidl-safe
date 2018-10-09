/// Supererror type SafeArray element conversion errors
#[derive(Debug, Fail)]
pub enum ElementError {
    /// Holds FromSafeArrElemErrors
    #[fail(display = "{}", _0)]
    From(Box<FromSafeArrElemError>),
    /// Holds IntoSafeArrElemErrors
    #[fail(display = "{}", _0)]
    Into(Box<IntoSafeArrElemError>), 
}

/// Errors for converting from C/C++ data structure to Rust types
#[derive(Copy, Clone, Debug, Fail)]
pub enum FromSafeArrElemError {
    /// The unsafe call to SafeArrayGetElement failed - HRESULT stored within tells why
    #[fail(display = "SafeArrayGetElement failed with HRESULT=0x{:x}", hr)]
    GetElementFailed { 
        /// Holds an HRESULT value
        hr: i32 
    },
    /// VARIANT pointer during conversion was null
    #[fail(display = "VARIANT pointer is null")]
    VariantPtrNull, 
    /// The call to `.into_variant()` failed for some reason
    #[fail(display = "conversion from variant failed")]
    FromVariantFailed, 
    /// IUnknown pointer during conversion was null
    #[fail(display = "IUnknown pointer is null")]
    UnknownPtrNull,
    /// IDispatch pointer during conversion was null
    #[fail(display = "IDispatch pointer is null")]
    DispatchPtrNull,
}

/// Errors for converting into C/C++ data structures from Rust types
#[derive(Debug, Fail)]
pub enum IntoSafeArrElemError {
    /// `SysAllocStringLen` failed with len
    #[fail(display = "BSTR allocation failed for len: {}", len)]
    BStringAllocFailed{
        /// The len used that failed.
        len: usize
    },
    /// `VARIANT` allocation failed
    #[fail(display = "VARIANT allocation failed for vartype: {}", vartype)]
    VariantAllocFailed{
        /// vartype that failed
        vartype: u32
    },
    /// `SafeArrayPutElement` failed with `HRESULT`
    #[fail(display = "SafeArrayPutElement failed with HRESULT = 0x{}", hr)]
    PutElementFailed { 
        /// HRESULT returned by SafeArrayPutElement call
        hr: i32 
    }, 
    /// Encapsulates a `IntoVariantError`
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

/// Supererror for SafeArray errors
#[derive(Debug, Fail)]
pub enum SafeArrayError {
    /// From wrapper for `FromSafeArrayError`
    #[fail(display = "{}", _0)]
    From(Box<FromSafeArrayError>),
    /// Into wrapper for `IntoSafeArrayError`
    #[fail(display = "{}", _0)]
    Into(Box<IntoSafeArrayError>), 
}

/// Represents the different ways converting from `SAFEARRAY` can fail
#[derive(Debug, Fail)]
pub enum FromSafeArrayError{
    /// Either the safe array dimensions = 0 or > 1
    /// multi-dimensional arrays are *not* handled.
    #[fail(display = "Safe array dimensions are invalid: {}", sa_dims)]
    SafeArrayDimsInvalid {
        /// safe array dimensions that was wrong
        sa_dims: u32
    },
    /// Expected vartype did not match found vartype - runtime consistency check
    #[fail(display = "expected vartype was not found - expected: {} - found: {}", expected, found)]
    VarTypeDoesNotMatch {
        /// The expected vartype
        expected: u32, 
        /// the found vartype
        found: u32
    },
    /// Call to SafeArrayGetLBound failed
    #[fail(display = "SafeArrayGetLBound failed with HRESULT = 0x{}", hr)]
    SafeArrayLBoundFailed {
        /// HRESULT returned
        hr: i32
    }, 
    /// Call to SafeArrayGetRBound failed
    #[fail(display = "SafeArrayGetRBound failed with HRESULT = 0x{}", hr)]
    SafeArrayRBoundFailed {
        /// HRESULT returned
        hr: i32
    },
    /// Call to SafeArrayGetVartype failed
    #[fail(display = "SafeArrayGetVartype failed with HRESULT = 0x{}", hr)]
    SafeArrayGetVartypeFailed {
        /// HRESULT returned
        hr: i32
    },
    /// Encapsulates the `ElementError` that occurred during conversion
    #[fail(display = "element conversion failed at index {} with {}", index, element)]
    ElementConversionFailed {
        /// the index the conversion failed at
        index: usize, 
        /// The element error encapsulating the failure
        element: Box<ElementError>
    }
}

/// Represents the different ways converting into `SAFEARRAY` can fail
#[derive(Debug, Fail)]
pub enum IntoSafeArrayError {
    /// Encapsulates the `ElementError` that occurred during conversion
    #[fail(display = "element conversion failed at index {} with {}", index, element)]
    ElementConversionFailed {
       /// the index the conversion failed at
        index: usize, 
        /// The element error encapsulating the failure
        element: Box<ElementError>
    },
    /// The called to `SafeArrayCreate` failed
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
    /// converts an `ElementError` into a `FromSafeArrayError`
    /// Need the index so a From impl doesn't apply
    pub fn from_element_err<E: Into<ElementError>>(ee: E, index: usize) -> FromSafeArrayError {
        FromSafeArrayError::ElementConversionFailed{index: index, element: Box::new(ee.into())}
    }
}

impl IntoSafeArrayError {
    /// converts an `ElementError` into a `FromSafeArrayError`
    /// Need the index so a From impl doesn't apply
    pub fn from_element_err<E: Into<ElementError>>(ee: E, index: usize) -> IntoSafeArrayError {
        IntoSafeArrayError::ElementConversionFailed{index: index, element: Box::new(ee.into())}
    }
}

/// Ways BString can fail. Currently just one way.
#[derive(Clone, Copy, Debug, Fail)]
pub enum BStringError {
    /// SysAllocStringLen failed
    #[fail(display = "BSTR allocation failed for len: {}", len)]
    AllocateFailed {
        /// len which was used for allocation
        len: usize
    },    
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

/// Encapsulates the ways converting from a `VARIANT` can fail.
#[derive(Debug, Fail)]
pub enum FromVariantError {
    /// Expected vartype did not match found vartype - runtime consistency check
    #[fail(display = "expected vartype was not found - expected: {} - found: {}", expected, found)]
    VarTypeDoesNotMatch {
        /// The expected vartype
        expected: u32, 
        /// the found vartype
        found: u32
    },
    /// Encapsulates BString errors
    #[fail(display = "{}", _0)]
    AllocBStr(BStringError),
    /// `IUnknown` pointer during conversion was null
    #[fail(display = "IUnknown pointer is null")]
    UnknownPtrNull,
    /// `IDispatch` pointer during conversion was null
    #[fail(display = "IDispatch pointer is null")]
    DispatchPtrNull,
    /// `VARIANT` pointer during conversion was null
    #[fail(display = "VARIANT pointer is null")]
    VariantPtrNull,
    /// `SAFEARRAY` pointer during conversion was null
    #[fail(display = "SAFEARRAY pointer is null")]
    ArrayPtrNull, 
    /// `*mut c_void` pointer during conversion was null
    #[fail(display = "void pointer is null")]
    CVoidPtrNull,
    /// Conversion into `SAFEARRAY` failed.
    #[fail(display = "Safe array conversion failed: {}", _0)]
    SafeArrConvFailed(Box<SafeArrayError>),
}

/// Encapsulates errors that can occur during conversion into VARIANT
#[derive(Debug, Fail)]
pub enum IntoVariantError {
    /// Encapsulates a `BStringError`
    #[fail(display = "{}", _0)]
    AllocBStrFailed(BStringError),
    /// Encapsulates a `SafeArrayError`
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