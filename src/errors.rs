use std::error::Error;

use thiserror::Error;

/// Supererror type for SafeArray element conversion errors
#[derive(Debug, Error)]
pub enum ElementError {
    /// Holds FromSafeArrElemErrors in a box
    #[error(transparent)]
    From(Box<FromSafeArrElemError>),
    /// Holds IntoSafeArrElemErrors in a box
    #[error(transparent)]
    Into(Box<IntoSafeArrElemError>),
}

impl From<FromSafeArrElemError> for ElementError {
    /// Holds a [`FromSafeArrElemError`] with a box. This means conversion is free.
    fn from(fsaee: FromSafeArrElemError) -> ElementError {
        ElementError::From(Box::new(fsaee))
    }
}

impl From<IntoSafeArrElemError> for ElementError {
    /// Holds a [`IntoSafeArrElemError`] with a box. This means conversion is free.
    fn from(isaee: IntoSafeArrElemError) -> ElementError {
        ElementError::Into(Box::new(isaee))
    }
}

impl From<IntoVariantError> for ElementError {
    /// Uses From impls on [`IntoSafeArrElemError] and [`ElementError`] to convert the error.
    fn from(ive: IntoVariantError) -> Self {
        ElementError::from(IntoSafeArrElemError::from(ive))
    }
}

impl From<FromVariantError> for ElementError {
    /// Uses From impls on [`FromSafeArrElemError] and [`ElementError`] to convert the error.
    fn from(fve: FromVariantError) -> Self {
        ElementError::from(FromSafeArrElemError::from(fve))
    }
}

/// Errors for converting from C/C++ data structure to Rust types
#[derive(Debug, Error)]
pub enum FromSafeArrElemError {
    /// The unsafe call to SafeArrayGetElement failed - HRESULT stored within tells why
    #[error("SafeArrayGetElement failed with HRESULT=0x{hr:x}")]
    GetElementFailed {
        /// Holds an HRESULT value
        hr: i32,
    },
    /// Holds a [`BStringError`] in a box.
    #[error(transparent)]
    BStringFailed(Box<BStringError>),
    /// [`from_variant`] failed somehow. Error is stored in a box.
    #[error("from variant failure")]
    FromVarError(#[source] Box<FromVariantError>),
}

impl From<FromVariantError> for FromSafeArrElemError {
    /// Boxes a [`FromVariantError`] into a [`FromSafeArrElemError`] which means the conversion is free.
    fn from(fve: FromVariantError) -> FromSafeArrElemError {
        FromSafeArrElemError::FromVarError(Box::new(fve))
    }
}

/// Errors for converting into C/C++ data structures from Rust types
#[derive(Debug, Error)]
pub enum IntoSafeArrElemError {
    /// `SysAllocStringLen` failed with len
    #[error(transparent)]
    BStringFailed(Box<BStringError>),
    /// `SafeArrayPutElement` failed with `HRESULT`
    #[error("SafeArrayPutElement failed with HRESULT = 0x{hr:x}")]
    PutElementFailed {
        /// HRESULT returned by SafeArrayPutElement call
        hr: i32,
    },
    /// Encapsulates a `IntoVariantError`
    #[error("IntoVariantError")]
    IntoVariantError(#[source] Box<IntoVariantError>),
}

impl From<IntoVariantError> for IntoSafeArrElemError {
    /// Boxes an [`IntoVariantError`] into an [`IntoSafeArrElemError`]
    fn from(ive: IntoVariantError) -> IntoSafeArrElemError {
        IntoSafeArrElemError::IntoVariantError(Box::new(ive))
    }
}

impl From<BStringError> for IntoSafeArrElemError {
    /// Boxes a [`BStringError`] into an [`IntoSafeArrElemError`]. This means the conversion is free.
    fn from(bse: BStringError) -> IntoSafeArrElemError {
        IntoSafeArrElemError::BStringFailed(Box::new(bse))
    }
}

/// Supererror for SafeArray errors
#[derive(Debug, Error)]
pub enum SafeArrayError {
    /// From wrapper for `FromSafeArrayError`
    #[error(transparent)]
    From(Box<FromSafeArrayError>),
    /// Into wrapper for `IntoSafeArrayError`
    #[error(transparent)]
    Into(Box<IntoSafeArrayError>),
}

impl From<FromSafeArrayError> for SafeArrayError {
    /// Holds a [`FromSafeArrayError`] with a box. This means conversion is free.
    fn from(fsae: FromSafeArrayError) -> SafeArrayError {
        SafeArrayError::From(Box::new(fsae))
    }
}

impl From<IntoSafeArrayError> for SafeArrayError {
    /// Holds a [`IntoSafeArrayError`] with a box. This means conversion is free.
    fn from(isae: IntoSafeArrayError) -> SafeArrayError {
        SafeArrayError::Into(Box::new(isae))
    }
}

impl From<FromVariantError> for SafeArrayError {
    /// Uses From impls on [`FromSafeArrayError] and [`SafeArrayError`] to convert the error.
    fn from(fve: FromVariantError) -> Self {
        SafeArrayError::from(FromSafeArrayError::from(fve))
    }
}

/// Represents the different ways converting from `SAFEARRAY` can fail
#[derive(Debug, Error)]
pub enum FromSafeArrayError {
    /// Either the safe array dimensions = 0 or > 1
    /// multi-dimensional arrays are *not* handled.
    #[error("Safe array dimensions are invalid: {}", sa_dims)]
    SafeArrayDimsInvalid {
        /// safe array dimensions that was wrong
        sa_dims: u32,
    },
    /// Expected vartype did not match found vartype - runtime consistency check
    #[error("expected vartype was not found - expected: {expected} - found: {found}")]
    VarTypeDoesNotMatch {
        /// The expected vartype
        expected: u32,
        /// the found vartype
        found: u32,
    },
    /// Call to SafeArrayGetLBound failed
    #[error("SafeArrayGetLBound failed with HRESULT = 0x{:x}", hr)]
    SafeArrayLBoundFailed {
        /// HRESULT returned
        hr: i32,
    },
    /// Call to SafeArrayGetRBound failed
    #[error("SafeArrayGetRBound failed with HRESULT = 0x{:x}", hr)]
    SafeArrayRBoundFailed {
        /// HRESULT returned
        hr: i32,
    },
    /// Call to SafeArrayGetVartype failed
    #[error("SafeArrayGetVartype failed with HRESULT = 0x{:x}", hr)]
    SafeArrayGetVartypeFailed {
        /// HRESULT returned
        hr: i32,
    },
    /// Encapsulates the `ElementError` that occurred during conversion
    #[error("element conversion failed at index {index} with {element}")]
    ElementConversionFailed {
        /// the index the conversion failed at
        index: usize,
        /// The element error encapsulating the failure
        #[source]
        element: Box<ElementError>,
    },
    /// [`from_variant`] call failed
    #[error("from variant failure: {}", _0)]
    FromVariantError(Box<FromVariantError>),
}

impl From<FromVariantError> for FromSafeArrayError {
    /// Boxes a [`FromVariantError`] into a [`FromSafeArrElemError`] which means the conversion is free.
    fn from(fve: FromVariantError) -> Self {
        FromSafeArrayError::FromVariantError(Box::new(fve))
    }
}

impl FromSafeArrayError {
    /// Boxes an [`ElementError`] into a [`FromSafeArrayError`].
    ///
    /// Need the index so a From impl doesn't apply.
    pub fn from_element_err<E: Into<ElementError>>(ee: E, index: usize) -> FromSafeArrayError {
        FromSafeArrayError::ElementConversionFailed {
            index: index,
            element: Box::new(ee.into()),
        }
    }
}

/// Represents the different ways converting into `SAFEARRAY` can fail
#[derive(Debug, Error)]
pub enum IntoSafeArrayError {
    /// Encapsulates the [`ElementError`] that occurred during conversion
    #[error("element conversion failed at index {index} with {element}")]
    ElementConversionFailed {
        /// the index the conversion failed at
        index: usize,
        /// The element error encapsulating the failure
        #[source]
        element: Box<ElementError>,
    },
    /// into_variant call failed
    #[error("into variant failure")]
    IntoVariantError(#[source] Box<IntoVariantError>),
}

impl From<IntoVariantError> for IntoSafeArrayError {
    /// Boxes a [`FromVariantError`] into a [`FromSafeArrElemError`] which means the conversion is free.
    fn from(ive: IntoVariantError) -> Self {
        IntoSafeArrayError::IntoVariantError(Box::new(ive))
    }
}

impl IntoSafeArrayError {
    /// Boxes an [`ElementError`] into a [`IntoSafeArrayError`].
    ///
    /// Need the index so a From impl doesn't apply.
    pub fn from_element_err<E: Into<ElementError>>(ee: E, index: usize) -> IntoSafeArrayError {
        IntoSafeArrayError::ElementConversionFailed {
            index: index,
            element: Box::new(ee.into()),
        }
    }
}

/// Ways BString can fail. Currently just one way.
#[derive(Clone, Copy, Debug, Error)]
pub enum BStringError {
    /// SysAllocStringLen failed
    #[error("BSTR allocation failed for len: {len}")]
    AllocateFailed {
        /// The len which was used for allocation
        len: usize,
    },
}

/// Encapsulates the ways converting from a `VARIANT` can fail.
#[derive(Copy, Clone, Debug, Error)]
pub enum FromVariantError {
    /// `VARIANT` pointer during conversion was null
    #[error("VARIANT pointer is null")]
    VariantPtrNull,
    /// Unknown VT for
    #[error("Variants does not support this vartype: {0:p}")]
    UnknownVarType(u16),
}

/// Encapsulates errors that can occur during conversion into VARIANT
#[derive(Debug, Error)]
pub enum IntoVariantError {
    /// Encapsulates a `BStringError`
    #[error(transparent)]
    AllocBStrFailed(Box<BStringError>),
    /// Encapsulates a `SafeArrayError`
    #[error("SafeArray conversion failed")]
    SafeArrConvFailed(#[source] Box<SafeArrayError>),
    ///
    #[error("Can't convert &dyn CVariantWrappers into Ptr<VARIANTS>")]
    CVarWrapper,
}

impl From<BStringError> for IntoVariantError {
    /// Boxes a [`BStringError`] into a [`FromSafeArrElemError`]. This means the conversion is free.
    fn from(bse: BStringError) -> IntoVariantError {
        IntoVariantError::AllocBStrFailed(Box::new(bse))
    }
}

/// Errors which can arise primarily from using `Conversion::convert` calls
#[derive(Debug, Error)]
pub enum ConversionError {
    /// Ptr being used was null
    #[error("pointer was null")]
    PtrWasNull,
    /// General purpose holder of `std::error::Error` values
    #[error(transparent)]
    General(Box<dyn Error>),
}
