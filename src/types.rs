//! # Types
//! Convenience wrapper types and conversion logic for the winapi-defined types:
//!   * CY
//!   * DATE
//!   * DECIMAL
//!   * VARIANT_BOOL
//!
//! Also implements TryFrom when feature "impl_tryfrom" is turned on in configuration. Only works on Nightly!
//!
use std::fmt;

#[cfg(feature = "impl_tryfrom")]
use std::convert::TryFrom;

#[cfg(feature = "impl_tryfrom")]
use std::num::TryFromIntError;

use std::error::Error;

use rust_decimal::Decimal;

use winapi::shared::wtypes::{CY, DECIMAL, DECIMAL_NEG, VARIANT_BOOL, VARIANT_TRUE};

use super::errors::{
    ConversionError, ElementError, FromVariantError, IntoVariantError, SafeArrayError,
};

/// Pseudo-`From` trait because of orphan rules
pub trait TryConvert<T, F>
where
    Self: Sized,
    F: Error,
{
    /// Utility method which can fail.
    fn try_convert(val: T) -> Result<Self, F>;
}

impl<T, F> TryConvert<T, F> for T
where
    T: From<T>,
    F: Error,
{
    /// Blanket TryConvert implementation wherever a From<T> is implemented for T. (Which is all types.)
    /// This avoids repetitive code. The compiler monomorphizes the code for F.
    /// And because its always an Ok, should optimize this code away.
    fn try_convert(val: T) -> Result<Self, F> {
        Ok(val)
    }
}

macro_rules! impl_conv_for_box_wrapper {
    ($(#[$attrs:meta])* $inner:ident, $wrapper:ident) => {
        impl TryConvert<Box<$wrapper>,IntoVariantError> for *mut $inner {
            $(#[$attrs])*
            fn try_convert(b: Box<$wrapper>) -> Result<Self,IntoVariantError> {
                let b = *b;
                let inner = $inner::from(b);
                Ok(Box::into_raw(Box::new(inner)))
            }
        }

        impl TryConvert<*mut $inner,FromVariantError> for Box<$wrapper> {
            $(#[$attrs])*
            fn try_convert(inner: *mut $inner) -> Result<Self,FromVariantError> {
                if inner.is_null() {
                    return Err(FromVariantError::VariantPtrNull);
                }
                let inner = unsafe {*inner};
                let wrapper = $wrapper::from(inner);
                Ok(Box::new(wrapper))
            }
        }
    };
}

macro_rules! wrapper_conv_impl {
    ($inner:ident, $wrapper:ident) => {
        impl From<$inner> for $wrapper {
            fn from(i: $inner) -> Self {
                $wrapper(i)
            }
        }

        impl<'f> From<&'f $inner> for $wrapper {
            fn from(i: &'f $inner) -> Self {
                $wrapper(*i)
            }
        }

        impl<'f> From<&'f mut $inner> for $wrapper {
            fn from(i: &'f mut $inner) -> Self {
                $wrapper(*i)
            }
        }

        impl From<$wrapper> for $inner {
            fn from(o: $wrapper) -> Self {
                o.0
            }
        }

        impl<'f> From<&'f $wrapper> for $inner {
            fn from(o: &'f $wrapper) -> Self {
                o.0
            }
        }

        impl<'f> From<&'f mut $wrapper> for $inner {
            fn from(o: &'f mut $wrapper) -> Self {
                o.0
            }
        }

        conversions_impl!($inner, $wrapper);
        impl_conv_for_box_wrapper!($inner, $wrapper);
    };
}

macro_rules! conversions_impl {
    ($inner:ident, $wrapper:ident) => {
        impl TryConvert<$inner, FromVariantError> for $wrapper {
            /// Does not return any errors.
            fn try_convert(val: $inner) -> Result<Self, FromVariantError> {
                Ok($wrapper::from(val))
            }
        }

        impl TryConvert<$wrapper, IntoVariantError> for $inner {
            /// Does not return any errors.
            fn try_convert(val: $wrapper) -> Result<Self, IntoVariantError> {
                Ok($inner::from(val))
            }
        }

        impl<'c> TryConvert<&'c $inner, FromVariantError> for $wrapper {
            /// Does not return any errors.
            fn try_convert(val: &'c $inner) -> Result<Self, FromVariantError> {
                Ok($wrapper::from(val))
            }
        }

        impl<'c> TryConvert<&'c $wrapper, IntoVariantError> for $inner {
            /// Does not return any errors.
            fn try_convert(val: &'c $wrapper) -> Result<Self, IntoVariantError> {
                Ok($inner::from(val))
            }
        }

        impl<'c> TryConvert<&'c mut $inner, FromVariantError> for $wrapper {
            /// Does not return any errors.
            fn try_convert(val: &'c mut $inner) -> Result<Self, FromVariantError> {
                Ok($wrapper::from(val))
            }
        }

        impl<'c> TryConvert<&'c mut $wrapper, IntoVariantError> for $inner {
            /// Does not return any errors.
            fn try_convert(val: &'c mut $wrapper) -> Result<Self, IntoVariantError> {
                Ok($inner::from(val))
            }
        }

        impl TryConvert<$inner, SafeArrayError> for $wrapper {
            /// Does not return any errors.
            fn try_convert(val: $inner) -> Result<Self, SafeArrayError> {
                Ok($wrapper::from(val))
            }
        }
        impl TryConvert<$wrapper, SafeArrayError> for $inner {
            /// Does not return any errors.
            fn try_convert(val: $wrapper) -> Result<Self, SafeArrayError> {
                Ok($inner::from(val))
            }
        }

        impl TryConvert<$inner, ElementError> for $wrapper {
            /// Does not return any errors.
            fn try_convert(val: $inner) -> Result<Self, ElementError> {
                Ok($wrapper::from(val))
            }
        }
        impl TryConvert<$wrapper, ElementError> for $inner {
            /// Does not return any errors.
            fn try_convert(val: $wrapper) -> Result<Self, ElementError> {
                Ok($inner::from(val))
            }
        }
    };
}

/// Helper type for the OLE/COM+ type CY
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Clone, Copy, Debug, Eq, Hash, PartialOrd, PartialEq)]
pub struct Currency(i64);

impl From<CY> for Currency {
    /// Converts CY to Currency. Consumes the CY, and copies the internal value.
    fn from(cy: CY) -> Currency {
        Currency(cy.int64)
    }
}
impl<'c> From<&'c CY> for Currency {
    /// Converts CY to Currency. Copies the internal value from the reference.
    fn from(cy: &'c CY) -> Currency {
        Currency(cy.int64)
    }
}
impl<'c> From<&'c mut CY> for Currency {
    /// Converts CY to Currency. Copies the internal value from the reference.
    fn from(cy: &'c mut CY) -> Currency {
        Currency(cy.int64)
    }
}

impl From<Currency> for CY {
    /// Converts Currency to CY. Consumes the Currency and copies the internal value to CY.
    fn from(cy: Currency) -> CY {
        CY { int64: cy.0 }
    }
}
impl<'c> From<&'c Currency> for CY {
    /// Converts Currency to CY. Copies the internal value from the reference.
    fn from(cy: &'c Currency) -> CY {
        CY { int64: cy.0 }
    }
}
impl<'c> From<&'c mut Currency> for CY {
    /// Converts Currency to CY. Copies the internal value from the reference.
    fn from(cy: &'c mut Currency) -> CY {
        CY { int64: cy.0 }
    }
}

impl From<i64> for Currency {
    fn from(i: i64) -> Self {
        Currency(i)
    }
}

impl AsRef<i64> for Currency {
    fn as_ref(&self) -> &i64 {
        &self.0
    }
}
conversions_impl!(CY, Currency);
impl_conv_for_box_wrapper!(CY, Currency);

/// Helper type for the OLE/COM+ type DATE
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, PartialOrd, PartialEq)]
pub struct Date(f64); //DATE <--> F64

impl AsRef<f64> for Date {
    fn as_ref(&self) -> &f64 {
        &self.0
    }
}

wrapper_conv_impl!(f64, Date);

/// Helper type for the OLE/COM+ type DECIMAL
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, Eq, Hash, Ord, PartialEq, PartialOrd)]
pub struct DecWrapper(Decimal);

impl DecWrapper {
    /// wraps a `Decimal` from rust_decimal
    pub fn new(dec: Decimal) -> DecWrapper {
        DecWrapper(dec)
    }

    /// Get access to the internal value, consuming it in the process
    pub fn unwrap(self) -> Decimal {
        self.0
    }

    /// Get borrow of internal value
    pub fn borrow(&self) -> &Decimal {
        &self.0
    }

    /// Get mutable borrow of internal value
    pub fn borrow_mut(&mut self) -> &mut Decimal {
        &mut self.0
    }
    // Internal conversion function
    fn build_c_decimal(dec: &Decimal) -> DECIMAL {
        let scale = dec.scale() as u8;
        let sign = if dec.is_sign_positive() {
            0
        } else {
            DECIMAL_NEG
        };
        let serial = dec.serialize();
        let lo: u64 = (serial[4] as u64)
            + ((serial[5] as u64) << 8)
            + ((serial[6] as u64) << 16)
            + ((serial[7] as u64) << 24)
            + ((serial[8] as u64) << 32)
            + ((serial[9] as u64) << 40)
            + ((serial[10] as u64) << 48)
            + ((serial[11] as u64) << 56);
        let hi: u32 = (serial[12] as u32)
            + ((serial[13] as u32) << 8)
            + ((serial[14] as u32) << 16)
            + ((serial[15] as u32) << 24);
        DECIMAL {
            wReserved: 0,
            scale: scale,
            sign: sign,
            Hi32: hi,
            Lo64: lo,
        }
    }

    fn build_rust_decimal(dec: &DECIMAL) -> Decimal {
        let sign = if dec.sign == DECIMAL_NEG { true } else { false };
        Decimal::from_parts(
            (dec.Lo64 & 0xFFFFFFFF) as u32,
            ((dec.Lo64 >> 32) & 0xFFFFFFFF) as u32,
            dec.Hi32,
            sign,
            dec.scale as u32,
        )
    }
}
//Conversions between triad of types:
//                    DECIMAL    |   DecWrapper   |   Decimal
//   DECIMAL    |      N/A       | owned, &, &mut | orphan rules
//   DecWrapper | owned, &, &mut |     N/A        | owned, &, &mut
//   Decimal    | orphan rules   | owned, &, &mut |     N/A
//
// Ophan rules mean that I can't apply traits from other crates
// to types that come from still other traits.

//DECIMAL to DecWrapper conversions
impl From<DECIMAL> for DecWrapper {
    /// Converts DECIMAL into a Decimal wrapped in a DecWrapper.
    /// Allocates a new Decimal.
    fn from(d: DECIMAL) -> DecWrapper {
        DecWrapper(DecWrapper::build_rust_decimal(&d))
    }
}
impl<'d> From<&'d DECIMAL> for DecWrapper {
    /// Converts DECIMAL into a Decimal wrapped in a DecWrapper
    /// Allocates a new Decimal.
    fn from(d: &'d DECIMAL) -> DecWrapper {
        DecWrapper(DecWrapper::build_rust_decimal(d))
    }
}
impl<'d> From<&'d mut DECIMAL> for DecWrapper {
    /// Converts DECIMAL into a Decimal wrapped in a DecWrapper
    /// Allocates a new Decimal.
    fn from(d: &'d mut DECIMAL) -> DecWrapper {
        DecWrapper(DecWrapper::build_rust_decimal(d))
    }
}

//DecWrapper to DECIMAL conversions
impl From<DecWrapper> for DECIMAL {
    /// Converts a DecWrapper into  DECIMAL.
    /// Allocates a new DECIMAL
    fn from(d: DecWrapper) -> DECIMAL {
        DecWrapper::build_c_decimal(&d.0)
    }
}
impl<'d> From<&'d DecWrapper> for DECIMAL {
    /// Converts a DecWrapper into  DECIMAL.
    /// Allocates a new DECIMAL
    fn from(d: &'d DecWrapper) -> DECIMAL {
        DecWrapper::build_c_decimal(&d.0)
    }
}
impl<'d> From<&'d mut DecWrapper> for DECIMAL {
    /// Converts a DecWrapper into  DECIMAL.
    /// Allocates a new DECIMAL
    fn from(d: &'d mut DecWrapper) -> DECIMAL {
        DecWrapper::build_c_decimal(&d.0)
    }
}

//DecWrapper to Decimal conversions
impl From<DecWrapper> for Decimal {
    /// Converts a DecWrapper into Decimal.
    /// Zero cost or allocations.
    fn from(dw: DecWrapper) -> Decimal {
        dw.0
    }
}
impl<'w> From<&'w DecWrapper> for Decimal {
    /// Converts a DecWrapper into Decimal.
    /// Zero cost or allocations.
    fn from(dw: &'w DecWrapper) -> Decimal {
        dw.0
    }
}
impl<'w> From<&'w mut DecWrapper> for Decimal {
    /// Converts a DecWrapper into Decimal.
    /// Zero cost or allocations.
    fn from(dw: &'w mut DecWrapper) -> Decimal {
        dw.0
    }
}

//Decimal to DecWrapper conversions
impl From<Decimal> for DecWrapper {
    /// Converts a Decimal into a DecWrapper.
    /// Zero cost or allocations.
    fn from(dec: Decimal) -> DecWrapper {
        DecWrapper(dec)
    }
}
impl<'d> From<&'d Decimal> for DecWrapper {
    /// Converts a Decimal into a DecWrapper.
    /// Zero cost or allocations.
    fn from(dec: &'d Decimal) -> DecWrapper {
        DecWrapper(dec.clone())
    }
}
impl<'d> From<&'d mut Decimal> for DecWrapper {
    /// Converts a Decimal into a DecWrapper.
    /// Zero cost or allocations.
    fn from(dec: &'d mut Decimal) -> DecWrapper {
        DecWrapper(dec.clone())
    }
}

impl AsRef<Decimal> for DecWrapper {
    fn as_ref(&self) -> &Decimal {
        &self.0
    }
}
conversions_impl!(DECIMAL, DecWrapper);

/// Helper type for the OLE/COM+ type VARIANT_BOOL
///
/// A VARIANT_Bool represents true as 0x0, and false as 0xFFFF.
///
/// This means there's a bit of conversion logic required.
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, Eq, Hash, Ord, PartialEq, PartialOrd)]
pub struct VariantBool(bool);

impl From<VariantBool> for VARIANT_BOOL {
    /// Converts a VariantBool into a VARIANT_BOOL.
    fn from(vb: VariantBool) -> VARIANT_BOOL {
        if vb.0 {
            VARIANT_TRUE
        } else {
            0
        }
    }
}
impl<'v> From<&'v VariantBool> for VARIANT_BOOL {
    /// Converts a VariantBool into a VARIANT_BOOL.
    fn from(vb: &'v VariantBool) -> VARIANT_BOOL {
        if vb.0 {
            VARIANT_TRUE
        } else {
            0
        }
    }
}
impl<'v> From<&'v mut VariantBool> for VARIANT_BOOL {
    /// Converts a VariantBool into a VARIANT_BOOL.
    fn from(vb: &'v mut VariantBool) -> VARIANT_BOOL {
        if vb.0 {
            VARIANT_TRUE
        } else {
            0
        }
    }
}

impl From<VARIANT_BOOL> for VariantBool {
    /// Converts a VARIANT_BOOL into a VariantBool.
    fn from(vb: VARIANT_BOOL) -> VariantBool {
        VariantBool(vb < 0)
    }
}
impl<'v> From<&'v VARIANT_BOOL> for VariantBool {
    /// Converts a VARIANT_BOOL into a VariantBool.
    fn from(vb: &'v VARIANT_BOOL) -> VariantBool {
        VariantBool(*vb < 0)
    }
}
impl<'v> From<&'v mut VARIANT_BOOL> for VariantBool {
    /// Converts a VARIANT_BOOL into a VariantBool.
    fn from(vb: &'v mut VARIANT_BOOL) -> VariantBool {
        VariantBool(*vb < 0)
    }
}

impl From<bool> for VariantBool {
    fn from(b: bool) -> Self {
        VariantBool(b)
    }
}

impl<'b> From<&'b bool> for VariantBool {
    fn from(b: &'b bool) -> Self {
        VariantBool(*b)
    }
}

impl<'b> From<&'b mut bool> for VariantBool {
    fn from(b: &'b mut bool) -> Self {
        VariantBool(*b)
    }
}

impl From<VariantBool> for bool {
    fn from(b: VariantBool) -> Self {
        b.0
    }
}
impl<'v> From<&'v VariantBool> for bool {
    fn from(b: &'v VariantBool) -> Self {
        b.0
    }
}
impl<'v> From<&'v mut VariantBool> for bool {
    fn from(b: &'v mut VariantBool) -> Self {
        b.0
    }
}

impl AsRef<bool> for VariantBool {
    fn as_ref(&self) -> &bool {
        &self.0
    }
}

conversions_impl!(bool, VariantBool);
conversions_impl!(VARIANT_BOOL, VariantBool);
impl_conv_for_box_wrapper!(VARIANT_BOOL, VariantBool);

impl TryConvert<*mut VARIANT_BOOL, ConversionError> for Box<VariantBool> {
    fn try_convert(p: *mut VARIANT_BOOL) -> Result<Self, ConversionError> {
        if p.is_null() {
            return Err(ConversionError::PtrWasNull);
        }
        Ok(Box::new(VariantBool::from(unsafe { *p })))
    }
}

/// Helper type for the OLE/COM+ type INT
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, Eq, Hash, Ord, PartialEq, PartialOrd)]
pub struct Int(i32);

impl AsRef<i32> for Int {
    fn as_ref(&self) -> &i32 {
        &self.0
    }
}

impl fmt::UpperHex for Int {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{:X}", self.0)
    }
}

impl fmt::LowerHex for Int {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{:x}", self.0)
    }
}

impl fmt::Octal for Int {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{:o}", self.0)
    }
}

impl fmt::Binary for Int {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{:b}", self.0)
    }
}

wrapper_conv_impl!(i32, Int);

#[cfg(feature = "impl_tryfrom")]
impl TryFrom<i64> for Int {
    type Error = TryFromIntError;
    fn try_from(value: i64) -> Result<Self, Self::Error> {
        Ok(Int(i32::try_from(value)?))
    }
}

#[cfg(feature = "impl_tryfrom")]
impl TryFrom<i128> for Int {
    type Error = TryFromIntError;
    fn try_from(value: i128) -> Result<Self, Self::Error> {
        Ok(Int(i32::try_from(value)?))
    }
}

#[cfg(feature = "impl_tryfrom")]
impl TryFrom<i16> for Int {
    type Error = TryFromIntError;
    fn try_from(value: i16) -> Result<Self, Self::Error> {
        Ok(Int(value as i32))
    }
}

#[cfg(feature = "impl_tryfrom")]
impl TryFrom<i8> for Int {
    type Error = TryFromIntError;
    fn try_from(value: i8) -> Result<Self, Self::Error> {
        Ok(Int(value as i32))
    }
}

/// Helper type for the OLE/COM+ type UINT
#[derive(Debug, Clone, Copy, Eq, Hash, Ord, PartialEq, PartialOrd)]
pub struct UInt(u32);

impl AsRef<u32> for UInt {
    fn as_ref(&self) -> &u32 {
        &self.0
    }
}

impl fmt::UpperHex for UInt {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{:X}", self.0)
    }
}

impl fmt::LowerHex for UInt {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{:x}", self.0)
    }
}

impl fmt::Octal for UInt {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{:o}", self.0)
    }
}

impl fmt::Binary for UInt {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{:b}", self.0)
    }
}

wrapper_conv_impl!(u32, UInt);

#[cfg(feature = "impl_tryfrom")]
impl TryFrom<u64> for UInt {
    type Error = TryFromIntError;
    fn try_from(value: u64) -> Result<Self, Self::Error> {
        Ok(UInt(u32::try_from(value)?))
    }
}

#[cfg(feature = "impl_tryfrom")]
impl TryFrom<u128> for UInt {
    type Error = TryFromIntError;
    fn try_from(value: u128) -> Result<Self, Self::Error> {
        Ok(UInt(u32::try_from(value)?))
    }
}

#[cfg(feature = "impl_tryfrom")]
impl TryFrom<u16> for UInt {
    type Error = TryFromIntError;
    fn try_from(value: u16) -> Result<Self, Self::Error> {
        Ok(UInt(value as u32))
    }
}

#[cfg(feature = "impl_tryfrom")]
impl TryFrom<u8> for UInt {
    type Error = TryFromIntError;
    fn try_from(value: u8) -> Result<Self, Self::Error> {
        Ok(UInt(value as u32))
    }
}

/// Helper type for the OLE/COM+ type SCODE
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Clone, Copy, Eq, Hash, Ord, PartialEq, PartialOrd)]
pub struct SCode(i32);

impl AsRef<i32> for SCode {
    fn as_ref(&self) -> &i32 {
        &self.0
    }
}

impl fmt::UpperHex for SCode {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "0X{:X}", self.0)
    }
}

impl fmt::LowerHex for SCode {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "0x{:x}", self.0)
    }
}

impl fmt::Octal for SCode {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{:o}", self.0)
    }
}

impl fmt::Binary for SCode {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{:b}", self.0)
    }
}
wrapper_conv_impl!(i32, SCode);

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn c_decimal() {
        let d = Decimal::new(0xFFFFFFFFFFFF, 0);
        let d = d * Decimal::new(0xFFFFFFFF, 0);
        assert_eq!(d.is_sign_positive(), true);
        assert_eq!(format!("{}", d), "1208925819333149903028225");

        let c = DecWrapper::build_c_decimal(&d);
        //println!("({}, {}, {}, {})", c.Hi32, c.Lo64, c.scale, c.sign);
        //println!("{:?}", d.serialize());
        assert_eq!(c.Hi32, 65535);
        assert_eq!(c.Lo64, 18446462594437873665);
        assert_eq!(c.scale, 0);
        assert_eq!(c.sign, 0);
    }

    #[test]
    fn rust_decimal_from() {
        let d = DECIMAL {
            wReserved: 0,
            scale: 0,
            sign: 0,
            Hi32: 65535,
            Lo64: 18446462594437873665,
        };
        let new_d = DecWrapper::build_rust_decimal(&d);
        //println!("{:?}", new_d.serialize());
        // assert_eq!(new_d.is_sign_positive(), true);
        assert_eq!(format!("{}", new_d), "1208925819333149903028225");
    }

    #[test]
    fn variant_bool() {
        let vb = VariantBool::from(true);
        let pvb = VARIANT_BOOL::from(vb);
        assert_eq!(VARIANT_TRUE, pvb);

        let vb = VariantBool::from(false);
        let pvb = VARIANT_BOOL::from(vb);
        assert_ne!(VARIANT_TRUE, pvb);
    }

    #[test]
    fn test_send() {
        fn assert_send<T: Send>() {}
        assert_send::<Currency>();
        assert_send::<Date>();
        assert_send::<DecWrapper>();
        assert_send::<Int>();
        assert_send::<SCode>();
        assert_send::<UInt>();
        assert_send::<VariantBool>();
    }

    #[test]
    fn test_sync() {
        fn assert_sync<T: Sync>() {}
        assert_sync::<Currency>();
        assert_sync::<Date>();
        assert_sync::<DecWrapper>();
        assert_sync::<Int>();
        assert_sync::<SCode>();
        assert_sync::<UInt>();
        assert_sync::<VariantBool>();
    }

    #[cfg(feature = "impl_tryfrom")]
    #[cfg_attr(feature = "impl_tryfrom", test)]
    fn test_tryfrom() {
        let v = Int::try_from(999999999999999i64);
        assert!(v.is_err());
    }
}
