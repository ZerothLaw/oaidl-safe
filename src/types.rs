
//! # Types
//! Convenience wrapper types and conversion logic for the winapi-defined types: 
//!   * CY
//!   * DATE
//!   * DECIMAL
//! 

use rust_decimal::Decimal;

use winapi::shared::wtypes::{CY, DATE, DECIMAL, DECIMAL_NEG, VARIANT_BOOL, VARIANT_TRUE};

#[derive(Clone, Copy, Debug, Eq,  Hash, PartialOrd, PartialEq)]
pub struct Currency(pub i64);

impl From<i64> for Currency {
    fn from(i: i64) -> Currency {
        Currency(i)
    }
}
impl<'i> From<&'i i64> for Currency {
    fn from(i: &i64) -> Currency {
        Currency(*i)
    }
}
impl<'i> From<&'i mut i64> for Currency {
    fn from(i: &mut i64) -> Currency {
        Currency(*i)
    }
}

impl From<CY> for Currency {
    fn from(cy: CY) -> Currency {
        Currency(cy.int64)
    }
}
impl<'c> From<&'c CY> for Currency {
    fn from(cy: &CY) -> Currency {
        Currency(cy.int64)
    }
}
impl<'c> From<&'c mut CY> for Currency {
    fn from(cy: &mut CY) -> Currency {
        Currency(cy.int64)
    }
}

impl From<Currency> for CY {
    fn from(cy: Currency) -> CY {
        CY {int64: cy.0}
    }
}
impl<'c> From<&'c Currency> for CY {
    fn from(cy: &Currency) -> CY {
        CY {int64: cy.0}
    }
}
impl<'c> From<&'c mut Currency> for CY {
    fn from(cy: &mut Currency) -> CY {
        CY {int64: cy.0}
    }
}

#[derive(Debug, Clone, Copy, PartialOrd, PartialEq)]
pub struct Date(pub f64); //DATE <--> F64

impl From<DATE> for Date {
    fn from(f: DATE) -> Date {
        Date(f)
    }
}
impl<'f> From<&'f DATE> for Date {
    fn from(f: &DATE) -> Date {
        Date(*f)
    }
}
impl<'f> From<&'f mut DATE> for Date {
    fn from(f: &mut DATE) -> Date {
        Date(*f)
    }
}

impl From<Date> for DATE {
    fn from(d: Date) -> DATE {
        d.0
    }
}

impl<'f> From<&'f Date> for DATE {
    fn from(d: &Date) -> DATE {
        d.0
    }
}

impl<'f> From<&'f mut Date> for DATE {
    fn from(d: &mut Date) -> DATE {
        d.0
    }
}

#[derive(Debug, Clone, Copy, Eq, Hash, Ord, PartialEq, PartialOrd)]
pub struct DecWrapper(Decimal);

impl DecWrapper {
    pub fn new(dec: Decimal) -> DecWrapper {
        DecWrapper(dec)
    }

    pub fn unwrap(self) -> Decimal {
        self.0
    }

    pub fn borrow(&self) -> &Decimal {
        &self.0
    }

    pub fn borrow_mut(&mut self) -> &mut Decimal {
        &mut self.0
    }

    fn build_c_decimal(dec: Decimal) -> DECIMAL {
        let scale = dec.scale() as u8;
        let sign = if dec.is_sign_positive() {0} else {DECIMAL_NEG};
        let serial = dec.serialize();
        let lo: u64 = (serial[4]  as u64)        + 
                    ((serial[5]  as u64) << 8)  + 
                    ((serial[6]  as u64) << 16) + 
                    ((serial[7]  as u64) << 24) + 
                    ((serial[8]  as u64) << 32) +
                    ((serial[9]  as u64) << 40) +
                    ((serial[10] as u64) << 48) + 
                    ((serial[11] as u64) << 56);
        let hi: u32 = (serial[12] as u32)        +
                    ((serial[13] as u32) << 8)  +
                    ((serial[14] as u32) << 16) +
                    ((serial[15] as u32) << 24);
        DECIMAL {
            wReserved: 0, 
            scale: scale, 
            sign: sign, 
            Hi32: hi, 
            Lo64: lo
        }
    }

    fn build_rust_decimal(dec: DECIMAL) -> Decimal {
        let sign = if dec.sign == DECIMAL_NEG {true} else {false};
        Decimal::from_parts((dec.Lo64 & 0xFFFFFFFF) as u32, 
                            ((dec.Lo64 >> 32) & 0xFFFFFFFF) as u32, 
                            dec.Hi32, 
                            sign,
                            dec.scale as u32 ) 
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
    fn from(d: DECIMAL) -> DecWrapper {
        DecWrapper(DecWrapper::build_rust_decimal(d))
    }
}
impl<'d> From<&'d DECIMAL> for DecWrapper {
    fn from(d: &DECIMAL) -> DecWrapper {
        DecWrapper(DecWrapper::build_rust_decimal(d.clone()))
    }
}
impl<'d> From<&'d mut DECIMAL> for DecWrapper {
    fn from(d: &mut DECIMAL) -> DecWrapper {
        DecWrapper(DecWrapper::build_rust_decimal(d.clone()))
    }
}

//DecWrapper to DECIMAL conversions
impl From<DecWrapper> for DECIMAL {
    fn from(d: DecWrapper) -> DECIMAL {
        DecWrapper::build_c_decimal(d.0)
    }
}
impl<'d> From<&'d DecWrapper> for DECIMAL {
    fn from(d: &DecWrapper) -> DECIMAL {
        DecWrapper::build_c_decimal(d.0)
    }
}
impl<'d> From<&'d mut DecWrapper> for DECIMAL {
    fn from(d: & mut DecWrapper) -> DECIMAL {
        DecWrapper::build_c_decimal(d.0)
    }
}

//DecWrapper to Decimal conversions
impl From<DecWrapper> for Decimal {
    fn from(dw: DecWrapper) -> Decimal {
        dw.0
    }
}
impl<'w> From<&'w DecWrapper> for Decimal {
    fn from(dw: &DecWrapper) -> Decimal {
        dw.0
    }
}
impl<'w> From<&'w mut DecWrapper> for Decimal {
    fn from(dw: &mut DecWrapper) -> Decimal {
        dw.0
    }
}

//Decimal to DecWrapper conversions
impl From<Decimal> for DecWrapper {
    fn from(dec: Decimal) -> DecWrapper {
        DecWrapper(dec)
    }
}
impl<'d> From<&'d Decimal> for DecWrapper {
    fn from(dec: &Decimal) -> DecWrapper {
        DecWrapper(dec.clone())
    }
}
impl<'d> From<&'d mut Decimal> for DecWrapper {
    fn from(dec: &mut Decimal) -> DecWrapper {
        DecWrapper(dec.clone())
    }
}

#[derive(Debug, Clone, Copy, Eq, Hash, Ord, PartialEq, PartialOrd)]
pub struct VariantBool(bool);

impl From<VariantBool> for VARIANT_BOOL {
    fn from(vb: VariantBool) -> VARIANT_BOOL {
        if vb.0 {VARIANT_TRUE} else {0}
    }
}
impl<'v> From<&'v VariantBool> for VARIANT_BOOL {
    fn from(vb: &VariantBool) -> VARIANT_BOOL {
        if vb.0 {VARIANT_TRUE} else {0}
    }
}
impl<'v> From<&'v mut VariantBool> for VARIANT_BOOL {
    fn from(vb: &mut VariantBool) -> VARIANT_BOOL {
        if vb.0 {VARIANT_TRUE} else {0}
    }
}

impl From<VARIANT_BOOL> for VariantBool {
    fn from(vb: VARIANT_BOOL) -> VariantBool {
        VariantBool(vb < 0) 
    }
}
impl<'v> From<&'v VARIANT_BOOL> for VariantBool {
    fn from(vb: &VARIANT_BOOL) -> VariantBool {
        VariantBool(*vb < 0) 
    }
}
impl<'v> From<&'v mut VARIANT_BOOL> for VariantBool {
    fn from(vb: &mut VARIANT_BOOL) -> VariantBool {
        VariantBool(*vb < 0) 
    }
}

impl From<bool> for VariantBool {
    fn from(b: bool) -> Self {
        VariantBool(b)
    }
}

impl<'b> From<&'b bool> for VariantBool {
    fn from(b: &bool) -> Self {
        VariantBool(*b)
    }
}

impl<'b> From<&'b mut bool> for VariantBool {
    fn from(b: &mut bool) -> Self {
        VariantBool(*b)
    }
}

impl From<VariantBool> for bool {
    fn from(b: VariantBool) -> Self {
        b.0
    }
}
impl<'v> From<&'v VariantBool> for bool {
    fn from(b: &VariantBool) -> Self {
        b.0
    }
}
impl<'v> From<&'v mut VariantBool> for bool {
    fn from(b: &mut VariantBool) -> Self {
        b.0
    }
}

#[derive(Debug, Clone, Copy, Eq, Hash, Ord, PartialEq, PartialOrd)]
pub struct Int(pub i32);
#[derive(Debug, Clone, Copy, Eq, Hash, Ord, PartialEq, PartialOrd)]
pub struct UInt(pub u32);
#[derive(Debug, Clone, Copy, Eq, Hash, Ord, PartialEq, PartialOrd)]
pub struct SCode(pub i32);

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn c_decimal() {
        let d = Decimal::new(0xFFFFFFFFFFFF, 0);
        let d = d * Decimal::new(0xFFFFFFFF, 0);
        assert_eq!(d.is_sign_positive(), true);
        assert_eq!(format!("{}", d), "1208925819333149903028225" );
        
        let c = DecWrapper::build_c_decimal(d);
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
            Lo64: 18446462594437873665
        };
        let new_d = DecWrapper::build_rust_decimal(d);
        //println!("{:?}", new_d.serialize());
       // assert_eq!(new_d.is_sign_positive(), true);
        assert_eq!(format!("{}", new_d), "1208925819333149903028225"  );
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
}
