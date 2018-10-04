
//! # Types
//! Convenience wrapper types and conversion logic for the winapi-defined types: 
//!   * CY
//!   * DATE
//!   * DECIMAL
//! 

use rust_decimal::Decimal;

use winapi::shared::wtypes::{CY, DATE, DECIMAL, DECIMAL_NEG, VARIANT_BOOL};

#[derive(Clone, Copy, Debug, Eq,  Hash, PartialOrd, PartialEq)]
pub struct Currency(pub i64);

impl From<i64> for Currency {
    fn from(i: i64) -> Currency {
        Currency(i)
    }
}

impl From<CY> for Currency {
    fn from(cy: CY) -> Currency {
        Currency(cy.int64)
    }
}

impl From<Currency> for CY {
    fn from(cy: Currency) -> CY {
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

impl From<Date> for DATE {
    fn from(d: Date) -> DATE {
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
}

impl From<DECIMAL> for DecWrapper {
    fn from(d: DECIMAL) -> DecWrapper {
        DecWrapper(build_rust_decimal(d))
    }
}

impl From<DecWrapper> for DECIMAL {
    fn from(d: DecWrapper) -> DECIMAL {
        build_c_decimal(d.0)
    }
}

impl<'d> From<&'d mut DecWrapper> for DECIMAL {
    fn from(d: & mut DecWrapper) -> DECIMAL {
        build_c_decimal(d.0)
    }
}

impl From<DecWrapper> for Decimal {
    fn from(dw: DecWrapper) -> Decimal {
        dw.0
    }
}

impl From<Decimal> for DecWrapper {
    fn from(dec: Decimal) -> DecWrapper {
        DecWrapper(dec)
    }
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

#[derive(Debug, Clone, Copy, Eq, Hash, Ord, PartialEq, PartialOrd)]
pub struct VariantBool(bool);

impl From<VariantBool> for VARIANT_BOOL {
    fn from(vb: VariantBool) -> VARIANT_BOOL {
        if vb.0 {-1} else {0}
    }
}

impl From<VARIANT_BOOL> for VariantBool {
    fn from(vb: VARIANT_BOOL) -> VariantBool {
        VariantBool(vb < 0) 
    }
}

impl From<bool> for VariantBool {
    fn from(b: bool) -> Self {
        VariantBool(b)
    }
}

impl From<VariantBool> for bool {
    fn from(b: VariantBool) -> Self {
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
        
        let c = build_c_decimal(d);
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
        let new_d = build_rust_decimal(d);
        //println!("{:?}", new_d.serialize());
       // assert_eq!(new_d.is_sign_positive(), true);
        assert_eq!(format!("{}", new_d), "1208925819333149903028225"  );
    }
}
