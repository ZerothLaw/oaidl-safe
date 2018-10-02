use winapi::um::oaidl::IDispatch;
use winapi::um::unknwnbase::IUnknown;

use ptr::Ptr;
use types::{Currency, DecWrapper, Int, SCode, UInt,};
use variant::Variant;

pub enum RSafeArrayItem {
    Short(i16),
    Long(i32), 
    LongLong(i64), 
    Float(f32), 
    Double(f64), 
    Currency(Currency), 
    //BString(String),
    Dispatch(Ptr<IDispatch>), 
    SCode(SCode),
    Bool(bool), 
    Variant(Variant), 
    Unknown(Ptr<IUnknown>), 
    Decimal(DecWrapper), 
    //Record
    Char(i8), 
    Byte(u8), 
    UShort(u16),
    ULong(u32), 
    Int(Int),
    UInt(UInt), 
}

pub type Array = Vec<RSafeArrayItem>;
