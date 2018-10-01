use rust_decimal::Decimal;

use winapi::ctypes::c_void;
use winapi::um::oaidl::IDispatch;
use winapi::um::unknwnbase::IUnknown;

use ptr::Ptr;
use types::{Date, Currency, Int, UInt, DecWrapper};

/// Enum to wrap Rust-types for conversion to/from 
/// the C-level VARIANT data structure. 
/// 
/// Reference:
///typedef struct tagVARIANT {
///  union {
///    struct {
///      VARTYPE vt;
///      WORD    wReserved1;
///      WORD    wReserved2;
///      WORD    wReserved3;
///      union {
///        LONGLONG     llVal;
///        LONG         lVal;
///        BYTE         bVal;
///        SHORT        iVal;
///        FLOAT        fltVal;
///        DOUBLE       dblVal;
///        VARIANT_BOOL boolVal;
///        SCODE        scode;
///        CY           cyVal;
///        DATE         date;
///        BSTR         bstrVal;
///        IUnknown     *punkVal;
///        IDispatch    *pdispVal;
///        SAFEARRAY    *parray;
///        BYTE         *pbVal;
///        SHORT        *piVal;
///        LONG         *plVal;
///        LONGLONG     *pllVal;
///        FLOAT        *pfltVal;
///        DOUBLE       *pdblVal;
///        VARIANT_BOOL *pboolVal;
///        SCODE        *pscode;
///        CY           *pcyVal;
///        DATE         *pdate;
///        BSTR         *pbstrVal;
///        IUnknown     **ppunkVal;
///        IDispatch    **ppdispVal;
///        SAFEARRAY    **pparray;
///        VARIANT      *pvarVal;
///        PVOID        byref;
///        CHAR         cVal;
///        USHORT       uiVal;
///        ULONG        ulVal;
///        ULONGLONG    ullVal;
///        INT          intVal;
///        UINT         uintVal;
///        DECIMAL      *pdecVal;
///        CHAR         *pcVal;
///        USHORT       *puiVal;
///        ULONG        *pulVal;
///        ULONGLONG    *pullVal;
///        INT          *pintVal;
///        UINT         *puintVal;
///        struct {
///          PVOID       pvRecord;
///          IRecordInfo *pRecInfo;
///        } __VARIANT_NAME_4;
///      } __VARIANT_NAME_3;
///    } __VARIANT_NAME_2;
///    DECIMAL decVal;
///  } __VARIANT_NAME_1;
///} VARIANT;
pub enum Variant {
    Empty, 
    Null, 
    LongLong(i64), 
    Long(i32),
    Byte(u8), 
    Short(i16), 
    Float(f32), 
    Double(f64), 
    VBool(bool), 
    SCode(i32),
    Currency(Currency),
    Date(Date),
    BString(String), 
    Unknown(Ptr<IUnknown>),
    Dispatch(Ptr<IDispatch>),
    //SAFEARRAY, 
    PByte(Box<u8>), 
    PShort(Box<i16>), 
    PLong(Box<i32>), 
    PLongLong(Box<i64>), 
    PFloat(Box<f32>), 
    PDouble(Box<f64>),
    PVBool(Box<bool>), 
    PSCode(Box<i32>), 
    PCurrency(Box<Currency>), 
    PDate(Box<Date>),
    PBString(Box<String>), 
    PUnknown(Box<Ptr<IUnknown>>), 
    PDispatch(Box<Ptr<IDispatch>>),
    //*mut *mut SAFEARRAY
    PVariant(Box<Variant>), 
    ByRef(Box<c_void>), 
    Char(u8), 
    UShort(u16), 
    ULong(u32), 
    ULongLong(u64), 
    Int(Int), 
    UInt(UInt), 
    PDecimal(Box<DecWrapper>),
    PChar(Box<u8>), 
    PUShort(Box<u16>), 
    PULong(Box<u32>), 
    PULongLong(Box<u64>), 
    PInt(Box<Int>), 
    PUInt(Box<UInt>),
    PRecord, //TODO: handle BRecord
    Decimal(Decimal),
}

