use std::ptr::NonNull;

use rust_decimal::Decimal;

use winapi::ctypes::c_void;


use winapi::shared::wtypes::{
    VT_ARRAY, 
    VT_BSTR, 
    VT_BOOL,
    VT_BYREF, 
    VT_CY,
    VT_DATE,
    VT_DECIMAL, 
    VT_DISPATCH,
    VT_EMPTY, 
    VT_ERROR, 
    VT_I1, 
    VT_I2, 
    VT_I4,
    VT_I8,
    VT_INT,  
    VT_NULL,
    VT_R4, 
    VT_R8, 
    VT_RECORD,
    VT_UI1,
    VT_UI2,
    VT_UI4,
    VT_UI8,  
    VT_UINT, 
    VT_UNKNOWN, 
    VT_VARIANT, 
};

use winapi::um::oaidl::{IDispatch, VARIANT};
use winapi::um::unknwnbase::IUnknown;

use ptr::Ptr;
use types::{Date, DecWrapper, Currency, Int, SCode, UInt };

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
    Byte(i8), 
    Short(i16), 
    Float(f32), 
    Double(f64), 
    VBool(bool), 
    SCode(SCode),
    Currency(Currency),
    Date(Date),
    BString(String), 
    Unknown(Ptr<IUnknown>),
    Dispatch(Ptr<IDispatch>),
    //SAFEARRAY, 
    PByte(Box<i8>), 
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
    //PRecord, //TODO: handle BRecord
    Decimal(DecWrapper),
}


impl Variant {
    pub fn from_error_code(err: SCode) -> Variant {
        Variant::SCode(err)
    }

    pub fn from_perror_code(err: Box<SCode>) -> Variant {
        Variant::PSCode(err)
    }

    pub fn vartype(&self) -> u32 {
        /*
        * VARENUM usage key,
        *
        * * [V] - may appear in a VARIANT
        * * [T] - may appear in a TYPEDESC
        * * [P] - may appear in an OLE property set
        * * [S] - may appear in a Safe Array
        *
        *
        *  VT_EMPTY            [V]   [P]     nothing
        *  VT_NULL             [V]   [P]     SQL style Null
        *  VT_I2               [V][T][P][S]  2 byte signed int
        *  VT_I4               [V][T][P][S]  4 byte signed int
        *  VT_R4               [V][T][P][S]  4 byte real
        *  VT_R8               [V][T][P][S]  8 byte real
        *  VT_CY               [V][T][P][S]  currency
        *  VT_DATE             [V][T][P][S]  date
        *  VT_BSTR             [V][T][P][S]  OLE Automation string
        *  VT_DISPATCH         [V][T]   [S]  IDispatch *
        *  VT_ERROR            [V][T][P][S]  SCODE
        *  VT_BOOL             [V][T][P][S]  True=-1, False=0
        *  VT_VARIANT          [V][T][P][S]  VARIANT *
        *  VT_UNKNOWN          [V][T]   [S]  IUnknown *
        *  VT_DECIMAL          [V][T]   [S]  16 byte fixed point
        *  VT_RECORD           [V]   [P][S]  user defined type
        *  VT_I1               [V][T][P][s]  signed char
        *  VT_UI1              [V][T][P][S]  unsigned char
        *  VT_UI2              [V][T][P][S]  unsigned short
        *  VT_UI4              [V][T][P][S]  ULONG
        *  VT_I8                  [T][P]     signed 64-bit int
        *  VT_UI8                 [T][P]     unsigned 64-bit int
        *  VT_INT              [V][T][P][S]  signed machine int
        *  VT_UINT             [V][T]   [S]  unsigned machine int
        *  VT_INT_PTR             [T]        signed machine register size width
        *  VT_UINT_PTR            [T]        unsigned machine register size width
        *  VT_VOID                [T]        C style void
        *  VT_HRESULT             [T]        Standard return type
        *  VT_PTR                 [T]        pointer type
        *  VT_SAFEARRAY           [T]        (use VT_ARRAY in VARIANT)
        *  VT_CARRAY              [T]        C style array
        *  VT_USERDEFINED         [T]        user defined type
        *  VT_LPSTR               [T][P]     null terminated string
        *  VT_LPWSTR              [T][P]     wide null terminated string
        *  VT_FILETIME               [P]     FILETIME
        *  VT_BLOB                   [P]     Length prefixed bytes
        *  VT_STREAM                 [P]     Name of the stream follows
        *  VT_STORAGE                [P]     Name of the storage follows
        *  VT_STREAMED_OBJECT        [P]     Stream contains an object
        *  VT_STORED_OBJECT          [P]     Storage contains an object
        *  VT_VERSIONED_STREAM       [P]     Stream with a GUID version
        *  VT_BLOB_OBJECT            [P]     Blob contains an object 
        *  VT_CF                     [P]     Clipboard format
        *  VT_CLSID                  [P]     A Class ID
        *  VT_VECTOR                 [P]     simple counted array
        *  VT_ARRAY            [V]           SAFEARRAY*
        *  VT_BYREF            [V]           void* for local use
        *  VT_BSTR_BLOB                      Reserved for system use
        */
        match *self {
            Variant::Empty => VT_EMPTY, 
            Variant::Null => VT_NULL, 
            Variant::Short(_) => VT_I2,
            Variant::Long(_) => VT_I4, 
            Variant::Float(_) => VT_R4, 
            Variant::Double(_) => VT_R8,
            Variant::Currency(_) => VT_CY, 
            Variant::Date(_) => VT_DATE,
            Variant::BString(_) => VT_BSTR, 
            Variant::Dispatch(_) => VT_DISPATCH, 
            Variant::SCode(_) => VT_ERROR, 
            Variant::VBool(_) => VT_BOOL, 
            Variant::PVariant(_) => VT_VARIANT, 
            Variant::Unknown(_) => VT_UNKNOWN,
            Variant::Decimal(_) => VT_DECIMAL, 
            //Variant::PRecord => VT_RECORD, 
            Variant::Byte(_) => VT_I1, 
            Variant::Char(_) => VT_UI1, 
            Variant::UShort(_) => VT_UI2, 
            Variant::ULong(_) => VT_UI4,
            Variant::LongLong(_) => VT_I8,
            Variant::ULongLong(_) => VT_UI8, 
            Variant::Int(_) => VT_INT, 
            Variant::UInt(_) => VT_UINT, 
            Variant::ByRef(_) => VT_BYREF, 
            Variant::PByte(_) => VT_BYREF | VT_I1, 
            Variant::PShort(_) => VT_BYREF | VT_I2, 
            Variant::PLong(_) => VT_BYREF | VT_I4,
            Variant::PLongLong(_) => VT_BYREF | VT_I8, 
            Variant::PULongLong(_) => VT_BYREF | VT_UI8, 
            Variant::PFloat(_) => VT_BYREF | VT_R4, 
            Variant::PDouble(_) => VT_BYREF | VT_R8, 
            Variant::PVBool(_) => VT_BYREF | VT_BOOL, 
            Variant::PSCode(_) => VT_BYREF | VT_ERROR, 
            Variant::PCurrency(_) => VT_BYREF | VT_CY, 
            Variant::PDate(_) => VT_BYREF | VT_DATE, 
            Variant::PBString(_) => VT_BYREF | VT_BSTR, 
            Variant::PUnknown(_) => VT_BYREF | VT_UNKNOWN, 
            Variant::PDispatch(_) => VT_BYREF | VT_DISPATCH, 
            Variant::PDecimal(_) => VT_BYREF | VT_DECIMAL, 
            Variant::PChar(_) => VT_BYREF | VT_UI1, 
            Variant::PUShort(_) => VT_BYREF | VT_UI2, 
            Variant::PULong(_) => VT_BYREF | VT_UI4, 
            Variant::PInt(_) => VT_BYREF | VT_INT, 
            Variant::PUInt(_) => VT_BYREF | VT_UINT, 
            
        }
    }
}

macro_rules! from_impls {
    ($($t:ty => $v:ident),* $(,)*) => {
        $(
            impl From<$t> for Variant {
                fn from(t: $t) -> Variant {
                    Variant::$v(t)
                }
            }
        )*
    };
}

impl From<i64> for Variant {
    fn from(i: i64) -> Variant {
        Variant::LongLong(i)
    }
}

from_impls!{
    i32 => Long, 
    i8 =>Byte,
    i16 => Short, 
    f32 => Float, 
    f64 => Double, 
    bool => VBool, 
    Currency => Currency, 
    Date => Date, 
    String => BString, 
    Ptr<IUnknown> => Unknown, 
    Ptr<IDispatch> => Dispatch,
    Box<i8> => PByte,
    Box<i16> => PShort, 
    Box<i32> => PLong, 
    Box<i64> => PLongLong, 
    Box<f32> => PFloat, 
    Box<f64> => PDouble, 
    Box<bool> => PVBool,
    Box<Currency> => PCurrency, 
    Box<Date> => PDate, 
    Box<String> => PBString,
    Box<Ptr<IUnknown>> => PUnknown, 
    Box<Ptr<IDispatch>> => PDispatch,
    Box<Variant> => PVariant, 
    Box<c_void> => ByRef, 
    u8 => Char,
    u16 => UShort, 
    u32 => ULong, 
    u64 => ULongLong,
    Box<DecWrapper> => PDecimal, 
    Box<u8> => PChar, 
    Box<u16> => PUShort,  
    Box<u32> => PULong, 
    Box<u64> => PULongLong, 
    DecWrapper => Decimal,
}

impl From<*mut IDispatch> for Variant {
    fn from(p: *mut IDispatch) -> Variant {
        match NonNull::new(p) {
            Some(p) => Variant::Dispatch(Ptr::new(p)), 
            None => Variant::Null,
        }
    }
}

impl From<*mut IUnknown> for Variant {
    fn from(p: *mut IUnknown) -> Variant {
        match NonNull::new(p) {
            Some(p) => Variant::Unknown(Ptr::new(p)), 
            None => Variant::Null,
        }
    }
}

// impl From<Variant> for VARIANT {
//     fn from(v: Variant) -> VARIANT {

//     }
// }

// impl From<VARIANT> for Variant {
//     fn from(v: VARIANT) -> Variant {

//     }
// }