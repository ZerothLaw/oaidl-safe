use std::mem;
use std::ptr::NonNull;

use winapi::ctypes::c_void;

use winapi::shared::wtypes::{
    CY, DATE, DECIMAL,
    VARIANT_BOOL,
    //VT_ARRAY, 
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
    //VT_RECORD,
    VT_UI1,
    VT_UI2,
    VT_UI4,
    VT_UI8,  
    VT_UINT, 
    VT_UNKNOWN, 
    VT_VARIANT, 
    
};

use winapi::um::oaidl::{IDispatch, __tagVARIANT, VARIANT, VARIANT_n3, VARIANT_n1};
use winapi::um::unknwnbase::IUnknown;

use ptr::Ptr;
use types::{Date, DecWrapper, Currency, Int, SCode, UInt, VariantBool };

pub const VT_PUI1: u32 = VT_BYREF | VT_UI1;
pub const VT_PI2: u32 = VT_BYREF | VT_I2;
pub const VT_PI4: u32 = VT_BYREF | VT_I4;
pub const VT_PI8: u32 = VT_BYREF | VT_I8;
pub const VT_PUI8: u32 = VT_BYREF | VT_UI8;
pub const VT_PR4: u32 = VT_BYREF | VT_R4;
pub const VT_PR8: u32 = VT_BYREF | VT_R8;
pub const VT_PBOOL: u32 = VT_BYREF | VT_BOOL;
pub const VT_PERROR: u32 = VT_BYREF | VT_ERROR;
pub const VT_PCY: u32 = VT_BYREF | VT_CY;
pub const VT_PDATE: u32 = VT_BYREF | VT_DATE;
//pub const VT_PBSTR: u32 = VT_BYREF | VT_BSTR;
pub const VT_PUNKNOWN: u32 = VT_BYREF | VT_UNKNOWN;
pub const VT_PDISPATCH: u32 = VT_BYREF | VT_DISPATCH;
pub const VT_PDECIMAL: u32 = VT_BYREF | VT_DECIMAL;
pub const VT_PI1: u32 = VT_BYREF | VT_I1;
pub const VT_PUI2: u32 = VT_BYREF | VT_UI2;
pub const VT_PUI4: u32 = VT_BYREF | VT_UI4;
pub const VT_PINT: u32 = VT_BYREF | VT_INT;
pub const VT_PUINT: u32 = VT_BYREF | VT_UINT;

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
    SCode(SCode),
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
    ByRef(Ptr<c_void>), 
    Char(i8), 
    UShort(u16), 
    ULong(u32), 
    ULongLong(u64), 
    Int(Int), 
    UInt(UInt), 
    PDecimal(Box<DecWrapper>),
    PChar(Box<i8>), 
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
            Variant::Byte(_) => VT_UI1, 
            Variant::Char(_) => VT_I1, 
            Variant::UShort(_) => VT_UI2, 
            Variant::ULong(_) => VT_UI4,
            Variant::LongLong(_) => VT_I8,
            Variant::ULongLong(_) => VT_UI8, 
            Variant::Int(_) => VT_INT, 
            Variant::UInt(_) => VT_UINT, 
            Variant::ByRef(_) => VT_BYREF, 
            Variant::PByte(_) => VT_BYREF | VT_UI1, 
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
            Variant::PChar(_) => VT_BYREF | VT_I1, 
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
    u8 =>Byte,
    i16 => Short, 
    f32 => Float, 
    f64 => Double, 
    bool => VBool, 
    Currency => Currency, 
    Date => Date, 
    String => BString, 
    Ptr<IUnknown> => Unknown, 
    Ptr<IDispatch> => Dispatch,
    Box<u8> => PByte,
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
    Ptr<c_void> => ByRef, 
    i8 => Char,
    u16 => UShort, 
    u32 => ULong, 
    u64 => ULongLong,
    Box<DecWrapper> => PDecimal, 
    Box<i8> => PChar, 
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

impl From<Variant> for VARIANT {
    fn from(v: Variant) -> VARIANT {
        let vt = v.vartype();
        let mut n3: VARIANT_n3 = unsafe {mem::zeroed()};
        let mut n1: VARIANT_n1 = unsafe {mem::zeroed()};

        match v {
            Variant::Empty => {}, 
            Variant::Null => {}, 
            Variant::LongLong(val) => unsafe {
                let mut n_ptr = n3.llVal_mut();
                *n_ptr = val;
            },
            Variant::Long(val) => unsafe {
                let mut n_ptr = n3.lVal_mut();
                *n_ptr = val;
            },
            Variant::Byte(val) => unsafe {
                let mut n_ptr = n3.bVal_mut();
                *n_ptr = val;
            },
            Variant::Short(val) => unsafe {
                let mut n_ptr = n3.iVal_mut();
                *n_ptr = val;
            },
            Variant::Float(val) => unsafe {
                let mut n_ptr = n3.fltVal_mut();
                *n_ptr = val;
            },
            Variant::Double(val) => unsafe {
                let mut n_ptr = n3.dblVal_mut();
                *n_ptr = val;
            },
            Variant::VBool(v) => unsafe {
                let mut n_ptr = n3.boolVal_mut();
                *n_ptr = VARIANT_BOOL::from(v)
            }, 
            Variant::SCode(v) => unsafe {
                let mut n_ptr = n3.scode_mut();
                *n_ptr = v;
            }
            Variant::Currency(val) => unsafe {
                let mut n_ptr = n3.cyVal_mut();
                *n_ptr = CY::from(val);
            },
            Variant::Date(v) => unsafe {
                let mut n_ptr = n3.date_mut();
                *n_ptr = DATE::from(v);
            }, 
            Variant::BString(_) => {}, 
            Variant::Unknown(ptr) => unsafe {
                let mut n_ptr = n3.punkVal_mut();
                *n_ptr = ptr.as_ptr();
            }
            Variant::Dispatch(ptr) => unsafe {
                let mut n_ptr = n3.pdispVal_mut();
                *n_ptr = ptr.as_ptr();
            }, 
            Variant::PByte(bptr) => unsafe {
                let mut n_ptr = n3.pbVal_mut();
                *n_ptr = Box::into_raw(bptr);
            }, 
            Variant::PShort(bptr) => unsafe {
                let mut n_ptr = n3.piVal_mut();
                *n_ptr = Box::into_raw(bptr);
            }, 
            Variant::PLong(bptr) => unsafe {
                let mut n_ptr = n3.plVal_mut();
                *n_ptr = Box::into_raw(bptr);
            }, 
            Variant::PLongLong(bptr) => unsafe {
                let mut n_ptr = n3.pllVal_mut();
                *n_ptr = Box::into_raw(bptr);
            }, 
            Variant::PFloat(bptr) => unsafe {
                let mut n_ptr = n3.pfltVal_mut();
                *n_ptr = Box::into_raw(bptr);
            }
            Variant::PDouble(bptr) => unsafe {
                let mut n_ptr = n3.pdblVal_mut();
                *n_ptr = Box::into_raw(bptr);
            }
            Variant::PVBool(bptr) => unsafe {
                let mut n_ptr = n3.pboolVal_mut();
                let b_val = *bptr;
                let bptr = Box::new(VARIANT_BOOL::from(b_val));
                *n_ptr = Box::into_raw(bptr);
            }, 
            Variant::PSCode(bptr) => unsafe {
                let mut n_ptr = n3.pscode_mut();
                *n_ptr = Box::into_raw(bptr);
            }, 
            Variant::PCurrency(bptr) => unsafe {
                let mut n_ptr = n3.pcyVal_mut();
                let cy = *bptr;
                let bptr = Box::new(CY::from(cy));
                *n_ptr = Box::into_raw(bptr);
            }, 
            Variant::PDate(bptr) => unsafe {
                let mut n_ptr = n3.pdate_mut();
                let dt = (*bptr).0;
                let bptr = Box::new(DATE::from(dt));
                *n_ptr = Box::into_raw(bptr);
            }, 
            Variant::PBString(_) => {},
            Variant::PUnknown(bptr) => unsafe {
                let mut n_ptr = n3.ppunkVal_mut();
                let bptr = Box::new((*bptr).as_ptr());
                *n_ptr = Box::into_raw(bptr);
            }, 
            Variant::PDispatch(bptr) => unsafe {
                let mut n_ptr = n3.ppdispVal_mut();
                let bptr = Box::new((*bptr).as_ptr());
                *n_ptr = Box::into_raw(bptr);
            }, 
            Variant::PVariant(bptr) => unsafe {
                let mut n_ptr = n3.pvarVal_mut();
                let mut bptr = Box::new(VARIANT::from(*bptr));
                *n_ptr = Box::into_raw(bptr);
            }, 
            Variant::ByRef(bptr) => unsafe {
                let mut n_ptr = n3.byref_mut();
                *n_ptr = bptr.as_ptr();
            }, 
            Variant::Char(v) => {
                let mut n_ptr = unsafe {n3.cVal_mut()};
                *n_ptr = v;
            }, 
            Variant::UShort(v) => {
                let mut n_ptr = unsafe {n3.uiVal_mut()};
                *n_ptr = v;
            }, 
            Variant::ULong(v) => unsafe {
                let mut n_ptr = n3.ulVal_mut();
                *n_ptr = v;
            }, 
            Variant::ULongLong(v) => unsafe {
                let mut n_ptr = n3.ullVal_mut();
                *n_ptr = v;
            }, 
            Variant::Int(v) => unsafe {
                let mut n_ptr = n3.intVal_mut();
                *n_ptr = v;
            }, 
            Variant::UInt(v) => unsafe {
                let mut n_ptr = n3.uintVal_mut();
                *n_ptr = v;
            }, 
            Variant::PDecimal(bptr) => unsafe {
                let mut n_ptr = n3.pdecVal_mut();
                let bptr = Box::new(DECIMAL::from(*bptr));
                *n_ptr = Box::into_raw(bptr);
            }, 
            Variant::PChar(bptr) => unsafe {
                let mut n_ptr = n3.pcVal_mut();
                *n_ptr = Box::into_raw(bptr);
            }, 
            Variant::PUShort(bptr) => unsafe {
                let mut n_ptr = n3.puiVal_mut();
                *n_ptr = Box::into_raw(bptr);
            }, 
            Variant::PULong(bptr) => unsafe {
                let mut n_ptr = n3.pulVal_mut();
                *n_ptr = Box::into_raw(bptr);
            }, 
            Variant::PULongLong(bptr) => unsafe {
                let mut n_ptr = n3.pullVal_mut();
                *n_ptr = Box::into_raw(bptr);
            }, 
            Variant::PInt(bptr) => unsafe {
                let mut n_ptr = n3.pintVal_mut();
                *n_ptr = Box::into_raw(bptr);
            }, 
            Variant::PUInt(bptr) => unsafe {
                let mut n_ptr = n3.puintVal_mut();
                *n_ptr = Box::into_raw(bptr);
            }, 
            Variant::Decimal(d) => unsafe {
                let mut n_ptr = n1.decVal_mut();
                *n_ptr = DECIMAL::from(d);
            }
        };

        let tv = __tagVARIANT { vt: vt as u16, 
                                wReserved1: 0, 
                                wReserved2: 0, 
                                wReserved3: 0, 
                                n3: n3};
        unsafe {
            let n_ptr = n1.n2_mut();
            *n_ptr = tv;
        };
        VARIANT {n1: n1}
    }
}

impl From<VARIANT> for Variant {
    fn from(v: VARIANT) -> Variant {
        let mut n1 = v.n1;

        let vt = unsafe {
            n1.n2_mut().vt
        };

        let mut n3 = unsafe {
            n1.n2_mut().n3
        };

        match vt as u32 {
            VT_I8 => unsafe {
                let mut n_ptr = n3.llVal();
                Variant::LongLong(*n_ptr)
            }, 
            VT_I4 => unsafe {
                let mut n_ptr = n3.lVal();
                Variant::Long(*n_ptr)
            }, 
            VT_UI1 => unsafe {
                let mut n_ptr = n3.bVal();
                Variant::Byte(*n_ptr)
            }, 
            VT_R4 => unsafe {
                let mut n_ptr = n3.fltVal();
                Variant::Float(*n_ptr)
            },
            VT_R8 => unsafe {
                let mut n_ptr = n3.dblVal();
                Variant::Double(*n_ptr)
            },
            VT_BOOL => unsafe {
                let mut n_ptr = n3.boolVal();
                Variant::VBool(bool::from(VariantBool::from(*n_ptr)))
            }, 
            VT_ERROR => unsafe {
                let mut n_ptr = n3.scode();
                Variant::SCode(*n_ptr)
            }, 
            VT_CY => unsafe {
                let mut n_ptr = n3.cyVal();
                Variant::Currency(Currency::from(*n_ptr))
            }, 
            VT_DATE => unsafe {
                let mut n_ptr = n3.date();
                Variant::Date(Date::from(*n_ptr))
            }, 
            //VT_BSTR
            VT_UNKNOWN => unsafe {
                let mut n_ptr = n3.punkVal();
                match NonNull::new(*n_ptr) {
                    Some(nn) => Variant::Unknown(Ptr::new(nn)), 
                    None => Variant::Null,
                }
            }, 
            VT_DISPATCH => unsafe {
                let mut n_ptr = n3.pdispVal();
                match NonNull::new(*n_ptr) {
                    Some(nn) => Variant::Dispatch(Ptr::new(nn)), 
                    None => Variant::Null,
                }
            },
            //VT_ARRAY
            VT_PUI1 => unsafe {
                let mut n_ptr = n3.pbVal();
                Variant::PByte(Box::new(**n_ptr))
            }, 
            VT_PI2 => unsafe {
                let mut n_ptr = n3.piVal();
                Variant::PShort(Box::new(**n_ptr))
            }, 
            VT_PI4 => unsafe {
                let mut n_ptr = n3.plVal();
                Variant::PLong(Box::new(**n_ptr))
            }, 
            VT_PI8 => unsafe {
                let mut n_ptr = n3.pllVal();
                Variant::PLongLong(Box::new(**n_ptr))
            }, 
            VT_PR4 => unsafe {
                let mut n_ptr = n3.pfltVal();
                Variant::PFloat(Box::new(**n_ptr))
            }, 
            VT_PR8 => unsafe {
                let mut n_ptr = n3.pdblVal();
                Variant::PDouble(Box::new(**n_ptr))
            }, 
            VT_PBOOL => unsafe {
                let mut n_ptr = n3.pboolVal();
                Variant::PVBool(Box::new(bool::from(VariantBool::from(**n_ptr))))
            }, 
            VT_PERROR => unsafe {
                let mut n_ptr = n3.pscode();
                Variant::PSCode(Box::new(**n_ptr))
            }, 
            VT_PCY => unsafe {
                let mut n_ptr = n3.pcyVal();
                Variant::PCurrency(Box::new(Currency::from(**n_ptr)))
            }, 
            VT_PDATE => unsafe {
                let mut n_ptr = n3.pdate();
                Variant::PDate(Box::new(Date(**n_ptr)))
            },
            //VT_PBSTR 
            VT_PUNKNOWN => unsafe {
                let mut n_ptr = n3.ppunkVal_mut();
                match NonNull::new((**n_ptr).clone()) {
                    Some(nn) => Variant::PUnknown(Box::new(Ptr::new(nn))), 
                    None => Variant::Null,
                }
            },
            VT_PDISPATCH => unsafe {
                let mut n_ptr = n3.ppdispVal_mut();
                match NonNull::new((**n_ptr).clone()) {
                    Some(nn) => Variant::PDispatch(Box::new(Ptr::new(nn))), 
                    None => Variant::Null
                }
            }, 
            //VT_PARRAY
            VT_BYREF => unsafe {
                let mut n_ptr = n3.byref();
                match NonNull::new(*n_ptr) {
                    Some(nn) => Variant::ByRef(Ptr::new(nn)), 
                    None => Variant::Null
                }
            }, 
            VT_I1 => unsafe {
                let mut n_ptr = n3.cVal();
                Variant::Char(*n_ptr)
            }, 
            VT_UI2 => unsafe {
                let mut n_ptr = n3.uiVal();
                Variant::UShort(*n_ptr)
            }, 
            VT_UI4 => unsafe {
                let mut n_ptr = n3.ulVal();
                Variant::ULong(*n_ptr)
            }, 
            VT_UI8 => unsafe {
                let mut n_ptr = n3.ullVal();
                Variant::ULongLong(*n_ptr)
            }, 
            VT_INT => unsafe {
                Variant::Int(*(n3.intVal()))
            }, 
            VT_UINT => unsafe {
                Variant::UInt(*(n3.uintVal()))
            }, 
            VT_PDECIMAL => unsafe {
                Variant::PDecimal(Box::new(DecWrapper::from(**(n3.pdecVal()))))
            }, 
            VT_PI1 => unsafe {
                Variant::PChar(Box::new(**(n3.pcVal())))
            }, 
            VT_PUI2 => unsafe {
                Variant::PUShort(Box::new(**(n3.puiVal())))
            }, 
            VT_PUI4 => unsafe {
                Variant::PULong(Box::new(**(n3.pulVal())))
            }, 
            VT_PUI8 => unsafe {
                Variant::PULongLong(Box::new(**(n3.pullVal())))
            }, 
            VT_PINT => unsafe {
                Variant::PInt(Box::new(**(n3.pintVal())))
            }, 
            VT_PUINT => unsafe {
                Variant::PUInt(Box::new(**(n3.puintVal())))
            }, 
            VT_DECIMAL => unsafe {
                Variant::Decimal(DecWrapper::from(*(n1.decVal())))
            }, 
            VT_EMPTY => Variant::Empty, 
            VT_NULL => Variant::Null,
            _ => unimplemented!()
        }

        //Variant::Null
    }
}