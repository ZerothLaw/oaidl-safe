# Oaidl change log

## 0.1.5 Release

**Generics**
Updated `SafeArrayExt` to be implemented on `ExactSizeIterator<Item=SafeArrayElement>`
This necessitated a change to the interfaces of `.into_safearray` and `.into_variant` from `&mut self` to `self`. This means the original value will be consumed. 

## 0.1.4 Release (Published) Oct-8-2018
Initial feature set released. 

**To/From BSTR**

 * `String` (through `U16String` from the widestring crate.)

**To/From SAFEARRAY**

 * `Vec<T>` where T is:
   * `i16`
   * `i32`
   * `f32`
   * `f64`
   * `Currency`
   * `Date`
   * `String`
   * `Ptr<IDispatch>`
   * `SCode`
   * `bool`
   * `Variant<T>` where `T` is supported by VARIANT
   * `Ptr<IUnknown>`
   * `Decimal`, `DecWrapper`
   * `i8`
   * `u8`
   * `u16`
   * `u32`
   * `Int`
   * `UInt`

**To/From VARIANT**

 * `i64`
 * `i32`
 * `u8`
 * `i16`
 * `f32`
 * `f64`
 * `bool`
 * `SCode`
 * `Currency`
 * `Date`
 * `String`
 * `Ptr<IUnknown>`
 * `Ptr<IDispatch>`
 * `Box<u8>`
 * `Box<i16>`
 * `Box<i32>`
 * `Box<i64>`
 * `Box<f32>`
 * `Box<f64>`
 * `Box<bool>`
 * `Box<SCode>`
 * `Box<Currency>`
 * `Box<Date>`
 * `Box<String>`
 * `Box<Ptr<IUnknown>>`
 * `Box<Ptr<IDispatch>>`
 * `Variant<T>` where `T` is any of these supported types
 * `Vec<T>` where `T` is any of the SafeArray supported types
 * `Ptr<c_void>`
 * `i8`
 * `u16`
 * `u32`
 * `u32`
 * `Int`
 * `UInt`
 * `Box<Decimal>`
 * `Box<i8>`
 * `Box<u16>`
 * `Box<u32>`
 * `Box<u64>`
 * `Box<Int>`
 * `Box<UInt>`
 * `DecWrapper`, `Decimal`

**New types for convenience usage**
 
 * `Ptr<T>` which encapsulates a pointer of type `*mut T`
 * `Currency` which can be converted to and from the COM/OLE type CY
 * `Date` which can be converted to and from the COM/OLE type DATE
 * `DecWrapper` which wraps the type `Decimal` from the crate rust_decimal and converts into COM/OLE type DECIMAL
 * `Int` which is used for VT_INT conversions
 * `UInt` which is used for VT_UINT conversions
 * `SCode` which is used for VT_ERROR conversions
 * `VariantBool` which can convert between the `bool` <==> `VARIANT_BOOL` types
 * `Variant<T>` holds a value of type `T` for VT_VARIANT conversions
 * `VtEmpty` which is used for VT_EMPTY conversions
 * `VtNull` which is used for VT_NULL conversions