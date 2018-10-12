# oaidl
[![Crates.io]](https://crates.io/crates/oaidl)[![docs.rs(https://docs.rs/oaidl/badge.svg)]](https://docs.rs/oaidl/)

A crate to convert common Rust types to common COM/OLE types, primarily for use
in FFI - `BSTR`, `SAFEARRAY`, and `VARIANT` are the three implemented here. 

This crate provides traits and trait implementations to make it easy and safe to
convert between Rust types and the FFI-compatible data types. 

For reference, a `SAFEARRAY` of `VARIANTs` corresponds to a C# `object[]`. A 
`VARIANT` is considered an `object` by C# interop. 

## Documentation 

 - [Crate API Reference](https://docs.rs/oaidl/) 

## License

This project is distributed under the terms of the MIT license ([LICENSE-MIT](LICENSE-MIT) or
[http://opensource.org/licenses/MIT](http://opensource.org/licenses/MIT))

### Contributing

Unless you explicitly state otherwise, any contribution intentionally submitted for inclusion in the work by you shall be licensed as above, without any additional terms or conditions.