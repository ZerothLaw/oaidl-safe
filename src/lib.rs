#![cfg_attr(feature = "impl_tryfrom", feature(try_from))]
#![cfg(windows)]
//Enable lints for specific cases
#![deny(future_incompatible)]
#![deny(missing_copy_implementations)]
#![deny(missing_docs)]
#![deny(nonstandard_style)]
#![deny(single_use_lifetimes)]
#![deny(trivial_casts)]
#![deny(trivial_numeric_casts)]
#![deny(unreachable_pub)]
#![deny(unused)]

//Turn these warnings into errors
#![deny(const_err)]
#![deny(dead_code)]
#![deny(deprecated)]
#![deny(improper_ctypes)]
#![deny(overflowing_literals)]

#![doc(html_root_url = "https://docs.rs/oaidl/0.2.0/x86_64-pc-windows-msvc/oaidl/")]
//! # Introduction
//! 
//! A module to handle conversion to and from common OLE/COM types - VARIANT, SAFEARRAY, and BSTR. 
//! 
//! This module provides some convenience types as well as traits and trait implementations for 
//! built in rust types - [`u8`], [`i8`], [`u16`], [`i16`], [`u32`], [`i32`], [`u64`], [`i64`], [`f32`], [`f64`], [`String`], [`bool`] 
//! to and from `VARIANT` structures. 
//! In addition, [`Vec<T>`] can be converted into a `SAFEARRAY` where T is one of
//! `i8`, `u8`, `u16`, `i16`, `u32`, `i32`, `String`, `f32`, `f64`, `bool`.
//! 
//! As well, `IUnknown`, `IDispatch` pointers can be marshalled back and forth across boundaries.
//! 
//! There are some convenience types provided for further types that VARIANT/SAFEARRAY support:
//! [`SCode`], [`Int`], [`UInt`], [`Currency`], [`Date`], [`DecWrapper`], [`VtEmpty`], [`VtNull`]
//! 
//! The relevant traits to use are: [`BStringExt`], [`SafeArrayElement`], [`SafeArrayExt`], and [`VariantExt`]
//! 
//! ## Examples
//! 
//! An example of how to use the module:
//! 
//! ```rust
//! extern crate oaidl;
//! extern crate widestring;
//! extern crate winapi;
//! 
//! use widestring::U16String;
//! use winapi::um::oaidl::VARIANT;
//! 
//! use oaidl::{BStringExt, IntoVariantError, VariantExt};
//! 
//! //simulate an FFI function
//! unsafe fn c_masq(s: *mut VARIANT, p: *mut VARIANT) {
//!     println!("vt of s: {}", (*s).n1.n2_mut().vt);
//!     println!("vt of p: {}", (*p).n1.n2_mut().vt);
//!     
//!     //free the memory allocated for these VARIANTS
//!     let s = *s;
//!     let p = *p;
//! }
//! 
//! fn main() -> Result<(), IntoVariantError> {
//!     let mut u = 1337u32;
//!     let mut sr = U16String::from_str("Turing completeness.");
//!     let p = VariantExt::<u32>::into_variant(u)?;
//!     let s = VariantExt::<*mut u16>::into_variant(sr)?;
//!     unsafe {c_masq(s.as_ptr(), p.as_ptr())};
//!     Ok(())
//! } 
//! 
//! ```
//! 
//! [`i8`]: https://doc.rust-lang.org/std/i8/index.html
//! [`u8`]: https://doc.rust-lang.org/std/u8/index.html
//! [`f32`]: https://doc.rust-lang.org/std/f32/index.html
//! [`f64`]: https://doc.rust-lang.org/std/f64/index.html
//! [`i16`]: https://doc.rust-lang.org/std/i16/index.html
//! [`i32`]: https://doc.rust-lang.org/std/i32/index.html
//! [`i64`]: https://doc.rust-lang.org/std/i64/index.html
//! [`u16`]: https://doc.rust-lang.org/std/u16/index.html
//! [`u32`]: https://doc.rust-lang.org/std/u32/index.html
//! [`u64`]: https://doc.rust-lang.org/std/u64/index.html
//! [`String`]: https://doc.rust-lang.org/std/string/struct.String.html
//! [`bool`]: https://doc.rust-lang.org/std/primitive.bool.html
//! [`Vec<T>`]: https://doc.rust-lang.org/std/vec/struct.Vec.html
//! [`SCode`]: struct.SCode.html
//! [`Currency`]: struct.Currency.html
//! [`Date`]: struct.Date.html
//! [`Int`]: struct.Int.html
//! [`UInt`]: struct.UInt.html 
//! [`DecWrapper`]: struct.DecWrapper.html
//! [`VtEmpty`]: struct.VtEmpty.html
//! [`VtNull`]: struct.VtNull.html
//! [`BStringExt`]: trait.BStringExt.html
//! [`SafeArrayElement`]: trait.SafeArrayElement.html
//! [`SafeArrayExt`]: trait.SafeArrayExt.html
//! [`VariantExt`]: trait.VariantExt.html
#[macro_use] extern crate failure;

extern crate rust_decimal;

#[cfg(feature="serde")]
#[macro_use]
extern crate serde;

extern crate widestring;

extern crate winapi;

mod array;
mod bstr;
mod errors;
mod ptr;
mod types;
mod variant;

// Types = DroppableBString, Ptr, Currency, Date, DecWrapper, Int, SCode, UInt, VariantBool, 
//  Variant, VtEmpty, VtNull, 
// Traits = BStringExt, SafeArrayElement, SafeArrayExt, VariantExt
pub use self::array::{SafeArrayElement, SafeArrayExt};
pub use self::bstr::{BStringExt, DroppableBString};
pub use self::errors::*;
pub use self::ptr::Ptr;
pub use self::types::{Currency, Date, DecWrapper, Int, SCode, TryConvert, UInt, VariantBool};
pub use self::variant::{Variant, VariantExt, VtEmpty, VtNull};