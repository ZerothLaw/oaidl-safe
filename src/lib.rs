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

// Types = Ptr, Currency, Date, DecWrapper, Int, SCode, UInt, VariantBool, 
//  Variant, VtEmpty, VtNull
// Traits = SafeArrayElement, SafeArrayExt, VariantExt
pub use array::{SafeArrayElement, SafeArrayExt};
pub use ptr::{Ptr};
pub use types::{Currency, Date, DecWrapper,Int, SCode, UInt, VariantBool};
pub use variant::{Variant, VariantExt, VtEmpty, VtNull};