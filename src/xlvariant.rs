#![allow(non_snake_case, non_camel_case_types, non_upper_case_globals, dead_code)]

use crate::xlcall::*;
// use crate::entrypoint::excel_free;

const xltypeMask : u16 = !(xlbitDLLFree as u16 | xlbitXLFree as u16);
const xltypeStr_xlbitDLLFree: u16 = xltypeStr as u16| xlbitDLLFree as u16;
const xltypeMulti_xlbitDLLFree: u16 = xltypeMulti as u16| xlbitDLLFree as u16;

const xltypeMulti_u16: u16 = xltypeMulti as u16;
const xltypeMissing_u16: u16 = xltypeMissing as u16;
const xltypeSRef_u16: u16 = xltypeSRef as u16;
const xltypeRef_u16: u16 = xltypeRef as u16;
const xltypeErr_u16: u16 =xltypeErr as u16;
const xltypeInt_u16: u16 = xltypeInt as u16;
const xltypeNil_u16: u16 = xltypeNil as u16;
const xltypeNum_u16: u16 = xltypeNum as u16;
const xltypeStr_u16: u16 = xltypeStr as u16;

/// Variant is a wrapper around an XLOPER. It can contain a string, i32 or f64, or a
/// two dimensional of any mixture of these. Basically, it can contain anything that an
/// Excel cell or array of cells can contain.
pub struct Variant(XLOPER);

impl Variant {
    /// Construct a variant containing nil. This is used in Excel to represent cells that have
    /// nothing in them. It is also a sensible starting state for an uninitialized variant.
    pub fn new() -> Variant {
        Variant(XLOPER { xltype : xltypeNil as u16, val: xloper__bindgen_ty_1 { w: 0 } })
    }

    /// Construct a variant from an LPXLOPER, for example supplied by Excel. The assumption
    /// is that Excel continues to own the XLOPER and its lifetime is greater than that of
    /// the Variant we construct here. For example, the LPXLOPER may be an argument to one
    /// of our functions. We therefore do not want to own any of the data in this variant, so
    /// we clear all ownership bits. This means we treat it as a kind of dynamic mut ref. 
    pub fn from_xloper(xloper: LPXLOPER) -> Variant {
        let mut result = Variant(unsafe { *xloper });
        result.0.xltype &= xltypeMask ;    // no ownership bits
        result
    }

    /// Construct a variant containing an int (i32)
    pub fn from_int(w: i16) -> Variant {
        Variant(XLOPER { xltype : xltypeInt as u16, val: xloper__bindgen_ty_1 { w: w  } })
    }

    /// Construct a variant containing a float (f64)
    pub fn from_float(num: f64) -> Variant {
        Variant(XLOPER { xltype : xltypeNum as u16, val: xloper__bindgen_ty_1 { num: num } })
    }

    /// Construct a variant containing a missing entry. This is used in function calls to
    /// signal that a parameter should be defaulted.
    pub fn missing() -> Variant {
        Variant(XLOPER { xltype : xltypeMissing as u16, val: xloper__bindgen_ty_1 { w: 0 } })
    }

    /// Construct a variant containing an error. This is used in Excel to represent standard errors
    /// that are shown as #DIV0 etc. Currently supported error codes are:
    /// xlerrNull, xlerrDiv0, xlerrValue, xlerrRef, xlerrName, xlerrNum, xlerrNA, xlerrGettingData
    pub fn from_err(xlerr: u32) -> Variant {
        Variant(XLOPER { xltype : xltypeErr as u16, val: xloper__bindgen_ty_1 { err: xlerr as u16 } })
    }

    /// Construct a variant containing a string. Strings in Excel (at least after Excel 97) are 16bit
    /// Unicode starting with a 16-bit length. The length is treated as signed, which means that
    /// strings can be no longer than 32k characters. If a string longer than this is supplied, or a 
    /// string that is not valid 16bit Unicode, an xlerrValue error is stored instead.
    pub fn from_str(s: &str) -> Variant {
        let mut cstr : Vec<u8> = s.as_bytes().into_iter().map(|char| *char).collect();
        let len = cstr.len();
        if len > 32767 {
            return Variant::from_err(xlerrValue)
        }

        // Pascal-style string with length at the start. Forget the string so we do not delete it.
        // We are now relying on the drop method of Variant to clean it up for us. Note that the
        // shrink_to_fit is essential, so the capacity is the same as the length. We have no way
        // of storing the capacity otherwise.
        cstr.insert(0, len as u8);
        cstr.shrink_to_fit();
        let p = cstr.as_mut_ptr();
        std::mem::forget(cstr);
  
        Variant(XLOPER { xltype : xltypeStr as u16 + xlbitDLLFree as u16, val: xloper__bindgen_ty_1 { str: p as *mut i8 } })
    }

    /// Constructs an XLOPER that contains an array of XLOPERs. The data should match the rows
    /// and columns.
    pub fn from_array(cols: usize, rows: usize, data: &[Variant]) -> Variant {
        let size = cols * rows;
        let data_len = data.len();
        if data_len > size {
            return Variant::from_str("Error: variant data size greater than array size")
        }
        let mut array = vec![Variant::from_err(xlerrNA); size];
        for (i, src) in data.iter().enumerate() {
            let xloper = Box::new(src.clone());
            let raw_xloper = Box::into_raw(xloper) as LPXLOPER;
            array[i] = Variant::from_xloper(raw_xloper);
        }

        let lparray = array.as_mut_ptr() as LPXLOPER;
        std::mem::forget(array);

        Variant(XLOPER { 
            xltype : xltypeMulti as u16 + xlbitDLLFree as u16, 
            val: xloper__bindgen_ty_1 {
                array: xloper__bindgen_ty_1__bindgen_ty_3 {
                    lparray, rows : rows as u16, columns : cols as u16 } } })
    }

    /// Construct a variant containing an array from a slice of other variants. The variants
    /// may contain arrays or scalar strings or numbers, which are treated like single-cell 
    /// arrays. They are glued either horizontally (horiz=true) or vertically. If the arrays
    /// do not match sizes in the other dimension, they are padded with blanks.
    pub fn concat(from: &[Variant], horiz: bool) -> Variant {

        // first find the size of the resulting array
        let mut columns : usize = 0;
        let mut rows : usize = 0;
        for xloper in from.iter() {
            let dim = xloper.dim();
            if horiz {
                columns += dim.0;
                rows = rows.max(dim.1);
            } else {
                columns = columns.max(dim.0);
                rows += dim.1;
            }
        }

        // Zero-sized arrays cause Excel to crash. Arrays with a dimension of
        // one (either rows or cols) are confusing to Excel, which repeats them
        // when using array paste. Solve both problems by padding with NA and
        // setting the min rows or cols to two.
        rows = rows.max(2);
        columns = columns.max(2); 

        // If the array is too big, return an error string
        if rows > 1048576 || columns > 16384 {
            return Self::from_str("#ERR resulting array is too big")
        }

        // now clone the components into place
        let size = rows * columns;
        let mut array = vec![Variant::from_err(xlerrNA); size];
        let mut col = 0;
        let mut row = 0;
        for var in from.iter() {
            match var.0.xltype & xltypeMask {
                xltypeMulti_u16   => unsafe {
                    let p = var.0.val.array.lparray;
                    let var_cols = var.0.val.array.columns as usize;
                    let var_rows = var.0.val.array.rows as usize;
                    for x in 0..var_cols {
                        for y in 0..var_rows {
                            let src = (y * var_cols + x) as isize;
                            let dest = (row + y) * columns + col + x;
                            array[dest] = Variant::from_xloper(p.offset(src)).clone();
                        }
                    }

                    if horiz {
                        col += var_cols;
                    } else {
                        row += var_rows;
                    }
                },
                xltypeMissing_u16 => {},
                _ => {
                    let dest = row * columns + col;
                    array[dest] = var.clone();
                    if horiz {
                        col += 1;
                    } else {
                        row += 1;
                    }
                }
            }
        }

        let lparray = array.as_mut_ptr() as LPXLOPER;
        std::mem::forget(array);

        Variant(XLOPER { 
            xltype : xltypeMulti as u16, 
            val: xloper__bindgen_ty_1 {
                array: xloper__bindgen_ty_1__bindgen_ty_3 {
                    lparray, rows : rows as u16, columns : columns as u16 } } })
    }

    /// Creates a transposed clone of this Variant. If this Variant is a scalar type,
    /// simply returns it unchanged.
    pub fn transpose(&self) -> Variant {
        // simply clone any scalar type, including errors
        if (self.0.xltype & xltypeMask) != xltypeMulti as u16 {
            return self.clone()
        }

        // We have an array that we need to transpose. Create a vector of
        // Variant to contain the elements.
        let dim = self.dim();
        if dim.0 > 1048576 || dim.1 > 16384 {
            return Self::from_str("#ERR resulting array is too big")
        }

        let len = dim.0 * dim.1;
        let mut array = Vec::with_capacity(len);

        // Copy the elements transposed, cloning each one
        for i in 0..dim.1 {
            for j in 0..dim.0 {
                array.push(self.at(j, i));
            }
        }

        // Return as a Variant
        let lparray = array.as_mut_ptr() as LPXLOPER;
       std::mem::forget(array);

        Variant(XLOPER { 
            xltype : xltypeMulti as u16, 
            val: xloper__bindgen_ty_1 {
                array: xloper__bindgen_ty_1__bindgen_ty_3 {
                    lparray, rows : dim.0 as u16, columns : dim.1 as u16 } } })
    }

    /// Converts this variant to a string. Alternatively, you can use Display or to_string,
    /// which both go through this call if the variant contains a string. Guaranteed to return
    /// Some(...) if this object is of type xltypeStr. Always returns None if this object is
    /// of any other type. If the string contains a unicode string that is misformed, return
    /// the error message.
    pub fn as_string(&self) -> Option<String> {
        if (self.0.xltype & xltypeMask) != xltypeStr as u16 {
             None
        } else {
            let cstr_slice = unsafe {
                let cstr: *mut i8 = self.0.val.str;
                let cstr_len = *cstr.offset(0) as usize;
                std::slice::from_raw_parts(cstr.offset(1), cstr_len) };

            match String::from_utf8(cstr_slice.to_vec().into_iter().map(|char| char as u8).collect()) {
                Ok(s) => Some(s),
                Err(e) => Some(e.to_string())
            }
        }
    }

    /// Converts this variant to an int. If we do not contain an int, return None. Note that
    /// Excel cells do not ever contain ints, so this would only come from a non-Excel user
    /// creating an XLOPER, for example the result of a call into Excel.
    pub fn as_i32(&self) -> Option<i32> {
        if (self.0.xltype & xltypeMask) != xltypeInt as u16{
            None
        } else {
            Some(unsafe { self.0.val.w as i32 })
        }
    }
    
    /// Converts this variant to a float. If we do not contain a float, return None.
    pub fn as_f64(&self) -> Option<f64> {
        if (self.0.xltype & xltypeMask) != xltypeNum as u16{
            None
        } else {
            Some(unsafe { self.0.val.num })
        }
    }

    /// Does this variant represent a missing entry?
    pub fn is_missing(&self) -> bool {
        return self.0.xltype & xltypeMissing as u16 == xltypeMissing as u16
    }

    /// Exposes the underlying XLOPER
    pub fn as_mut_xloper(&mut self) -> *mut XLOPER {
        Box::into_raw(Box::new(self.0)) 
    }

    /// Is this a cell reference?
    pub fn is_ref(&self) -> bool {
        let xltype = self.0.xltype & xltypeMissing as u16;
        return xltype == xltypeRef as u16 || xltype == xltypeSRef as u16
    }

    /// Gets the count of rows and columns. Scalars are treated as 1x1. Missing values are
    /// treated as 0x0.
    pub fn dim(&self) -> (usize, usize) {
        match self.0.xltype & xltypeMask {
            xltypeMulti_u16 => unsafe { (self.0.val.array.columns as usize, self.0.val.array.rows as usize) },
            xltypeMissing_u16 => (0, 0),
            xltypeSRef_u16 => get_sref_dim(unsafe { &self.0.val.sref.ref_ }),
            xltypeRef_u16 => get_mref_dim(unsafe { self.0.val.mref.lpmref }),

            _ => (1, 1)
        }
    }

    /// Gets the element at the given column and row. If this is a scalar, treat it as a one-element
    /// array. If the column or row is out of bounds, return NA. The returned element is always cloned
    /// so it can be returned as a value
    pub fn at(&self, column: usize, row: usize) -> Variant {
        if (self.0.xltype & xltypeMask) != xltypeMulti as u16 {
            if column == 0 && row == 0 {
                self.clone()
            } else {
                Self::from_err(xlerrNA)
            }
        } else {
            let (columns, rows) = unsafe {
                (self.0.val.array.columns as usize, self.0.val.array.rows as usize) };
            if column >= columns || row >= rows {
                Self::from_err(xlerrNA)
            } else {
                let index = row * columns + column;
                Self::from_xloper( unsafe {
                    self.0.val.array.lparray.offset(index as isize) }).clone()
            }
        }
    }
}

// Gets the array size of a multi-cell reference. If the reference is badly formed,
// returns (0, 0)
fn get_mref_dim(mref: * const XLMREF) -> (usize, usize) {
    // currently we only handle single contiguous references
    if mref.is_null() || unsafe { (*mref).count } != 1 {
        return (0, 0)
    }

    return get_sref_dim(unsafe { &(*mref).reftbl[0] })
}

// Gets the array size of a single-cell reference
fn get_sref_dim(sref: &XLREF) -> (usize, usize) {
    let rows = 1 + (sref.rwLast - sref.rwFirst) as usize;
    let cols = 1 + (sref.colLast - sref.colFirst) as usize;
    (cols, rows)
}

/// Implement Display, which means we do not need a method for converting to strings. Just use
/// to_string.
impl std::fmt::Display for Variant {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        match self.0.xltype & xltypeMask {
            xltypeErr_u16 => match unsafe { self.0.val.err } as u32 {
                xlerrNull => write!(f, "#NULL"),
                xlerrDiv0 => write!(f, "#DIV0"),
                xlerrValue => write!(f, "#VALUE"),
                xlerrRef => write!(f, "#REF"),
                xlerrName => write!(f, "#NAME"),
                xlerrNum => write!(f, "#NUM"),
                xlerrNA => write!(f, "#NA"),
                xlerrGettingData => write!(f, "#DATA"),
                _ => write!(f, "#BAD_ERR")
            }
            xltypeInt_u16 => write!(f, "{}", unsafe { self.0.val.w }),
            xltypeMissing_u16 => write!(f, "#MISSING"),
            xltypeMulti_u16 => write!(f, "#MULTI"),
            xltypeNil_u16 => write!(f, "#NIL"),
            xltypeNum_u16 => write!(f, "{}", unsafe { self.0.val.num }),
            xltypeStr_u16 => write!(f, "{}", self.as_string().unwrap()),
            _ => write!(f, "#BAD_XLOPER")
        }
    }
}

// /// We need to implement Drop, as Variant is a wrapper around a union type that does
// /// not know how to handle its contained pointers.
// impl Drop for Variant {
//     fn drop(&mut self) {
//         if (self.0.xltype & xlbitXLFree) != 0 {
//             excel_free(&mut self.0);
//             return
//         }
        
//         match self.0.xltype {
//             xltypeStr_xlbitDLLFree => {
//                 // We have a 16bit string that was originally allocated as a vector
//                 // but then forgotten. Reconstruct the vector, so its drop method
//                 // will clean up the memory for us.
//                 unsafe {
//                     let p = self.0.val.str;
//                     let len = *p as usize + 1;
//                     let cap = len;
//                     Vec::from_raw_parts(p, len, cap);
//                 }
//             },
//             xltypeMulti_xlbitDLLFree => {
//                 // We have an array that was originally allocated as a vector of
//                 // Variant but then forgotten. Reconstruct the vector, so its drop method
//                 // will clean up the vector and its elements for us.
//                 unsafe {
//                     let p = self.0.val.array.lparray as *mut Variant;
//                     let len = (self.0.val.array.rows * self.0.val.array.columns) as usize;
//                     let cap = len;
//                     Vec::from_raw_parts(p, len, cap);
//                 }
//             },
//             _ => {
//                 // nothing to do
//             }
//         }
//     }
// }

/// We need to hand-code Clone, because of the ownership issues for strings and multi.
impl Clone for Variant {
    fn clone(&self) -> Variant {
        // a simple copy is good enough for most variant types, but make sure the addin
        // is the owner
        let mut copy = Variant(self.0.clone());
        copy.0.xltype &= !xlbitXLFree as u16;
        copy.0.xltype |= xlbitDLLFree as u16;

        // Special handling for string and mult, to avoid double delete of the member
        match copy.0.xltype {
            xltypeStr_xlbitDLLFree => {

                // We have a 16bit string that was originally allocated as a vector
                // but then forgotten. Reconstruct the vector, so we can clone it.
                unsafe {
                    let p = copy.0.val.str;
                    let len = *p as usize + 1;
                    let cap = len;
                    let string_vec = Vec::from_raw_parts(p, len, cap);
                    let mut cloned = string_vec.clone();
                    copy.0.val.str = cloned.as_mut_ptr();

                    // now forget everything -- we do not want either string deallocated
                   std::mem::forget(string_vec);
                   std::mem::forget(cloned);
                }
            },
            xltypeMulti_xlbitDLLFree => {

                // We have an array that was originally allocated as a vector
                // but then forgotten. Reconstruct the vector, so we can clone it.
                unsafe {
                    let p = self.0.val.array.lparray as *mut Variant;
                    let len = (self.0.val.array.rows * self.0.val.array.columns) as usize;
                    let cap = len;
                    let array = Vec::from_raw_parts(p, len, cap);
                    let mut cloned = array.clone();
                    copy.0.val.array.lparray = cloned.as_mut_ptr() as LPXLOPER;

                    // now forget everything -- we do not want either string deallocated
                   std::mem::forget(array);
                   std::mem::forget(cloned);
                }
            },
            _ => {
                // nothing to do
            }
        }

        copy
    }
}