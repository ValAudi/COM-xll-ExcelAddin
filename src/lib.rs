use std::error::Error;

use windows::Win32::System::Variant::*;
use crate::variant::*;

pub mod automation;
pub mod rot;
pub mod workbook;
pub mod dispatch;
pub mod variant;
pub mod worksheet;
pub mod range;
pub mod data;
pub mod ribbon;
pub mod menu;


#[no_mangle]
#[allow(non_snake_case, unused_variables)]
extern "system" fn xlAutoOpen() {
    // Function implementation goes here
    // You can return an integer as in the original signature.
    // If this function is a placeholder, you can just return 0.\

    let _r = set_sheetname();
}

pub fn add(left: usize, right: usize) -> usize {
    left + right
}


pub fn test_ribbon_ui() -> Result<(), Box<dyn Error>> {
    let _t = automation::excel_automation::modify_ui_ribbon()?;
    Ok(())
}

pub fn test_command_bar() -> Result<(), Box<dyn Error>> {
    let _t = automation::excel_automation::modify_ui_menu()?;
    Ok(())
}

pub fn test_variant() {
    let t = variant_initialize(None, VT_I2, VariantType::VT_I2(2));
    println!("{:#?}", unsafe { t.Anonymous.Anonymous.vt } );
    
    println!("{:#?}", unsafe { t.Anonymous.Anonymous.Anonymous.iVal} );
}

pub fn get_sheetname() -> Result<(), Box<dyn Error>> {
    let _t = automation::excel_automation::retrieve_worksheet_name()?;
    Ok(())
}

pub fn set_sheetname() -> Result<(), Box<dyn Error>> {
    let _t = automation::excel_automation::set_worksheet_name()?;
    Ok(())
}

pub fn insert() -> Result<(), Box<dyn Error>> {
    let _t = automation::excel_automation::insert_value()?;
    Ok(())
}

pub fn retrieve() -> Result<(), Box<dyn Error>> {
    let _t = automation::excel_automation::retrieve_value()?;
    Ok(())
}

pub fn retrieve_array() -> Result<(), Box<dyn Error>> {
    let _t = automation::excel_automation::retrieve_range_values()?;
    Ok(())
}

pub fn set_values_array() -> Result<(), Box<dyn Error>> {
    let _t = automation::excel_automation::set_range_values()?;
    Ok(())
}

#[cfg(test)]
mod tests {

    use super::*;

    #[test]
    fn it_works() {
        let result = add(2, 2);
        assert_eq!(result, 4);
    }

    #[test]
    fn ribbon_ui_test()  -> Result<(), Box<dyn Error>> {
        let _ts = test_ribbon_ui()?;   
        Ok(())
    }

    #[test]
    fn menu_bar_test()  -> Result<(), Box<dyn Error>> {
        let _ts = test_command_bar()?;   
        Ok(())
    }

    #[test]
    fn works() {
        test_variant();   
    }

    #[test]
    fn test_com_set_array_values() -> Result<(), Box<dyn Error>> {
        let result = set_values_array()?;
        assert_eq!(result, ());
        Ok(())
    }

    #[test]
    fn test_com_get_array_values() -> Result<(), Box<dyn Error>> {
        let result = retrieve_array()?;
        assert_eq!(result, ());
        Ok(())
    }

    #[test]
    fn test_com_get_worksheet_name() -> Result<(), Box<dyn Error>> {
        let result = get_sheetname()?;
        assert_eq!(result, ());
        Ok(())
    }
    
    #[test]
    fn test_com_set_worksheet_name() -> Result<(), Box<dyn Error>> {
        let result = set_sheetname()?;
        assert_eq!(result, ());
        Ok(())
    }

    #[test]
    fn test_com_get_value() -> Result<(), Box<dyn Error>> {
        let result = retrieve()?;
        assert_eq!(result, ());
        Ok(())
    }

    #[test]
    fn test_com_set_value() -> Result<(), Box<dyn Error>> {
        let result = insert()?;
        assert_eq!(result, ());
        Ok(())
    }

}
