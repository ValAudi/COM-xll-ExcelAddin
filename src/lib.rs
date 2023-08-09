use std::error::Error;

pub mod automation;
pub mod rot;
pub mod workbook;
pub mod dispatch;
pub mod variant;
pub mod worksheet;
pub mod range;
pub mod data;

pub fn add(left: usize, right: usize) -> usize {
    left + right
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

#[cfg(test)]
mod tests {

    use super::*;

    #[test]
    fn it_works() {
        let result = add(2, 2);
        assert_eq!(result, 4);
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
