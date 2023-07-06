use std::error::Error;

pub mod automation;

pub fn add(left: usize, right: usize) -> usize {
    left + right
}

pub fn insert() -> Result<(), Box<dyn Error>> {
    let _t = automation::excel_automation::insert_value()?;
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
    fn test_com() -> Result<(), Box<dyn Error>> {
        let result = insert()?;
        assert_eq!(result, ());
        Ok(())
    }
}
