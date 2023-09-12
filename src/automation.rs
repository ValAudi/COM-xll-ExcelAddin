pub mod excel_automation {

    use windows::{Win32::System::Com::*, core::*};
    use crate::rot;
    use crate::dispatch;
    use crate::workbook;
    use crate::worksheet;
    use crate::range;
    use crate::data;
    use crate::ribbon;

    pub fn modify_ui_ribbon() -> Result<()>
    {
        // Initialize a COM 
        unsafe { 
            CoInitializeEx(None, COINIT_APARTMENTTHREADED) 
        }?;

        let _res = ribbon::load_customized_ribbon()?;

        unsafe {
            CoUninitialize()
        };
        Ok(())
    } 

    pub fn set_range_values() -> Result<()>
    {
        // Initialize a COM 
        unsafe { 
            CoInitializeEx(None, COINIT_APARTMENTTHREADED) 
        }?;

        let excel_dispatch = rot::ole_active_object()?;
        let workbook_dispacth = dispatch::get_dispatch_interface(excel_dispatch, workbook::ACTIVE_WORKBOOK_ID)?;
        let worksheet_dispatch = dispatch::get_dispatch_interface(workbook_dispacth, worksheet::ACTIVE_WORKSHEET_ID)?;
        let range_dispatch = dispatch::get_range_interface(worksheet_dispatch, range::CELL_RANGE_ID)?;
        let var = range::set_range_data();
        range::set_range_array(range_dispatch, range::RANGE_VALUES_ID, var)?;
        unsafe {
            CoUninitialize()
        };
        Ok(())
    } 

    pub fn retrieve_range_values() -> Result<()>
    {
        // Initialize a COM 
        unsafe { 
            CoInitializeEx(None, COINIT_APARTMENTTHREADED) 
        }?;

        let excel_dispatch = rot::ole_active_object()?;
        let workbook_dispacth = dispatch::get_dispatch_interface(excel_dispatch, workbook::ACTIVE_WORKBOOK_ID)?;
        let worksheet_dispatch = dispatch::get_dispatch_interface(workbook_dispacth, worksheet::ACTIVE_WORKSHEET_ID)?;
        let range_dispatch = dispatch::get_range_interface(worksheet_dispatch, range::CELL_RANGE_ID)?;
        let array_data = range::get_range_array(range_dispatch, range::RANGE_VALUES_ID)?;
        range::get_range_data(array_data);
        unsafe {
            CoUninitialize()
        };
        Ok(())
    } 

    pub fn retrieve_worksheet_name() -> Result<()>
    {
        // Initialize a COM 
        unsafe { 
            CoInitializeEx(None, COINIT_APARTMENTTHREADED) 
        }?;

        let excel_dispatch = rot::ole_active_object()?;
        let workbook_dispacth = dispatch::get_dispatch_interface(excel_dispatch, workbook::ACTIVE_WORKBOOK_ID)?;
        let worksheet_dispatch = dispatch::get_dispatch_interface(workbook_dispacth, worksheet::ACTIVE_WORKSHEET_ID)?;
        let name = worksheet::get_sheetname(worksheet_dispatch, worksheet::ACTIVE_WORKSHEET_NAME_ID)?;
        println!("{:#?}", name );
        unsafe {
            CoUninitialize()
        };
        Ok(())
    }

    pub fn retrieve_value() -> Result<()>
    {
        // Initialize a COM 
        unsafe { 
            CoInitializeEx(None, COINIT_APARTMENTTHREADED) 
        }?;

        let excel_dispatch = rot::ole_active_object()?;
        let workbook_dispacth = dispatch::get_dispatch_interface(excel_dispatch, workbook::ACTIVE_WORKBOOK_ID)?;
        let worksheet_dispatch = dispatch::get_dispatch_interface(workbook_dispacth, worksheet::ACTIVE_WORKSHEET_ID)?;
        let range_dispatch = dispatch::get_range_interface(worksheet_dispatch, range::CELL_RANGE_ID)?;
       let _get_value = data::get_cell_data(range_dispatch, data::CELL_VALUES_ID)?;
        unsafe {
            CoUninitialize()
        };
        Ok(())
    } 

    pub fn insert_value() -> Result<()>
    {
        // Initialize a COM 
        unsafe { 
            CoInitializeEx(None, COINIT_APARTMENTTHREADED) 
        }?;

        let excel_dispatch = rot::ole_active_object()?;
        let workbook_dispacth = dispatch::get_dispatch_interface(excel_dispatch, workbook::ACTIVE_WORKBOOK_ID)?;
        let worksheet_dispatch = dispatch::get_dispatch_interface(workbook_dispacth, worksheet::ACTIVE_WORKSHEET_ID)?;
        let range_dispatch = dispatch::get_range_interface(worksheet_dispatch, range::CELL_RANGE_ID)?;
       let _set_value = data::set_cell_value(range_dispatch, data::CELL_VALUES_ID)?;
        unsafe {
            CoUninitialize()
        };
        Ok(())
    }

    pub fn set_worksheet_name() -> Result<()>
    {
        // Initialize a COM 
        unsafe { 
            CoInitializeEx(None, COINIT_APARTMENTTHREADED) 
        }?;

        let excel_dispatch = rot::ole_active_object()?;
        let workbook_dispacth = dispatch::get_dispatch_interface(excel_dispatch, workbook::ACTIVE_WORKBOOK_ID)?;
        let worksheet_dispatch = dispatch::get_dispatch_interface(workbook_dispacth, worksheet::ACTIVE_WORKSHEET_ID)?;
        let name = worksheet::set_sheetname(worksheet_dispatch, worksheet::ACTIVE_WORKSHEET_NAME_ID)?;
        println!("{:#?}", name );
        unsafe {
            CoUninitialize()
        };
        Ok(())
    }

}