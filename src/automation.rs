pub mod excel_automation {

    use windows::{Win32::System::Com::*, core::*};
    use crate::rot;
    use crate::dispatch;
    use crate::workbook;
    use crate::worksheet;
    use crate::range;
    use crate::data;
    use crate::menu;
    use crate::ribbon;
    // use crate::com;
    use crate::registry;

    pub fn registering() -> Result<()> 
    {
        unsafe { 
            CoInitializeEx(None, COINIT_APARTMENTTHREADED) 
        }?;
        let _ = registry::create_typelibrary()?;
        unsafe {
            CoUninitialize()
        };
        Ok(())
    }

    pub fn modify_ui_ribbon() -> Result<()>
    {
        // Initialize a COM 
        unsafe { 
            CoInitializeEx(None, COINIT_APARTMENTTHREADED) 
        }?;
        let ribbon_stream = ribbon::read_ribbon_customizations()?;
        println!("{:#?}", ribbon_stream);
        let _loading = ribbon::load_customized_ribbon()?;
        unsafe {
            CoUninitialize()
        };
        Ok(())
    } 


    pub fn modify_ui_menu() -> Result<()>
    {
        // Initialize a COM 
        unsafe { 
            CoInitializeEx(None, COINIT_APARTMENTTHREADED) 
        }?;

        let excel_dispatch = rot::ole_active_object()?;
        let command_bars_dispacth = dispatch::get_dispatch_interface(&excel_dispatch, menu::COMMAND_BARS_ID)?;
        let count = menu::get_int_values(&command_bars_dispacth, menu::COMMAND_BARS_COUNT)?;
        for i in 204..=204 {
            let command_bar_dispacth = dispatch::get_dispatch_interface_passing_in_paramters(&command_bars_dispacth, menu::COMMAND_BARS_ITEM, i)?;
            let name = menu::get_text_values(&command_bar_dispacth, menu::COMMAND_BAR_NAME)?;
            let visible = menu::get_visibility(&command_bar_dispacth, menu::COMMAND_BAR_VISIBILITY)?;
            let position = menu::get_int_values(&command_bar_dispacth, menu::COMMAND_BAR_POSITION)?;

            if visible == true {
                println!("-------------------------------------------------");
                println!(" Number of Menu Items in the Command Bar: \t{:#?}", count);
                println!(" Index position of Menu Item: \t{:#?}", i);
                println!(" Name of Command Bar Item: \t{:#?}", name);
                println!(" Visibility of Command Bar Item: \t{:#?}", visible);
                println!(" MSO Position of Command Bar Item: \t{:#?}", position);
                println!("-------------------------------------------------");

                let command_bar_controls = dispatch::get_dispatch_interface(&command_bar_dispacth, menu::COMMAND_BAR_CONTROLS)?;
                let count_of_controls = menu::get_int_values(&command_bar_controls, menu::COMMAND_BAR_CONTROLS_COUNT)?;
                println!("COUNT OF COMMAND BAR CONTROLS: \t{}", count_of_controls);
                for j in 1..=1 {
                    let command_bar_controls_dispacth = dispatch::get_dispatch_interface_passing_in_paramters(&command_bars_dispacth, menu::COMMAND_BARS_ITEM, j)?;

                    let ti = unsafe { &command_bar_controls_dispacth.GetTypeInfo(0, 0)? };
                    let ptr_func: *mut TYPEATTR = unsafe { ti.GetTypeAttr()? };
                    let func = unsafe { *ptr_func };
                    println!("{:#?}", func.guid);

                    let item_name = menu::get_text_values(&command_bar_controls_dispacth, menu::COMMAND_BAR_NAME)?;
                    let visible1 = menu::get_visibility(&command_bar_controls_dispacth, menu::COMMAND_BAR_VISIBILITY)?;

                    println!("-------------------------------------------------");
                    println!(" Index position of Menu Item: \t{:#?}", j);
                    println!(" Name of Command Bar Item: \t{:#?}", item_name);
                    println!(" Visibility of Command Bar Item: \t{:#?}", visible1);
                    println!("-------------------------------------------------");

                    let layer_two_dispatch = dispatch::get_dispatch_interface(&command_bar_controls_dispacth, menu::COMMAND_BAR_CONTROLS)?;

                    let two = unsafe { &layer_two_dispatch.GetTypeInfo(0, 0)? };
                    let ptr_func_one: *mut TYPEATTR = unsafe { two.GetTypeAttr()? };
                    let func = unsafe { *ptr_func_one };
                    println!("{:#?}", func.guid);

                    let office_count_of_controls = menu::get_int_values(&layer_two_dispatch, menu::COMMAND_BAR_CONTROLS_COUNT)?;
                    println!("COUNT OF COMMAND BAR CONTROLS: \t{}", office_count_of_controls);

                    for k in 9..=9 {
                        let office_command_bar_controls_dispacth = dispatch::get_dispatch_interface_passing_in_paramters(&layer_two_dispatch, menu::COMMAND_BARS_ITEM, k)?;
                        let three = unsafe { &office_command_bar_controls_dispacth.GetTypeInfo(0, 0)? };
                        let ptr_func_two: *mut TYPEATTR = unsafe { three.GetTypeAttr()? };
                        let func_two = unsafe { *ptr_func_two };
                        println!("-------------------------------------------------");
                        println!("{:#?}", func_two.guid);
                        println!(" Index position of Menu Item: \t{:#?}", k);

                        let visible_popup = menu::get_visibility(&office_command_bar_controls_dispacth, menu::COMMAND_BAR_POPUP_VISIBILITY)?;
                        let tooltip_popup = menu::get_text_values(&office_command_bar_controls_dispacth, menu::COMMAND_BAR_TOOLTIP)?;
                        let desc_text_popup = menu::get_text_values(&office_command_bar_controls_dispacth, menu::COMMAND_BAR_DESCRIPTION)?;
                        let tag_popup = menu::get_text_values(&office_command_bar_controls_dispacth, menu::COMMAND_BAR_TAG)?;
                        let caption_popup = menu::get_text_values(&office_command_bar_controls_dispacth, menu::COMMAND_BAR_CAPTION)?;
                        println!(" Visibility of Command Bar Popup Item: \t{:#?}", visible_popup);
                        println!(" Description Text of Command Bar Popup Item: \t{:#?}", desc_text_popup);
                        println!(" Tooltip of Command Bar Popup Item: \t{:#?}", tooltip_popup);
                        println!(" Tag of Command Bar Popup Item: \t{:#?}", tag_popup);
                        println!(" Caption of Command Bar Popup Item: \t{:#?}", caption_popup);
                        println!("-------------------------------------------------");

                        let layer_three_dispatch = dispatch::get_dispatch_interface(&office_command_bar_controls_dispacth, menu::COMMAND_BAR_POPUP_CONTROLS)?;
                        let three = unsafe { &layer_three_dispatch.GetTypeInfo(0, 0)? };
                        let ptr_func_two: *mut TYPEATTR = unsafe { three.GetTypeAttr()? };
                        let func_three = unsafe { *ptr_func_two };
                        println!("Third Layer GUID : {:#?}", func_three.guid);

                        let layer_three_count_of_controls = menu::get_int_values(&layer_three_dispatch, menu::COMMAND_BAR_CONTROLS_COUNT)?;
                        println!("COUNT OF COMMAND BAR CONTROLS: \t{}", layer_three_count_of_controls);

                        for l in 7..=7 {
                            let into_popup_command_bar_controls_dispacth = dispatch::get_dispatch_interface_passing_in_paramters(&layer_three_dispatch, menu::COMMAND_BARS_ITEM, l)?;
                            let four = unsafe { &into_popup_command_bar_controls_dispacth.GetTypeInfo(0, 0)? };
                            let ptr_func_three: *mut TYPEATTR = unsafe { four.GetTypeAttr()? };
                            let func_four = unsafe { *ptr_func_three };
                            println!("-------------------------------------------------");
                            println!("{:#?}", func_four.guid);
                            println!(" Index position of Menu Item in LAYER THREE: \t{:#?}", l);

                            let visible_bar_button = menu::get_visibility(&into_popup_command_bar_controls_dispacth, menu::COMMAND_BAR_POPUP_VISIBILITY)?;
                            let tooltip_bar_button = menu::get_text_values(&into_popup_command_bar_controls_dispacth, menu::COMMAND_BAR_TOOLTIP)?;
                            let desc_text_bar_button = menu::get_text_values(&into_popup_command_bar_controls_dispacth, menu::COMMAND_BAR_DESCRIPTION)?;
                            let tag_bar_button = menu::get_text_values(&into_popup_command_bar_controls_dispacth, menu::COMMAND_BAR_TAG)?;
                            let caption_bar_button = menu::get_text_values(&into_popup_command_bar_controls_dispacth, menu::COMMAND_BAR_CAPTION)?;
                            let onaction_bar_button = menu::get_text_values(&into_popup_command_bar_controls_dispacth, menu::COMMAND_BAR_ONACTION)?;
                            let parameter_bar_button = menu::get_text_values(&into_popup_command_bar_controls_dispacth, menu::COMMAND_BAR_BUTTON_PARAMETER)?;
                            let cmd_type = menu::get_int_values(&into_popup_command_bar_controls_dispacth, menu::COMMAND_BAR_BUTTON_TYPE)?;
                            let hyperlink_type = menu::get_int_values(&into_popup_command_bar_controls_dispacth, menu::COMMAND_BAR_BUTTON_HYPERLINK_TYPE)?;
                            let style = menu::get_int_values(&into_popup_command_bar_controls_dispacth, menu::COMMAND_BAR_BUTTON_STYLE)?;
                            let state = menu::get_int_values(&into_popup_command_bar_controls_dispacth, menu::COMMAND_BAR_BUTTON_STATE)?;

                            println!("-------------------------------------------------");
                            println!(" Visibility of Command Bar Button Item: \t{:#?}", visible_bar_button);
                            println!(" Description Text of Command Bar Button Item: \t{:#?}", desc_text_bar_button);
                            println!(" Tooltip of Command Bar Button Item: \t{:#?}", tooltip_bar_button);
                            println!(" Tag of Command Bar Button Item: \t{:#?}", tag_bar_button);
                            println!(" Caption of Command Bar Button Item: \t{:#?}", caption_bar_button);
                            println!(" On Action of Command Bar Button Item: \t{:#?}", onaction_bar_button);
                            println!(" Parameter of Command Bar Button Item: \t{:#?}", parameter_bar_button);
                            println!(" Type of Command Bar Button Item: \t{:#?}", cmd_type);
                            println!(" Hyperlink Type of Command Bar Button Item: \t{:#?}", hyperlink_type);
                            println!(" Style of Command Bar Button Item: \t{:#?}", style);
                            println!(" State of Command Bar Button Item: \t{:#?}", state);
                            println!("-------------------------------------------------");

                            // let command_bar_button_control_dispacth = dispatch::get_dispatch_interface(&into_popup_command_bar_controls_dispacth, menu::CONTROL_INTERFACE_ID)?;
                            // let five = unsafe { &command_bar_button_control_dispacth.GetTypeInfo(0, 0)? };
                            // let ptr_func_four: *mut TYPEATTR = unsafe { five.GetTypeAttr()? };
                            // let func_five = unsafe { *ptr_func_four };
                            // println!("-------------------------------------------------");
                            // println!("{:#?}", func_five.guid);

                        }

                        // let visible_popup = menu::get_visibility(&layer_three_dispatch, menu::COMMAND_BAR_POPUP_VISIBILITY)?;
                        // let tooltip_popup = menu::get_text_values(&layer_three_dispatch, menu::COMMAND_BAR_TOOLTIP)?;
                        // let desc_text_popup = menu::get_text_values(&layer_three_dispatch, menu::COMMAND_BAR_DESCRIPTION)?;
                        // let tag_popup = menu::get_text_values(&layer_three_dispatch, menu::COMMAND_BAR_TAG)?;
                        // let caption_popup = menu::get_text_values(&layer_three_dispatch, menu::COMMAND_BAR_CAPTION)?;
                        // println!(" Visibility of Command Bar Popup Item: \t{:#?}", visible_popup);
                        // println!(" Description Text of Command Bar Popup Item: \t{:#?}", desc_text_popup);
                        // println!(" Tooltip of Command Bar Popup Item: \t{:#?}", tooltip_popup);
                        // println!(" Tag of Command Bar Popup Item: \t{:#?}", tag_popup);
                        // println!(" Caption of Command Bar Popup Item: \t{:#?}", caption_popup);
                        // println!("-------------------------------------------------");

                    }

                }
            }
            // let _del = menu::command_bar_delete(&command_bar_dispacth, menu::COMMAND_BAR_DELETE)?;
            // let command_control_interface = dispatch::get_dispatch_interface_from_fn(command_bar_dispacth, menu::COMMAND_BAR_CONTROL)?;
            // let control_interface = dispatch::get_dispatch_interface(command_control_interface, menu::CONTROL_INTERFACE_ID)?;
            // let ti = unsafe { control_interface.GetTypeInfo(0, 0)? };
            // let ptr_func: *mut TYPEATTR = unsafe { ti.GetTypeAttr()? };
            // let func = unsafe { *ptr_func };
            // // let _res = menu::load_customized_menu()?;
            // println!("{:#?}", func.guid);
            
        }

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
        let workbook_dispacth = dispatch::get_dispatch_interface(&excel_dispatch, workbook::ACTIVE_WORKBOOK_ID)?;
        let worksheet_dispatch = dispatch::get_dispatch_interface(&workbook_dispacth, worksheet::ACTIVE_WORKSHEET_ID)?;
        let range_dispatch = range::get_range_interface(&worksheet_dispatch, range::CELL_RANGE_ID)?;
        let var = data::set_range_data();
        range::set_range_array(&range_dispatch, range::RANGE_VALUES_ID, var)?;
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
        let workbook_dispacth = dispatch::get_dispatch_interface(&excel_dispatch, workbook::ACTIVE_WORKBOOK_ID)?;
        let worksheet_dispatch = dispatch::get_dispatch_interface(&workbook_dispacth, worksheet::ACTIVE_WORKSHEET_ID)?;
        let range_dispatch = range::get_range_interface(&worksheet_dispatch, range::CELL_RANGE_ID)?;
        let array_data = range::get_range_array(&range_dispatch, range::RANGE_VALUES_ID)?;
        data::get_range_data(array_data);
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
        let workbook_dispacth = dispatch::get_dispatch_interface(&excel_dispatch, workbook::ACTIVE_WORKBOOK_ID)?;
        let worksheet_dispatch = dispatch::get_dispatch_interface(&workbook_dispacth, worksheet::ACTIVE_WORKSHEET_ID)?;
        let name = worksheet::get_sheetname(&worksheet_dispatch, worksheet::ACTIVE_WORKSHEET_NAME_ID)?;
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
        let workbook_dispacth = dispatch::get_dispatch_interface(&excel_dispatch, workbook::ACTIVE_WORKBOOK_ID)?;
        let worksheet_dispatch = dispatch::get_dispatch_interface(&workbook_dispacth, worksheet::ACTIVE_WORKSHEET_ID)?;
        let range_dispatch = range::get_range_interface(&worksheet_dispatch, range::CELL_RANGE_ID)?;
       let _get_value = data::get_cell_data(&range_dispatch, data::CELL_VALUES_ID)?;
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
        let workbook_dispacth = dispatch::get_dispatch_interface(&excel_dispatch, workbook::ACTIVE_WORKBOOK_ID)?;
        let worksheet_dispatch = dispatch::get_dispatch_interface(&workbook_dispacth, worksheet::ACTIVE_WORKSHEET_ID)?;
        let range_dispatch = range::get_range_interface(&worksheet_dispatch, range::CELL_RANGE_ID)?;
       let _set_value = data::set_cell_value(&range_dispatch, data::CELL_VALUES_ID)?;
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
        let workbook_dispacth = dispatch::get_dispatch_interface(&excel_dispatch, workbook::ACTIVE_WORKBOOK_ID)?;
        let worksheet_dispatch = dispatch::get_dispatch_interface(&workbook_dispacth, worksheet::ACTIVE_WORKSHEET_ID)?;
        let name = worksheet::set_sheetname(&worksheet_dispatch, worksheet::ACTIVE_WORKSHEET_NAME_ID)?;
        println!("{:#?}", name );
        unsafe {
            CoUninitialize()
        };
        Ok(())
    }

}