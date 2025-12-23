// Adobe Dump list view with Excel Import
// Just 5 lines to add import functionality!
frappe.listview_settings['adobe_dump'] = {
    onload: function (listview) {
        // Use global utility - auto-detects doctype
        add_excel_import_button(listview);
    }
};
