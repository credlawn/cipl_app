// Copyright (c) 2025, Credlawn India Private Limited and contributors
// For license information, please see license.txt

frappe.ui.form.on("excel_field_mapping", {
    refresh(frm) {
        setup_field_autocomplete(frm);

        // Filter doctype_name to show only PB module DocTypes
        frm.set_query('doctype_name', function () {
            return {
                filters: {
                    'module': 'Cipl App'
                }
            };
        });
    },

    doctype_name(frm) {
        // Update autocomplete options when doctype changes
        setup_field_autocomplete(frm);

        // Clear child table when doctype changes to avoid confusion
        if (frm.doc.mapped_field && frm.doc.mapped_field.length > 0) {
            frappe.confirm(
                __('Changing the DocType will clear all mapped fields. Do you want to continue?'),
                function () {
                    frm.clear_table('mapped_field');
                    frm.refresh_field('mapped_field');
                },
                function () {
                    // Revert to previous value if user cancels
                    frm.reload_doc();
                }
            );
        }
    }
});

// Child table event to refresh autocomplete when a field is selected
frappe.ui.form.on("field_mapping_child", {
    doctype_field_name(frm, cdt, cdn) {
        // Refresh autocomplete to exclude newly selected field
        setup_field_autocomplete(frm);
    },

    excel_field_name(frm, cdt, cdn) {
        // Auto-populate doctype_field_name if matching field found
        if (!frm.doc.doctype_name) {
            return;
        }

        // Use frappe.get_doc to reliably get the row data
        let row = frappe.get_doc(cdt, cdn);

        if (!row || !row.excel_field_name) {
            return;
        }

        // Get all fields from selected doctype
        frappe.model.with_doctype(frm.doc.doctype_name, function () {
            let fields = frappe.get_meta(frm.doc.doctype_name).fields;

            // Normalize excel field name: lowercase, trim, replace spaces with underscores
            let excel_field_normalized = row.excel_field_name
                .toLowerCase()
                .trim()
                .replace(/\s+/g, '_');  // Replace spaces with underscores

            // Try to find matching field (case-insensitive, space-insensitive)
            let matched_field = null;
            fields.forEach(function (field) {
                if (field.fieldname && field.fieldname.toLowerCase() === excel_field_normalized) {
                    matched_field = field.fieldname;
                }
            });

            // Auto-populate if match found and doctype_field_name is empty
            if (matched_field && !row.doctype_field_name) {
                frappe.model.set_value(cdt, cdn, 'doctype_field_name', matched_field);
                // Refresh autocomplete after setting value
                setTimeout(function () {
                    setup_field_autocomplete(frm);
                }, 100);
            }
        });
    },

    mapped_field_remove(frm, cdt, cdn) {
        // Refresh autocomplete when a row is removed
        setup_field_autocomplete(frm);
    }
});

function setup_field_autocomplete(frm) {
    if (!frm.doc.doctype_name) {
        return;
    }

    // Fetch fields from the selected doctype
    frappe.model.with_doctype(frm.doc.doctype_name, function () {
        let fields = frappe.get_meta(frm.doc.doctype_name).fields;
        let all_field_options = [];

        // Create list of all fieldnames
        fields.forEach(function (field) {
            if (field.fieldname) {
                all_field_options.push(field.fieldname);
            }
        });

        // Get already selected fields to exclude them
        let selected_fields = [];
        if (frm.doc.mapped_field) {
            frm.doc.mapped_field.forEach(function (row) {
                if (row.doctype_field_name) {
                    selected_fields.push(row.doctype_field_name);
                }
            });
        }

        // Filter out already selected fields
        let available_options = all_field_options.filter(function (field) {
            return !selected_fields.includes(field);
        });

        // Set autocomplete options for the child table field
        frm.fields_dict.mapped_field.grid.update_docfield_property(
            'doctype_field_name',
            'options',
            available_options
        );

        // Refresh the grid to apply changes
        frm.refresh_field('mapped_field');
    });
}
