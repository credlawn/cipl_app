/**
 * Global Excel Import Utility for Cipl App
 * Modern, reusable Excel import dialog with auto-doctype detection
 * 
 * Usage in any doctype_list.js:
 * frappe.listview_settings['your_doctype'] = {
 *     onload: function(listview) {
 *         add_excel_import_button(listview);
 *     }
 * };
 */

// Global state management to prevent cross-doctype glitches
window.excel_import_state = {
    current_dialog: null,
    current_doctype: null,
    realtime_listener: null,

    reset: function () {
        // Cleanup previous state
        if (this.realtime_listener) {
            frappe.realtime.off('excel_import_progress', this.realtime_listener);
            this.realtime_listener = null;
        }
        if (this.current_dialog) {
            this.current_dialog.hide();
            this.current_dialog = null;
        }
        this.current_doctype = null;
    }
};

/**
 * Add Excel Import button to any list view
 * @param {Object} listview - Frappe listview object
 */
window.add_excel_import_button = function (listview) {
    // Auto-detect current doctype from listview
    const current_doctype = listview.doctype;

    listview.page.add_inner_button(__('Import Excel'), function () {
        show_excel_import_dialog(current_doctype);
    }).addClass('btn-primary');
};

/**
 * Show Excel Import Dialog with modern UI
 * @param {string} target_doctype - Doctype to import data into
 */
window.show_excel_import_dialog = function (target_doctype) {
    // Reset any previous state to prevent glitches
    window.excel_import_state.reset();
    window.excel_import_state.current_doctype = target_doctype;

    let selected_file = null;

    const dialog = new frappe.ui.Dialog({
        title: `
            <div style="display: flex; align-items: center; gap: 10px;">
                <i class="fa fa-file-excel-o" style="color: #10b981;"></i>
                <span>Import Excel Data</span>
            </div>
        `,
        fields: [
            {
                fieldtype: 'HTML',
                options: `
                    <div style="
                        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                        color: white;
                        padding: 15px;
                        border-radius: 8px;
                        margin-bottom: 20px;
                        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
                    ">
                        <div style="font-size: 14px; font-weight: 500; margin-bottom: 5px;">
                            📊 Importing to: <strong>${target_doctype}</strong>
                        </div>
                        <div style="font-size: 12px; opacity: 0.9;">
                            Upload your Excel file to import records
                        </div>
                    </div>
                `
            },
            {
                fieldname: 'attach_file',
                label: __('Select Excel File'),
                fieldtype: 'Attach',
                reqd: 1,
                description: 'Upload .xlsx or .xls file',
                onchange: function () {
                    selected_file = dialog.get_value('attach_file');
                    check_and_show_import_button();
                }
            },
            {
                fieldtype: 'Section Break'
            },
            {
                fieldname: 'import_controls',
                fieldtype: 'HTML',
                options: '<div id="excel-import-controls"></div>'
            }
        ],
        primary_action_label: __('Close'),
        primary_action: function () {
            window.excel_import_state.reset();
        }
    });

    // Store dialog reference
    window.excel_import_state.current_dialog = dialog;

    dialog.show();

    // Attach field workaround - multiple event listeners for reliability
    setTimeout(function () {
        dialog.fields_dict.attach_file.$input.on('change', function () {
            setTimeout(function () {
                selected_file = dialog.get_value('attach_file');
                check_and_show_import_button();
            }, 300);
        });
    }, 500);

    function check_and_show_import_button() {
        if (selected_file) {
            const import_html = `
                <div class="excel-import-controls-wrapper" style="
                    background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
                    padding: 20px;
                    border-radius: 12px;
                    margin-top: 10px;
                    box-shadow: 0 8px 16px rgba(0,0,0,0.1);
                ">
                    <button class="btn btn-primary btn-lg" id="excel-start-import-btn" style="
                        width: 100%;
                        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                        border: none;
                        padding: 12px;
                        font-size: 16px;
                        font-weight: 600;
                        border-radius: 8px;
                        box-shadow: 0 4px 12px rgba(102, 126, 234, 0.4);
                        transition: all 0.3s ease;
                    ">
                        <i class="fa fa-rocket"></i> Start Import
                    </button>
                    
                    <div id="excel-progress-container" style="display: none; margin-top: 20px;">
                        <div class="progress" style="
                            height: 30px;
                            border-radius: 15px;
                            background: #e0e7ff;
                            overflow: hidden;
                            box-shadow: inset 0 2px 4px rgba(0,0,0,0.1);
                        ">
                            <div class="progress-bar progress-bar-striped progress-bar-animated" 
                                 role="progressbar" 
                                 id="excel-import-progress-bar"
                                 style="
                                     width: 0%;
                                     background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
                                     font-size: 14px;
                                     font-weight: 600;
                                     line-height: 30px;
                                     box-shadow: 0 2px 8px rgba(102, 126, 234, 0.5);
                                 ">
                                0%
                            </div>
                        </div>
                        <div id="excel-progress-message" style="
                            margin-top: 12px;
                            font-size: 14px;
                            color: #4b5563;
                            text-align: center;
                            font-weight: 500;
                        "></div>
                    </div>
                </div>
            `;

            $('#excel-import-controls').html(import_html);

            // Add hover effect
            $('#excel-start-import-btn').hover(
                function () {
                    $(this).css('transform', 'translateY(-2px)');
                    $(this).css('box-shadow', '0 6px 16px rgba(102, 126, 234, 0.6)');
                },
                function () {
                    $(this).css('transform', 'translateY(0)');
                    $(this).css('box-shadow', '0 4px 12px rgba(102, 126, 234, 0.4)');
                }
            );

            // Bind click event
            $('#excel-start-import-btn').off('click').on('click', function () {
                start_import_process();
            });
        }
    }

    function start_import_process() {
        // First, validate and preview missing fields
        $('#excel-start-import-btn')
            .prop('disabled', true)
            .html('<i class="fa fa-spinner fa-spin"></i> Checking fields...')
            .css('opacity', '0.7');

        frappe.call({
            method: 'cipl_app.cipl_app.scripts.excel_import.validate_import_preview',
            args: {
                doctype_name: target_doctype,
                file_url: selected_file
            },
            callback: function (r) {
                // Re-enable button
                $('#excel-start-import-btn')
                    .prop('disabled', false)
                    .html('<i class="fa fa-rocket"></i> Start Import')
                    .css('opacity', '1');

                if (r.message && r.message.status === 'success') {
                    const preview = r.message;

                    // Show confirmation dialog with field info
                    let message_html = '<div style="font-size: 13px;">';

                    if (preview.mapped_fields && preview.mapped_fields.length > 0) {
                        message_html += `<p><strong style="color: #059669;">✓ ${preview.mapped_fields.length} fields will be imported:</strong></p>`;
                        message_html += '<ul style="max-height: 150px; overflow-y: auto; margin-left: 20px;">';
                        preview.mapped_fields.forEach(field => {
                            message_html += `<li style="color: #059669;">${field}</li>`;
                        });
                        message_html += '</ul>';
                    }

                    if (preview.missing_fields && preview.missing_fields.length > 0) {
                        message_html += `<p><strong style="color: #f59e0b;">⚠ ${preview.missing_fields.length} Excel columns not mapped (will be skipped):</strong></p>`;
                        message_html += '<ul style="max-height: 100px; overflow-y: auto; margin-left: 20px;">';
                        preview.missing_fields.forEach(field => {
                            message_html += `<li style="color: #f59e0b;">${field}</li>`;
                        });
                        message_html += '</ul>';
                    }

                    if (preview.missing_in_excel && preview.missing_in_excel.length > 0) {
                        message_html += `<p><strong style="color: #dc2626;">✗ ${preview.missing_in_excel.length} mapped fields not found in Excel:</strong></p>`;
                        message_html += '<ul style="max-height: 100px; overflow-y: auto; margin-left: 20px;">';
                        preview.missing_in_excel.forEach(field => {
                            message_html += `<li style="color: #dc2626;">${field}</li>`;
                        });
                        message_html += '</ul>';
                    }

                    message_html += '</div>';

                    // Add checkboxes for create/update options
                    message_html += `
                        <div style="margin-top: 15px; padding: 10px; background: #f3f4f6; border-radius: 6px;">
                            <p style="margin-bottom: 8px; font-weight: 600;">Import Options:</p>
                            <label style="display: block; margin-bottom: 5px; cursor: pointer;">
                                <input type="checkbox" id="allow-create-checkbox" checked style="margin-right: 8px;">
                                Create new records
                            </label>
                            <label style="display: block; cursor: pointer;">
                                <input type="checkbox" id="allow-update-checkbox" checked style="margin-right: 8px;">
                                Update existing records (using unique field)
                            </label>
                        </div>
                    `;

                    // Show confirmation dialog
                    frappe.confirm(
                        message_html,
                        function () {
                            // Get checkbox values
                            const allow_create = $('#allow-create-checkbox').is(':checked') ? 1 : 0;
                            const allow_update = $('#allow-update-checkbox').is(':checked') ? 1 : 0;

                            // User clicked "Yes" - proceed with import
                            proceed_with_import(allow_create, allow_update);
                        },
                        function () {
                            // User clicked "No" - do nothing
                            frappe.show_alert({
                                message: __('Import cancelled'),
                                indicator: 'orange'
                            }, 3);
                        }
                    );
                } else {
                    frappe.msgprint({
                        title: __('Validation Error'),
                        indicator: 'red',
                        message: r.message.message || __('Failed to validate file')
                    });
                }
            },
            error: function (r) {
                $('#excel-start-import-btn')
                    .prop('disabled', false)
                    .html('<i class="fa fa-rocket"></i> Start Import')
                    .css('opacity', '1');

                frappe.msgprint({
                    title: __('Validation Error'),
                    indicator: 'red',
                    message: r.message || __('Failed to validate file')
                });
            }
        });
    }


    function proceed_with_import(allow_create, allow_update) {
        // Disable button with loading state
        $('#excel-start-import-btn')
            .prop('disabled', true)
            .html('<i class="fa fa-spinner fa-spin"></i> Starting Import...')
            .css('opacity', '0.7');

        // Show progress container
        $('#excel-progress-container').slideDown(300);

        // Setup realtime listener (cleanup old one first)
        if (window.excel_import_state.realtime_listener) {
            frappe.realtime.off('excel_import_progress', window.excel_import_state.realtime_listener);
        }

        window.excel_import_state.realtime_listener = function (data) {
            update_progress(data);
        };

        frappe.realtime.on('excel_import_progress', window.excel_import_state.realtime_listener);

        // Call backend API
        frappe.call({
            method: 'cipl_app.cipl_app.scripts.excel_import.start_excel_import',
            args: {
                doctype_name: target_doctype,
                file_url: selected_file,
                allow_create: allow_create,
                allow_update: allow_update
            },
            callback: function (r) {
                if (r.message && r.message.status === 'success') {
                    frappe.show_alert({
                        message: __('Import started successfully'),
                        indicator: 'green'
                    }, 3);
                }
            },
            error: function (r) {
                frappe.msgprint({
                    title: __('Import Error'),
                    indicator: 'red',
                    message: r.message || __('Failed to start import')
                });

                // Reset button
                $('#excel-start-import-btn')
                    .prop('disabled', false)
                    .html('<i class="fa fa-rocket"></i> Start Import')
                    .css('opacity', '1');
                $('#excel-progress-container').hide();
            }
        });
    }

    function update_progress(data) {
        const progress = data.progress || 0;
        const message = data.message || '';
        const status = data.status || 'in_progress';

        // Update progress bar
        $('#excel-import-progress-bar')
            .css('width', progress + '%')
            .text(progress + '%');

        // Update message
        $('#excel-progress-message').text(message);

        // Handle completion
        if (status === 'completed') {
            $('#excel-import-progress-bar')
                .removeClass('progress-bar-animated progress-bar-striped')
                .css('background', 'linear-gradient(90deg, #10b981 0%, #059669 100%)');

            $('#excel-progress-message')
                .css('color', '#059669')
                .html('<i class="fa fa-check-circle"></i> ' + message);

            frappe.show_alert({
                message: message,
                indicator: 'green'
            }, 5);

            // Refresh current list view
            setTimeout(function () {
                if (cur_list && cur_list.doctype === target_doctype) {
                    cur_list.refresh();
                }
                window.excel_import_state.reset();
            }, 2000);

        } else if (status === 'error') {
            $('#excel-import-progress-bar')
                .removeClass('progress-bar-animated progress-bar-striped')
                .css('background', 'linear-gradient(90deg, #ef4444 0%, #dc2626 100%)');

            $('#excel-progress-message')
                .css('color', '#dc2626')
                .html('<i class="fa fa-exclamation-circle"></i> ' + message);

            frappe.msgprint({
                title: __('Import Error'),
                indicator: 'red',
                message: message
            });

            // Reset button
            $('#excel-start-import-btn')
                .prop('disabled', false)
                .html('<i class="fa fa-rocket"></i> Start Import')
                .css('opacity', '1');
        }
    }
};
