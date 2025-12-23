import frappe
from frappe import _
import pandas as pd
import numpy as np
from frappe.utils import now_datetime
import os


@frappe.whitelist()
@frappe.validate_and_sanitize_search_inputs
def get_cipl_app_doctypes(doctype, txt, searchfield, start, page_len, filters):
    """
    Get all doctypes from Cipl App module excluding system doctypes
    This is used as a query method for Link field
    """
    excluded_doctypes = ['excel_field_mapping', 'field_mapping_child']
    
    # Build filters
    filter_conditions = {
        'module': 'Cipl App',
        'istable': 0,
        'name': ['not in', excluded_doctypes]
    }
    
    # Add search text filter if provided
    if txt:
        filter_conditions['name'] = ['like', f'%{txt}%']
    
    doctypes = frappe.get_all(
        'DocType',
        filters=filter_conditions,
        fields=['name'],
        order_by='name',
        limit_start=start,
        limit_page_length=page_len
    )
    
    # Return in the format expected by query methods: list of tuples
    return [[d.name] for d in doctypes]




@frappe.whitelist()
def validate_import_preview(doctype_name, file_url):
    """
    Validate Excel file and return preview of missing fields
    """
    try:
        # Validate file
        file_path = validate_excel_file(file_url)
        
        # Get field mapping
        field_mapping = get_field_mapping(doctype_name)
        
        if not field_mapping:
            frappe.throw(_("No field mappings found for {0}").format(doctype_name))
        
        # Read Excel headers
        df = pd.read_excel(file_path, engine='openpyxl', nrows=0)
        excel_columns = df.columns.tolist()
        
        # Find missing fields (Excel columns not in mapping)
        missing_fields = []
        for col in excel_columns:
            if col not in field_mapping:
                missing_fields.append(col)
        
        # Find mapped fields that exist
        mapped_fields = []
        for excel_field in field_mapping.keys():
            if excel_field in excel_columns:
                mapped_fields.append(excel_field)
        
        # Find expected fields not in Excel
        missing_in_excel = []
        for excel_field in field_mapping.keys():
            if excel_field not in excel_columns:
                missing_in_excel.append(excel_field)
        
        return {
            'status': 'success',
            'total_excel_columns': len(excel_columns),
            'mapped_fields': mapped_fields,
            'missing_fields': missing_fields,
            'missing_in_excel': missing_in_excel,
            'has_missing': len(missing_fields) > 0 or len(missing_in_excel) > 0
        }
        
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), _("Import Preview Error"))
        return {
            'status': 'error',
            'message': str(e)
        }


@frappe.whitelist()
def start_excel_import(doctype_name, file_url, allow_create=1, allow_update=1):
    """
    Start the Excel import process
    """
    try:
        # Validate inputs
        if not doctype_name or not file_url:
            frappe.throw(_("Doctype name and file URL are required"))
        
        # Check if field mapping exists
        if not frappe.db.exists('excel_field_mapping', doctype_name):
            frappe.throw(_("Please map fields first for this doctype"))
        
        # Validate file
        file_path = validate_excel_file(file_url)
        
        # Get field mapping and unique field
        field_mapping = get_field_mapping(doctype_name)
        unique_field = get_unique_field(doctype_name)
        
        if not field_mapping:
            frappe.throw(_("No field mappings found for {0}").format(doctype_name))
        
        # Enqueue background job
        frappe.enqueue(
            process_excel_import,
            queue='long',
            timeout=3000,
            doctype_name=doctype_name,
            file_path=file_path,
            field_mapping=field_mapping,
            unique_field=unique_field,
            file_url=file_url,
            allow_create=int(allow_create),
            allow_update=int(allow_update),
            user=frappe.session.user
        )
        
        return {
            'status': 'success',
            'message': _('Import started successfully')
        }
        
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), _("Excel Import Start Error"))
        frappe.throw(str(e))


def validate_excel_file(file_url):
    """
    Validate Excel file and return file path
    """
    # Remove leading slash if present
    file_url = file_url.lstrip('/')
    
    # Get site path
    site_path = frappe.get_site_path()
    
    # Construct full file path
    if file_url.startswith('private/files/') or file_url.startswith('files/'):
        file_path = os.path.join(site_path, file_url)
    else:
        file_path = os.path.join(site_path, 'private', 'files', os.path.basename(file_url))
    
    # Check if file exists
    if not os.path.exists(file_path):
        # Try public files
        file_path = os.path.join(site_path, 'public', 'files', os.path.basename(file_url))
        if not os.path.exists(file_path):
            frappe.throw(_("File not found: {0}").format(file_url))
    
    # Validate file extension
    valid_extensions = ['.xlsx', '.xls']
    file_ext = os.path.splitext(file_path)[1].lower()
    
    if file_ext not in valid_extensions:
        frappe.throw(_("Invalid file format. Please upload Excel file (.xlsx or .xls)"))
    
    return file_path


def get_field_mapping(doctype_name):
    """
    Get field mapping for the doctype
    Returns dict: {excel_field_name: doctype_field_name}
    """
    try:
        mapping_doc = frappe.get_doc('excel_field_mapping', doctype_name)
        
        field_mapping = {}
        for row in mapping_doc.mapped_field:
            field_mapping[row.excel_field_name] = row.doctype_field_name
        
        return field_mapping
        
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), _("Get Field Mapping Error"))
        return {}


def get_unique_field(doctype_name):
    """
    Get unique field for duplicate detection
    """
    try:
        mapping_doc = frappe.get_doc('excel_field_mapping', doctype_name)
        return mapping_doc.unique_field if mapping_doc.unique_field else None
    except:
        return None


def process_excel_import(doctype_name, file_path, field_mapping, unique_field, file_url, allow_create, allow_update, user):
    """
    Process Excel import in background using pandas for better performance
    """
    frappe.set_user(user)
    
    try:
        # Load Excel file using pandas (much faster than openpyxl)
        df = pd.read_excel(file_path, engine='openpyxl')
        
        # Get total rows
        total_rows = len(df)
        
        if total_rows <= 0:
            publish_progress(0, 0, "No data found in Excel file", "error")
            return
        
        # Process rows
        success_count = 0
        created_count = 0
        updated_count = 0
        skipped_count = 0
        error_count = 0
        skipped_fields_log = set()  # Use set to avoid duplicates
        
        # Get Excel column names
        excel_columns = df.columns.tolist()
        
        # Identify which Excel columns are not in mapping
        for col in excel_columns:
            if col not in field_mapping:
                skipped_fields_log.add(col)
        
        # Iterate through DataFrame rows
        for idx, row in df.iterrows():
            try:
                # Create document data
                doc_data = {'doctype': doctype_name}
                unique_value = None
                
                # Map Excel columns to doctype fields
                for excel_field, doctype_field in field_mapping.items():
                    if excel_field in excel_columns:
                        value = row[excel_field]
                        
                        # Handle NaN values (pandas uses NaN for empty cells)
                        if pd.notna(value):
                            # Convert numpy types to Python types
                            if isinstance(value, (np.integer, np.floating)):
                                value = value.item()
                            elif isinstance(value, pd.Timestamp):
                                value = value.to_pydatetime()
                            
                            doc_data[doctype_field] = value
                            
                            # Track unique field value
                            if unique_field and doctype_field == unique_field:
                                unique_value = value
                
                # Check if record exists (if unique_field is set)
                existing_doc = None
                if unique_field and unique_value:
                    existing_doc = frappe.db.get_value(
                        doctype_name, 
                        {unique_field: unique_value}, 
                        'name'
                    )
                
                # Decide action based on flags
                if existing_doc:
                    # Record exists
                    if allow_update:
                        # Update existing record
                        doc = frappe.get_doc(doctype_name, existing_doc)
                        for field, value in doc_data.items():
                            if field != 'doctype':
                                setattr(doc, field, value)
                        doc.save(ignore_permissions=True)
                        updated_count += 1
                        success_count += 1
                    else:
                        # Skip update (intentional, no error)
                        skipped_count += 1
                else:
                    # Record doesn't exist
                    if allow_create:
                        # Create new record
                        doc = frappe.get_doc(doc_data)
                        doc.insert(ignore_permissions=True)
                        created_count += 1
                        success_count += 1
                    else:
                        # Skip create (intentional, no error)
                        skipped_count += 1
                
                # Publish progress every 10 rows or on last row
                if (idx + 1) % 10 == 0 or (idx + 1) == total_rows:
                    progress = int(((idx + 1) / total_rows) * 100)
                    publish_progress(
                        progress, 
                        total_rows, 
                        f"Processing row {idx + 1} of {total_rows}",
                        "in_progress"
                    )
                
            except Exception as e:
                error_count += 1
                frappe.log_error(
                    f"Row {idx + 1}: {str(e)}\n{frappe.get_traceback()}", 
                    f"Excel Import Row Error - {doctype_name}"
                )
                
                # Continue processing other rows
                continue
        
        # Commit changes
        frappe.db.commit()
        
        # Delete the uploaded file after successful import
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                
                # Also delete the File doctype record if exists
                file_name = os.path.basename(file_url)
                if frappe.db.exists('File', {'file_url': file_url}):
                    frappe.delete_doc('File', frappe.db.get_value('File', {'file_url': file_url}, 'name'))
                    frappe.db.commit()
        except Exception as e:
            frappe.log_error(f"File deletion error: {str(e)}", "Excel Import File Cleanup")
        
        
        # Final message - always show created, updated, skipped counts
        message = f"Import completed: {created_count} created, {updated_count} updated, {skipped_count} skipped"
        if error_count > 0:
            message += f", {error_count} errors"
        
        publish_progress(100, total_rows, message, "completed")
        
    except Exception as e:
        frappe.log_error(frappe.get_traceback(), f"Excel Import Error - {doctype_name}")
        publish_progress(0, 0, f"Import failed: {str(e)}", "error")
        frappe.db.rollback()


def publish_progress(progress, total, message, status):
    """
    Publish real-time progress updates
    """
    frappe.publish_realtime(
        'excel_import_progress',
        {
            'progress': progress,
            'total': total,
            'message': message,
            'status': status
        },
        user=frappe.session.user
    )
