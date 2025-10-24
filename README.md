# Export Convert

A Python utility script to reshape Queue export files into the format expected by the appointment importer, with enhanced date handling for Jalali and Gregorian calendars.

## Features

### 📅 **Enhanced Date Handling**
- **Automatic Detection**: Automatically detects Jalali vs Gregorian dates
- **Accurate Conversion**: Uses `jdatetime` library for precise Jalali-Gregorian conversion
- **Time Handling**: Defaults to `00:00:00` when no time is specified
- **Multiple Formats**: Provides different date formats for database storage and UI display

### 👥 **Patient Data Processing**
- **Name Processing**: Splits full names into first and last names
- **Gender Detection**: Automatic gender detection using Persian names database
- **Data Validation**: Comprehensive validation for national IDs and mobile numbers
- **Deduplication**: Enhanced deduplication with name completion from historical records

### 🏥 **Clinic & Appointment Management**
- **Clinic Mapping**: Maps clinic names to standardized tags
- **Appointment Types**: Categorizes appointment types (internet, phone, etc.)
- **Status Tracking**: Tracks patient status (showed up, cancelled, etc.)

## Installation

### Prerequisites
- Python 3.8 or higher
- pip (Python package installer)

### Setup
1. Clone or download this repository
2. Create a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Basic Usage
```bash
python convert_excel.py input_file.xlsx
```

### Multiple Files
```bash
python convert_excel.py file1.xlsx file2.xlsx file3.xlsx
```

### Custom Output
```bash
python convert_excel.py input_file.xlsx -o output_file.xlsx
```

### Overwrite Existing Files
```bash
python convert_excel.py input_file.xlsx --overwrite
```

### Verbose Logging
```bash
python convert_excel.py input_file.xlsx --log-level DEBUG
```

## Input Format

The script expects Excel files with the following columns (in Persian):

| Persian Column | English Description | Required |
|----------------|-------------------|----------|
| کدملی | National ID | ✅ Yes |
| بیمار | Patient Full Name | ✅ Yes |
| موبایل | Mobile Number | ❌ Optional |
| تاریخ اخذ | Visit Date | ❌ Optional |
| وضعیت | Status | ❌ Optional |
| نوع | Appointment Type | ❌ Optional |
| درمانگاه | Clinic | ❌ Optional |

## Output Format

The script generates multiple output files:

### 1. **Main Output** (`*_cleaned.xlsx`)
Clean, validated records with complete data:
- `national_id`: Validated national ID
- `first_name`: Patient's first name
- `last_name`: Patient's last name
- `gender`: Detected gender (male/female)
- `mobile`: Cleaned mobile number
- `visit_date`: Database storage format (Gregorian)
- `visit_date_ui`: UI display format (Jalali)
- `visit_datetime_ui`: UI display with time (Jalali)
- `visit_date_db`: ISO format for database storage
- `tags`: Generated tags for categorization

### 2. **Excluded Records** (`*_excluded.xlsx`)
Records with invalid or incomplete data that couldn't be processed.

### 3. **Duplicate Phone Records** (`*_duplicate_phone.xlsx`)
Records with duplicate phone numbers for manual review.

### 4. **Incomplete Name Records** (`*_incomplete_name.xlsx`)
Records where names couldn't be completed from historical data.

## Date Handling

### Supported Date Formats

#### Jalali Dates (Persian Calendar)
- `1403/05/15` - Date only
- `1403/05/15 14:30` - Date with time
- `1403/12/29` - Leap year dates

#### Gregorian Dates
- `2024/08/06` - Date only
- `2024/08/06 14:30` - Date with time
- `2024/02/29` - Leap year dates

### Date Processing Flow

1. **Input Detection**: Automatically detects Jalali vs Gregorian dates
2. **Conversion**: Jalali dates are converted to Gregorian for database storage
3. **Storage**: All dates stored in Gregorian format for data integrity
4. **Display**: Dates converted back to Jalali for UI display

### Database Storage Best Practices

- **Storage Format**: All dates stored in Gregorian format (`YYYY-MM-DD`)
- **ISO Format**: Full ISO datetime format for precise storage
- **Query Format**: Standard format for database queries
- **UI Display**: Jalali format for user-friendly display

## Configuration

### Persian Names Database
The script uses two sources for Persian names gender detection:

1. **CSV File** (`iranian_names_full.csv`): Primary source with comprehensive names
2. **JSON File** (`persian_names_gender.json`): Secondary source with additional names

### Clinic Mapping
Clinics are automatically mapped to standardized tags:
- `bariatric_surgery_clinic`
- `cardiology_clinic`
- `dermatology_hair_aesthetics_clinic`
- And many more...

### Status Mapping
Patient statuses are mapped to tags:
- `not_showed_patient`: ثبت نوبت
- `showup_patient`: چاپ نوبت
- `canceling_patient`: کنسل شده

## Examples

### Example 1: Multiple Files
```bash
python convert_excel.py file1.xlsx file2.xlsx -o merged_output.xlsx
```
Output: `merged_output.xlsx`

### Example 2: Debug Mode
```bash
python convert_excel.py input.xlsx --log-level DEBUG
```

## Data Quality Features

### Validation Rules
- **National ID**: 8-11 digits, not all the same
- **Mobile Numbers**: Iranian mobile format validation
- **Names**: Minimum 3 characters, no special characters
- **Dates**: Valid Jalali or Gregorian dates

### Deduplication Logic
1. **Primary**: Keep most recent record per national ID
2. **Name Completion**: Use historical records to complete incomplete names
3. **Phone Duplicates**: Separate handling for duplicate phone numbers

## Error Handling

The script provides comprehensive error handling:
- **Invalid Dates**: Gracefully handles invalid date formats
- **Missing Data**: Handles missing required fields
- **Conversion Errors**: Fallback mechanisms for date conversion
- **File Errors**: Proper error messages for file operations

## Dependencies

- `pandas`: Data manipulation and analysis
- `openpyxl`: Excel file reading/writing
- `jdatetime`: Jalali calendar support

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

This project is licensed under the MIT License.

## Support

For issues and questions:
1. Check the error logs with `--log-level DEBUG`
2. Verify input file format matches expected columns
3. Ensure all dependencies are installed correctly

## Changelog

### Version 2.0
- ✅ Enhanced date handling with jdatetime library
- ✅ Automatic Jalali/Gregorian detection
- ✅ Multiple date format outputs
- ✅ Improved data validation
- ✅ Enhanced deduplication logic

### Version 1.0
- ✅ Basic Excel conversion
- ✅ Name processing
- ✅ Data validation
- ✅ Tag generation
