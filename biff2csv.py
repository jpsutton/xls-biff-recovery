#!/usr/bin/env python

import struct
import csv
import logging
import argparse
import os
from datetime import datetime, timedelta
from enum import Enum

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("biff_converter.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("BIFFConverter")

class BIFFRecordType(Enum):
    # Workbook/Worksheet Structure
    BOF = 0x0809              # Beginning of file/worksheet
    EOF = 0x000A              # End of file/worksheet
    BOUNDSHEET = 0x0085       # Sheet information
    INDEX = 0x020B            # Index record
    DIMENSION = 0x0200        # Worksheet dimensions
    WINDOW1 = 0x003D          # Workbook window information
    WINDOW2 = 0x023E          # Worksheet window information

    # Cell Content Records
    INTEGER_CELL = 0x0002     # Integer cell
    NUMBER_CELL = 0x0203      # Float cell
    LABEL_CELL = 0x0204       # String cell (up to 255 chars)
    FORMULA_CELL = 0x0006     # Formula cell
    BOOLERR = 0x0205          # Boolean or error value
    BLANK = 0x0201            # Blank cell
    MULRK = 0x00BD            # Multiple RK values
    RK = 0x027E               # RK value (compressed integer or float)
    LABELSST = 0x00FD         # String cell with SST index
    MULBLANK = 0x00BE         # Multiple blank cells
    RSTRING = 0x00D6          # Rich string cell

    # Shared Strings
    SST = 0x00FC              # Shared string table
    CONTINUE = 0x003C         # Continues a previous record
    STRING_FORMULA_RESULT = 0x0207  # Result from string formula

    # Formatting Records
    FORMAT = 0x041E           # Number format definition
    XF = 0x00E0               # Extended format
    FONT = 0x0031             # Font definition
    STYLE = 0x0293            # Style definition
    PALETTE = 0x0092          # Color palette

    # Row/Column Information
    ROW = 0x0208              # Row record
    COLINFO = 0x007D          # Column information
    DEFAULTROWHEIGHT = 0x0225 # Default row height
    DEFCOLWIDTH = 0x0055      # Default column width

    # Named Ranges/Areas
    NAME = 0x0018             # Named range
    ARRAY = 0x0221            # Array formula

    # Protection Records
    PROTECT = 0x0012          # Protection flag
    PASSWORD = 0x0013         # Password record

    # Miscellaneous
    EXTERNSHEET = 0x0017      # External sheet references
    EXTERNNAME = 0x0023       # External name
    MERGEDCELLS = 0x00E5      # Merged cell ranges
    HYPERLINK = 0x01B8        # Hyperlink
    COUNTRY = 0x008C          # Country code
    HEADER = 0x0014           # Page header
    FOOTER = 0x0015           # Page footer
    CODEPAGE = 0x0042         # Code page
    FILEPASS = 0x002F         # File protection
    CALCMODE = 0x000D         # Calculation mode
    PRECISION = 0x000E        # Precision flag
    BACKUP = 0x0040           # Backup flag

    # Internal Records
    UNCALCED = 0x005E         # Uncalculated formulas
    SAVERECALC = 0x005F       # Recalculate on save
    BOOKBOOL = 0x00DA         # Workbook options

    # Worksheet Graphics
    OBJPROJ = 0x00D3          # Object protection
    MSODRAWING = 0x00EC       # Office drawing object
    MSODRAWINGGROUP = 0x00EB  # Office drawing group

# Function to convert Excel date serial number to datetime
def excel_date_to_datetime(excel_date):
    """Convert Excel date value to Python datetime"""
    if not isinstance(excel_date, (int, float)):
        return None
    
    try:
        # Excel's date system starts on January 0, 1900, which is actually December 31, 1899
        base_date = datetime(1899, 12, 30)
        
        # Excel has a leap year bug: it thinks 1900 was a leap year
        days = int(excel_date)
        if days > 60:  # February 29, 1900 (which doesn't exist)
            days -= 1
        
        date_value = base_date + timedelta(days=days)
        
        # Handle time component (fractional part)
        frac = excel_date - int(excel_date)
        if frac > 0:
            seconds = int(frac * 86400)  # 86400 seconds in a day
            hours, remainder = divmod(seconds, 3600)
            minutes, seconds = divmod(remainder, 60)
            date_value = date_value.replace(hour=hours, minute=minutes, second=seconds)
        
        return date_value
    except:
        return None

class BIFFConverter:
    def __init__(self, input_file, output_dir, date_columns=None):
        self.input_file = input_file
        self.output_dir = output_dir
        self.data = None
        self.position = 0
        self.worksheets = []
        self.current_worksheet = None
        self.shared_strings = []
        self.max_rows = 0
        self.max_cols = 0
        self.cell_data = {}  # Format: {(row, col): value}
        self.cell_formats = {}  # Format: {(row, col): format_id}
        self.format_codes = {}  # Format: {format_id: format_code}
        # List of column indices that should be forced as dates
        self.date_columns = date_columns or []
        self.xf_records = list()
        self.date_formats = set()

    def set_cell_value(self, row, col, value, type_code=None, raw_value=None):
        """
        Helper method to set cell value and log debug information
        
        Args:
            row: Row index (0-based)
            col: Column index (0-based)
            value: The processed value to store
            type_code: The BIFFRecordType or other identifier of the data type
            raw_value: The raw value before conversion (for debug purposes)
        """
        # Convert row/col to Excel-style reference (e.g., A1, B2)
        col_letter = ""
        temp_col = col
        while temp_col >= 0:
            col_letter = chr(65 + (temp_col % 26)) + col_letter
            temp_col = (temp_col // 26) - 1
        cell_ref = f"{col_letter}{row+1}"
        
        # Get human-readable type name
        type_name = "UNKNOWN"
        if type_code is not None:
            if isinstance(type_code, BIFFRecordType):
                type_name = type_code.name
            elif isinstance(type_code, int) and type_code in [e.value for e in BIFFRecordType]:
                type_name = BIFFRecordType(type_code).name
            else:
                type_name = str(type_code)
        
        # Store the value
        self.cell_data[(row, col)] = value
        
        # Log debug information
        raw_value_hex = ""
        if raw_value is not None:
            if isinstance(raw_value, bytes):
                raw_value_hex = raw_value.hex()
            else:
                try:
                    raw_value_hex = f"{raw_value:x}" if isinstance(raw_value, int) else str(raw_value)
                except:
                    raw_value_hex = str(raw_value)
                    
        logger.debug(f"Cell {cell_ref} ({row},{col}) set to {value} [Type: {type_name}, Raw: {raw_value_hex}]")

    def read_biff_file(self):
        """Read the BIFF file into memory"""
        try:
            with open(self.input_file, 'rb') as f:
                self.data = f.read()
            logger.info(f"Successfully read BIFF file: {self.input_file}, size: {len(self.data)} bytes")
            return True
        except Exception as e:
            logger.error(f"Failed to read BIFF file: {e}")
            return False

    def read_record_header(self):
        """Read a BIFF record header (type and size)"""
        if self.position + 4 > len(self.data):
            return None, 0

        try:
            record_type, record_size = struct.unpack("<HH", self.data[self.position:self.position+4])
            self.position += 4
            return record_type, record_size
        except struct.error as e:
            logger.error(f"Error reading record header at position {self.position}: {e}")
            return None, 0

    def read_bytes(self, size):
        """Read a specified number of bytes from the current position"""
        if self.position + size > len(self.data):
            logger.warning(f"Attempted to read past end of file at position {self.position}, size {size}")
            return None

        data = self.data[self.position:self.position+size]
        self.position += size
        return data

    def skip_bytes(self, size):
        """Skip a specified number of bytes"""
        self.position += size
        if self.position > len(self.data):
            logger.warning(f"Skipped past end of file to position {self.position}")
            self.position = len(self.data)

    def process_bof_record(self, size):
        """Process BOF record to identify sheet structure"""
        if size < 8:
            logger.warning(f"BOF record too small: {size} bytes")
            self.skip_bytes(size)
            return

        data = self.read_bytes(size)
        if data is None:
            return

        version, doc_type = struct.unpack("<HH", data[0:4])
        logger.debug(f"BOF record: version={version}, type={doc_type}")

        # Doc type: 5 = workbook, 10 = worksheet
        if doc_type == 10:
            logger.info(f"Found worksheet BOF at position {self.position - size - 4}")
            if self.current_worksheet:
                logger.debug(f"Setting current worksheet: {self.current_worksheet}")

    def process_dimensions_record(self, size):
        """Process DIMENSION record to get worksheet dimensions"""
        if size < 14:
            logger.warning(f"DIMENSION record too small: {size} bytes")
            self.skip_bytes(size)
            return

        data = self.read_bytes(size)
        if data is None:
            return

        # Unpack dimensions
        first_row, last_row, first_col, last_col = struct.unpack("<IIHH", data[0:12])

        logger.info(f"Worksheet dimensions: rows={first_row}-{last_row}, cols={first_col}-{last_col}")

        # Update max dimensions
        self.max_rows = max(self.max_rows, last_row + 1)
        self.max_cols = max(self.max_cols, last_col + 1)

    def process_boundsheet_record(self, size):
        """Process BOUNDSHEET record to get worksheet names"""
        if size < 6:
            logger.warning(f"BOUNDSHEET record too small: {size} bytes")
            self.skip_bytes(size)
            return

        data = self.read_bytes(size)
        if data is None:
            return

        # Position of BOF record for the sheet
        position = struct.unpack("<I", data[0:4])[0]

        # Sheet type and visibility
        sheet_type = data[4]
        visibility = data[5]

        # Sheet name (length in first byte, followed by name)
        name_length = data[6]
        sheet_name = data[7:7+name_length].decode('utf-8', errors='replace')

        logger.info(f"Found worksheet: '{sheet_name}', position={position}, type={sheet_type}, visibility={visibility}")
        self.worksheets.append({
            'name': sheet_name,
            'position': position,
            'type': sheet_type,
            'visibility': visibility
        })

    def process_sst_record(self, size):
        """Process SST (Shared String Table) record"""
        data = self.read_bytes(size)
        if data is None:
            return

        if size < 8:
            logger.warning(f"SST record too small: {size} bytes")
            return

        # Total strings and unique strings
        total_strings, unique_strings = struct.unpack("<II", data[0:8])
        logger.info(f"SST: {unique_strings} unique strings, {total_strings} total")

        # SST processing is complex due to CONTINUE records
        # This is a simplified implementation
        position = 8
        strings_processed = 0

        while strings_processed < unique_strings and position < size:
            if position + 3 > size:
                logger.warning(f"Reached end of SST record while processing strings")
                break

            # String length is 2 bytes, followed by option flags
            str_length = struct.unpack("<H", data[position:position+2])[0]
            options = data[position+2]
            position += 3

            compressed = (options & 0x01) == 0
            bytes_per_char = 1 if compressed else 2

            # Check if we have enough data
            if position + (str_length * bytes_per_char) > size:
                logger.warning(f"String length exceeds available data. Expected {str_length * bytes_per_char} bytes, have {size - position}")
                break

            # Extract the string
            if compressed:
                string_value = data[position:position+str_length].decode('latin1', errors='replace')
            else:
                string_value = data[position:position+(str_length*2)].decode('utf-16-le', errors='replace')

            position += str_length * bytes_per_char
            self.shared_strings.append(string_value)
            strings_processed += 1

        logger.info(f"Processed {strings_processed} out of {unique_strings} shared strings")

    def process_format_record(self, size):
        """Process FORMAT record to get cell format codes"""
        if size < 4:
            logger.warning(f"FORMAT record too small: {size} bytes")
            self.skip_bytes(size)
            return

        data = self.read_bytes(size)
        if data is None:
            return

        # Format code index
        format_index = struct.unpack("<H", data[0:2])[0]
        
        # Format string length
        format_length = struct.unpack("<H", data[2:4])[0]
        
        if 4 + format_length > size:
            logger.warning(f"FORMAT: String length exceeds record size")
            return
            
        # Format string
        format_string = data[4:4+format_length].decode('latin1', errors='replace')
        
        self.format_codes[format_index] = format_string
        logger.debug(f"FORMAT: index={format_index}, format='{format_string}'")

    def process_xf_record(self, record_size):
        xf_data = self.data[self.position:self.position+record_size]
        
        # XF index is the position in the array
        xf_index = len(self.xf_records)
        
        # Format index is at offset 2, 2 bytes
        if len(xf_data) >= 4:
            format_index = struct.unpack('<H', xf_data[2:4])[0]
            
            # Built-in Excel date format numbers (14-22 are date formats)
            if 14 <= format_index <= 22:
                self.date_formats.add(xf_index)
            
            # Also check custom formats in self.format_codes
            if format_index in self.format_codes:
                format_code = self.format_codes[format_index]
                # Check if this is likely a date format (contains y, m, d, h, etc.)
                date_indicators = ['yy', 'mm', 'dd', 'h', 'm', 's', 'am/pm', 'AM/PM']
                if any(indicator in format_code.lower() for indicator in date_indicators):
                    self.date_formats.add(xf_index)
        
        self.xf_records.append(xf_data)
        self.position += record_size

    def process_labelsst_record(self, size):
        """Process LABELSST (label with shared string) record"""
        if size < 10:
            logger.warning(f"LABELSST record too small: {size} bytes")
            self.skip_bytes(size)
            return

        data = self.read_bytes(size)
        if data is None:
            return

        row, col, xf = struct.unpack("<HHH", data[0:6])
        sst_index = struct.unpack("<I", data[6:10])[0]

        if sst_index < len(self.shared_strings):
            cell_value = self.shared_strings[sst_index]
            self.set_cell_value(row, col, cell_value, BIFFRecordType.LABELSST, data[6:10])
            self.cell_formats[(row, col)] = xf
            logger.debug(f"LABELSST: row={row}, col={col}, value='{cell_value}'")
        else:
            logger.warning(f"LABELSST: Invalid SST index {sst_index} at row={row}, col={col}")

    def process_label_record(self, size):
        """Process LABEL (non-shared string) record"""
        if size < 8:
            logger.warning(f"LABEL record too small: {size} bytes")
            self.skip_bytes(size)
            return

        data = self.read_bytes(size)
        if data is None:
            return

        row, col, xf = struct.unpack("<HHH", data[0:6])
        str_length = struct.unpack("<H", data[6:8])[0]

        if 8 + str_length > size:
            logger.warning(f"LABEL: String length exceeds record size at row={row}, col={col}")
            return

        cell_value = data[8:8+str_length].decode('latin1', errors='replace')
        self.set_cell_value(row, col, cell_value, BIFFRecordType.LABEL_CELL, data[8:8+str_length])
        self.cell_formats[(row, col)] = xf
        logger.debug(f"LABEL: row={row}, col={col}, value='{cell_value}'")

    def process_number_record(self, record_size):
        data = self.data[self.position:self.position+record_size]
        
        # Extract row, column, and XF info
        row = struct.unpack('<H', data[0:2])[0]
        col = struct.unpack('<H', data[2:4])[0]
        xf = struct.unpack('<H', data[4:6])[0]
        
        # Extract the IEEE-754 double-precision floating point value
        value = struct.unpack('<d', data[6:14])[0]
        
        # Move position past this record
        self.position += record_size
        
        # Check if this is a date format based on column or XF format
        is_date_format = (col in self.date_columns) or (xf in self.date_formats)
        
        if is_date_format:
            # Check if this is a date format and convert
            date_value = excel_date_to_datetime(value)
            if isinstance(date_value, datetime):
                self.set_cell_value(row, col, date_value.strftime("%Y-%m-%d %H:%M:%S"), BIFFRecordType.NUMBER_CELL, data[6:14])
                logger.debug(f"Converted date at ({row}, {col}): {date_value.strftime('%Y-%m-%d %H:%M:%S')}")
            else:
                # Handle integer values nicely
                if value == int(value):
                    value = int(value)
                self.set_cell_value(row, col, value, BIFFRecordType.NUMBER_CELL, data[6:14])
        else:
            # Handle integer values nicely
            if value == int(value):
                value = int(value)
            self.set_cell_value(row, col, value, BIFFRecordType.NUMBER_CELL, data[6:14])
            
        self.cell_formats[(row, col)] = xf
        logger.debug(f"NUMBER: row={row}, col={col}, final value={self.cell_data[(row, col)]}")

    def process_rk_record(self, record_size):
        data = self.data[self.position:self.position+record_size]
        
        # Extract row, column, and XF info
        row = struct.unpack('<H', data[0:2])[0]
        col = struct.unpack('<H', data[2:4])[0]
        xf = struct.unpack('<H', data[4:6])[0]
        rk_value = struct.unpack('<I', data[6:10])[0]
        
        # Move position past this record
        self.position += record_size
        
        # Decode the RK value
        value = self.decode_rk(rk_value)
        
        # Check if this is a date format based on column or XF format
        is_date_format = (col in self.date_columns) or (xf in self.date_formats)
        
        if is_date_format or col in self.date_columns:
            # Check if this might be a date
            date_value = excel_date_to_datetime(value)
            if isinstance(date_value, datetime):
                self.set_cell_value(row, col, date_value.strftime("%Y-%m-%d %H:%M:%S"), BIFFRecordType.RK, rk_value)
            else:
                self.set_cell_value(row, col, value, BIFFRecordType.RK, rk_value)
        else:
            self.set_cell_value(row, col, value, BIFFRecordType.RK, rk_value)
            
        self.cell_formats[(row, col)] = xf
        logger.debug(f"RK: row={row}, col={col}, final value={self.cell_data[(row, col)]}")

    def decode_rk(self, rk_value):
        """
        Decode an RK value according to Excel BIFF specification.
        
        Args:
            rk_value: The 4-byte RK value (as integer)
        
        Returns:
            The decoded number
        """
        # Check if the number is an integer (bit 0)
        is_integer = (rk_value & 0x01) != 0
        
        # Check if the number should be divided by 100 (bit 1)
        is_divided_by_100 = (rk_value & 0x02) != 0
        
        # Clear the last 2 bits
        value = rk_value & 0xFFFFFFFC  # Mask out the last 2 bits
        
        if is_integer:
            # Handle as 30-bit signed integer (shift right to remove the 2 flag bits)
            value = value >> 2
            # Convert to signed integer if necessary
            if value & 0x20000000:  # Check if the sign bit is set
                value = value | 0xC0000000  # Sign extend to 32 bits
            result = value
        else:
            # Handle as IEEE-754 double
            # For floating-point, don't shift - the upper 30 bits are already in place
            # The 30 significant bits are the high bytes of an 8-byte IEEE double
            double_bytes = bytearray(8)
            # Copy the upper 4 bytes with the 2 LSBs cleared
            double_bytes[4:8] = value.to_bytes(4, byteorder='little')
            # Lower 4 bytes are all zeros
            double_bytes[0:4] = b'\x00\x00\x00\x00'
            result = struct.unpack('<d', bytes(double_bytes))[0]
        
        if is_divided_by_100:
            result /= 100
        
        return result
    
    def col_to_excel_col(self, col):
        """
        Convert a zero-based column index to Excel column letters (A, B, ..., Z, AA, AB, ...)
        
        Args:
            col: The zero-based column index
            
        Returns:
            A string representing the Excel column letter(s)
        """
        col_letter = ""
        temp_col = col
        while temp_col >= 0:
            col_letter = chr(65 + (temp_col % 26)) + col_letter
            temp_col = (temp_col // 26) - 1
        return col_letter

    def process_mulrk_record(self, record_size):
        data = self.data[self.position:self.position+record_size]
        
        # First row field (2 bytes)
        row = struct.unpack('<H', data[0:2])[0]
        
        # First column field (2 bytes)
        first_col = struct.unpack('<H', data[2:4])[0]
        
        # Last column field (2 bytes) - at the end of the record
        last_col = struct.unpack('<H', data[record_size-2:record_size])[0]
        
        # Move position past the header
        pos = 4
        
        # Process each RK value
        for current_col in range(first_col, last_col + 1):
            # XF record (2 bytes)
            xf = struct.unpack('<H', data[pos:pos+2])[0]
            pos += 2
            
            # RK value (4 bytes)
            rk_bytes = data[pos:pos+4]
            rk_value = struct.unpack('<I', rk_bytes)[0]
            pos += 4
            
            # Decode the RK value
            value = self.decode_rk(rk_value)
            
            # Get Excel-style cell reference
            excel_ref = f"{self.col_to_excel_col(current_col)}{row+1}"
            
            # Add additional checks to determine if this should be a date
            # Is this column explicitly marked as a date column?
            is_date_column = current_col in self.date_columns
            
            # Is this cell's format (XF) marked as a date format?
            is_date_format = xf in self.date_formats
            
            # For more intelligent detection, consider checking if the value is likely a date
            # A date near 1.0 is likely a small integer rather than a date
            is_small_integer = (isinstance(value, (int, float)) and abs(value) < 10 and value == int(value))
            
            # Only convert to date if explicitly requested AND not likely a small integer
            if (is_date_column or is_date_format) and not is_small_integer:
                date_value = excel_date_to_datetime(value)
                if isinstance(date_value, datetime):
                    self.set_cell_value(row, current_col, date_value.strftime("%Y-%m-%d %H:%M:%S"), 
                                        BIFFRecordType.MULRK, rk_bytes.hex())
                    logger.debug(f"Converted to date: Cell {excel_ref} ({row},{current_col}): {date_value}")
                else:
                    self.set_cell_value(row, current_col, value, BIFFRecordType.MULRK, rk_bytes.hex())
            else:
                # Not a date or a small integer that shouldn't be treated as a date
                self.set_cell_value(row, current_col, value, BIFFRecordType.MULRK, rk_bytes.hex())
            
            self.cell_formats[(row, current_col)] = xf
        
        # Move position past this record
        self.position += record_size

    def process_blank_record(self, size):
        """Process BLANK record (empty cell)"""
        if size < 6:
            logger.warning(f"BLANK record too small: {size} bytes")
            self.skip_bytes(size)
            return

        data = self.read_bytes(size)
        if data is None:
            return

        row, col, xf = struct.unpack("<HHH", data[0:6])
        self.set_cell_value(row, col, "", BIFFRecordType.BLANK, None)
        self.cell_formats[(row, col)] = xf
        logger.debug(f"BLANK: row={row}, col={col}")

    def process_boolerr_record(self, size):
        """Process BOOLERR record (boolean or error value)"""
        if size < 8:
            logger.warning(f"BOOLERR record too small: {size} bytes")
            self.skip_bytes(size)
            return

        data = self.read_bytes(size)
        if data is None:
            return

        row, col, xf = struct.unpack("<HHH", data[0:6])
        value = data[6]
        is_error = data[7]

        if is_error:
            error_values = {
                0x00: "#NULL!",
                0x07: "#DIV/0!",
                0x0F: "#VALUE!",
                0x17: "#REF!",
                0x1D: "#NAME?",
                0x24: "#NUM!",
                0x2A: "#N/A",
                0x2B: "#GETTING_DATA"
            }
            cell_value = error_values.get(value, f"#ERROR:{value}")
        else:
            cell_value = bool(value)

        self.set_cell_value(row, col, cell_value, BIFFRecordType.BOOLERR, data[6:8])
        self.cell_formats[(row, col)] = xf
        logger.debug(f"BOOLERR: row={row}, col={col}, value={cell_value}, is_error={is_error}")

    def process_biff_stream(self):
        """Process the entire BIFF stream"""
        if not self.read_biff_file():
            return False

        logger.info("Starting to process BIFF stream...")

        # First pass - collect workbook info
        while self.position < len(self.data):
            record_type, record_size = self.read_record_header()
            if record_type is None:
                break

            record_type_name = BIFFRecordType(record_type).name if record_type in [e.value for e in BIFFRecordType] else f"UNKNOWN(0x{record_type:04X})"
            logger.debug(f"Record: {record_type_name}, size={record_size}, position={self.position-4}")

            # Process specific record types
            try:
                if record_type == BIFFRecordType.BOF.value:
                    self.process_bof_record(record_size)
                elif record_type == BIFFRecordType.EOF.value:
                    self.read_bytes(record_size)  # Usually empty
                elif record_type == BIFFRecordType.DIMENSION.value:
                    self.process_dimensions_record(record_size)
                elif record_type == BIFFRecordType.BOUNDSHEET.value:
                    self.process_boundsheet_record(record_size)
                elif record_type == BIFFRecordType.SST.value:
                    self.process_sst_record(record_size)
                elif record_type == BIFFRecordType.FORMAT.value:
                    self.process_format_record(record_size)
                elif record_type == BIFFRecordType.XF.value:
                    self.process_xf_record(record_size)
                elif record_type == BIFFRecordType.LABELSST.value:
                    self.process_labelsst_record(record_size)
                elif record_type == BIFFRecordType.LABEL_CELL.value:
                    self.process_label_record(record_size)
                elif record_type == BIFFRecordType.NUMBER_CELL.value:
                    self.process_number_record(record_size)
                elif record_type == BIFFRecordType.RK.value:
                    self.process_rk_record(record_size)
                elif record_type == BIFFRecordType.MULRK.value:
                    self.process_mulrk_record(record_size)
                elif record_type == BIFFRecordType.BLANK.value:
                    self.process_blank_record(record_size)
                elif record_type == BIFFRecordType.BOOLERR.value:
                    self.process_boolerr_record(record_size)
                else:
                    # Skip records we don't process
                    self.skip_bytes(record_size)
            except Exception as e:
                logger.error(f"Error processing record type 0x{record_type:04X}: {e}")
                self.skip_bytes(record_size)

        logger.info(f"BIFF stream processing complete. Found {len(self.worksheets)} worksheets, {len(self.cell_data)} cells.")
        return True

    def write_csv(self):
        """Write the data to CSV files, one per worksheet"""
        logger.info(f"Writing data to CSV files in {self.output_dir}")
        os.makedirs(self.output_dir, exist_ok=True)

        # If no worksheets were detected, create a default one
        if not self.worksheets:
            logger.warning("No worksheets detected, creating a default worksheet")
            self.worksheets.append({'name': 'Sheet1'})

        # If we have cell data but no max dimensions, estimate them
        if self.cell_data and (self.max_rows == 0 or self.max_cols == 0):
            max_row = max([row for row, _ in self.cell_data.keys()], default=0)
            max_col = max([col for _, col in self.cell_data.keys()], default=0)
            self.max_rows = max(self.max_rows, max_row + 1)
            self.max_cols = max(self.max_cols, max_col + 1)
            logger.info(f"Estimated dimensions from cell data: rows={self.max_rows}, cols={self.max_cols}")

        # For simplicity, we'll output all data to a single worksheet for now
        worksheet_name = self.worksheets[0]['name'] if self.worksheets else 'Sheet1'
        output_file = os.path.join(self.output_dir, f"{worksheet_name}.csv")

        try:
            with open(output_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)

                # Generate rows
                for row in range(self.max_rows):
                    csv_row = []
                    for col in range(self.max_cols):
                        cell_value = self.cell_data.get((row, col), "")
                        csv_row.append(cell_value)
                    writer.writerow(csv_row)

            logger.info(f"Successfully wrote CSV to {output_file}")
            return True
        except Exception as e:
            logger.error(f"Error writing CSV file: {e}")
            return False

def main():
    parser = argparse.ArgumentParser(description='Convert BIFF stream to CSV')
    parser.add_argument('input_file', help='Path to the BIFF stream file')
    parser.add_argument('output_dir', help='Directory to write CSV files to')
    parser.add_argument('--debug', action='store_true', help='Enable debug logging')
    parser.add_argument('--date-columns', type=str, help='Comma-separated list of column indices (0-based) to force date conversion, e.g. "2,5,7"')
    args = parser.parse_args()

    if args.debug:
        logger.setLevel(logging.DEBUG)

    start_time = datetime.now()
    logger.info(f"Starting BIFF to CSV conversion at {start_time}")
    logger.info(f"Input file: {args.input_file}")
    logger.info(f"Output directory: {args.output_dir}")
    
    # Process date columns if provided
    date_columns = []
    if args.date_columns:
        try:
            date_columns = [int(col.strip()) for col in args.date_columns.split(',')]
            logger.info(f"Forcing date conversion for columns: {date_columns}")
        except ValueError as e:
            logger.error(f"Error parsing date columns: {e}")

    converter = BIFFConverter(args.input_file, args.output_dir, date_columns=date_columns)

    if converter.process_biff_stream():
        converter.write_csv()

    end_time = datetime.now()
    elapsed = end_time - start_time
    logger.info(f"Conversion completed in {elapsed.total_seconds():.2f} seconds")

if __name__ == "__main__":
    main()
