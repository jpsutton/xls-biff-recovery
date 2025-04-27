#!/usr/bin/env python

import struct
import logging
import argparse
import os
from datetime import datetime
from enum import Enum

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("biff_transplanter.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("BIFFTransplanter")

class BIFFRecordType(Enum):
    BOF = 0x0809
    EOF = 0x000A
    FILEPASS = 0x002F  # Password protection record

class BIFFTransplanter:
    """Class to transplant a BIFF stream into a new XLS file"""
    def __init__(self, biff_file, output_file, template_file=None):
        self.biff_file = biff_file
        self.output_file = output_file
        self.template_file = template_file
        self.biff_data = None
        self.template_data = None
        self.workbook_sector = None
        self.workbook_size = None

    def read_biff_stream(self):
        """Read the BIFF stream file into memory"""
        try:
            with open(self.biff_file, 'rb') as f:
                self.biff_data = f.read()
            logger.info(f"Successfully read BIFF stream: {self.biff_file}, size: {len(self.biff_data)} bytes")
            return True
        except Exception as e:
            logger.error(f"Failed to read BIFF stream file: {e}")
            return False

    def read_template_file(self):
        """Read the template XLS file into memory"""
        try:
            with open(self.template_file, 'rb') as f:
                self.template_data = f.read()
            logger.info(f"Successfully read template XLS: {self.template_file}, size: {len(self.template_data)} bytes")
            return True
        except Exception as e:
            logger.error(f"Failed to read template XLS file: {e}")
            return False

    def verify_biff_structure(self):
        """Verify that the BIFF stream begins with a BOF record"""
        if len(self.biff_data) < 4:
            logger.error("BIFF stream is too short to contain a valid header")
            return False

        # Check for BOF record at the beginning
        record_type, record_size = struct.unpack("<HH", self.biff_data[0:4])

        if record_type != BIFFRecordType.BOF.value:
            logger.error(f"BIFF stream does not begin with BOF record. Found record type: 0x{record_type:04X}")
            return False

        # Additional verification of the BOF record
        if record_size < 8 or len(self.biff_data) < record_size + 4:
            logger.error(f"BOF record too small or truncated, size: {record_size}")
            return False
            
        # Parse BOF record to get more info
        if record_size >= 8:
            version, doctype = struct.unpack("<HH", self.biff_data[4:8])
            logger.info(f"BIFF stream version: 0x{version:04X}, document type: {doctype}")
            
            # Version should be 0x0600 for BIFF8 (Excel 97-2003)
            if version != 0x0600:
                logger.warning(f"BIFF version 0x{version:04X} is not standard Excel 97-2003 format (0x0600)")
                
            # Document type should be 5 for workbook globals
            if doctype != 5:
                logger.warning(f"Document type {doctype} is not standard workbook globals (5)")

        logger.info(f"BIFF stream begins with BOF record, size: {record_size}")
        return True

    def check_for_password_protection(self):
        """Check if the BIFF stream contains FILEPASS records indicating password protection"""
        position = 0
        while position < len(self.biff_data) - 4:
            record_type, record_size = struct.unpack("<HH", self.biff_data[position:position+4])
            
            if record_type == BIFFRecordType.FILEPASS.value:
                logger.warning("BIFF stream contains FILEPASS record - worksheet is password protected")
                return True
                
            position += 4 + record_size
            
        return False

    def find_workbook_stream(self):
        """Find the Workbook stream in the template XLS file"""
        if not self.template_data:
            logger.error("Template data not loaded")
            return False
            
        # Verify OLE signature
        if self.template_data[0:8] != b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1':
            logger.error("Template file is not a valid OLE document")
            return False
            
        # Get sector size (typically 512 bytes)
        sector_size = 2 ** struct.unpack("<H", self.template_data[30:32])[0]
        logger.info(f"Template OLE sector size: {sector_size} bytes")
        
        # Get first directory sector
        first_dir_sector = struct.unpack("<I", self.template_data[48:52])[0]
        
        # Read directory entries (128 bytes each)
        dir_sector_offset = (first_dir_sector + 1) * sector_size
        
        # Search for "Workbook" or "Book" stream
        current_offset = dir_sector_offset
        found = False
        
        # Skip root entry (first entry)
        current_offset += 128
        
        while current_offset < len(self.template_data):
            # Read entry name (first 64 bytes, Unicode)
            entry_name_raw = self.template_data[current_offset:current_offset+64]
            
            # Convert to string until null termination
            entry_name = ""
            for i in range(0, len(entry_name_raw), 2):
                if entry_name_raw[i] == 0 and entry_name_raw[i+1] == 0:
                    break
                if entry_name_raw[i+1] == 0:  # Skip high byte if 0
                    entry_name += chr(entry_name_raw[i])
            
            # Check entry type (2 = stream)
            entry_type = struct.unpack("<H", self.template_data[current_offset+66:current_offset+68])[0]
            
            if entry_type == 2 and (entry_name == "Workbook" or entry_name == "Book"):
                # Found Workbook stream
                self.workbook_sector = struct.unpack("<I", self.template_data[current_offset+116:current_offset+120])[0]
                self.workbook_size = struct.unpack("<I", self.template_data[current_offset+120:current_offset+124])[0]
                logger.info(f"Found {entry_name} stream: sector {self.workbook_sector}, size {self.workbook_size} bytes")
                found = True
                break
                
            current_offset += 128
            
        if not found:
            logger.error("Could not find Workbook stream in template file")
            return False
            
        return True

    def analyze_biff_stream(self):
        """Analyze the BIFF stream to ensure it has the correct structure for Excel"""
        if not self.biff_data:
            logger.error("BIFF data not loaded")
            return False
            
        # Quick check for binary data that doesn't look like BIFF
        if len(self.biff_data) < 8:
            logger.error("BIFF stream is too short to be valid")
            return False
            
        # Check for OLE signature at the beginning (might be full file, not just BIFF stream)
        if self.biff_data[0:8] == b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1':
            logger.warning("Data appears to be a complete OLE document, not a raw BIFF stream")
            return False
            
        # Check the first record - should be BOF
        record_type, record_length = struct.unpack("<HH", self.biff_data[0:4])
        
        if record_type != BIFFRecordType.BOF.value:
            logger.error(f"First record is not BOF (0x0809), found: 0x{record_type:04X}")
            # Try offset + 4 to see if there's a header
            if len(self.biff_data) >= 8:
                offset_record_type = struct.unpack("<H", self.biff_data[4:6])[0]
                if offset_record_type == BIFFRecordType.BOF.value:
                    logger.warning("BOF record found at offset 4, file may have a 4-byte header")
                    
            return False
            
        # Check BOF record content - should have version and document type
        if record_length >= 8 and len(self.biff_data) >= 8:
            version, doc_type = struct.unpack("<HH", self.biff_data[4:8])
            logger.info(f"BIFF version: 0x{version:04X}, Document type: {doc_type}")
            
            # Excel 97-2003 uses BIFF8 (version 0x0600)
            if version != 0x0600:
                logger.warning(f"Version 0x{version:04X} is not Excel 97-2003 (BIFF8) format")
                
            # Document type 5 is Workbook Globals
            if doc_type != 5:
                logger.warning(f"Document type {doc_type} is not Workbook Globals (5)")
        else:
            logger.warning("BOF record is truncated or too small")
            
        # Scan records to verify structure
        position = 0
        record_count = 0
        found_eof = False
        sheet_count = 0
        
        important_records = {
            0x0085: "Sheet",
            0x00FC: "SST (Shared String Table)",
            0x00E2: "MSO Drawing",
            0x00EC: "MsoDrawingGroup",
            0x00EB: "Picture"
        }
        
        found_records = set()
        
        # Record ID frequency analysis
        record_frequency = {}
        
        while position + 4 <= len(self.biff_data):
            record_type, record_size = struct.unpack("<HH", self.biff_data[position:position+4])
            
            # Track record frequency
            if record_type in record_frequency:
                record_frequency[record_type] += 1
            else:
                record_frequency[record_type] = 1
                
            # Track important records
            if record_type in important_records:
                found_records.add(record_type)
                if record_type == 0x0085:  # Sheet record
                    sheet_count += 1
            
            if record_type == BIFFRecordType.EOF.value:
                logger.info(f"Found EOF record at position {position}")
                found_eof = True
                # In Excel files, we often have multiple EOF records (one per sheet/substream)
                
            # Skip to next record
            position += 4 + record_size
            record_count += 1
            
            # Sanity check for corrupted lengths
            if record_size > 8192:  # Most BIFF records are smaller than 8K
                logger.warning(f"Unusually large record at position {position-record_size-4}: "
                               f"type=0x{record_type:04X}, size={record_size}")
                
            if position > len(self.biff_data):
                logger.warning("Record extends beyond end of data - truncated BIFF stream?")
                break
        
        logger.info(f"BIFF analysis: {record_count} records, {sheet_count} sheets")
        
        # Log important records found
        if found_records:
            for rec_id in found_records:
                logger.info(f"Found {important_records[rec_id]} records: {record_frequency.get(rec_id, 0)}")
                
        if not found_eof:
            logger.warning("No EOF record found - BIFF stream may be truncated")
        
        return True

    def create_xls_from_template(self):
        """Create the XLS file with transplanted BIFF stream using the template file"""
        if not self.read_biff_stream():
            return False
            
        if not self.verify_biff_structure():
            logger.warning("BIFF stream structure verification failed, but attempting to continue")
            
        if self.check_for_password_protection():
            logger.warning("Password protected BIFF stream detected. The resulting XLS may not open correctly.")
        
        if not self.read_template_file():
            return False
            
        if not self.find_workbook_stream():
            return False
            
        # Get sector size from template
        sector_size = 2 ** struct.unpack("<H", self.template_data[30:32])[0]
        
        try:
            # Calculate sectors needed for BIFF data
            biff_sectors_needed = (len(self.biff_data) + sector_size - 1) // sector_size
            old_sectors_used = (self.workbook_size + sector_size - 1) // sector_size
            
            # If new BIFF data is larger than old workbook data, check if we need to resize
            if biff_sectors_needed > old_sectors_used:
                logger.warning(f"New BIFF data ({biff_sectors_needed} sectors) is larger than original workbook ({old_sectors_used} sectors)")
                
                # Check if template has enough space
                if (self.workbook_sector + 1) * sector_size + len(self.biff_data) > len(self.template_data):
                    logger.warning("Template file doesn't have enough space. Creating new file structure.")
                    return self.create_expanded_xls_from_template()
                
                logger.info("Template file has enough space for the larger BIFF stream.")
            
            # Create a copy of the template
            output_data = bytearray(self.template_data)
            
            # Find directory entry for Workbook again and update the size
            first_dir_sector = struct.unpack("<I", self.template_data[48:52])[0]
            dir_sector_offset = (first_dir_sector + 1) * sector_size
            
            # Skip root entry
            current_offset = dir_sector_offset + 128
            
            while current_offset < len(self.template_data):
                # Get entry name
                entry_name_raw = self.template_data[current_offset:current_offset+64]
                
                # Convert to string until null termination
                entry_name = ""
                for i in range(0, len(entry_name_raw), 2):
                    if entry_name_raw[i] == 0 and entry_name_raw[i+1] == 0:
                        break
                    if entry_name_raw[i+1] == 0:  # Skip high byte if 0
                        entry_name += chr(entry_name_raw[i])
                
                # Check entry type (2 = stream)
                entry_type = struct.unpack("<H", self.template_data[current_offset+66:current_offset+68])[0]
                
                if entry_type == 2 and (entry_name == "Workbook" or entry_name == "Book"):
                    # Update workbook size in directory entry
                    struct.pack_into("<I", output_data, current_offset+120, len(self.biff_data))
                    logger.info(f"Updated {entry_name} size in directory to {len(self.biff_data)} bytes")
                    break
                    
                current_offset += 128
            
            # Calculate offset to workbook stream
            workbook_offset = (self.workbook_sector + 1) * sector_size
            
            # Replace workbook stream with new BIFF data
            # We pad the data to fill complete sectors
            padded_biff = self.biff_data + (b'\x00' * (sector_size - (len(self.biff_data) % sector_size) if len(self.biff_data) % sector_size != 0 else 0))
            
            # Replace the workbook stream
            output_data[workbook_offset:workbook_offset+len(padded_biff)] = padded_biff
                
            # Write the modified file
            with open(self.output_file, 'wb') as f:
                f.write(output_data)
                
            logger.info(f"Successfully created XLS file: {self.output_file}")
            logger.info(f"File size: {os.path.getsize(self.output_file)} bytes")
            return True
            
        except Exception as e:
            logger.error(f"Error creating XLS file from template: {e}")
            return False

    def create_expanded_xls_from_template(self):
        """Create a new XLS file with expanded structure to accommodate larger BIFF data"""
        logger.info("Creating expanded XLS file structure based on template")
        
        # Analyze the BIFF stream to ensure it's correct
        if not self.analyze_biff_stream():
            logger.warning("BIFF stream analysis found issues but continuing anyway")
        
        try:
            # For Excel 97-2003, we'll create a standard structure
            # Standard sector size is 512 bytes
            sector_size = 512
            
            # Create a new minimal Excel 97-2003 OLE document
            # This uses a much simpler approach focused on Excel compatibility
            
            # Calculate required sizes
            padded_biff_size = len(self.biff_data)
            # Ensure BIFF data ends on a sector boundary
            if padded_biff_size % sector_size != 0:
                padding_size = sector_size - (padded_biff_size % sector_size)
                padded_biff_size += padding_size
            
            # Create minimal directory with Root and Workbook entries
            root_entry = bytearray()
            
            # Root entry name (Root Entry - 16-bit unicode, null-terminated)
            root_name = "Root Entry".encode('utf-16-le') + b'\x00\x00'
            root_name = root_name.ljust(64, b'\x00')  # Pad to 64 bytes
            root_entry.extend(root_name)
            
            # Entry type (0x05 for Root Storage)
            root_entry.extend(struct.pack("<H", 0x05))
            
            # Entry color (0x01 for black)
            root_entry.extend(struct.pack("<B", 0x01))
            
            # Left sibling, right sibling, child (all -1 for none, except child points to Workbook)
            root_entry.extend(struct.pack("<i", -1))  # Left
            root_entry.extend(struct.pack("<i", -1))  # Right
            root_entry.extend(struct.pack("<i", 1))   # Child - points to Workbook entry (index 1)
            
            # CLSID (all zeros for root)
            root_entry.extend(b'\x00' * 16)
            
            # State bits (0)
            root_entry.extend(struct.pack("<I", 0))
            
            # Creation time, modification time (all zeros)
            root_entry.extend(b'\x00' * 16)
            
            # Starting sector (0 - unused for storage)
            root_entry.extend(struct.pack("<I", 0))
            
            # Stream size (0 for storage)
            root_entry.extend(struct.pack("<Q", 0))
            
            # Workbook Entry
            wb_entry = bytearray()
            
            # Excel may use either "Workbook" or "Book" as the stream name
            # Use "Workbook" for Excel 97-2003
            wb_name = "Workbook".encode('utf-16-le') + b'\x00\x00'
            wb_name = wb_name.ljust(64, b'\x00')  # Pad to 64 bytes
            wb_entry.extend(wb_name)
            
            # Entry type (0x02 for Stream)
            wb_entry.extend(struct.pack("<H", 0x02))
            
            # Entry color (0x01 for black)
            wb_entry.extend(struct.pack("<B", 0x01))
            
            # Left sibling, right sibling, child (all -1 for stream)
            wb_entry.extend(struct.pack("<i", -1))  # Left
            wb_entry.extend(struct.pack("<i", -1))  # Right
            wb_entry.extend(struct.pack("<i", -1))  # Child
            
            # CLSID (all zeros for stream)
            wb_entry.extend(b'\x00' * 16)
            
            # State bits (0)
            wb_entry.extend(struct.pack("<I", 0))
            
            # Creation time, modification time (all zeros)
            wb_entry.extend(b'\x00' * 16)
            
            # Starting sector (after FAT and header)
            # For simplicity, we'll put the workbook stream right after the FAT sector
            workbook_start_sector = 2
            wb_entry.extend(struct.pack("<I", workbook_start_sector))
            
            # Stream size (size of BIFF data) - CRITICAL for Excel to read correctly
            wb_entry.extend(struct.pack("<I", len(self.biff_data)))
            wb_entry.extend(struct.pack("<I", 0))  # High 32 bits (0 for smaller files)
            
            # Combine directory entries
            directory = root_entry + wb_entry
            
            # Pad directory to sector size
            dir_padding_size = sector_size - (len(directory) % sector_size)
            if dir_padding_size < sector_size:
                directory.extend(b'\x00' * dir_padding_size)
            
            # Calculate number of sectors needed for each component
            fat_sectors = 1  # We'll use just one FAT sector for simplicity
            biff_sectors = (padded_biff_size + sector_size - 1) // sector_size
            dir_sectors = len(directory) // sector_size
            
            # Directory starts after BIFF data
            dir_start_sector = workbook_start_sector + biff_sectors
            
            # Create FAT (File Allocation Table)
            fat = bytearray()
            
            # Add FAT entries for the header (DIFAT)
            fat.extend(struct.pack("<I", 0xFFFFFFFD))  # Header is a reserved sector
            
            # FAT sector itself
            fat.extend(struct.pack("<I", 0xFFFFFFFE))  # FAT is an FAT sector
            
            # Chain the BIFF data sectors
            for i in range(workbook_start_sector, workbook_start_sector + biff_sectors - 1):
                fat.extend(struct.pack("<I", i + 1))  # Point to next sector
            
            # Last BIFF sector is end of chain
            fat.extend(struct.pack("<I", 0xFFFFFFFE))
            
            # Chain the directory sectors
            for i in range(dir_start_sector, dir_start_sector + dir_sectors - 1):
                fat.extend(struct.pack("<I", i + 1))  # Point to next sector
            
            # Last directory sector is end of chain
            fat.extend(struct.pack("<I", 0xFFFFFFFE))
            
            # Fill rest of FAT sector with free sector markers
            entries_so_far = 1 + 1 + biff_sectors + dir_sectors  # Header, FAT, BIFF, Directory
            entries_per_fat = sector_size // 4
            
            for i in range(entries_so_far, entries_per_fat):
                fat.extend(struct.pack("<I", 0xFFFFFFFF))  # FREESECT
            
            # Create OLE header with fixed size (512 bytes)
            ole_header = bytearray([
                # OLE header signature
                0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1,
                # CLSID (all zeros)
                0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                # Minor version
                0x3E, 0x00,
                # Major version (3 - for Excel 97-2003)
                0x03, 0x00,
                # Byte order (Little Endian)
                0xFE, 0xFF,
                # Sector size (512 bytes = 2^9)
                0x09, 0x00,
                # Mini-sector size (64 bytes = 2^6)
                0x06, 0x00,
                # Reserved
                0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                # Number of directory sectors (usually 0 for smaller files)
                0x00, 0x00, 0x00, 0x00,
                # Number of FAT sectors
                0x01, 0x00, 0x00, 0x00,  # 1 FAT sector
                # First directory sector location
                0x00, 0x00, 0x00, 0x00,  # Will be filled in later
                # Transaction signature (0 for no transactions)
                0x00, 0x00, 0x00, 0x00,
                # Mini-stream cutoff size (4096 bytes)
                0x00, 0x10, 0x00, 0x00,
                # First Mini FAT sector location (none)
                0xFE, 0xFF, 0xFF, 0xFF,
                # Number of Mini FAT sectors (0)
                0x00, 0x00, 0x00, 0x00,
                # First DIFAT sector location (none)
                0xFE, 0xFF, 0xFF, 0xFF,
                # Number of DIFAT sectors (0)
                0x00, 0x00, 0x00, 0x00
            ])
            
            # Ensure the header is large enough before writing to offset 76
            # The array needs to be at least size 76 + 4 (for a 32-bit integer)
            # Make sure we have enough space for DIFAT entries (109 entries of 4 bytes each starting at offset 76)
            while len(ole_header) < 76 + (109 * 4):
                ole_header.append(0)
                
            # Update the first directory sector location
            struct.pack_into("<I", ole_header, 48, dir_start_sector)
            
            # Set up DIFAT entries in header (points to FAT sectors)
            # First DIFAT entry points to the first FAT sector
            struct.pack_into("<I", ole_header, 76, 1)  # First FAT sector is at 1
            
            # Rest of DIFAT entries are empty
            for i in range(1, 109):
                struct.pack_into("<I", ole_header, 76 + i * 4, 0xFFFFFFFF)
            
            # Pad header to sector size
            if len(ole_header) < sector_size:
                ole_header.extend(b'\x00' * (sector_size - len(ole_header)))
            
            # Pad BIFF data to sector size
            padded_biff = bytearray(self.biff_data)
            biff_padding_size = sector_size - (len(padded_biff) % sector_size)
            if biff_padding_size < sector_size:
                padded_biff.extend(b'\x00' * biff_padding_size)
            
            # Write the file
            with open(self.output_file, 'wb') as f:
                # Write OLE header
                f.write(ole_header)
                
                # Write FAT
                f.write(fat)
                
                # Write BIFF data - write raw bytes exactly as they were extracted
                f.write(self.biff_data)
                
                # Pad to sector boundary if needed
                if len(self.biff_data) % sector_size != 0:
                    pad_size = sector_size - (len(self.biff_data) % sector_size)
                    f.write(b'\x00' * pad_size)
                
                # Write directory
                f.write(directory)
            
            logger.info(f"Successfully created Excel file: {self.output_file}")
            logger.info(f"File size: {os.path.getsize(self.output_file)} bytes")
            return True
            
        except Exception as e:
            logger.error(f"Error creating Excel file: {e}")
            logger.exception(e)  # Log full exception trace for debugging
            return False

    def create_ole_header(self):
        """Create a minimal OLE header for XLS file"""
        # This creates a very minimal OLE Container structure
        # A real implementation would create a more complete OLE structure
        
        # Fixed header for OLE compound document
        ole_header = bytearray([
            # OLE header signature (D0CF11E0A1B11AE1)
            0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1,
            # CLSID (all zeros)
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            # Minor version
            0x3E, 0x00,
            # Major version (3 - for Excel 97-2003)
            0x03, 0x00,
            # Byte order (Little Endian)
            0xFE, 0xFF,
            # Sector size (512 bytes = 2^9)
            0x09, 0x00,
            # Mini-sector size (64 bytes = 2^6)
            0x06, 0x00,
            # Reserved
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            # Number of directory sectors (usually 0 for smaller files)
            0x00, 0x00, 0x00, 0x00,
            # Number of FAT sectors (placeholder)
            0x01, 0x00, 0x00, 0x00,
            # First directory sector location (placeholder)
            0x01, 0x00, 0x00, 0x00,
            # Transaction signature (0 for no transactions)
            0x00, 0x00, 0x00, 0x00,
            # Mini-stream cutoff size (4096 bytes)
            0x00, 0x10, 0x00, 0x00,
            # First Mini FAT sector location (placeholder)
            0xFE, 0xFF, 0xFF, 0xFF,
            # Number of Mini FAT sectors (placeholder)
            0x00, 0x00, 0x00, 0x00,
            # First DIFAT sector location (placeholder)
            0xFE, 0xFF, 0xFF, 0xFF,
            # Number of DIFAT sectors (placeholder)
            0x00, 0x00, 0x00, 0x00
        ])
        
        # Initialize the first 109 DIFAT entries (-1 = unused)
        for i in range(109):
            ole_header.extend(struct.pack("<i", -1))
            
        # Add padding to make header sector-sized (512 bytes)
        padding_size = 512 - len(ole_header)
        ole_header.extend(b'\x00' * padding_size)
        
        logger.info(f"Created OLE header of size {len(ole_header)} bytes")
        return ole_header

    def create_minimal_directory_structure(self):
        """Create a minimal directory structure for the OLE file"""
        # Directory Entry for Root Storage
        root_entry = bytearray()
        
        # Root entry name (Root Entry - 16-bit unicode, null-terminated)
        root_name = "Root Entry".encode('utf-16-le') + b'\x00\x00'
        root_name = root_name.ljust(64, b'\x00')  # Pad to 64 bytes
        root_entry.extend(root_name)
        
        # Entry type (0x05 for Root Storage)
        root_entry.extend(struct.pack("<H", 0x05))
        
        # Entry color (0x01 for black)
        root_entry.extend(struct.pack("<B", 0x01))
        
        # Left sibling, right sibling, child (all -1 for none)
        root_entry.extend(struct.pack("<i", -1))  # Left
        root_entry.extend(struct.pack("<i", -1))  # Right
        root_entry.extend(struct.pack("<i", 1))   # Child - points to Workbook entry
        
        # CLSID (all zeros for root)
        root_entry.extend(b'\x00' * 16)
        
        # State bits (0)
        root_entry.extend(struct.pack("<I", 0))
        
        # Creation time, modification time (all zeros)
        root_entry.extend(b'\x00' * 16)
        
        # Starting sector (0)
        root_entry.extend(struct.pack("<I", 0))
        
        # Stream size (0 for storage)
        root_entry.extend(struct.pack("<Q", 0))
        
        # Directory Entry for Workbook Stream
        wb_entry = bytearray()
        
        # Workbook entry name (Workbook - 16-bit unicode, null-terminated)
        wb_name = "Workbook".encode('utf-16-le') + b'\x00\x00'
        wb_name = wb_name.ljust(64, b'\x00')  # Pad to 64 bytes
        wb_entry.extend(wb_name)
        
        # Entry type (0x02 for Stream)
        wb_entry.extend(struct.pack("<H", 0x02))
        
        # Entry color (0x01 for black)
        wb_entry.extend(struct.pack("<B", 0x01))
        
        # Left sibling, right sibling, child (all -1 for none)
        wb_entry.extend(struct.pack("<i", -1))  # Left
        wb_entry.extend(struct.pack("<i", -1))  # Right
        wb_entry.extend(struct.pack("<i", -1))  # Child
        
        # CLSID (all zeros for stream)
        wb_entry.extend(b'\x00' * 16)
        
        # State bits (0)
        wb_entry.extend(struct.pack("<I", 0))
        
        # Creation time, modification time (all zeros)
        wb_entry.extend(b'\x00' * 16)
        
        # Starting sector (2 - after header and FAT)
        wb_entry.extend(struct.pack("<I", 2))
        
        # Stream size (size of BIFF data)
        wb_entry.extend(struct.pack("<I", len(self.biff_data)))
        wb_entry.extend(struct.pack("<I", 0))  # High 32 bits (0 for smaller files)
        
        directory = root_entry + wb_entry
        
        # Pad to sector size
        padding_size = 512 - (len(directory) % 512)
        if padding_size < 512:
            directory.extend(b'\x00' * padding_size)
            
        logger.info(f"Created directory structure of size {len(directory)} bytes")
        return directory

    def create_fat(self, total_sectors):
        """Create File Allocation Table (FAT)"""
        # Calculate sectors needed for BIFF data
        biff_sectors = (len(self.biff_data) + 511) // 512
        
        # Create FAT
        fat = bytearray()
        
        # Special sectors
        fat.extend(struct.pack("<I", 0xFFFFFFFD))  # Header (DIFAT)
        fat.extend(struct.pack("<I", 0xFFFFFFFE))  # FAT sector itself
        
        # BIFF data sectors - all chained together
        for i in range(2, 2 + biff_sectors - 1):
            fat.extend(struct.pack("<I", i + 1))
        
        # Last BIFF sector is end of chain
        fat.extend(struct.pack("<I", 0xFFFFFFFE))
        
        # Directory sector is end of chain
        fat.extend(struct.pack("<I", 0xFFFFFFFE))
        
        # Pad to sector size
        padding_size = 512 - (len(fat) % 512)
        if padding_size < 512:
            fat.extend(struct.pack("<I", 0xFFFFFFFF) * (padding_size // 4))
            
        logger.info(f"Created FAT of size {len(fat)} bytes")
        return fat

    def create_xls(self):
        """Create the XLS file with transplanted BIFF stream"""
        # If template file is provided, use it instead of creating from scratch
        if self.template_file:
            return self.create_xls_from_template()
            
        if not self.read_biff_stream():
            return False
            
        if not self.verify_biff_structure():
            logger.warning("BIFF stream structure verification failed, but attempting to continue")
            
        if self.check_for_password_protection():
            logger.warning("Password protected BIFF stream detected. The resulting XLS may not open correctly.")
            
        # Create OLE header
        ole_header = self.create_ole_header()
            
        # Calculate sectors for BIFF data
        biff_sectors = (len(self.biff_data) + 511) // 512
        
        # Create directory structure
        directory = self.create_minimal_directory_structure()
        
        # Total sectors: header(1) + FAT(1) + BIFF data + directory(1)
        total_sectors = 1 + 1 + biff_sectors + 1
        
        # Create FAT
        fat = self.create_fat(total_sectors)
        
        # Pad BIFF data to sector size
        biff_padded = self.biff_data
        padding_size = 512 - (len(biff_padded) % 512)
        if padding_size < 512:
            biff_padded = biff_padded + (b'\x00' * padding_size)
            
        try:
            with open(self.output_file, 'wb') as f:
                # Write OLE header
                f.write(ole_header)
                
                # Write FAT
                f.write(fat)
                
                # Write BIFF data
                f.write(biff_padded)
                
                # Write directory structure
                f.write(directory)
                
            logger.info(f"Successfully created XLS file: {self.output_file}")
            logger.info(f"File size: {os.path.getsize(self.output_file)} bytes")
            return True
        except Exception as e:
            logger.error(f"Error creating XLS file: {e}")
            return False

def create_test_biff():
    """Create a very minimal test BIFF stream for testing"""
    biff = bytearray()
    
    # BOF record for workbook
    bof = bytearray()
    bof.extend(struct.pack("<HH", 0x0809, 8))  # Record type, length
    bof.extend(struct.pack("<HH", 0x0600, 0x0005))  # Version, type (5=workbook)
    bof.extend(struct.pack("<HH", 0x07CD, 0x0000))  # Build ID, year (0=unused)
    biff.extend(bof)
    
    # EOF record
    eof = bytearray()
    eof.extend(struct.pack("<HH", 0x000A, 0))  # Record type, length
    biff.extend(eof)
    
    # Write test BIFF to file
    with open("test_biff.bin", "wb") as f:
        f.write(biff)
        
    logger.info("Created test BIFF file: test_biff.bin")
    return "test_biff.bin"

def main():
    parser = argparse.ArgumentParser(description='Transplant BIFF stream into a new XLS file')
    parser.add_argument('biff_file', help='Path to the BIFF stream file')
    parser.add_argument('output_file', help='Path to the output XLS file')
    parser.add_argument('--template', help='Path to template XLS file to use as a basis')
    parser.add_argument('--debug', action='store_true', help='Enable debug logging')
    parser.add_argument('--test', action='store_true', help='Create and use a test BIFF stream')
    args = parser.parse_args()
    
    if args.debug:
        logger.setLevel(logging.DEBUG)
    
    start_time = datetime.now()
    logger.info(f"Starting BIFF to XLS transplantation at {start_time}")
    
    if args.test:
        args.biff_file = create_test_biff()
        
    logger.info(f"BIFF stream file: {args.biff_file}")
    logger.info(f"Output XLS file: {args.output_file}")
    
    if args.template:
        logger.info(f"Using template XLS file: {args.template}")
        
    transplanter = BIFFTransplanter(args.biff_file, args.output_file, args.template)
    success = transplanter.create_xls()
    
    end_time = datetime.now()
    elapsed = end_time - start_time
    
    if success:
        logger.info(f"Transplantation completed successfully in {elapsed.total_seconds():.2f} seconds")
    else:
        logger.error(f"Transplantation failed after {elapsed.total_seconds():.2f} seconds")

if __name__ == "__main__":
    main()
