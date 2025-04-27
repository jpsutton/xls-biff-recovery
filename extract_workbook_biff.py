#!/usr/bin/env python
import struct
import io
import sys
import os
import logging
import traceback
from typing import Dict, List, Tuple, Optional, BinaryIO

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("xls_extractor_debug.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("XLSExtractor")

class XLSExtractor:
    """Minimal XLS file parser that extracts worksheet data from potentially corrupt files."""

    # OLE2 constants
    SECTOR_SIZE = 512
    MINI_SECTOR_SIZE = 64
    HEADER_SIZE = 512

    # OLE2 special sector values
    FREESECT = 0xFFFFFFFF  # Free sector
    ENDOFCHAIN = 0xFFFFFFFE  # End of chain marker
    FATSECT = 0xFFFFFFFD  # FAT sector
    DIFSECT = 0xFFFFFFFC  # DIFAT sector

    # Signature at the beginning of OLE files
    OLE_SIGNATURE = b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'

    def __init__(self, file_path: str, debug_mode: bool = False):
        self.file_path = file_path
        self.debug_mode = debug_mode
        self.file_size = 0
        self.sector_size = self.SECTOR_SIZE
        self.mini_sector_size = self.MINI_SECTOR_SIZE
        self.fat_sectors = []
        self.directory_sectors = []
        self.mini_fat_sectors = []
        self.difat = []
        self.streams = {}  # name -> (size, start_sector)
        self.debug_info = {}  # Store debug info for each step

    def _debug_dump(self, section: str, data: bytes, description: str) -> None:
        """Dump binary data to a file for debugging."""
        if not self.debug_mode:
            return

        try:
            # Create debug directory if it doesn't exist
            debug_dir = os.path.join(os.path.dirname(self.file_path), "xls_debug")
            os.makedirs(debug_dir, exist_ok=True)

            # Generate a debug filename
            base_name = os.path.basename(self.file_path)
            debug_file = os.path.join(debug_dir, f"{base_name}_{section}.bin")

            # Write the data
            with open(debug_file, 'wb') as f:
                f.write(data)

            logger.debug(f"Dumped {len(data)} bytes to {debug_file}: {description}")
        except Exception as e:
            logger.error(f"Failed to dump debug data: {e}")

    def _log_hex_dump(self, data: bytes, offset: int = 0, length: int = 64) -> None:
        """Create a hex dump of binary data for logging."""
        if not data:
            logger.debug("No data to dump")
            return

        length = min(length, len(data))
        hex_bytes = ' '.join(f"{b:02X}" for b in data[:length])
        ascii_chars = ''.join(chr(b) if 32 <= b <= 126 else '.' for b in data[:length])

        logger.debug(f"Offset {offset}, Length {length}:")
        logger.debug(f"HEX: {hex_bytes}")
        logger.debug(f"ASCII: {ascii_chars}")

    def open(self) -> bool:
        """Open and validate the XLS file, parse basic structure."""
        try:
            logger.info(f"Opening file: {self.file_path}")

            with open(self.file_path, 'rb') as f:
                # Read header
                header = f.read(self.HEADER_SIZE)
                if not header.startswith(self.OLE_SIGNATURE):
                    logger.error("Not a valid OLE2 file (wrong signature)")
                    self._log_hex_dump(header[:16], 0)
                    return False

                logger.info("OLE2 signature verified")
                self._debug_dump("header", header, "OLE2 Header")

                self.file_size = self._get_file_size(f)
                logger.info(f"File size: {self.file_size} bytes")

                # Parse header
                try:
                    self._parse_header(header)
                except Exception as e:
                    logger.error(f"Error parsing header: {e}")
                    logger.error(traceback.format_exc())
                    return False

                # Read FAT
                try:
                    self._read_fat(f)
                except Exception as e:
                    logger.error(f"Error reading FAT: {e}")
                    logger.error(traceback.format_exc())
                    return False

                # Read directory
                try:
                    self._read_directory(f)
                except Exception as e:
                    logger.error(f"Error reading directory: {e}")
                    logger.error(traceback.format_exc())
                    return False

                logger.info(f"Found streams: {list(self.streams.keys())}")
                return True

        except FileNotFoundError:
            logger.error(f"File not found: {self.file_path}")
            return False
        except PermissionError:
            logger.error(f"Permission denied when opening: {self.file_path}")
            return False
        except Exception as e:
            logger.error(f"Unexpected error opening file: {e}")
            logger.error(traceback.format_exc())
            return False

    def _get_file_size(self, file_obj: BinaryIO) -> int:
        """Get the size of the file."""
        current_pos = file_obj.tell()
        file_obj.seek(0, io.SEEK_END)
        size = file_obj.tell()
        file_obj.seek(current_pos)
        return size

    def _parse_header(self, header: bytes) -> None:
        """Parse the OLE2 header to extract essential metadata."""
        logger.info("Parsing OLE2 header")

        # Sector size: 2^x, typically 512 bytes (2^9)
        sector_shift = struct.unpack('<H', header[30:32])[0]
        self.sector_size = 2 ** sector_shift
        logger.info(f"Sector size: {self.sector_size} bytes (2^{sector_shift})")

        # Mini sector size: 2^y, typically 64 bytes (2^6)
        mini_sector_shift = struct.unpack('<H', header[32:34])[0]
        self.mini_sector_size = 2 ** mini_sector_shift
        logger.info(f"Mini sector size: {self.mini_sector_size} bytes (2^{mini_sector_shift})")

        # Number of directory sectors
        num_dir_sectors = struct.unpack('<I', header[40:44])[0]
        logger.info(f"Number of directory sectors: {num_dir_sectors}")

        # Number of FAT sectors
        num_fat_sectors = struct.unpack('<I', header[44:48])[0]
        logger.info(f"Number of FAT sectors: {num_fat_sectors}")

        # First directory sector
        first_dir_sector = struct.unpack('<I', header[48:52])[0]
        logger.info(f"First directory sector: {first_dir_sector}")
        self.debug_info["first_dir_sector"] = first_dir_sector

        # First mini FAT sector
        first_mini_fat_sector = struct.unpack('<I', header[60:64])[0]
        logger.info(f"First mini FAT sector: {first_mini_fat_sector}")

        # Number of mini FAT sectors
        num_mini_fat_sectors = struct.unpack('<I', header[64:68])[0]
        logger.info(f"Number of mini FAT sectors: {num_mini_fat_sectors}")

        # First DIFAT sector
        first_difat_sector = struct.unpack('<I', header[68:72])[0]
        logger.info(f"First DIFAT sector: {first_difat_sector}")

        # Number of DIFAT sectors
        num_difat_sectors = struct.unpack('<I', header[72:76])[0]
        logger.info(f"Number of DIFAT sectors: {num_difat_sectors}")

        # Read the first 109 DIFAT entries from the header
        self.difat = []
        valid_difat_entries = 0

        for i in range(109):
            entry_pos = 76 + i * 4
            if entry_pos + 4 <= len(header):
                entry = struct.unpack('<I', header[entry_pos:entry_pos+4])[0]
                if entry != self.FREESECT:  # Not a free sector
                    self.difat.append(entry)
                    valid_difat_entries += 1

        logger.info(f"DIFAT entries from header: {valid_difat_entries}")
        if valid_difat_entries > 0:
            logger.debug(f"First few DIFAT entries: {self.difat[:min(5, len(self.difat))]}")

        # TODO: If needed, follow DIFAT chain for files with many FAT sectors

    def _read_fat(self, file_obj: BinaryIO) -> None:
        """Read the File Allocation Table sectors."""
        logger.info("Reading FAT sectors")
        self.fat_sectors = []

        if not self.difat:
            logger.warning("No DIFAT entries found, cannot read FAT")
            return

        fat_sectors_count = 0

        # Read each FAT sector referenced in the DIFAT
        for idx, fat_sector_id in enumerate(self.difat):
            logger.debug(f"Reading FAT sector #{idx} (ID: {fat_sector_id})")

            if fat_sector_id >= (self.file_size // self.sector_size):
                logger.warning(f"FAT sector ID {fat_sector_id} exceeds file bounds, skipping")
                continue

            try:
                # Seek to the sector position (+1 for header)
                fat_sector_pos = (fat_sector_id + 1) * self.sector_size
                file_obj.seek(fat_sector_pos)

                # Read sector data
                fat_sector_data = file_obj.read(self.sector_size)
                if len(fat_sector_data) < self.sector_size:
                    logger.warning(f"Incomplete FAT sector read: expected {self.sector_size}, got {len(fat_sector_data)}")

                # Debug dump
                if self.debug_mode:
                    self._debug_dump(f"fat_sector_{idx}", fat_sector_data, f"FAT Sector {idx}")

                # Each FAT sector contains sector IDs
                for i in range(0, len(fat_sector_data), 4):
                    if i + 4 <= len(fat_sector_data):
                        entry = struct.unpack('<I', fat_sector_data[i:i+4])[0]
                        self.fat_sectors.append(entry)

                fat_sectors_count += 1

            except Exception as e:
                logger.error(f"Error reading FAT sector {fat_sector_id}: {e}")
                logger.error(traceback.format_exc())

        logger.info(f"Read {fat_sectors_count} FAT sectors, total entries: {len(self.fat_sectors)}")

        # Validate FAT entries
        valid_entries = sum(1 for entry in self.fat_sectors if entry != self.FREESECT)
        logger.info(f"Valid FAT entries: {valid_entries}")

        if valid_entries == 0:
            logger.warning("No valid FAT entries found - file may be severely corrupted")

    def _read_directory(self, file_obj: BinaryIO) -> None:
        """Read the directory entries to find stream information."""
        logger.info("Reading directory entries")

        # Use the first directory sector stored in the header
        if "first_dir_sector" not in self.debug_info:
            logger.error("No directory sector information available")
            return

        dir_sector = self.debug_info["first_dir_sector"]
        if dir_sector == self.FREESECT or dir_sector == self.ENDOFCHAIN:
            logger.error(f"Invalid directory sector: {dir_sector}")
            return

        try:
            # Seek to the directory sector (+1 for header)
            dir_pos = (dir_sector + 1) * self.sector_size
            logger.debug(f"Directory sector position: {dir_pos}")

            file_obj.seek(dir_pos)
            directory_data = file_obj.read(self.sector_size)

            if len(directory_data) < self.sector_size:
                logger.warning(f"Incomplete directory sector read: expected {self.sector_size}, got {len(directory_data)}")

            # Debug dump
            if self.debug_mode:
                self._debug_dump("directory", directory_data, "Directory Sector")

            # Process directory entries (128 bytes each)
            for i in range(0, len(directory_data), 128):
                if i + 128 > len(directory_data):
                    logger.warning(f"Incomplete directory entry at offset {i}, skipping")
                    continue

                entry_data = directory_data[i:i+128]
                entry_idx = i // 128

                try:
                    # Get entry name length (in bytes)
                    name_len = struct.unpack('<H', entry_data[64:66])[0]

                    # Skip invalid entries
                    if name_len == 0 or name_len > 64:
                        continue

                    # Get entry name (Unicode string)
                    name_bytes = entry_data[0:name_len]
                    try:
                        name = name_bytes.decode('utf-16-le').rstrip('\x00')
                    except UnicodeDecodeError:
                        logger.warning(f"Failed to decode entry name at index {entry_idx}")
                        name = f"Unknown-{entry_idx}"

                    # Get entry type
                    entry_type = entry_data[66]
                    type_str = {
                        0: "Unknown",
                        1: "Storage",
                        2: "Stream",
                        5: "Root"
                    }.get(entry_type, f"Unknown({entry_type})")

                    # Get stream size
                    stream_size = struct.unpack('<Q', entry_data[120:128])[0]

                    # Get starting sector
                    start_sector = struct.unpack('<I', entry_data[116:120])[0]

                    logger.debug(f"Entry {entry_idx}: Name='{name}', Type={type_str}, Size={stream_size}, Sector={start_sector}")

                    # Store stream information if it's a valid stream
                    if entry_type == 2 and stream_size > 0 and start_sector != self.FREESECT:
                        self.streams[name] = (stream_size, start_sector)

                        # Special logging for workbook streams
                        if name in ["Workbook", "Book"]:
                            logger.info(f"Found {name} stream: Size={stream_size}, Start Sector={start_sector}")

                except Exception as e:
                    logger.error(f"Error processing directory entry {entry_idx}: {e}")
                    logger.error(traceback.format_exc())

            logger.info(f"Found {len(self.streams)} valid streams")

        except Exception as e:
            logger.error(f"Error reading directory: {e}")
            logger.error(traceback.format_exc())

    def extract_workbook_stream(self) -> Optional[bytes]:
        """Extract the Workbook/Book stream containing worksheet data."""
        logger.info("Attempting to extract workbook stream")

        workbook_stream = None

        # Look for common workbook stream names
        for name in ["Workbook", "Book"]:
            if name in self.streams:
                logger.info(f"Found '{name}' stream, extracting...")
                stream_data = self._extract_stream(name)
                if stream_data:
                    logger.info(f"Successfully extracted {len(stream_data)} bytes of {name} data")
                    return stream_data

        logger.warning("No workbook stream found")
        return None

    def _extract_stream(self, stream_name: str) -> Optional[bytes]:
        """Extract a stream by following its sector chain."""
        if stream_name not in self.streams:
            logger.warning(f"Stream '{stream_name}' not found")
            return None

        stream_size, start_sector = self.streams[stream_name]
        logger.info(f"Extracting stream '{stream_name}': Size={stream_size}, Start Sector={start_sector}")

        if start_sector == self.FREESECT or start_sector == self.ENDOFCHAIN:
            logger.error(f"Invalid start sector for stream '{stream_name}': {start_sector}")
            return None

        try:
            with open(self.file_path, 'rb') as f:
                # For simplicity, we're assuming the stream uses standard sectors
                # (not mini streams which would be handled differently)
                current_sector = start_sector
                stream_data = bytearray()
                sectors_read = 0

                while (current_sector != self.ENDOFCHAIN and
                       current_sector != self.FREESECT and
                       len(stream_data) < stream_size and
                       sectors_read < 10000):  # Safety limit

                    # Validate sector ID
                    if current_sector >= len(self.fat_sectors):
                        logger.error(f"Invalid sector ID {current_sector} (FAT has {len(self.fat_sectors)} entries)")
                        break

                    # Seek to the sector position
                    sector_pos = (current_sector + 1) * self.sector_size

                    # Check if we're going beyond file bounds
                    if sector_pos + self.sector_size > self.file_size:
                        logger.error(f"Sector position {sector_pos} exceeds file size {self.file_size}")
                        break

                    f.seek(sector_pos)

                    # Read sector data
                    sector_data = f.read(self.sector_size)

                    if len(sector_data) < self.sector_size:
                        logger.warning(f"Incomplete sector read: expected {self.sector_size}, got {len(sector_data)}")

                    # Add to stream data
                    stream_data.extend(sector_data)
                    sectors_read += 1

                    # Debug log
                    if sectors_read % 100 == 0:
                        logger.debug(f"Read {sectors_read} sectors ({len(stream_data)} bytes)")

                    # Get next sector from FAT
                    current_sector = self.fat_sectors[current_sector]

                logger.info(f"Finished reading stream: {sectors_read} sectors, {len(stream_data)} bytes")

                if sectors_read >= 10000:
                    logger.warning("Reached safety limit of 10000 sectors - possible circular reference")

                # Trim to the actual size
                final_data = bytes(stream_data[:stream_size])

                if self.debug_mode:
                    self._debug_dump(f"stream_{stream_name}", final_data, f"Stream {stream_name}")

                # Validate we got the right amount of data
                if len(final_data) < stream_size:
                    logger.warning(f"Extracted data is smaller than expected: {len(final_data)} < {stream_size}")

                return final_data

        except Exception as e:
            logger.error(f"Error extracting stream '{stream_name}': {e}")
            logger.error(traceback.format_exc())
            return None

    def dump_workbook_stream(self, output_path: str) -> bool:
        """Extract and save the workbook stream to a file."""
        logger.info(f"Extracting workbook stream to {output_path}")

        workbook_data = self.extract_workbook_stream()
        if not workbook_data:
            logger.error("Failed to extract workbook stream")
            return False

        try:
            with open(output_path, 'wb') as f:
                f.write(workbook_data)
            logger.info(f"Successfully wrote {len(workbook_data)} bytes to {output_path}")
            return True
        except Exception as e:
            logger.error(f"Error writing output file: {e}")
            logger.error(traceback.format_exc())
            return False

    def debug_dump_all(self) -> None:
        """Create a comprehensive debug dump of the file structure."""
        if not self.debug_mode:
            return

        try:
            # Create debug directory
            debug_dir = os.path.join(os.path.dirname(self.file_path), "xls_debug")
            os.makedirs(debug_dir, exist_ok=True)

            # Create a debug info file
            debug_info_path = os.path.join(debug_dir, f"{os.path.basename(self.file_path)}_debug_info.txt")

            with open(debug_info_path, 'w') as f:
                f.write(f"XLS File: {self.file_path}\n")
                f.write(f"File Size: {self.file_size} bytes\n")
                f.write(f"Sector Size: {self.sector_size} bytes\n")
                f.write(f"Mini Sector Size: {self.mini_sector_size} bytes\n")
                f.write(f"DIFAT Entries: {len(self.difat)}\n")
                f.write(f"FAT Entries: {len(self.fat_sectors)}\n")
                f.write(f"Streams: {len(self.streams)}\n\n")

                # List streams
                f.write("Streams:\n")
                for name, (size, sector) in self.streams.items():
                    f.write(f"  - {name}: Size={size}, Start Sector={sector}\n")

            logger.info(f"Created debug info file: {debug_info_path}")

        except Exception as e:
            logger.error(f"Error creating debug dump: {e}")
            logger.error(traceback.format_exc())

def extract_xls_stream(input_path: str, output_path: str, debug_mode: bool = False) -> bool:
    """Extract the workbook stream from an XLS file."""
    extractor = XLSExtractor(input_path, debug_mode=debug_mode)

    logger.info(f"Starting extraction: {input_path} -> {output_path}")

    if extractor.open():
        # If in debug mode, do a complete debug dump
        if debug_mode:
            extractor.debug_dump_all()

        # Extract the workbook stream
        result = extractor.dump_workbook_stream(output_path)

        if result:
            logger.info(f"Successfully extracted workbook stream to {output_path}")
        else:
            logger.error("Failed to extract workbook stream")

        return result
    else:
        logger.error("Failed to open and parse the XLS file")
        return False

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python xls_extractor.py input.xls output.bin [--debug]")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]
    debug_mode = "--debug" in sys.argv

    if debug_mode:
        print("Debug mode enabled - additional output will be saved to the xls_debug directory")

    if extract_xls_stream(input_file, output_file, debug_mode):
        print(f"Successfully extracted workbook stream to {output_file}")
    else:
        print("Failed to extract workbook stream - check xls_extractor_debug.log for details")
        sys.exit(1)
