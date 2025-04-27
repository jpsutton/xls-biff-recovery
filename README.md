# XLS BIFF Recovery Tools

A set of Python tools for recovering data from corrupted or damaged Excel (.xls) files by working directly with the Binary Interchange File Format (BIFF) structures.

## Overview

This project provides utilities to extract worksheet data from XLS files that may be corrupted or not properly accessible through standard Excel libraries. It works by:

1. Extracting the raw BIFF workbook stream from XLS files
2. Directly parsing the BIFF format to extract cell values
3. Converting the recovered data to CSV format

## Tools Included

### `extract_workbook_biff.py`

This tool extracts the raw BIFF workbook stream from an OLE2 compound document (XLS file). It implements a minimal OLE2 parser that can navigate through the file structure even when parts of the file are damaged.

Features:
- OLE2 compound document parsing
- Sector chain traversal
- FAT/DIFAT processing
- Stream extraction
- Debug mode for diagnostics

### `biff2csv.py`

This tool parses raw BIFF workbook streams and converts them to CSV format. It works by interpreting the various BIFF record types directly instead of relying on Excel libraries.

Features:
- BIFF record parsing
- Support for multiple data types (strings, numbers, dates, formulas)
- Date value conversion
- Shared string table support
- CSV output for recovered data

## Installation

Clone the repository:

```bash
git clone https://github.com/jpsutton/xls-biff-recovery.git
cd xls-biff-recovery
```

Requirements:
- Python 3.6+
- No external dependencies (uses only standard libraries)

## Usage

### Recovery Process

The recovery process usually involves two steps:

1. Extract the workbook stream from the XLS file:

```bash
python extract_workbook_biff.py damaged_file.xls workbook.bin --debug
```

2. Convert the workbook stream to CSV:

```bash
python biff2csv.py workbook.bin output_directory
```

### Advanced Usage

#### Extract Workbook Stream

```bash
python extract_workbook_biff.py [input.xls] [output.bin] [--debug]
```

- `input.xls`: Path to the damaged XLS file
- `output.bin`: Path to save the extracted BIFF stream
- `--debug`: (Optional) Enable debug mode, which creates detailed logs and dumps binary data for analysis

#### Convert BIFF to CSV

```bash
python biff2csv.py [input_file] [output_dir] [--debug] [--date-columns=0,1,2]
```

- `input_file`: Path to the BIFF stream file
- `output_dir`: Directory to write CSV files to (one per worksheet)
- `--debug`: (Optional) Enable debug logging
- `--date-columns`: (Optional) Comma-separated list of column indices (0-based) to force date conversion

## How It Works

### OLE2 Structure

Excel XLS files are OLE2 Compound Documents that store data in "streams" within a structured file system:

```
OLE2 Document
├── Header
├── File Allocation Table (FAT)
├── DIFAT
├── Directory
│   └── Stream entries
└── Sectors
    └── Stream data
```

The `extract_workbook_biff.py` tool navigates this structure to locate and extract the "Workbook" or "Book" stream.

### BIFF Format

The BIFF (Binary Interchange File Format) is the binary format used to store Excel worksheet data:

```
BIFF Stream
├── BOF (Beginning of File) record
├── Workbook records
│   └── BOUNDSHEET records
├── Worksheet records
│   ├── BOF record
│   ├── DIMENSION record
│   ├── Cell data records
│   │   ├── NUMBER
│   │   ├── LABEL
│   │   ├── FORMULA
│   │   └── ...
│   └── EOF record
└── EOF (End of File) record
```

The `biff2csv.py` tool interprets these records to extract and convert cell values.

## Debugging and Troubleshooting

When using the `--debug` flag, both tools create detailed log files:
- `xls_extractor_debug.log` - For OLE2 extraction
- `biff_converter.log` - For BIFF parsing

Additionally, when using the `--debug` flag with `extract_workbook_biff.py`, it creates an `xls_debug` directory with:
- Binary dumps of key structures
- Detailed information about file structure
- Hex dumps of problematic sections

## Limitations

- These tools focus on data recovery rather than complete Excel functionality
- Formula results may not be available if the file is corrupted
- Some complex formatting information is not preserved
- Not all BIFF record types are fully supported

## License

This project is licensed under the GNU General Purpose License (version 2) - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Acknowledgments

- Microsoft's [MS-XLS](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/cd03cb5f-ca02-4934-a391-bb674cb8aa06) documentation
- OpenOffice.org's documentation on the [Microsoft Compound Document File Format](https://www.openoffice.org/sc/compdocfileformat.pdf)
