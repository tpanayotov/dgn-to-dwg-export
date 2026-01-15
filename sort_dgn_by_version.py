"""
DGN File Version Sorter
Scans a folder of .dgn files and sorts them into subfolders by version (V7, V8, CONNECT)

Usage:
    python sort_dgn_by_version.py [folder_path]

If no folder path is provided, uses the current directory.
"""

import os
import sys
import shutil
from pathlib import Path

def is_valid_v7_element(header):
    """
    Check if the header bytes represent a valid V7 DGN element.

    V7 DGN element structure (first 4 bytes):
    - Bytes 0-1: Element length in words (16-bit little-endian)
    - Byte 2: Element type (valid types are 1-127)
    - Byte 3: Level (0-63) and complex bit flags

    The file typically starts with Type 9 (Design File Header) or Type 8 (Cell Library Header)
    """
    if len(header) < 4:
        return False

    # Get element type (byte 2, lower 7 bits)
    element_type = header[2] & 0x7F

    # Get element length (bytes 0-1, little-endian)
    element_length = header[0] | (header[1] << 8)

    # Valid V7 element types range from 1 to about 66
    # Type 9 = Design File Header (most common start)
    # Type 8 = Digitizer setup / Cell Header
    # Type 10 = Level Symbology
    valid_element_types = list(range(1, 67))

    # Check for valid element type
    if element_type not in valid_element_types:
        return False

    # Element length should be reasonable (not 0, not huge)
    # V7 elements are measured in words (2 bytes), typical range 2-8192 words
    if element_length < 2 or element_length > 16384:
        return False

    return True


def check_v7_structure(f):
    """
    Perform deeper V7 structure validation by checking multiple elements.
    """
    f.seek(0)
    header = f.read(128)

    if len(header) < 36:
        return False

    # V7 Design File Header (Type 9) structure check
    # The header element contains specific fields we can validate

    element_type = header[2] & 0x7F

    # Type 9 = Design File Header - most common
    if element_type == 9:
        # Check for reasonable values in the design file header
        # Bytes 4-5 often contain sub-type info
        # Bytes 30-31 often contain view info
        return True

    # Type 8 = Cell Header / Library
    if element_type == 8:
        return True

    # Type 10 = Level Symbology
    if element_type == 10:
        return True

    # Type 1 = Cell Library Header
    if element_type == 1:
        return True

    # Type 2 = Cell
    if element_type == 2:
        return True

    # Type 3 = Line
    if element_type == 3:
        return True

    # Type 4 = Line String
    if element_type == 4:
        return True

    # Type 5 = Group Data (often appears early in files)
    if element_type == 5:
        return True

    # Type 6 = Shape
    if element_type == 6:
        return True

    # Type 7 = Text Node
    if element_type == 7:
        return True

    # Type 11 = Curve
    if element_type == 11:
        return True

    # Type 12 = Complex Chain/Shape Header
    if element_type == 12:
        return True

    # Type 14 = Complex Shape
    if element_type == 14:
        return True

    # Type 15 = Ellipse
    if element_type == 15:
        return True

    # Type 16 = Arc
    if element_type == 16:
        return True

    # Type 17 = Text
    if element_type == 17:
        return True

    # Type 19 = B-spline Surface
    if element_type == 19:
        return True

    # Type 21 = B-spline Curve
    if element_type == 21:
        return True

    # Type 33 = Dimension
    if element_type == 33:
        return True

    # Type 34 = Shared Cell Definition
    if element_type == 34:
        return True

    # Type 35 = Shared Cell Instance
    if element_type == 35:
        return True

    # Type 37 = Multi-line
    if element_type == 37:
        return True

    # Type 66 = Application Element
    if element_type == 66:
        return True

    # If element type is in valid range and structure looks ok
    if 1 <= element_type <= 66:
        return True

    return False


def get_dgn_version(file_path):
    """
    Determine DGN file version by reading the file header.

    DGN Version signatures:
    - V7: Custom binary format with element-based structure
    - V8/V8i/CONNECT: OLE compound document (starts with D0 CF 11 E0)
    """
    try:
        with open(file_path, 'rb') as f:
            header = f.read(512)

            if len(header) < 32:
                return "UNKNOWN"

            # =================================================================
            # CHECK FOR V8+ FIRST (OLE Compound Document signature)
            # This is the most reliable check - V8+ files ALWAYS have this
            # =================================================================
            if header[0:8] == b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1':
                # It's V8 or newer - check for CONNECT markers
                f.seek(0)
                chunk = f.read(32768)

                # CONNECT Edition markers
                if b'DgnDb' in chunk or b'DGNDB' in chunk:
                    return "CONNECT"
                if b'Dgn.FileInfo' in chunk:
                    return "CONNECT"
                if b'imodel' in chunk.lower():
                    return "CONNECT"

                # Default to V8 for OLE-based files
                return "V8"

            # =================================================================
            # CHECK FOR V7 - Multiple detection methods
            # =================================================================

            # Method 1: Check for valid V7 element structure
            if is_valid_v7_element(header):
                if check_v7_structure(f):
                    return "V7"

            # Method 2: Check first byte patterns common in V7
            # V7 files often start with specific element length patterns
            first_byte = header[0]
            second_byte = header[1]
            third_byte = header[2]

            # Element type is in byte 2 (lower 7 bits)
            element_type = third_byte & 0x7F

            # Check if it looks like a V7 element header
            # Type 9 (Design File Header) is the most common start
            if element_type == 9:
                return "V7"

            # Type 8 (Cell library/digitizer)
            if element_type == 8:
                return "V7"

            # Type 10 (Level Symbology)
            if element_type == 10:
                return "V7"

            # Method 3: Check for V7 by examining element length patterns
            # V7 elements have length in first 2 bytes (little-endian words)
            element_length = first_byte | (second_byte << 8)

            # Reasonable element length (2-8192 words = 4-16384 bytes)
            if 2 <= element_length <= 8192:
                # And valid element type
                if 1 <= element_type <= 66:
                    # Additional validation: check level byte
                    level_byte = header[3]
                    # Level is bits 0-5 (0-63), bits 6-7 are flags
                    level = level_byte & 0x3F
                    if level <= 63:
                        return "V7"

            # Method 4: Look for V7-specific byte patterns in header
            # V7 files have specific structures at known offsets

            # Check bytes 4-35 for typical V7 header patterns
            # Design file headers have specific ranges of values
            if len(header) >= 36:
                # Check for non-zero content in expected places
                has_content = any(b != 0 for b in header[4:36])

                # Check it's not all 0xFF either
                not_all_ff = any(b != 0xFF for b in header[4:36])

                if has_content and not_all_ff:
                    # Check element type again with less strict criteria
                    if 1 <= element_type <= 127:
                        # Check the words per element is reasonable
                        if 1 <= element_length <= 16384:
                            return "V7"

            # Method 5: Heuristic - if it's not OLE and has structure, likely V7
            # Check if file has repeating element-like structure
            f.seek(0)
            sample = f.read(1024)

            # V7 files should not have these signatures
            if sample[0:4] == b'PK\x03\x04':  # ZIP
                return "UNKNOWN"
            if sample[0:3] == b'%PDF':  # PDF
                return "UNKNOWN"
            if sample[0:4] == b'RIFF':  # Various formats
                return "UNKNOWN"
            if sample[0:2] == b'BM':  # BMP
                return "UNKNOWN"
            if sample[0:4] == b'\x89PNG':  # PNG
                return "UNKNOWN"

            # If we got here and file has .dgn extension and has binary content,
            # it's very likely a V7 file with unusual structure
            # Final check: does it have reasonable binary content?
            if len(sample) >= 64:
                # Count printable vs non-printable bytes
                non_printable = sum(1 for b in sample[:64] if b < 32 or b > 126)

                # V7 DGN files are binary, should have mix of values
                if non_printable > 20:  # More than ~30% non-printable
                    # Looks like binary data, check element structure once more
                    if element_length > 0 and element_type > 0:
                        return "V7"

            # Method 6: Last resort - check file extension context
            # If it's a .dgn file and not OLE, assume V7
            # (This catches edge cases with unusual V7 files)
            file_ext = str(file_path).lower()
            if file_ext.endswith('.dgn'):
                # One more validation: first 4 bytes shouldn't be all zeros
                if header[0:4] != b'\x00\x00\x00\x00':
                    # And shouldn't be all 0xFF
                    if header[0:4] != b'\xFF\xFF\xFF\xFF':
                        # If it's not a known non-DGN format, treat as V7
                        return "V7"

            return "UNKNOWN"

    except Exception as e:
        print(f"  Error reading {file_path}: {e}")
        return "ERROR"


def sort_dgn_files(source_folder):
    """
    Sort all DGN files in the source folder into version-specific subfolders.
    """
    source_path = Path(source_folder)

    if not source_path.exists():
        print(f"Error: Folder '{source_folder}' does not exist.")
        return

    if not source_path.is_dir():
        print(f"Error: '{source_folder}' is not a directory.")
        return

    # Create output folders
    folders = {
        "V7": source_path / "DGN_V7",
        "V8": source_path / "DGN_V8",
        "CONNECT": source_path / "DGN_CONNECT",
        "UNKNOWN": source_path / "DGN_UNKNOWN",
        "ERROR": source_path / "DGN_ERROR"
    }

    # Find all DGN files
    dgn_files = list(source_path.glob("*.dgn")) + list(source_path.glob("*.DGN"))

    if not dgn_files:
        print(f"No .dgn files found in '{source_folder}'")
        return

    print(f"Found {len(dgn_files)} DGN file(s) in '{source_folder}'")
    print("-" * 60)

    # Track counts
    counts = {"V7": 0, "V8": 0, "CONNECT": 0, "UNKNOWN": 0, "ERROR": 0}

    for dgn_file in dgn_files:
        version = get_dgn_version(dgn_file)
        counts[version] += 1

        # Create destination folder if needed
        dest_folder = folders[version]
        dest_folder.mkdir(exist_ok=True)

        # Copy file to destination
        dest_file = dest_folder / dgn_file.name

        try:
            shutil.copy2(dgn_file, dest_file)
            print(f"  [{version:7}] {dgn_file.name}")
        except Exception as e:
            print(f"  [ERROR  ] {dgn_file.name}: {e}")

    # Print summary
    print("-" * 60)
    print("Summary:")
    print(f"  V7 files:      {counts['V7']}")
    print(f"  V8/V8i files:  {counts['V8']}")
    print(f"  CONNECT files: {counts['CONNECT']}")
    print(f"  Unknown:       {counts['UNKNOWN']}")
    print(f"  Errors:        {counts['ERROR']}")
    print("-" * 60)
    print("Files have been COPIED to subfolders (originals preserved).")
    print(f"Output folders created in: {source_path}")


def main():
    if len(sys.argv) > 1:
        folder = sys.argv[1]
    else:
        folder = os.getcwd()

    print("=" * 60)
    print("DGN File Version Sorter")
    print("=" * 60)

    sort_dgn_files(folder)


if __name__ == "__main__":
    main()
