"""
DWG Cleaner - Extract schematic drawing from DWG files

This script:
1. Opens DWG files
2. Finds the drawing bounded by the 4 longest lines (2 horizontal + 2 vertical)
3. Keeps only entities inside those bounds (including the border lines)
4. Saves to a CLEAN subfolder
5. Generates CSV and HTML reports

Usage:
    python clean_dwg.py <input_dwg_or_folder>
"""

import os
import sys
import time
import math
import csv
from datetime import datetime
from pathlib import Path

try:
    import win32com.client
    import pythoncom
except ImportError:
    print("Error: pywin32 is required. Install with: pip install pywin32")
    sys.exit(1)


class ProcessingResult:
    """Store result of processing a single file."""
    def __init__(self, filename, input_path, output_path=None):
        self.filename = filename
        self.input_path = str(input_path)
        self.output_path = str(output_path) if output_path else ""
        self.status = "pending"  # pending, success, failed, review
        self.error_message = ""
        self.entities_before = 0
        self.entities_after = 0
        self.entities_deleted = 0
        self.border_found = False
        self.border_type = ""  # "polyline", "lines", "3dface", "border_layer"
        self.border_width = 0.0
        self.border_height = 0.0
        self.processing_time = 0.0
        self.needs_review = False  # Flag for files that may have issues


def generate_csv_report(results, output_folder):
    """Generate CSV report of processing results."""
    report_path = output_folder / "clean_dwg_report.csv"

    with open(report_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([
            'Filename', 'Status', 'Needs Review', 'Error Message',
            'Entities Before', 'Entities After', 'Entities Deleted', 'Delete %',
            'Border Found', 'Border Type', 'Border Width', 'Border Height',
            'Processing Time (s)', 'Input Path', 'Output Path'
        ])

        for r in results:
            delete_pct = (r.entities_deleted / r.entities_before * 100) if r.entities_before > 0 else 0
            writer.writerow([
                r.filename, r.status, r.needs_review, r.error_message,
                r.entities_before, r.entities_after, r.entities_deleted, f"{delete_pct:.1f}%",
                r.border_found, r.border_type, f"{r.border_width:.2f}", f"{r.border_height:.2f}",
                f"{r.processing_time:.2f}", r.input_path, r.output_path
            ])

    print(f"CSV report saved: {report_path}")
    return report_path


def generate_html_report(results, output_folder):
    """Generate HTML report of processing results."""
    report_path = output_folder / "clean_dwg_report.html"

    success_count = sum(1 for r in results if r.status == "success")
    failed_count = sum(1 for r in results if r.status == "failed")
    review_count = sum(1 for r in results if r.status == "review")
    total_count = len(results)
    total_deleted = sum(r.entities_deleted for r in results)

    html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>DWG Cleaner Report</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }}
        h1 {{ color: #333; }}
        .summary {{ background: #fff; padding: 20px; border-radius: 8px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
        .summary-stats {{ display: flex; gap: 20px; margin-top: 10px; flex-wrap: wrap; }}
        .stat {{ padding: 15px 25px; border-radius: 6px; color: white; font-size: 18px; }}
        .stat-total {{ background: #2196F3; }}
        .stat-success {{ background: #4CAF50; }}
        .stat-failed {{ background: #f44336; }}
        .stat-review {{ background: #FF9800; }}
        .stat-deleted {{ background: #9c27b0; }}
        table {{ border-collapse: collapse; width: 100%; background: white; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
        th, td {{ padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }}
        th {{ background: #333; color: white; position: sticky; top: 0; }}
        tr:hover {{ background: #f5f5f5; }}
        .status-success {{ color: #4CAF50; font-weight: bold; }}
        .status-failed {{ color: #f44336; font-weight: bold; }}
        .status-review {{ color: #FF9800; font-weight: bold; }}
        .row-review {{ background: #FFF3E0; }}
        .error-msg {{ color: #f44336; font-size: 12px; max-width: 300px; }}
        .timestamp {{ color: #666; font-size: 14px; }}
        .deleted-highlight {{ color: #9c27b0; font-weight: bold; }}
        .file-link {{ color: #1976D2; text-decoration: none; }}
        .file-link:hover {{ text-decoration: underline; }}
    </style>
</head>
<body>
    <h1>DWG Cleaner Report</h1>
    <p class="timestamp">Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>

    <div class="summary">
        <h2>Summary</h2>
        <div class="summary-stats">
            <div class="stat stat-total">Total: {total_count}</div>
            <div class="stat stat-success">Success: {success_count}</div>
            <div class="stat stat-failed">Failed: {failed_count}</div>
            <div class="stat stat-review">Needs Review: {review_count}</div>
            <div class="stat stat-deleted">Entities Removed: {total_deleted}</div>
        </div>
    </div>

    <table>
        <tr>
            <th>#</th>
            <th>Filename</th>
            <th>Status</th>
            <th>Error</th>
            <th>Entities Before</th>
            <th>Entities After</th>
            <th>Removed</th>
            <th>Border</th>
            <th>Size (WxH)</th>
            <th>Time (s)</th>
        </tr>
"""

    for i, r in enumerate(results, 1):
        if r.status == "success":
            status_class = "status-success"
        elif r.status == "review":
            status_class = "status-review"
        else:
            status_class = "status-failed"
        row_class = "row-review" if r.status == "review" else ""
        border_info = f"{r.border_type}" if r.border_found else "Not found"
        size_info = f"{r.border_width:.0f} x {r.border_height:.0f}" if r.border_found else "-"
        error_display = f'<span class="error-msg">{r.error_message}</span>' if r.error_message else ""
        deleted_class = "deleted-highlight" if r.entities_deleted > 0 else ""
        # Create file link - use output path if available, otherwise input path
        file_path = r.output_path if r.output_path and r.status in ("success", "review") else r.input_path
        file_link = f'<a href="file:///{file_path.replace(chr(92), "/")}" class="file-link">{r.filename}</a>'

        html += f"""        <tr class="{row_class}">
            <td>{i}</td>
            <td>{file_link}</td>
            <td class="{status_class}">{r.status.upper()}</td>
            <td>{error_display}</td>
            <td>{r.entities_before}</td>
            <td>{r.entities_after}</td>
            <td class="{deleted_class}">{r.entities_deleted}</td>
            <td>{border_info}</td>
            <td>{size_info}</td>
            <td>{r.processing_time:.2f}</td>
        </tr>
"""

    html += """    </table>
</body>
</html>
"""

    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"HTML report saved: {report_path}")
    return report_path


def get_autocad():
    """Connect to AutoCAD or start it if not running."""
    pythoncom.CoInitialize()
    try:
        acad = win32com.client.dynamic.Dispatch("AutoCAD.Application")
        print("Connected to AutoCAD instance")
    except:
        print("Starting AutoCAD...")
        acad = win32com.client.Dispatch("AutoCAD.Application")
        acad.Visible = True
        time.sleep(5)
    return acad


def get_line_length(entity):
    """Get the length of a line entity."""
    try:
        start = entity.StartPoint
        end = entity.EndPoint
        dx = end[0] - start[0]
        dy = end[1] - start[1]
        return math.sqrt(dx*dx + dy*dy)
    except:
        return 0


def lines_connect(line1, line2, tolerance=5.0):
    """Check if two lines connect at their endpoints (within tolerance)."""
    endpoints1 = [line1['start'], line1['end']]
    endpoints2 = [line2['start'], line2['end']]

    for p1 in endpoints1:
        for p2 in endpoints2:
            dist = math.sqrt((p1[0] - p2[0])**2 + (p1[1] - p2[1])**2)
            if dist <= tolerance:
                return True
    return False


def find_rectangle_from_lines(horizontal_lines, vertical_lines, tolerance=5.0):
    """
    Find 4 lines (2 horizontal, 2 vertical) that form a rectangle.
    Returns the 4 lines if found, or None.
    """
    # Try combinations of horizontal and vertical lines to find a rectangle
    for h1 in horizontal_lines:
        for h2 in horizontal_lines:
            if h1['handle'] == h2['handle']:
                continue

            for v1 in vertical_lines:
                for v2 in vertical_lines:
                    if v1['handle'] == v2['handle']:
                        continue

                    # Check if these 4 lines form a rectangle
                    # Each horizontal should connect to both verticals
                    # Each vertical should connect to both horizontals
                    h1_v1 = lines_connect(h1, v1, tolerance)
                    h1_v2 = lines_connect(h1, v2, tolerance)
                    h2_v1 = lines_connect(h2, v1, tolerance)
                    h2_v2 = lines_connect(h2, v2, tolerance)

                    if h1_v1 and h1_v2 and h2_v1 and h2_v2:
                        # Found a rectangle!
                        return [h1, h2, v1, v2]

    return None


def get_polyline_bounds(entity):
    """Get the bounding box of a polyline."""
    try:
        min_pt, max_pt = entity.GetBoundingBox()
        return (min_pt[0], min_pt[1], max_pt[0], max_pt[1])
    except:
        return None


def is_rectangular_polyline(entity, tolerance=5.0):
    """
    Check if a polyline is roughly rectangular (4 corners, right angles).
    Returns True if it looks like a rectangle.
    """
    try:
        # Get coordinates
        coords = entity.Coordinates

        # For 2D polyline, coords is a flat array [x1,y1,x2,y2,...]
        # For 3D, it's [x1,y1,z1,x2,y2,z2,...]

        # Try to get number of vertices
        if hasattr(entity, 'NumberOfVertices'):
            num_verts = entity.NumberOfVertices
        else:
            # Estimate from coordinate count
            num_verts = len(coords) // 2

        # A rectangle has 4 vertices (or 5 if closed with repeated first point)
        if num_verts < 4 or num_verts > 5:
            return False

        return True
    except:
        return False


def find_border_block_reference(doc):
    """
    Find a block reference that represents the drawing border.
    Block references with names containing 'BORDER' or 'FRAME' are common border indicators.
    Returns (bounds, handles) or (None, set()).
    """
    model_space = doc.ModelSpace
    best_block = None
    best_area = 0
    best_bounds = None

    for i in range(model_space.Count):
        try:
            entity = model_space.Item(i)
            entity_type = entity.EntityName

            # Check for block references (AcDbBlockReference)
            if "BlockReference" in entity_type or entity_type == "AcDbBlockReference":
                try:
                    # Get block name
                    block_name = entity.Name.upper() if hasattr(entity, 'Name') else ""

                    # Check if this is a border block
                    if "BORDER" in block_name or "FRAME" in block_name or "TITLE" in block_name:
                        min_pt, max_pt = entity.GetBoundingBox()
                        bounds = (min_pt[0], min_pt[1], max_pt[0], max_pt[1])
                        width = bounds[2] - bounds[0]
                        height = bounds[3] - bounds[1]
                        area = width * height

                        # Must be reasonably large
                        if area > best_area and width > 100 and height > 100:
                            # Check aspect ratio
                            aspect = max(width, height) / min(width, height) if min(width, height) > 0 else 999
                            if aspect < 5:
                                best_area = area
                                best_block = entity
                                best_bounds = bounds
                                print(f"  Found border block '{entity.Name}': {width:.1f} x {height:.1f}, area={area:.1f}")
                except:
                    pass
        except:
            pass

    if best_block:
        return best_bounds, {best_block.Handle}
    return None, set()


def find_border_by_layer(doc):
    """
    Find the drawing border by looking for entities on a BORDER, Defpoints, or Level layer.
    Many drawings have border lines on layers named like '02_BORDER', 'BORDER', 'Defpoints', 'Level 60', etc.
    Returns (bounds, handles) or (None, set()).
    """
    model_space = doc.ModelSpace
    border_entities = []

    for i in range(model_space.Count):
        try:
            entity = model_space.Item(i)
            layer_name = entity.Layer.upper()

            # Check if entity is on a border layer, Defpoints layer, or Level layer (DGN conversion)
            is_border_layer = ("BORDER" in layer_name or "FRAME" in layer_name or
                              layer_name == "DEFPOINTS" or layer_name.startswith("LEVEL "))
            if is_border_layer:
                try:
                    min_pt, max_pt = entity.GetBoundingBox()
                    border_entities.append({
                        'entity': entity,
                        'handle': entity.Handle,
                        'min_x': min_pt[0],
                        'min_y': min_pt[1],
                        'max_x': max_pt[0],
                        'max_y': max_pt[1]
                    })
                except:
                    pass
        except:
            pass

    if not border_entities:
        return None, set()

    # Calculate combined bounding box of all border entities
    min_x = min(e['min_x'] for e in border_entities)
    min_y = min(e['min_y'] for e in border_entities)
    max_x = max(e['max_x'] for e in border_entities)
    max_y = max(e['max_y'] for e in border_entities)

    width = max_x - min_x
    height = max_y - min_y
    area = width * height

    # Must be reasonably sized (at least 100x100 units)
    if width < 100 or height < 100:
        return None, set()

    # Check aspect ratio
    aspect = max(width, height) / min(width, height) if min(width, height) > 0 else 999
    if aspect > 5:  # Too extreme aspect ratio
        return None, set()

    print(f"  Found border/defpoints/level layer entities: {width:.1f} x {height:.1f}, area={area:.1f}")

    bounds = (min_x, min_y, max_x, max_y)
    handles = {e['handle'] for e in border_entities}

    return bounds, handles


def find_largest_3dface(doc):
    """
    Find the largest 3D Face entity that could be a drawing border.
    3D Faces (AcDbFace) are commonly used as drawing frames in DGN-converted files.
    Returns (bounds, handle) or (None, None).
    """
    model_space = doc.ModelSpace
    best_face = None
    best_area = 0
    best_bounds = None

    for i in range(model_space.Count):
        try:
            entity = model_space.Item(i)
            entity_type = entity.EntityName

            # Check for 3D Face entities
            # AcDbFace is the internal name for 3D Face in AutoCAD
            if entity_type == "AcDbFace" or "Face" in entity_type:
                try:
                    min_pt, max_pt = entity.GetBoundingBox()
                    bounds = (min_pt[0], min_pt[1], max_pt[0], max_pt[1])
                    width = bounds[2] - bounds[0]
                    height = bounds[3] - bounds[1]
                    area = width * height

                    # Must be reasonably large and roughly rectangular aspect ratio
                    if area > best_area and width > 0 and height > 0:
                        # Check aspect ratio (not too extreme for a drawing border)
                        aspect = max(width, height) / min(width, height)
                        if aspect < 3:  # Drawing borders are usually close to A-size ratios
                            best_area = area
                            best_face = entity
                            best_bounds = bounds
                            print(f"  Found 3D Face: {width:.1f} x {height:.1f}, area={area:.1f}")
                except:
                    pass
        except Exception as e:
            pass

    if best_face:
        return best_bounds, {best_face.Handle}
    return None, set()


def find_largest_rectangular_polyline(doc):
    """
    Find the largest closed polyline that looks like a rectangle.
    Returns (bounds, handle) or (None, None).
    """
    model_space = doc.ModelSpace
    best_polyline = None
    best_area = 0
    best_bounds = None

    for i in range(model_space.Count):
        try:
            entity = model_space.Item(i)
            entity_type = entity.EntityName

            # Check for polylines
            if "Polyline" in entity_type or "POLYLINE" in entity_type.upper():
                # Check if closed
                is_closed = False
                try:
                    is_closed = entity.Closed
                except:
                    pass

                bounds = get_polyline_bounds(entity)
                if bounds:
                    width = bounds[2] - bounds[0]
                    height = bounds[3] - bounds[1]
                    area = width * height

                    # Must be reasonably large and roughly rectangular aspect ratio
                    if area > best_area and width > 0 and height > 0:
                        # Check aspect ratio (not too extreme)
                        aspect = max(width, height) / min(width, height)
                        if aspect < 10:  # Reasonable aspect ratio for a drawing
                            best_area = area
                            best_polyline = entity
                            best_bounds = bounds
                            print(f"  Found polyline: {width:.1f} x {height:.1f}, area={area:.1f}, closed={is_closed}")
        except Exception as e:
            pass

    if best_polyline:
        return best_bounds, {best_polyline.Handle}
    return None, set()


def find_border_from_longest_lines(doc):
    """
    Find the drawing border by looking for 4 lines that form a rectangle.
    Prioritizes longer lines.

    Returns (bounds, border_line_handles) where bounds is (min_x, min_y, max_x, max_y)
    """
    model_space = doc.ModelSpace
    horizontal = []
    vertical = []
    all_lines = []

    # Collect all lines and classify as horizontal or vertical
    for i in range(model_space.Count):
        try:
            entity = model_space.Item(i)
            entity_type = entity.EntityName

            if "Line" in entity_type and "Polyline" not in entity_type:
                length = get_line_length(entity)
                if length > 0:
                    start = entity.StartPoint
                    end = entity.EndPoint

                    dx = abs(end[0] - start[0])
                    dy = abs(end[1] - start[1])

                    line_data = {
                        'entity': entity,
                        'handle': entity.Handle,
                        'length': length,
                        'start': (start[0], start[1]),
                        'end': (end[0], end[1]),
                        'min_x': min(start[0], end[0]),
                        'max_x': max(start[0], end[0]),
                        'min_y': min(start[1], end[1]),
                        'max_y': max(start[1], end[1]),
                        'center_x': (start[0] + end[0]) / 2,
                        'center_y': (start[1] + end[1]) / 2
                    }

                    all_lines.append(line_data)

                    if dx > dy * 2:  # Horizontal (more relaxed)
                        horizontal.append(line_data)
                    elif dy > dx * 2:  # Vertical (more relaxed)
                        vertical.append(line_data)
        except:
            pass

    print(f"Found {len(horizontal)} horizontal, {len(vertical)} vertical, {len(all_lines)} total lines")

    if len(horizontal) < 2 or len(vertical) < 2:
        print(f"Not enough H/V lines found")
        return None, set()

    # Sort by length (longest first)
    horizontal.sort(key=lambda x: x['length'], reverse=True)
    vertical.sort(key=lambda x: x['length'], reverse=True)

    # Take top candidates (longest lines)
    top_horizontal = horizontal[:min(30, len(horizontal))]
    top_vertical = vertical[:min(30, len(vertical))]

    print(f"Top 5 horizontal lengths: {[f'{l['length']:.1f}' for l in top_horizontal[:5]]}")
    print(f"Top 5 vertical lengths: {[f'{l['length']:.1f}' for l in top_vertical[:5]]}")

    # Try to find a rectangle from these lines
    # Start with largest tolerance and decrease if needed
    rectangle = None
    for tolerance in [50.0, 20.0, 10.0, 5.0, 2.0, 1.0]:
        rectangle = find_rectangle_from_lines(top_horizontal, top_vertical, tolerance)
        if rectangle:
            print(f"Found rectangle with tolerance {tolerance}")
            break

    if not rectangle:
        print("Could not find 4 lines forming a rectangle")
        # Fallback: use the 2 longest of each
        print("Falling back to longest lines method...")
        rectangle = [horizontal[0], horizontal[1], vertical[0], vertical[1]]

    # Extract the 4 lines
    h_lines = [l for l in rectangle if abs(l['end'][0] - l['start'][0]) > abs(l['end'][1] - l['start'][1])]
    v_lines = [l for l in rectangle if abs(l['end'][1] - l['start'][1]) > abs(l['end'][0] - l['start'][0])]

    if len(h_lines) < 2:
        h_lines = rectangle[:2]
    if len(v_lines) < 2:
        v_lines = rectangle[2:]

    # Sort to find positions
    h_lines.sort(key=lambda x: x['center_y'])
    v_lines.sort(key=lambda x: x['center_x'])

    bottom_line = h_lines[0]
    top_line = h_lines[-1]
    left_line = v_lines[0]
    right_line = v_lines[-1]

    print(f"Border lines selected:")
    print(f"  Bottom (Y={bottom_line['center_y']:.1f}): length {bottom_line['length']:.1f}")
    print(f"  Top (Y={top_line['center_y']:.1f}): length {top_line['length']:.1f}")
    print(f"  Left (X={left_line['center_x']:.1f}): length {left_line['length']:.1f}")
    print(f"  Right (X={right_line['center_x']:.1f}): length {right_line['length']:.1f}")

    # Calculate bounds - use the intersection of the lines
    min_x = left_line['center_x']
    max_x = right_line['center_x']
    min_y = bottom_line['center_y']
    max_y = top_line['center_y']

    bounds = (min_x, min_y, max_x, max_y)

    # Collect handles of border lines to preserve
    border_handles = {
        bottom_line['handle'],
        top_line['handle'],
        left_line['handle'],
        right_line['handle']
    }

    width = max_x - min_x
    height = max_y - min_y

    print(f"Rectangle bounds: {width:.2f} x {height:.2f}")
    print(f"  From ({min_x:.2f}, {min_y:.2f}) to ({max_x:.2f}, {max_y:.2f})")

    return bounds, border_handles


def find_drawing_border(doc):
    """
    Find the drawing border. Tries multiple methods:
    1. Look for block references named BORDER/FRAME
    2. Look for entities on a BORDER, Defpoints, or Level layer
    3. Look for a large 3D Face (common in DGN-converted files)
    4. Look for a large rectangular polyline
    5. Look for 4 lines forming a rectangle

    Returns (bounds, border_handles, border_type)
    """
    print("\n--- Looking for drawing border ---")

    # Collect all candidates and pick the largest valid one
    candidates = []

    # Method 1: Try to find a BORDER block reference
    print("Method 1: Looking for BORDER block reference...")
    bounds_block, handles_block = find_border_block_reference(doc)
    if bounds_block:
        area = (bounds_block[2] - bounds_block[0]) * (bounds_block[3] - bounds_block[1])
        candidates.append(('border_block', bounds_block, handles_block, area))

    # Method 2: Try to find entities on a BORDER, Defpoints, or Level layer
    print("Method 2: Looking for BORDER/Defpoints/Level layer...")
    bounds_layer, handles_layer = find_border_by_layer(doc)
    if bounds_layer:
        area = (bounds_layer[2] - bounds_layer[0]) * (bounds_layer[3] - bounds_layer[1])
        candidates.append(('border_layer', bounds_layer, handles_layer, area))

    # Method 3: Try to find a 3D Face (common in DGN conversions)
    print("Method 3: Looking for 3D Face border...")
    bounds_3dface, handles_3dface = find_largest_3dface(doc)
    if bounds_3dface:
        area = (bounds_3dface[2] - bounds_3dface[0]) * (bounds_3dface[3] - bounds_3dface[1])
        candidates.append(('3dface', bounds_3dface, handles_3dface, area))

    # Method 4: Try to find a rectangular polyline
    print("Method 4: Looking for rectangular polyline...")
    bounds_poly, handles_poly = find_largest_rectangular_polyline(doc)
    if bounds_poly:
        area = (bounds_poly[2] - bounds_poly[0]) * (bounds_poly[3] - bounds_poly[1])
        candidates.append(('polyline', bounds_poly, handles_poly, area))

    # Pick the largest candidate
    if candidates:
        # Sort by area descending
        candidates.sort(key=lambda x: x[3], reverse=True)
        best = candidates[0]
        border_type, bounds, handles, area = best
        width = bounds[2] - bounds[0]
        height = bounds[3] - bounds[1]
        print(f"Using {border_type} border (largest): {width:.2f} x {height:.2f}")
        return bounds, handles, border_type

    # Method 5: Look for 4 lines forming a rectangle (fallback)
    print("Method 5: Looking for lines forming a rectangle...")
    bounds, handles = find_border_from_longest_lines(doc)
    if bounds:
        return bounds, handles, "lines"

    print("No border found!")
    return None, set(), ""


def entity_in_bounds(entity, bounds, tolerance=1.0):
    """Check if an entity is within the given bounds."""
    if bounds is None:
        return True

    try:
        ent_bounds = entity.GetBoundingBox()
        ent_min = ent_bounds[0]
        ent_max = ent_bounds[1]

        min_x, min_y, max_x, max_y = bounds

        # Entity is inside if its bounding box is within the border bounds
        return (ent_min[0] >= min_x - tolerance and
                ent_min[1] >= min_y - tolerance and
                ent_max[0] <= max_x + tolerance and
                ent_max[1] <= max_y + tolerance)
    except:
        return True  # Keep if we can't determine bounds


def delete_entities_outside_bounds(doc, bounds, border_handles):
    """Delete all entities outside the drawing bounds, preserving border lines."""
    if bounds is None:
        print("No bounds specified - keeping all entities")
        return 0

    model_space = doc.ModelSpace
    entities_to_delete = []

    # Collect entities to delete
    for i in range(model_space.Count):
        try:
            entity = model_space.Item(i)

            # Don't delete the border lines
            if entity.Handle in border_handles:
                continue

            if not entity_in_bounds(entity, bounds):
                entities_to_delete.append(entity)
        except:
            pass

    # Delete collected entities
    deleted_count = 0
    for entity in entities_to_delete:
        try:
            entity.Delete()
            deleted_count += 1
        except:
            pass

    print(f"Deleted {deleted_count} entities outside drawing bounds")
    return deleted_count


def process_dwg_file(acad, dwg_path, output_folder, max_retries=3, retry_delay=5):
    """Process a single DWG file. Returns ProcessingResult."""
    dwg_path = Path(dwg_path)
    output_path = Path(output_folder) / dwg_path.name

    # Create result object
    result = ProcessingResult(dwg_path.name, dwg_path, output_path)
    start_time = time.time()

    print(f"\n{'='*60}")
    print(f"Processing: {dwg_path.name}")
    print(f"{'='*60}")

    # Get fresh connection
    acad = get_autocad()

    # Open the DWG file with retry mechanism
    doc = None
    last_error = None
    for attempt in range(1, max_retries + 1):
        try:
            doc = acad.Documents.Open(str(dwg_path))
            time.sleep(2)
            break  # Success, exit retry loop
        except Exception as e:
            last_error = str(e)
            if attempt < max_retries:
                print(f"Error opening file (attempt {attempt}/{max_retries}): {last_error}")
                print(f"Waiting {retry_delay} seconds before retry...")
                time.sleep(retry_delay)
                # Get fresh AutoCAD connection for retry
                acad = get_autocad()
            else:
                print(f"Error opening file (attempt {attempt}/{max_retries}): {last_error}")
                print(f"All {max_retries} attempts failed.")
                result.status = "failed"
                result.error_message = f"Failed to open after {max_retries} attempts: {last_error}"
                result.processing_time = time.time() - start_time
                return result

    # Count entities before
    try:
        result.entities_before = doc.ModelSpace.Count
        print(f"Entities before: {result.entities_before}")
    except:
        result.entities_before = 0

    # Find drawing bounds (polyline or lines)
    print("Finding drawing bounds...")
    bounds, border_handles, border_type = find_drawing_border(doc)

    if bounds:
        result.border_found = True
        result.border_type = border_type
        result.border_width = bounds[2] - bounds[0]
        result.border_height = bounds[3] - bounds[1]

        # Delete entities outside bounds
        print("Removing entities outside drawing bounds...")
        result.entities_deleted = delete_entities_outside_bounds(doc, bounds, border_handles)

        # Zoom to extents
        doc.SendCommand("_ZOOM _E \n")
        time.sleep(0.5)

        # Count entities after
        try:
            result.entities_after = doc.ModelSpace.Count
            print(f"Entities after: {result.entities_after}")
        except:
            result.entities_after = result.entities_before - result.entities_deleted

    # Purge unused elements
    doc.SendCommand("-PURGE\nA\n*\nN\n")
    time.sleep(2)

    # Ensure output directory exists
    output_folder.mkdir(parents=True, exist_ok=True)

    # Save to output folder
    output_path_str = str(output_path).replace("/", "\\")
    try:
        doc.SaveAs(output_path_str)
        print(f"Saved: {output_path_str}")
        result.status = "success"
    except Exception as e:
        error_msg = str(e)
        print(f"Error saving: {error_msg}")
        result.status = "failed"
        result.error_message = f"Failed to save: {error_msg}"

    # Close the document
    try:
        doc.Close(False)  # False = don't save again (already saved)
        print(f"Closed: {dwg_path.name}")
        time.sleep(1)  # Give AutoCAD time to fully close the document
    except Exception as e:
        print(f"Warning: Error closing document: {e}")
        # Try alternative close method
        try:
            doc.SendCommand("_.CLOSE\n_N\n")  # Close without saving
            time.sleep(2)
        except:
            pass

    # Check if file needs review (>50% entities deleted)
    if result.entities_before > 0:
        delete_percentage = (result.entities_deleted / result.entities_before) * 100
        if delete_percentage > 50:
            result.needs_review = True
            result.status = "review"
            print(f"WARNING: {delete_percentage:.1f}% entities deleted - flagged for review!")

    result.processing_time = time.time() - start_time
    return result


def process_folder(input_folder, output_folder):
    """Process all DWG files in a folder."""
    input_folder = Path(input_folder)
    output_folder = Path(output_folder)

    # Find all DWG files
    dwg_files = set()
    for f in input_folder.glob("*.dwg"):
        dwg_files.add(f)
    for f in input_folder.glob("*.DWG"):
        dwg_files.add(f)
    dwg_files = sorted(list(dwg_files))

    if not dwg_files:
        print(f"No DWG files found in: {input_folder}")
        return

    print(f"Found {len(dwg_files)} DWG file(s)")

    acad = get_autocad()
    results = []

    for i, dwg_file in enumerate(dwg_files):
        print(f"\n[{i+1}/{len(dwg_files)}]")
        try:
            result = process_dwg_file(acad, dwg_file, output_folder)
            results.append(result)
        except Exception as e:
            # Create a failed result for unexpected errors
            result = ProcessingResult(dwg_file.name, dwg_file, output_folder / dwg_file.name)
            result.status = "failed"
            result.error_message = f"Unexpected error: {str(e)}"
            results.append(result)
            print(f"Error processing {dwg_file.name}: {e}")
            time.sleep(2)

    # Ensure output folder exists for reports
    output_folder.mkdir(parents=True, exist_ok=True)

    # Generate reports
    print("\nGenerating reports...")
    generate_csv_report(results, output_folder)
    generate_html_report(results, output_folder)

    # Print summary
    success_count = sum(1 for r in results if r.status == "success")
    fail_count = sum(1 for r in results if r.status == "failed")
    review_count = sum(1 for r in results if r.status == "review")
    total_deleted = sum(r.entities_deleted for r in results)

    print(f"\n{'='*60}")
    print(f"Processing complete!")
    print(f"  Success:      {success_count}")
    print(f"  Failed:       {fail_count}")
    print(f"  Needs Review: {review_count}")
    print(f"  Total entities removed: {total_deleted}")
    print(f"{'='*60}")

    # List failed files
    if fail_count > 0:
        print("\nFailed files:")
        for r in results:
            if r.status == "failed":
                print(f"  - {r.filename}: {r.error_message}")

    # List files needing review
    if review_count > 0:
        print("\nFiles needing review (>50% entities deleted):")
        for r in results:
            if r.status == "review":
                pct = (r.entities_deleted / r.entities_before * 100) if r.entities_before > 0 else 0
                print(f"  - {r.filename}: {pct:.1f}% deleted ({r.entities_deleted}/{r.entities_before})")


def main():
    if len(sys.argv) < 2:
        print("Usage: python clean_dwg.py <input_dwg_or_folder>")
        print("")
        print("Examples:")
        print("  python clean_dwg.py drawing.dwg")
        print("  python clean_dwg.py C:\\DWG_Files")
        print("")
        print("Output will be saved to a CLEAN subfolder")
        sys.exit(1)

    input_path = Path(sys.argv[1])

    # Output folder is always CLEAN subfolder
    if input_path.is_file():
        output_folder = input_path.parent / "CLEAN"
    else:
        output_folder = input_path / "CLEAN"

    print("="*60)
    print("DWG Cleaner - Extract Schematic Drawing")
    print("="*60)
    print(f"Input:  {input_path}")
    print(f"Output: {output_folder}")
    print("="*60)

    if input_path.is_file():
        acad = get_autocad()
        process_dwg_file(acad, input_path, output_folder)
    elif input_path.is_dir():
        process_folder(input_path, output_folder)
    else:
        print(f"Error: '{input_path}' is not a valid file or folder")
        sys.exit(1)


if __name__ == "__main__":
    main()
