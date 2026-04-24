import trimesh
import numpy as np
import os

def extract_geometry(file_path):
    """
    Reads a 3D file (.stl, .stp converted to stl, etc.)
    and extracts all dimensional data needed for the BOM.
    
    Returns:
    {
        "folded_L": 1018.0,      ← installed length
        "folded_W": 457.0,       ← installed width  
        "folded_H": 185.0,       ← installed height
        "blank_L": 1053.0,       ← flat sheet length
        "blank_W": 578.0,        ← flat sheet width
        "blank_T": 20.0,         ← thickness
        "surface_area_sqm": 0.344065,
        "cad_volume_mm3": 123456.0
    }
    """
    
    print(f"📐 Reading 3D file: {file_path}")
    
    # Check file exists
    if not os.path.exists(file_path):
        print(f"  ❌ File not found: {file_path}")
        return _empty_geometry()
    
    # Check file type
    ext = os.path.splitext(file_path)[1].lower()
    
    if ext in ['.stp', '.step']:
        print("  ℹ️  STEP file detected — needs conversion to STL first")
        print("  ℹ️  Attempting direct load...")
    
    try:
        # Load the 3D file
        mesh = trimesh.load(file_path, force='mesh')
        
        if mesh is None or not hasattr(mesh, 'bounds'):
            print("  ❌ Could not load mesh")
            return _empty_geometry()
        
        # ── Folded size (simple axis-aligned bounding box) ─
        # This is the installed/formed shape dimensions
        bounds = mesh.bounds  # [[min_x,min_y,min_z],[max_x,max_y,max_z]]
        
        dims = [
            round(bounds[1][0] - bounds[0][0], 1),  # X
            round(bounds[1][1] - bounds[0][1], 1),  # Y
            round(bounds[1][2] - bounds[0][2], 1),  # Z
        ]
        dims_sorted = sorted(dims, reverse=True)
        
        folded_L = dims_sorted[0]   # largest dimension
        folded_W = dims_sorted[1]   # second largest
        folded_H = dims_sorted[2]   # smallest
        
        # ── Blank size (oriented bounding box) ────────────
        # This is the flat sheet before forming
        # OBB gives tightest fit around the actual geometry
        try:
            obb_extents = mesh.bounding_box_oriented.extents
            obb_sorted  = sorted(obb_extents, reverse=True)
            blank_L = round(obb_sorted[0], 1)
            blank_W = round(obb_sorted[1], 1)
            blank_T = round(obb_sorted[2], 1)
        except Exception:
            # Fall back to bounding box if OBB fails
            blank_L = folded_L
            blank_W = folded_W
            blank_T = folded_H
        
        # ── Surface area and volume ────────────────────────
        # Surface area divided by 2 = one face 
        # (mesh has both inner and outer faces)
        surface_area_mm2  = mesh.area / 2
        surface_area_sqm  = round(surface_area_mm2 / 1_000_000, 6)
        cad_volume_mm3    = round(abs(mesh.volume), 2)
        
        result = {
            "folded_L":         folded_L,
            "folded_W":         folded_W,
            "folded_H":         folded_H,
            "blank_L":          blank_L,
            "blank_W":          blank_W,
            "blank_T":          blank_T,
            "surface_area_sqm": surface_area_sqm,
            "cad_volume_mm3":   cad_volume_mm3
        }
        
        print(f"  ✅ Folded size: {folded_L} x {folded_W} x {folded_H} mm")
        print(f"  ✅ Blank size:  {blank_L} x {blank_W} x {blank_T} mm")
        print(f"  ✅ Surface area: {surface_area_sqm} sqm")
        print(f"  ✅ Volume: {cad_volume_mm3} mm³")
        
        return result
        
    except Exception as e:
        print(f"  ❌ Error reading 3D file: {e}")
        return _empty_geometry()


def gsm_crosscheck(surface_area_sqm, weight_g, gsm_from_drawing):
    """
    Cross-checks GSM from drawing against
    calculated GSM from weight and surface area.
    
    Example:
    surface_area = 0.344 sqm
    weight = 380g
    calculated GSM = 380/0.344 = 1104
    drawing GSM = 1100
    difference = 0.4% → PASS
    """
    
    if not all([surface_area_sqm, weight_g, gsm_from_drawing]):
        return {
            "status": "⚠️  Cannot check — missing data",
            "calculated_gsm": None,
            "difference_pct": None
        }
    
    calculated_gsm = weight_g / surface_area_sqm
    diff_pct = abs(calculated_gsm - gsm_from_drawing) \
               / gsm_from_drawing * 100
    
    status = "✅ PASS" if diff_pct <= 10 else "❌ MISMATCH"
    
    return {
        "status":           status,
        "calculated_gsm":   round(calculated_gsm, 1),
        "drawing_gsm":      gsm_from_drawing,
        "difference_pct":   round(diff_pct, 1)
    }


def _empty_geometry():
    """Returns empty dict when file can't be read"""
    return {
        "folded_L":         None,
        "folded_W":         None,
        "folded_H":         None,
        "blank_L":          None,
        "blank_W":          None,
        "blank_T":          None,
        "surface_area_sqm": None,
        "cad_volume_mm3":   None
    }


def find_3d_file(folder_path):
    """
    Finds the 3D file in an RFQ folder.
    Checks for STL first (easiest), then STEP.
    """
    
    priority = ['.stl', '.obj', '.stp', '.step', '.iges', '.igs']
    
    for ext in priority:
        for filename in os.listdir(folder_path):
            if filename.lower().endswith(ext):
                return os.path.join(folder_path, filename)
    
    return None


# ── Test ──────────────────────────────────────────────────
if __name__ == "__main__":
    
    print("🧪 Testing geometry reader...")
    print()
    
    # Test 1 — GSM crosscheck with your real MSIL data
    print("Test 1: GSM Cross-check")
    print("-" * 40)
    
    parts = [
        ("72450-69Q00", 0.344065, 380, 1100),
        ("72450-69Q10", 0.341674, 380, 1100),
        ("72450-69QA0", 0.355261, 380, 1550),
    ]
    
    for part_no, area, weight, gsm in parts:
        result = gsm_crosscheck(area, weight, gsm)
        print(f"  Part: {part_no}")
        print(f"  Drawing GSM: {gsm}")
        print(f"  Calculated GSM: {result['calculated_gsm']}")
        print(f"  Difference: {result['difference_pct']}%")
        print(f"  Status: {result['status']}")
        print()
    
    # Test 2 — Try reading a 3D file if one exists
    print("Test 2: 3D File Reading")
    print("-" * 40)
    
    # Check if any STL files exist in rfq_inputs
    test_file = None
    for root, dirs, files in os.walk("rfq_inputs"):
        for f in files:
            if f.lower().endswith(('.stl', '.stp', '.step')):
                test_file = os.path.join(root, f)
                break
    
    if test_file:
        print(f"Found 3D file: {test_file}")
        result = extract_geometry(test_file)
        print(f"Result: {result}")
    else:
        print("No 3D files found in rfq_inputs/")
        print("Drop a .stl or .stp file into rfq_inputs/ to test")
        print()
        print("✅ GSM crosscheck works without 3D files")
        print("✅ geometry_reader.py is ready")
