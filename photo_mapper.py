#!/usr/bin/env python3
import argparse
import os
import sys
from pathlib import Path
from typing import Optional, Tuple, Dict, Any

try:
    from PIL import Image, ExifTags
except ImportError:
    print("This script requires Pillow. Install with:  pip install pillow", file=sys.stderr)
    sys.exit(2)

# Optional HEIC/HEIF support
have_pillow_heif = False
try:
    import pillow_heif  # type: ignore
    pillow_heif.register_heif_opener()
    have_pillow_heif = True
except Exception:
    pass

# Optional EXIF-bytes parser (used for HEIC fallback)
have_piexif = False
try:
    import piexif  # type: ignore
    have_piexif = True
except Exception:
    pass

IMG_EXTS = {".jpg", ".jpeg", ".tif", ".tiff", ".heic", ".heif"}

# Build tag maps once
TAGS = {v: k for k, v in ExifTags.TAGS.items()}
GPSTAGS = {v: k for k, v in ExifTags.GPSTAGS.items()}

def rational_to_float(x):
    try:
        if hasattr(x, "numerator") and hasattr(x, "denominator"):
            return float(x.numerator) / float(x.denominator) if x.denominator else float(x.numerator)
        if isinstance(x, tuple) and len(x) == 2:
            num, den = x
            return float(num) / float(den) if den else float(num)
        return float(x)
    except Exception:
        return None

def dms_to_deg(dms, ref: Optional[str]) -> Optional[float]:
    try:
        if not dms or len(dms) != 3:
            return None
        d = rational_to_float(dms[0])
        m = rational_to_float(dms[1])
        s = rational_to_float(dms[2])
        if d is None or m is None or s is None:
            return None
        deg = d + (m / 60.0) + (s / 3600.0)
        if ref in ("S", "W"):
            deg = -deg
        return deg
    except Exception:
        return None

def extract_gps_from_pil_exif(exif: Dict[int, Any]) -> Optional[Tuple[float, float, Optional[float]]]:
    if not exif:
        return None
    gps_ifd_tag = TAGS.get("GPSInfo")
    if gps_ifd_tag is None or gps_ifd_tag not in exif:
        return None
    gps_ifd = exif[gps_ifd_tag]
    if not isinstance(gps_ifd, dict):
        return None

    gps = {GPSTAGS.get(k, k): v for k, v in gps_ifd.items()}
    lat = dms_to_deg(gps.get("GPSLatitude"), gps.get("GPSLatitudeRef"))
    lon = dms_to_deg(gps.get("GPSLongitude"), gps.get("GPSLongitudeRef"))
    alt = None
    if "GPSAltitude" in gps:
        a = rational_to_float(gps.get("GPSAltitude"))
        ref = gps.get("GPSAltitudeRef", 0)
        if a is not None:
            alt = -a if ref == 1 else a
    if lat is None or lon is None:
        return None
    return (lat, lon, alt)

def extract_gps_from_piexif_bytes(exif_bytes: bytes) -> Optional[Tuple[float, float, Optional[float]]]:
    if not (have_piexif and exif_bytes):
        return None
    try:
        ex = piexif.load(exif_bytes)
        gps_ifd = ex.get("GPS", {})
        if not gps_ifd:
            return None
        # Keys per EXIF spec
        lat = dms_to_deg(gps_ifd.get(piexif.GPSIFD.GPSLatitude),
                         gps_ifd.get(piexif.GPSIFD.GPSLatitudeRef, b"").decode("ascii", "ignore") or None)
        lon = dms_to_deg(gps_ifd.get(piexif.GPSIFD.GPSLongitude),
                         gps_ifd.get(piexif.GPSIFD.GPSLongitudeRef, b"").decode("ascii", "ignore") or None)
        alt = None
        if piexif.GPSIFD.GPSAltitude in gps_ifd:
            a = rational_to_float(gps_ifd[piexif.GPSIFD.GPSAltitude])
            ref = gps_ifd.get(piexif.GPSIFD.GPSAltitudeRef, 0)
            if a is not None:
                alt = -a if ref == 1 else a
        if lat is None or lon is None:
            return None
        return (lat, lon, alt)
    except Exception:
        return None

def get_exif_any(path: Path) -> Dict[int, Any]:
    """
    Try PIL's getexif(); if empty for HEIC, attempt to pull EXIF bytes and parse via piexif.
    Returns a PIL-style exif dict when possible; otherwise an empty dict (and we’ll use the piexif path).
    """
    try:
        with Image.open(path) as im:
            exif = im.getexif() or {}
            # If we already got GPS via PIL, we're done.
            if extract_gps_from_pil_exif(dict(exif) if exif else {}) is not None:
                return dict(exif)
            # Fallback: try raw EXIF bytes
            exif_bytes = None
            # Some formats expose EXIF bytes here:
            if hasattr(im, "info"):
                exif_bytes = im.info.get("exif")
            # HEIC-specific: pull from pillow-heif metadata
            if (not exif_bytes) and have_pillow_heif and path.suffix.lower() in (".heic", ".heif"):
                try:
                    h = pillow_heif.open_heif(str(path))
                    for md in h.metadata or []:
                        if md.get("type") == "Exif" and md.get("data"):
                            exif_bytes = md["data"]
                            break
                except Exception:
                    pass
            # If we found EXIF bytes, we won't convert to PIL dict here; we will parse for GPS later.
            if exif_bytes:
                # Stash bytes for caller by returning a sentinel dict containing raw bytes
                return {"__RAW_EXIF_BYTES__": exif_bytes}
            return dict(exif) if exif else {}
    except Exception:
        return {}

def build_kml(placemarks: list, doc_name: str) -> str:
    header = f'''<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>{xml_escape(doc_name)}</name>
'''
    body_parts = []
    for pm in placemarks:
        coords = f'{pm["lon"]},{pm["lat"]}'
        if pm.get("alt") is not None:
            coords = f'{coords},{pm["alt"]}'
        name = xml_escape(pm["name"])
        desc = xml_escape(pm.get("desc", ""))
        body_parts.append(f'''    <Placemark>
      <name>{name}</name>
      {"<description>"+desc+"</description>" if desc else ""}
      <Point><coordinates>{coords}</coordinates></Point>
    </Placemark>
''')
    footer = "  </Document>\n</kml>\n"
    return header + "".join(body_parts) + footer

def xml_escape(s: str) -> str:
    return (s.replace("&", "&amp;")
             .replace("<", "&lt;")
             .replace(">", "&gt;")
             .replace('"', "&quot;")
             .replace("'", "&apos;"))

def main():
    ap = argparse.ArgumentParser(description="Scan a folder for geotagged images and generate a KML with placemarks named after the image files.")
    ap.add_argument("folder", help="Path to the folder dropped onto the .bat (images will be scanned recursively).")
    ap.add_argument("-o", "--output", help="Optional output .kml path. Defaults to <foldername>_images.kml inside the folder.")
    args = ap.parse_args()

    folder = Path(args.folder).expanduser().resolve()
    if not folder.exists() or not folder.is_dir():
        print(f"Not a folder: {folder}", file=sys.stderr)
        return 3

    out_path = Path(args.output).expanduser().resolve() if args.output else (folder / f"{folder.name}_images.kml")

    placemarks = []
    total = 0
    with_gps = 0
    skipped = 0

    for root, _, files in os.walk(folder):
        for fname in files:
            p = Path(root) / fname
            if p.suffix.lower() not in IMG_EXTS:
                continue
            total += 1
            try:
                exif_or_raw = get_exif_any(p)

                gps = None
                # First try normal PIL exif dict
                if "__RAW_EXIF_BYTES__" not in exif_or_raw:
                    gps = extract_gps_from_pil_exif(exif_or_raw)
                # Fallback: parse raw EXIF bytes (works great for HEIC)
                if gps is None and "__RAW_EXIF_BYTES__" in exif_or_raw:
                    gps = extract_gps_from_piexif_bytes(exif_or_raw["__RAW_EXIF_BYTES__"])

                if gps is None:
                    skipped += 1
                    continue

                lat, lon, alt = gps
                placemarks.append({
                    "name": p.name,
                    "lat": lat,
                    "lon": lon,
                    "alt": alt,
                    "desc": str(p.relative_to(folder))
                })
                with_gps += 1
            except Exception as e:
                skipped += 1
                print(f"[WARN] {p}: {e}", file=sys.stderr)

    if not placemarks:
        print("No geotagged images found. No KML written.", file=sys.stderr)
        print(f"Scanned {total} images; {with_gps} with GPS; {skipped} skipped.", file=sys.stderr)
        return 4

    kml = build_kml(placemarks, f"{folder.name} (Geotagged Images)")
    try:
        out_path.write_text(kml, encoding="utf-8")
    except Exception as e:
        print(f"Failed to write KML: {e}", file=sys.stderr)
        return 5

    print(f"Wrote KML: {out_path}")
    print(f"Scanned {total} images; {with_gps} with GPS; {skipped} skipped.")
    return 0

if __name__ == "__main__":
    sys.exit(main())
