import os
import json
import base64
import re

base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
data_path = os.path.join(base_dir, "data", "parts.json")
images_dir = os.path.join(base_dir, "images")

os.makedirs(images_dir, exist_ok=True)

with open(data_path, "r", encoding="utf-8") as f:
    parts = json.load(f)

extracted_count = 0
total_bytes_saved = 0

for row in parts:
    lc_raw = row.get("LC Images")
    if not lc_raw:
        continue

    try:
        lc_map = json.loads(lc_raw) if isinstance(lc_raw, str) else lc_raw
    except Exception as e:
        print(f"Error parsing LC Images for {row.get('Part Name')}: {e}")
        continue

    if not isinstance(lc_map, dict):
        continue

    new_map = {}
    row_modified = False

    part_clean = re.sub(r"[^a-zA-Z0-9_-]", "_", str(row.get("Part Name", "PART")))
    ver_clean = re.sub(r"[^a-zA-Z0-9_-]", "_", str(row.get("Version", "VER")))
    date_str = str(row.get("Date", "20260101_000000"))
    date_clean = re.sub(r"[-:\s]", "", date_str)[:15]

    for lc, val in lc_map.items():
        if isinstance(val, str) and val.startswith("data:image"):
            # Extract header and base64 payload
            if "," in val:
                header, b64_payload = val.split(",", 1)
            else:
                b64_payload = val

            # Determine extension
            ext = "png"
            if "image/jpeg" in val or "image/jpg" in val:
                ext = "jpg"
            elif "image/webp" in val:
                ext = "webp"

            filename = f"{part_clean}_v{ver_clean}_{lc}_{date_clean}.{ext}"
            rel_path = f"images/{filename}"
            full_path = os.path.join(images_dir, filename)

            try:
                img_data = base64.b64decode(b64_payload)
                with open(full_path, "wb") as img_file:
                    img_file.write(img_data)

                new_map[lc] = rel_path
                extracted_count += 1
                total_bytes_saved += len(val)
                row_modified = True
                print(f"  [OK] Extracted {filename} ({len(img_data):,} bytes)")
            except Exception as ex:
                print(f"  [FAIL] Could not decode {filename}: {ex}")
                new_map[lc] = val
        else:
            new_map[lc] = val

    if row_modified:
        row["LC Images"] = json.dumps(new_map)

# Save updated parts.json
initial_size = os.path.getsize(data_path)
with open(data_path, "w", encoding="utf-8", newline="\n") as f:
    json.dump(parts, f, indent=2, ensure_ascii=False)

final_size = os.path.getsize(data_path)

print("=" * 60)
print(f"Extraction Complete!")
print(f"Total images extracted: {extracted_count}")
print(f"Initial parts.json size: {initial_size / (1024*1024):.2f} MB ({initial_size:,} bytes)")
print(f"Final parts.json size:   {final_size / 1024:.2f} KB ({final_size:,} bytes)")
print(f"Size reduction:          {(1 - final_size / initial_size) * 100:.2f}%")
print("=" * 60)
