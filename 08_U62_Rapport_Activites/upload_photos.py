"""
Upload selected construction photos to temporary hosting for Canva integration.
Resizes large images before upload to speed up the process.
"""
import os, sys, io, json, subprocess
from PIL import Image

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PHOTOS_DIR = r"D:\PREPA BTS MEC\08_U62_Rapport_Activites\Documents_Entreprise\Photos_Chantier"
ANNEXES_DIR = r"D:\PREPA BTS MEC\08_U62_Rapport_Activites\Annexes\Images"

# Photos to upload with their intended use
PHOTOS = {
    # Cover & Contact pages (large background)
    "cover": os.path.join(PHOTOS_DIR, "tom-shamberger--KAoVUNv9W0-unsplash.jpg"),
    # Introduction page (portrait/site)
    "intro": os.path.join(ANNEXES_DIR, "photo_bahafid.jpg"),
    # Page 4 - Parcours: Formation Maroc
    "maroc": os.path.join(PHOTOS_DIR, "20180330_153333.jpg"),
    # Page 4 - Parcours: Chantier France
    "france": os.path.join(PHOTOS_DIR, "IMG_20250718_155345[1].jpg"),
    # Page 4 - Parcours: BIMCO
    "bimco": os.path.join(PHOTOS_DIR, "d-c-NFVIq89r_8Q-unsplash.jpg"),
    # Page 5 - Conseil Regional banner
    "conseil": os.path.join(PHOTOS_DIR, "20181206_125854.jpg"),
    # Page 6 - Competences bottom
    "competences_bottom": os.path.join(PHOTOS_DIR, "20180523_102152.jpg"),
    # Page 6 - 6 skill icons
    "skill_revit": os.path.join(PHOTOS_DIR, "20180605_122447.jpg"),
    "skill_python": os.path.join(PHOTOS_DIR, "20180604_133353.jpg"),
    "skill_excel": os.path.join(PHOTOS_DIR, "20180511_155634.jpg"),
    "skill_metres": os.path.join(PHOTOS_DIR, "20180503_164033.jpg"),
    "skill_analyse": os.path.join(PHOTOS_DIR, "20180403_152746.jpg"),
    "skill_suivi": os.path.join(PHOTOS_DIR, "20180405_115900.jpg"),
    # Page 7 - Project 1 (communes)
    "projet1": os.path.join(PHOTOS_DIR, "20180403_123714.jpg"),
    # Page 7 - Project 2 (route)
    "projet2": os.path.join(PHOTOS_DIR, "20181206_125636.jpg"),
    # Page 9 - Bilan financier
    "bilan": os.path.join(PHOTOS_DIR, "pexels-efeburakbaydar-35846752.jpg"),
    # Page 11 - Conclusion
    "conclusion": os.path.join(PHOTOS_DIR, "d-c-zEjnjuA_KBY-unsplash.jpg"),
    # BIMCO banner
    "banniere": os.path.join(ANNEXES_DIR, "banniere_bimco.png"),
}

MAX_SIZE = 1600  # Max dimension in pixels for resize

def resize_and_save_temp(filepath, max_dim=MAX_SIZE):
    """Resize image if needed and return bytes."""
    img = Image.open(filepath)
    w, h = img.size
    if max(w, h) > max_dim:
        ratio = max_dim / max(w, h)
        new_w, new_h = int(w * ratio), int(h * ratio)
        img = img.resize((new_w, new_h), Image.LANCZOS)

    buf = io.BytesIO()
    fmt = "PNG" if filepath.lower().endswith(".png") else "JPEG"
    img.save(buf, format=fmt, quality=85)
    buf.seek(0)
    return buf

def upload_to_0x0(filepath, name):
    """Upload file to 0x0.st and return URL."""
    # First resize
    buf = resize_and_save_temp(filepath)
    ext = ".png" if filepath.lower().endswith(".png") else ".jpg"
    tmp_path = os.path.join(os.environ.get("TEMP", "/tmp"), f"upload_{name}{ext}")

    with open(tmp_path, "wb") as f:
        f.write(buf.read())

    try:
        result = subprocess.run(
            ["curl", "-s", "-F", f"file=@{tmp_path}", "https://0x0.st"],
            capture_output=True, text=True, timeout=60
        )
        url = result.stdout.strip()
        if url.startswith("http"):
            return url
        else:
            print(f"  ERREUR upload {name}: {result.stdout} {result.stderr}")
            return None
    finally:
        os.remove(tmp_path)

results = {}
for name, path in PHOTOS.items():
    if not os.path.exists(path):
        print(f"SKIP {name}: fichier introuvable - {path}")
        continue
    print(f"Upload {name}...", end=" ", flush=True)
    url = upload_to_0x0(path, name)
    if url:
        results[name] = url
        print(f"OK -> {url}")
    else:
        print("ECHEC")

print(f"\n=== {len(results)}/{len(PHOTOS)} photos uploadees ===")
print(json.dumps(results, indent=2))
