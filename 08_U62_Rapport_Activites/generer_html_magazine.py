# -*- coding: utf-8 -*-
"""
Generateur Rapport U62 - PDF Magazine Style Constructys V2
BAHAFID Mohamed - BTS MEC Session 2026
Design fidele au rapport Constructys 2022 : formes geometriques larges,
capsules KPI, couleurs bold, saignee aux bords, badges multicolores
"""
import os, html, math, base64, io
from PIL import Image

from contenu_v2 import (
    CANDIDAT, PAGE_01, PAGE_02, PAGE_03, PAGE_04, PAGE_05, PAGE_06, PAGE_07,
    PAGE_08, PAGE_09, PAGE_10, PAGE_15, PAGE_16, PAGE_17, PAGE_18,
    PAGE_20, PAGE_21, PAGE_22, PAGE_23, PAGE_24, PAGE_25, PAGE_26,
    PAGE_27, PAGE_28, PAGE_29, PAGE_30,
    SITUATION_1, SITUATION_2, SITUATION_3, SITUATION_4, SITUATION_5,
)

BASE_DIR    = r"D:\PREPA BTS MEC\08_U62_Rapport_Activites"
PHOTOS_DIR  = os.path.join(BASE_DIR, "Documents_Entreprise", "Photos_Chantier")
ANNEXES_DIR = os.path.join(BASE_DIR, "Annexes", "Images")
OUT_HTML    = os.path.join(BASE_DIR, "Rapport_Redaction", "RAPPORT_U62_BAHAFID_MAGAZINE.html")
OUT_PDF     = os.path.join(BASE_DIR, "Rapport_Redaction", "RAPPORT_U62_BAHAFID_MAGAZINE.pdf")

PHOTO = {
    "bahafid":      os.path.join(ANNEXES_DIR,  "photo_bahafid.jpg"),
    "logo":         os.path.join(ANNEXES_DIR,  "logo_bimco.png"),
    "banniere":     os.path.join(ANNEXES_DIR,  "banniere_bimco.png"),
    "cao":          os.path.join(ANNEXES_DIR,  "convocation_cao.jpg"),
    "rejet":        os.path.join(ANNEXES_DIR,  "notification_rejet.jpg"),
    "cps":          os.path.join(ANNEXES_DIR,  "cps_marche46_signe.jpg"),
    "maroc":        os.path.join(PHOTOS_DIR,   "20180330_153333.jpg"),
    "france":       os.path.join(PHOTOS_DIR,   "IMG_20250718_155345[1].jpg"),
    "conseil":      os.path.join(PHOTOS_DIR,   "20181206_125854.jpg"),
    "chantier1":    os.path.join(PHOTOS_DIR,   "20180523_102152.jpg"),
    "chantier2":    os.path.join(PHOTOS_DIR,   "20180605_122447.jpg"),
    "chantier3":    os.path.join(PHOTOS_DIR,   "20180604_133353.jpg"),
    "chantier4":    os.path.join(PHOTOS_DIR,   "20180511_155634.jpg"),
    "chantier5":    os.path.join(PHOTOS_DIR,   "20180503_164033.jpg"),
    "chantier6":    os.path.join(PHOTOS_DIR,   "20180403_152746.jpg"),
    "chantier7":    os.path.join(PHOTOS_DIR,   "20180405_115900.jpg"),
    "route":        os.path.join(PHOTOS_DIR,   "20181206_130504.jpg"),
    "route2":       os.path.join(PHOTOS_DIR,   "20181206_141917.jpg"),
    "montagne":     os.path.join(PHOTOS_DIR,   "20181113_154208.jpg"),
    "terrassement": os.path.join(PHOTOS_DIR,   "20181206_133336.jpg"),
    "gros_oeuvre":  os.path.join(PHOTOS_DIR,   "20180228_144045.jpg"),
    "vrd":          os.path.join(PHOTOS_DIR,   "20180330_154009.jpg"),
    "chantier_urban": os.path.join(PHOTOS_DIR, "20180503_105806.jpg"),
    "projet1":      os.path.join(PHOTOS_DIR,   "20180403_123714.jpg"),
    "bimco_tech":   os.path.join(PHOTOS_DIR,   "d-c-NFVIq89r_8Q-unsplash.jpg"),
    "bilan":        os.path.join(PHOTOS_DIR,   "pexels-efeburakbaydar-35846752.jpg"),
    "cover":        os.path.join(PHOTOS_DIR,   "tom-shamberger--KAoVUNv9W0-unsplash.jpg"),
}

def b64img(name, max_w=700, max_h=550, quality=78):
    """Encode une photo en data:URI base64 (JPEG, redim Pillow)."""
    path = PHOTO.get(name, "")
    if not path or not os.path.exists(path):
        return ""
    try:
        with Image.open(path) as im:
            im = im.convert("RGB")
            w, h = im.size
            ratio = min(max_w / w, max_h / h, 1.0)
            if ratio < 1.0:
                im = im.resize((int(w * ratio), int(h * ratio)), Image.LANCZOS)
            buf = io.BytesIO()
            mime = "image/jpeg"
            if path.lower().endswith(".png"):
                im.save(buf, format="PNG", optimize=True); mime = "image/png"
            else:
                im.save(buf, format="JPEG", quality=quality, optimize=True)
            return f"data:{mime};base64,{base64.b64encode(buf.getvalue()).decode()}"
    except Exception as e:
        print(f"  img {name}: {e}"); return ""

def img_tag(name, style="width:100%;height:30mm;object-fit:cover;border-radius:10px",
            alt="", max_w=700, max_h=550):
    """<img> data-URI ou fallback .ph."""
    uri = b64img(name, max_w=max_w, max_h=max_h)
    if not uri:
        return f'<div class="ph" style="min-height:30mm">{alt or name}</div>'
    return f'<img src="{uri}" alt="{alt or name}" style="{style};display:block">'

def img_circle(name, size_mm=50):
    """Photo ronde (cercle) pour couverture/profil."""
    uri = b64img(name, max_w=400, max_h=400)
    if not uri:
        return f'<div style="width:{size_mm}mm;height:{size_mm}mm;border-radius:50%;background:rgba(255,255,255,0.15)"></div>'
    return (f'<div style="width:{size_mm}mm;height:{size_mm}mm;border-radius:50%;overflow:hidden;'
            f'border:2.5px solid rgba(255,255,255,0.6);box-shadow:0 4px 20px rgba(0,0,0,0.3)">'
            f'<img src="{uri}" style="width:100%;height:100%;object-fit:cover;display:block"></div>')
H = html.escape

# -- Colors Constructys --
OG = "#F39200"   # orange
NV = "#1E3A5F"   # navy
TQ = "#5CC8C0"   # turquoise
WH = "#FFFFFF"
LG = "#F5F5F5"
PR = "#8B5CF6"   # purple
RD = "#EF4444"   # red
GN = "#10B981"   # green

# ━━━━━━ SVG Helpers ━━━━━━
def svg_donut(segs, txt, sub="", sz=200):
    cx=cy=sz/2; r=sz/2-26; sw=30; ci=2*math.pi*r
    arcs=[]; off=0
    for p,c in segs:
        d=ci*p/100; g=ci-d; rot=off*360/100-90
        arcs.append(f'<circle cx="{cx}" cy="{cy}" r="{r}" fill="none" stroke="{c}" stroke-width="{sw}" stroke-dasharray="{d:.1f} {g:.1f}" transform="rotate({rot:.1f} {cx} {cy})"/>')
        off+=p
    return f'''<svg viewBox="0 0 {sz} {sz}" width="{sz}" height="{sz}" xmlns="http://www.w3.org/2000/svg">
<circle cx="{cx}" cy="{cy}" r="{r}" fill="none" stroke="#EEE" stroke-width="{sw}"/>
{"".join(arcs)}
<circle cx="{cx}" cy="{cy}" r="{r-sw/2-5}" fill="white"/>
<text x="{cx}" y="{cy-4}" text-anchor="middle" font-size="17" font-weight="900" fill="{NV}" font-family="Montserrat,sans-serif">{H(txt)}</text>
<text x="{cx}" y="{cy+13}" text-anchor="middle" font-size="9" fill="#888" font-family="Inter,sans-serif">{H(sub)}</text>
</svg>'''

def svg_stripes(w,h,color="rgba(255,255,255,0.12)",gap=7,sw=1.5):
    lines="".join(f'<line x1="{x}" y1="{h}" x2="{x+h}" y2="0" stroke="{color}" stroke-width="{sw}"/>' for x in range(-h,w+h,gap))
    return f'<svg width="{w}" height="{h}" viewBox="0 0 {w} {h}" xmlns="http://www.w3.org/2000/svg">{lines}</svg>'

def svg_dots(w,h,color="rgba(255,255,255,0.35)",r=2.2,gap=8):
    dots="".join(f'<circle cx="{x}" cy="{y}" r="{r}" fill="{color}"/>' for x in range(0,w,gap) for y in range(0,h,gap))
    return f'<svg width="{w}" height="{h}" viewBox="0 0 {w} {h}" xmlns="http://www.w3.org/2000/svg">{dots}</svg>'

def svg_circle_deco(sz,color,sw=1.5):
    r=sz/2-sw
    return f'<svg width="{sz}" height="{sz}" viewBox="0 0 {sz} {sz}" xmlns="http://www.w3.org/2000/svg"><circle cx="{sz/2}" cy="{sz/2}" r="{r}" fill="none" stroke="{color}" stroke-width="{sw}" stroke-dasharray="4 3"/></svg>'

# Constructys-style geometric shapes
def svg_pill(w,h,color,rx=None):
    """Pill / capsule shape like Constructys KPI blocks"""
    if rx is None: rx=h/2
    return f'<svg width="{w}" height="{h}" viewBox="0 0 {w} {h}" xmlns="http://www.w3.org/2000/svg"><rect x="0" y="0" width="{w}" height="{h}" rx="{rx}" fill="{color}"/></svg>'

def svg_half_circle(sz,color,side="right"):
    """Half circle that bleeds off edge, like Constructys decorations"""
    if side=="right":
        return f'<svg width="{sz//2}" height="{sz}" viewBox="0 0 {sz//2} {sz}" xmlns="http://www.w3.org/2000/svg"><circle cx="0" cy="{sz//2}" r="{sz//2}" fill="{color}"/></svg>'
    elif side=="left":
        return f'<svg width="{sz//2}" height="{sz}" viewBox="0 0 {sz//2} {sz}" xmlns="http://www.w3.org/2000/svg"><circle cx="{sz//2}" cy="{sz//2}" r="{sz//2}" fill="{color}"/></svg>'
    elif side=="bottom":
        return f'<svg width="{sz}" height="{sz//2}" viewBox="0 0 {sz} {sz//2}" xmlns="http://www.w3.org/2000/svg"><circle cx="{sz//2}" cy="0" r="{sz//2}" fill="{color}"/></svg>'
    return f'<svg width="{sz}" height="{sz//2}" viewBox="0 0 {sz} {sz//2}" xmlns="http://www.w3.org/2000/svg"><circle cx="{sz//2}" cy="{sz//2}" r="{sz//2}" fill="{color}"/></svg>'

# ━━━━━━ SVG ICONS ━━━━━━
def _ico(path_d, sz=20, color="#fff", vb=24):
    return f'<svg width="{sz}" height="{sz}" viewBox="0 0 {vb} {vb}" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="{path_d}" fill="{color}"/></svg>'

def ico_briefcase(sz=20,c="#fff"):
    return _ico("M20 7h-4V5c0-1.1-.9-2-2-2h-4C8.9 3 8 3.9 8 5v2H4c-1.1 0-2 .9-2 2v5c0 .6.4 1 1 1h6v-1c0-.6.4-1 1-1h4c.6 0 1 .4 1 1v1h6c.6 0 1-.4 1-1V9c0-1.1-.9-2-2-2zM10 5h4v2h-4V5zm5 10h-6v1H3v3c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2v-3h-6v-1z",sz,c)

def ico_building(sz=20,c="#fff"):
    return _ico("M12 7V3H2v18h20V7H12zM6 19H4v-2h2v2zm0-4H4v-2h2v2zm0-4H4V9h2v2zm0-4H4V5h2v2zm4 12H8v-2h2v2zm0-4H8v-2h2v2zm0-4H8V9h2v2zm0-4H8V5h2v2zm10 12h-8v-2h2v-2h-2v-2h2v-2h-2V9h8v10zm-2-8h-2v2h2v-2zm0 4h-2v2h2v-2z",sz,c)

def ico_chart(sz=20,c="#fff"):
    return _ico("M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zM9 17H7v-7h2v7zm4 0h-2V7h2v10zm4 0h-2v-4h2v4z",sz,c)

def ico_tools(sz=20,c="#fff"):
    return _ico("M22.7 19l-9.1-9.1c.9-2.3.4-5-1.5-6.9-2-2-5-2.4-7.4-1.3L9 6 6 9 1.6 4.7C.4 7.1.9 10.1 2.9 12.1c1.9 1.9 4.6 2.4 6.9 1.5l9.1 9.1c.4.4 1 .4 1.4 0l2.3-2.3c.5-.4.5-1.1.1-1.4z",sz,c)

def ico_graduation(sz=20,c="#fff"):
    return _ico("M5 13.18v4L12 21l7-3.82v-4L12 17l-7-3.82zM12 3L1 9l11 6 9-4.91V17h2V9L12 3z",sz,c)

def ico_road(sz=20,c="#fff"):
    return _ico("M11 2H13V5H11V2ZM11 7H13V10H11V7ZM11 12H13V15H11V12ZM11 17H13V22H11V17ZM4 2L8 22H6L2 2H4ZM20 2L16 22H18L22 2H20Z",sz,c)

def ico_map_pin(sz=20,c="#fff"):
    return _ico("M12 2C8.13 2 5 5.13 5 9c0 5.25 7 13 7 13s7-7.75 7-13c0-3.87-3.13-7-7-7zm0 9.5c-1.38 0-2.5-1.12-2.5-2.5s1.12-2.5 2.5-2.5 2.5 1.12 2.5 2.5-1.12 2.5-2.5 2.5z",sz,c)

def ico_users(sz=20,c="#fff"):
    return _ico("M16 11c1.66 0 2.99-1.34 2.99-3S17.66 5 16 5c-1.66 0-3 1.34-3 3s1.34 3 3 3zm-8 0c1.66 0 2.99-1.34 2.99-3S9.66 5 8 5C6.34 5 5 6.34 5 8s1.34 3 3 3zm0 2c-2.33 0-7 1.17-7 3.5V19h14v-2.5c0-2.33-4.67-3.5-7-3.5zm8 0c-.29 0-.62.02-.97.05 1.16.84 1.97 1.97 1.97 3.45V19h6v-2.5c0-2.33-4.67-3.5-7-3.5z",sz,c)

def ico_target(sz=20,c="#fff"):
    return _ico("M12 2C6.49 2 2 6.49 2 12s4.49 10 10 10 10-4.49 10-10S17.51 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8zm3-8c0 1.66-1.34 3-3 3s-3-1.34-3-3 1.34-3 3-3 3 1.34 3 3zm-3-5c-2.76 0-5 2.24-5 5s2.24 5 5 5 5-2.24 5-5-2.24-5-5-5z",sz,c)

def ico_trophy(sz=20,c="#fff"):
    return _ico("M19 5h-2V3H7v2H5c-1.1 0-2 .9-2 2v1c0 2.55 1.92 4.63 4.39 4.94.63 1.5 1.98 2.63 3.61 2.96V19H7v2h10v-2h-4v-3.1c1.63-.33 2.98-1.46 3.61-2.96C19.08 12.63 21 10.55 21 8V7c0-1.1-.9-2-2-2zM5 8V7h2v3.82C5.84 10.4 5 9.3 5 8zm14 0c0 1.3-.84 2.4-2 2.82V7h2v1z",sz,c)

def ico_shield(sz=20,c="#fff"):
    return _ico("M12 1L3 5v6c0 5.55 3.84 10.74 9 12 5.16-1.26 9-6.45 9-12V5l-9-4zm-2 16l-4-4 1.41-1.41L10 14.17l6.59-6.59L18 9l-8 8z",sz,c)

def ico_clock(sz=20,c="#fff"):
    return _ico("M11.99 2C6.47 2 2 6.48 2 12s4.47 10 9.99 10C17.52 22 22 17.52 22 12S17.52 2 11.99 2zM12 20c-4.42 0-8-3.58-8-8s3.58-8 8-8 8 3.58 8 8-3.58 8-8 8zm.5-13H11v6l5.25 3.15.75-1.23-4.5-2.67V7z",sz,c)

def ico_check(sz=20,c="#fff"):
    return _ico("M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z",sz,c)

def ico_warning(sz=20,c="#fff"):
    return _ico("M1 21h22L12 2 1 21zm12-3h-2v-2h2v2zm0-4h-2v-4h2v4z",sz,c)

def ico_star(sz=20,c="#fff"):
    return _ico("M12 17.27L18.18 21l-1.64-7.03L22 9.24l-7.19-.61L12 2 9.19 8.63 2 9.24l5.46 4.73L5.82 21z",sz,c)

def ico_cube(sz=20,c="#fff"):
    return _ico("M21 16.5c0 .38-.21.71-.53.88l-7.9 4.44c-.16.12-.36.18-.57.18-.21 0-.41-.06-.57-.18l-7.9-4.44A.991.991 0 013 16.5v-9c0-.38.21-.71.53-.88l7.9-4.44c.16-.12.36-.18.57-.18.21 0 .41.06.57.18l7.9 4.44c.32.17.53.5.53.88v9zM12 4.15L6.04 7.5 12 10.85l5.96-3.35L12 4.15zM5 15.91l6 3.37v-6.73L5 9.18v6.73zm14 0v-6.73l-6 3.37v6.73l6-3.37z",sz,c)

def ico_bolt(sz=20,c="#fff"):
    return _ico("M11 21h-1l1-7H7.5c-.88 0-.33-.75-.31-.78C8.48 10.94 10.42 7.54 13.01 3h1l-1 7h3.51c.4 0 .62.19.4.66C12.97 17.55 11 21 11 21z",sz,c)

def ico_globe(sz=20,c="#fff"):
    return _ico("M11.99 2C6.47 2 2 6.48 2 12s4.47 10 9.99 10C17.52 22 22 17.52 22 12S17.52 2 11.99 2zm6.93 6h-2.95c-.32-1.25-.78-2.45-1.38-3.56 1.84.63 3.37 1.91 4.33 3.56zM12 4.04c.83 1.2 1.48 2.53 1.91 3.96h-3.82c.43-1.43 1.08-2.76 1.91-3.96zM4.26 14C4.1 13.36 4 12.69 4 12s.1-1.36.26-2h3.38c-.08.66-.14 1.32-.14 2 0 .68.06 1.34.14 2H4.26zm.82 2h2.95c.32 1.25.78 2.45 1.38 3.56-1.84-.63-3.37-1.91-4.33-3.56zm2.95-8H5.08c.96-1.65 2.49-2.93 4.33-3.56C8.81 5.55 8.35 6.75 8.03 8zM12 19.96c-.83-1.2-1.48-2.53-1.91-3.96h3.82c-.43 1.43-1.08 2.76-1.91 3.96zM14.34 14H9.66c-.09-.66-.16-1.32-.16-2 0-.68.07-1.35.16-2h4.68c.09.65.16 1.32.16 2 0 .68-.07 1.34-.16 2zm.25 5.56c.6-1.11 1.06-2.31 1.38-3.56h2.95c-.96 1.65-2.49 2.93-4.33 3.56zM16.36 14c.08-.66.14-1.32.14-2 0-.68-.06-1.34-.14-2h3.38c.16.64.26 1.31.26 2s-.1 1.36-.26 2h-3.38z",sz,c)

# Icon in colored circle badge
def ico_badge(icon_fn, bg_color, sz_circle=32, sz_icon=16):
    icon_svg = icon_fn(sz_icon, "#fff")
    return f'<div style="width:{sz_circle}px;height:{sz_circle}px;border-radius:50%;background:{bg_color};display:flex;align-items:center;justify-content:center;flex-shrink:0">{icon_svg}</div>'

# ━━━━━━ SVG Charts (new) ━━━━━━
def svg_progress_ring(pct, color, sz=60, sw=6, label=""):
    """Circular progress ring"""
    cx=cy=sz/2; r=sz/2-sw-2; ci=2*math.pi*r; d=ci*pct/100
    return f'''<svg width="{sz}" height="{sz}" viewBox="0 0 {sz} {sz}" xmlns="http://www.w3.org/2000/svg">
<circle cx="{cx}" cy="{cy}" r="{r}" fill="none" stroke="#EEE" stroke-width="{sw}"/>
<circle cx="{cx}" cy="{cy}" r="{r}" fill="none" stroke="{color}" stroke-width="{sw}" stroke-dasharray="{d:.1f} {ci-d:.1f}" transform="rotate(-90 {cx} {cy})" stroke-linecap="round"/>
<text x="{cx}" y="{cy+2}" text-anchor="middle" font-size="{sz//5}" font-weight="800" fill="{color}" font-family="Montserrat,sans-serif">{pct}%</text>
{f'<text x="{cx}" y="{cy+sz//5+3}" text-anchor="middle" font-size="{sz//9}" fill="#888" font-family="Inter,sans-serif">{label}</text>' if label else ''}
</svg>'''

def svg_bar_chart(data, w=280, h=140, bar_w=28):
    """Vertical bar chart. data = [(label, value, color), ...]"""
    mx = max(v for _,v,_ in data)
    n = len(data)
    gap = (w - n*bar_w) / (n+1)
    bars = ""
    for i,(lbl,val,clr) in enumerate(data):
        bh = val/mx * (h-30) if mx else 0
        x = gap + i*(bar_w+gap)
        y = h - 18 - bh
        bars += f'<rect x="{x:.1f}" y="{y:.1f}" width="{bar_w}" height="{bh:.1f}" rx="4" fill="{clr}"/>'
        bars += f'<text x="{x+bar_w/2:.1f}" y="{y-3}" text-anchor="middle" font-size="8" font-weight="700" fill="{clr}" font-family="Montserrat">{val}</text>'
        bars += f'<text x="{x+bar_w/2:.1f}" y="{h-4}" text-anchor="middle" font-size="6.5" fill="#888" font-family="Inter">{lbl}</text>'
    return f'<svg width="{w}" height="{h}" viewBox="0 0 {w} {h}" xmlns="http://www.w3.org/2000/svg"><line x1="0" y1="{h-18}" x2="{w}" y2="{h-18}" stroke="#EEE" stroke-width="1"/>{bars}</svg>'

def svg_radar(labels, values, color, sz=180):
    """Radar/spider chart. values 0-100."""
    n = len(labels); cx=cy=sz/2; r=sz/2-25
    angles = [i*2*math.pi/n - math.pi/2 for i in range(n)]
    # Grid
    grid = ""
    for lv in [0.25,0.5,0.75,1.0]:
        pts = " ".join(f"{cx+r*lv*math.cos(a):.1f},{cy+r*lv*math.sin(a):.1f}" for a in angles)
        grid += f'<polygon points="{pts}" fill="none" stroke="#E0E0E0" stroke-width="0.8"/>'
    # Axes
    for a in angles:
        grid += f'<line x1="{cx}" y1="{cy}" x2="{cx+r*math.cos(a):.1f}" y2="{cy+r*math.sin(a):.1f}" stroke="#E8E8E8" stroke-width="0.5"/>'
    # Data
    pts_d = " ".join(f"{cx+r*v/100*math.cos(a):.1f},{cy+r*v/100*math.sin(a):.1f}" for v,a in zip(values,angles))
    data_poly = f'<polygon points="{pts_d}" fill="{color}" fill-opacity="0.2" stroke="{color}" stroke-width="2"/>'
    # Dots + labels
    dots = ""
    for i,(v,a) in enumerate(zip(values,angles)):
        dx=cx+r*v/100*math.cos(a); dy=cy+r*v/100*math.sin(a)
        dots += f'<circle cx="{dx:.1f}" cy="{dy:.1f}" r="3" fill="{color}"/>'
        lx=cx+(r+14)*math.cos(a); ly=cy+(r+14)*math.sin(a)
        anch="middle"
        if math.cos(a) > 0.3: anch="start"
        elif math.cos(a) < -0.3: anch="end"
        dots += f'<text x="{lx:.1f}" y="{ly+3:.1f}" text-anchor="{anch}" font-size="7" fill="#555" font-family="Inter">{labels[i]}</text>'
    return f'<svg width="{sz}" height="{sz}" viewBox="0 0 {sz} {sz}" xmlns="http://www.w3.org/2000/svg">{grid}{data_poly}{dots}</svg>'

def svg_wave(w=800,h=40,color="#F39200",opacity=0.15):
    """Wave divider"""
    return f'<svg width="{w}" height="{h}" viewBox="0 0 {w} {h}" xmlns="http://www.w3.org/2000/svg"><path d="M0 {h/2} Q{w/8} 0 {w/4} {h/2} T{w/2} {h/2} T{3*w/4} {h/2} T{w} {h/2} V{h} H0Z" fill="{color}" opacity="{opacity}"/></svg>'

def svg_quote_mark(sz=40,color="#F39200"):
    """Large decorative quote mark"""
    return f'<svg width="{sz}" height="{sz}" viewBox="0 0 24 24" fill="{color}" opacity="0.2" xmlns="http://www.w3.org/2000/svg"><path d="M6 17h3l2-4V7H5v6h3zm8 0h3l2-4V7h-6v6h3z"/></svg>'

# ━━━━━━ CSS ━━━━━━
CSS = r"""
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;500;600;700;800;900&family=Inter:wght@300;400;500;600;700&display=swap');
@page{size:A4 portrait;margin:0}
*,*::before,*::after{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Inter',sans-serif;color:#2D2D2D;background:#fff;font-size:9.5pt;line-height:1.6;-webkit-font-smoothing:antialiased}
.m{font-family:'Montserrat',sans-serif}

/* PAGE */
.page{width:210mm;height:297mm;background:#fff;position:relative;overflow:hidden;page-break-after:always;page-break-inside:avoid}

/* FOOTER */
.ft{position:absolute;bottom:5mm;left:0;right:0;text-align:center;font-size:7pt;color:#999;z-index:50}
.ft b{color:#F39200;font-weight:600}

/* ACCENT BAR */
.bar-og{width:25mm;height:4px;background:#F39200;border-radius:2px;margin:3mm 0 4mm}

/* DOTTED SEPARATOR */
.dot-sep{border:none;border-top:2px dotted #DDD;margin:4mm 0}

/* BADGES - Constructys multicolor */
.badge{display:inline-block;padding:1.5mm 4mm;border-radius:4px;font-size:7pt;font-weight:800;text-transform:uppercase;letter-spacing:0.8px;font-family:'Montserrat',sans-serif;color:#fff}
.badge-og{background:#F39200}
.badge-nv{background:#1E3A5F}
.badge-tq{background:#5CC8C0}
.badge-pr{background:#8B5CF6}
.badge-rd{background:#EF4444}
.badge-gn{background:#10B981}

/* TEXT CLASSES */
.title-huge{font-family:'Montserrat',sans-serif;font-size:28pt;font-weight:900;color:#1E3A5F;line-height:1.08;text-transform:uppercase}
.title-sec{font-family:'Montserrat',sans-serif;font-size:20pt;font-weight:800;color:#1E3A5F;line-height:1.12;text-transform:uppercase}
.title-art{font-family:'Montserrat',sans-serif;font-size:13pt;font-weight:800;color:#F39200;line-height:1.25}
.body{font-size:9.5pt;line-height:1.65;text-align:justify;color:#444}
.body b{color:#2D2D2D}
.kpi-huge{font-family:'Montserrat',sans-serif;font-size:32pt;font-weight:900;color:#1E3A5F;line-height:1}
.kpi-label{font-size:9pt;color:#F39200;font-weight:600}
.label-caps{font-family:'Montserrat',sans-serif;font-size:7.5pt;font-weight:700;letter-spacing:2.5px;text-transform:uppercase;color:#F39200}

/* LAYOUTS */
.col2{display:flex;gap:5mm}
.col2>*{flex:1}
.col3{display:flex;gap:4mm}
.col3>*{flex:1}

/* CAPSULE KPI - Constructys style stacked pills */
.capsule{border-radius:50px;padding:5mm 8mm;display:flex;align-items:center;gap:4mm;position:relative;overflow:hidden}
.capsule-og{background:#F39200;color:#fff}
.capsule-nv{background:#1E3A5F;color:#fff}
.capsule-tq{background:#5CC8C0;color:#fff}
.capsule-lt{background:#fff;border:2px solid #EEE;color:#1E3A5F}
.capsule .num{font-family:'Montserrat',sans-serif;font-size:22pt;font-weight:900;line-height:1}
.capsule .lab{font-size:9pt;font-weight:600;line-height:1.3}

/* CARDS */
.card{background:#fff;border-radius:12px;box-shadow:0 2px 12px rgba(0,0,0,0.07);padding:5mm;position:relative;overflow:hidden}
.card-og{background:#F39200;border-radius:12px;padding:5mm;color:#fff;position:relative;overflow:hidden}
.card-nv{background:#1E3A5F;border-radius:12px;padding:5mm;color:#fff;position:relative;overflow:hidden}
.card-tq{background:#5CC8C0;border-radius:12px;padding:5mm;color:#fff;position:relative;overflow:hidden}
.card-light{background:#F5F5F5;border-radius:12px;padding:5mm;position:relative;overflow:hidden}

/* TABLE */
.tbl{width:100%;border-collapse:separate;border-spacing:0;font-size:8.5pt}
.tbl th{background:#1E3A5F;color:#fff;padding:3.5mm 3mm;font-weight:700;text-align:center;font-size:7.5pt;text-transform:uppercase;letter-spacing:0.5px;font-family:'Montserrat',sans-serif}
.tbl th:first-child{text-align:left;border-radius:6px 0 0 0}
.tbl th:last-child{border-radius:0 6px 0 0}
.tbl td{padding:3.5mm 3mm;border-bottom:1px solid #EEE;vertical-align:middle}
.tbl tr:nth-child(even){background:#F9F9F9}

/* PROGRESS BAR */
.pbar{display:flex;align-items:center;gap:2mm}
.pbar-t{width:24mm;height:6mm;background:#EEE;border-radius:3px;overflow:hidden}
.pbar-f{height:100%;border-radius:3px}
.pbar-l{font-size:7.5pt;font-weight:700}

/* HORIZONTAL BAR */
.hb{display:flex;align-items:center;margin-bottom:3mm}
.hb-l{width:36mm;font-size:8.5pt;font-weight:600;text-align:right;padding-right:3mm;color:#1E3A5F}
.hb-t{flex:1;height:9mm;background:#F0F0F0;border-radius:5px;overflow:hidden}
.hb-f{height:100%;border-radius:5px;display:flex;align-items:center;padding-left:3mm}
.hb-v{font-size:7.5pt;color:#fff;font-weight:700}
.hb-r{width:32mm;font-size:8pt;color:#666;padding-left:3mm}

/* CPAR */
.cpar-hd{background:#1E3A5F;padding:8mm 15mm;display:flex;align-items:center;gap:5mm;position:relative;overflow:hidden}
.cpar-n{font-family:'Montserrat',sans-serif;font-size:56pt;font-weight:900;color:rgba(255,255,255,0.1);line-height:1}
.cpar-i{flex:1}
.cpar-i .lb{font-size:7.5pt;color:#5CC8C0;text-transform:uppercase;letter-spacing:3px;font-family:'Montserrat',sans-serif;font-weight:600}
.cpar-i .tt{font-size:16pt;font-weight:800;color:#fff;margin-top:1mm;font-family:'Montserrat',sans-serif}
.cpar-k{background:rgba(255,255,255,0.08);border:1.5px solid rgba(255,255,255,0.2);padding:4mm 5mm;border-radius:10px;text-align:center}
.cpar-k .v{font-size:20pt;font-weight:900;color:#F39200;font-family:'Montserrat',sans-serif}
.cpar-k .l{font-size:7pt;color:rgba(255,255,255,0.65);margin-top:1mm}

.cpar-bd{padding:4mm 15mm}
.cpar-bk{display:flex;gap:3.5mm;padding:5mm;margin-bottom:3mm;border-radius:10px;border-left:4px solid;background:#fff;box-shadow:0 1px 6px rgba(0,0,0,0.05)}
.cpar-lt{width:12mm;height:12mm;border-radius:50%;color:#fff;font-weight:900;font-size:14pt;flex-shrink:0;display:flex;align-items:center;justify-content:center;font-family:'Montserrat',sans-serif}
.cpar-ct{flex:1}
.cpar-ct h4{font-size:8pt;font-weight:800;text-transform:uppercase;letter-spacing:1px;margin-bottom:2mm;font-family:'Montserrat',sans-serif}
.cpar-ct p{font-size:9pt;line-height:1.65;text-align:justify;color:#555}

.cb-c{border-color:#999}.cb-c .cpar-lt{background:#999}.cb-c h4{color:#666}
.cb-p{border-color:#F39200}.cb-p .cpar-lt{background:#F39200}.cb-p h4{color:#F39200}
.cb-a{border-color:#5CC8C0}.cb-a .cpar-lt{background:#5CC8C0}.cb-a h4{color:#5CC8C0}
.cb-r{border-color:#1E3A5F}.cb-r .cpar-lt{background:#1E3A5F}.cb-r h4{color:#1E3A5F}

.cpar-ft{background:#F5F5F5;border-top:3px solid #F39200;text-align:center;padding:3.5mm;font-size:8pt;font-weight:700;color:#1E3A5F;letter-spacing:2px;font-family:'Montserrat',sans-serif;text-transform:uppercase}

/* COMPARISON */
.cmp{display:flex;gap:3mm;margin-bottom:4mm;align-items:stretch}
.cmp-th{width:26mm;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:1.5mm;flex-shrink:0}
.cmp-c{width:11mm;height:11mm;border-radius:50%;background:#F39200;color:#fff;display:flex;align-items:center;justify-content:center;font-weight:900;font-size:12pt;font-family:'Montserrat',sans-serif}
.cmp-tl{font-size:7pt;color:#F39200;font-weight:700;text-align:center;font-family:'Montserrat',sans-serif}
.cmp-cd{flex:1;border-radius:10px;overflow:hidden;box-shadow:0 1px 6px rgba(0,0,0,0.06)}
.cmp-bn{padding:2.5mm 3mm;color:#fff;font-size:7.5pt;font-weight:800;text-align:center;text-transform:uppercase;letter-spacing:1px;font-family:'Montserrat',sans-serif}
.cmp-br{background:#F39200}.cmp-bb{background:#1E3A5F}
.cmp-tx{padding:4mm;font-size:8pt;text-align:center;color:#555;line-height:1.45;white-space:pre-line;background:#fff}

/* ORANGE BULLET LIST - Constructys colored dots */
.ol{list-style:none;padding:0}
.ol li{padding:2.5mm 0 2.5mm 6mm;position:relative;font-size:9.5pt;color:#444;line-height:1.55;text-align:justify}
.ol li::before{content:'';position:absolute;left:0;top:5.5mm;width:3mm;height:3mm;border-radius:50%;background:#F39200}

/* COLORED BULLET (Constructys uses different colors) */
.bl-tq li::before{background:#5CC8C0}
.bl-nv li::before{background:#1E3A5F}
.bl-pr li::before{background:#8B5CF6}

/* FICHE ROW */
.fr{display:flex;padding:3.5mm 0;border-bottom:1px solid #EEE}
.fr:last-child{border-bottom:none}
.fr .k{width:40mm;font-size:8.5pt;color:#F39200;font-weight:700;flex-shrink:0}
.fr .v{flex:1;font-size:9.5pt;color:#2D2D2D}

/* SOMMAIRE */
.som{padding:5mm 0;border-bottom:2px dotted #DDD}
.som:last-child{border-bottom:none}
.som .t{font-family:'Montserrat',sans-serif;font-size:12pt;font-weight:800;color:#1E3A5F;margin-bottom:1mm}
.som .p{font-size:9pt;color:#F39200;font-weight:700}

/* QUOTE */
.qt{border-left:4px solid #F39200;padding:6mm 7mm;background:#FFF8F0;border-radius:0 10px 10px 0;font-style:italic;color:#1E3A5F;font-size:10.5pt;line-height:1.6;text-align:center}

/* PLACEHOLDER */
.ph{background:#F5F5F5;border:1.5px dashed #DDD;border-radius:10px;padding:8mm 3mm;text-align:center;font-size:8pt;color:#AAA}

/* ICON ROW - icon + text side by side */
.ico-row{display:flex;align-items:center;gap:3mm;margin-bottom:3mm}
.ico-row .ico-text{flex:1}
.ico-row .ico-text h4{font-family:'Montserrat',sans-serif;font-size:9.5pt;font-weight:800;color:#1E3A5F;margin-bottom:1mm}
.ico-row .ico-text p{font-size:8.5pt;color:#555;line-height:1.5}

/* STAT CARD - icon + number + label */
.stat-card{background:#fff;border-radius:14px;box-shadow:0 2px 12px rgba(0,0,0,0.07);padding:4mm;display:flex;align-items:center;gap:3mm;overflow:hidden;position:relative}
.stat-card .sc-num{font-family:'Montserrat',sans-serif;font-size:18pt;font-weight:900;line-height:1}
.stat-card .sc-lab{font-size:8pt;color:#888;margin-top:0.5mm}

/* TIMELINE vertical */
.timeline{position:relative;padding-left:12mm}
.timeline::before{content:'';position:absolute;left:4.5mm;top:0;bottom:0;width:2px;background:#EEE}
.tl-item{position:relative;margin-bottom:4mm}
.tl-item::before{content:'';position:absolute;left:-8.5mm;top:2mm;width:9px;height:9px;border-radius:50%;border:2px solid #F39200;background:#fff;z-index:1}
.tl-item.tl-active::before{background:#F39200}
.tl-item .tl-date{font-family:'Montserrat',sans-serif;font-size:7.5pt;font-weight:700;color:#F39200}
.tl-item .tl-title{font-size:9pt;font-weight:700;color:#1E3A5F;margin-top:0.5mm}
.tl-item .tl-desc{font-size:8pt;color:#888;margin-top:0.5mm}

/* GRADIENT TEXT */
.grad-text{background:linear-gradient(135deg,#F39200,#E07800);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}

/* FANCY DIVIDER */
.fdiv{display:flex;align-items:center;gap:3mm;margin:4mm 0}
.fdiv::before,.fdiv::after{content:'';flex:1;height:1px;background:linear-gradient(to right,transparent,#DDD,transparent)}
.fdiv span{font-family:'Montserrat',sans-serif;font-size:7pt;font-weight:700;color:#F39200;letter-spacing:2px;text-transform:uppercase;white-space:nowrap}

/* UTILITY */
.mb2{margin-bottom:2mm}.mb3{margin-bottom:3mm}.mb4{margin-bottom:4mm}.mb5{margin-bottom:5mm}.mb6{margin-bottom:6mm}
.mt2{margin-top:2mm}.mt3{margin-top:3mm}.mt4{margin-top:4mm}.mt5{margin-top:5mm}
.tc{text-align:center}.tj{text-align:justify}
"""

# ━━━━━━ Helpers ━━━━━━
def ft(n):
    return f'<div class="ft"><b>BAHAFID Mohamed</b> | Rapport d\'activites U62 | BTS MEC 2026 | <b>{n}</b></div>'

def hbar(label,pct,color,value):
    return f'<div class="hb"><div class="hb-l">{H(label)}</div><div class="hb-t"><div class="hb-f" style="width:{pct:.0f}%;background:{color}"><span class="hb-v">{pct:.0f}%</span></div></div><div class="hb-r">{H(value)}</div></div>'

def sep_page(numero, titre, sous_titre, photo_name=""):
    t=titre.replace('\n','<br>')
    photo_uri = b64img(photo_name, max_w=900, max_h=600) if photo_name else ""
    if photo_uri:
        bg_style = f'background-image:url({photo_uri});background-size:cover;background-position:center'
        overlay  = '<div style="position:absolute;inset:0;background:linear-gradient(135deg,rgba(30,58,95,0.80),rgba(92,200,192,0.55))"></div>'
    else:
        bg_style = f'background:{TQ}'
        overlay  = ''
    return f'''<section class="page" style="{bg_style};overflow:hidden">
{overlay}
<div style="position:absolute;top:-30mm;right:-20mm;opacity:0.07">{svg_half_circle(300,WH,"left")}</div>
<div style="position:absolute;bottom:-20mm;left:-15mm;opacity:0.05">{svg_half_circle(200,OG,"right")}</div>
<div style="position:absolute;top:12mm;right:12mm;opacity:0.15">{svg_dots(90,90,"rgba(255,255,255,0.5)",2.5,8)}</div>
<div style="position:absolute;inset:0;opacity:0.03">{svg_stripes(800,1200,gap=8)}</div>
<div style="position:relative;z-index:10;height:100%;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;padding:30mm">
  <div class="m" style="font-size:80pt;font-weight:900;color:rgba(255,255,255,0.08);line-height:1;margin-bottom:5mm">{H(numero)}</div>
  <div class="m" style="font-size:30pt;font-weight:800;color:#fff;text-transform:uppercase;line-height:1.1;max-width:160mm;margin-bottom:4mm">{t}</div>
  <div style="width:25mm;height:3px;background:{OG};border-radius:2px;margin-bottom:4mm"></div>
  <div style="font-size:11pt;color:rgba(255,255,255,0.9);max-width:150mm;line-height:1.6">{H(sous_titre)}</div>
</div>
<div class="ft" style="color:rgba(255,255,255,0.5)"><b style="color:rgba(255,255,255,0.7)">BAHAFID Mohamed</b> | Rapport d'activites U62 | BTS MEC 2026</div>
</section>'''

def cpar_page(sit, num):
    cpar_icons={"C":ico_map_pin,"P":ico_warning,"A":ico_tools,"R":ico_check}
    bk=[("C","Contexte",sit["contexte"],"cb-c"),("P","Problematique",sit["probleme"],"cb-p"),("A","Action",sit["action"],"cb-a"),("R","Resultat",sit["resultat"],"cb-r")]
    bh="".join(f'<div class="cpar-bk {cl}"><div class="cpar-lt">{lt}</div><div class="cpar-ct"><h4>{cpar_icons[lt](10,"#fff")} {H(lb)}</h4><p>{H(tx)}</p></div></div>' for lt,lb,tx,cl in bk)
    return f'''<section class="page">
<div class="cpar-hd">
  <div style="position:absolute;top:0;right:0;width:140px;height:100%;opacity:0.05;overflow:hidden">{svg_stripes(140,200)}</div>
  <div class="cpar-n">{sit["numero"]}</div>
  <div class="cpar-i"><div class="lb">{ico_briefcase(11,TQ)} SITUATION {sit["numero"]}</div><div class="tt">{H(sit["titre"])}</div></div>
  <div class="cpar-k"><div class="v">{H(sit["chiffre_cle"])}</div><div class="l">{H(sit["chiffre_label"])}</div></div>
</div>
<div class="cpar-bd">{bh}</div>
<div class="cpar-ft">{ico_shield(12,"#1E3A5F")} {H(sit["competence"].upper())}</div>
{ft(num)}
</section>'''

# ━━━━━━ BUILD 30 PAGES ━━━━━━
def build():
    P=[]

    # ═══ P1 COVER - Constructys style with BOLD geometric shapes + icons ═══
    P.append(f'''<section class="page" style="overflow:hidden">
<!-- Right panel: turquoise with geometric shapes -->
<div style="position:absolute;top:0;right:0;width:42%;height:100%;background:{TQ};overflow:hidden">
  <div style="position:absolute;inset:0;opacity:0.06">{svg_stripes(400,1200,gap=8)}</div>
  <div style="position:absolute;top:15mm;left:10mm;opacity:0.15">{svg_circle_deco(120,"white",2)}</div>
  <div style="position:absolute;bottom:45mm;right:-15mm;opacity:0.12">{svg_dots(80,80,"rgba(255,255,255,0.5)",2.5,8)}</div>
  <!-- Photo Bahafid au centre du panel turquoise -->
  <div style="position:absolute;top:50%;left:50%;transform:translate(-50%,-55%);z-index:5">
    {img_circle("bahafid", 52)}
  </div>
  <!-- Floating icons en bas du panel turquoise -->
  <div style="position:absolute;bottom:35mm;left:15mm;opacity:0.1">{ico_tools(32,"#fff")}</div>
  <div style="position:absolute;bottom:20mm;right:15mm;opacity:0.1">{ico_cube(28,"#fff")}</div>
</div>
<!-- Large orange geometric block bleeding right -->
<div style="position:absolute;bottom:0;right:0;width:30%;height:35mm;background:{OG};border-radius:30px 0 0 0;overflow:hidden;opacity:0.85">
  <div style="position:absolute;inset:0;opacity:0.1">{svg_stripes(300,100,"rgba(255,255,255,0.3)",6,1.5)}</div>
</div>
<!-- Navy block accent -->
<div style="position:absolute;bottom:35mm;right:0;width:15%;height:20mm;background:{NV};border-radius:20px 0 0 20px;opacity:0.7"></div>
<!-- Content -->
<div style="position:relative;z-index:10;height:100%;display:flex;flex-direction:column;padding:18mm 0 18mm 18mm">

  <div style="margin-bottom:auto;display:flex;align-items:center;gap:3mm">
    {ico_badge(ico_graduation, OG, 28, 14)}
    <span class="badge badge-og" style="font-size:8pt;padding:2.5mm 6mm">SESSION 2026</span>
    <span style="font-size:8.5pt;color:#888">Academie de Lyon</span>
  </div>

  <div style="max-width:52%;margin-bottom:auto">
    <div class="label-caps" style="margin-bottom:4mm">Dossier de synthese</div>
    <div class="m" style="font-size:36pt;font-weight:900;color:{NV};line-height:1.05;text-transform:uppercase">
      RAPPORT<br>D'ACTIVITES<br><span style="color:{OG}">PROFESSIONNELLES</span>
    </div>
    <div class="bar-og" style="width:40mm;height:5px;margin:5mm 0"></div>
    <p class="body" style="max-width:120mm">
      <b>Analyse technique et economique de projets d'infrastructure au Maroc et en France.</b>
      82,5 M DH d'investissements analyses a travers 5 situations professionnelles et 7 marches publics.
    </p>
  </div>

  <div style="display:flex;gap:10mm;align-items:flex-end;max-width:52%">
    <div>
      <div class="label-caps" style="margin-bottom:2mm">Candidat</div>
      <div class="m" style="font-size:24pt;font-weight:900;color:{NV}">BAHAFID</div>
      <div class="m" style="font-size:24pt;font-weight:900;color:{OG}">Mohamed</div>
      <div style="font-size:8.5pt;color:#888;margin-top:1mm">N. {H(CANDIDAT["numero"])}</div>
    </div>
    <div style="text-align:right">
      <span class="badge badge-nv" style="font-size:7.5pt;padding:2.5mm 5mm">BTS MEC</span>
      <div style="font-size:8pt;color:#888;margin-top:2mm;line-height:1.4">Management Economique<br>de la Construction</div>
    </div>
  </div>

</div>
</section>''')

    # ═══ P2 FICHE CANDIDAT ═══
    rows="".join(f'<div class="fr"><div class="k">{H(k)}</div><div class="v">{H(v)}</div></div>' for k,v in PAGE_02["champs"])
    P.append(f'''<section class="page">
<!-- Constructys: turquoise header with geometric shapes bleeding right -->
<div style="background:{TQ};padding:15mm 18mm 12mm;position:relative;overflow:hidden;border-radius:0 0 20px 20px">
  <div style="position:absolute;top:0;right:0;width:120px;height:100%;opacity:0.1;overflow:hidden">{svg_stripes(120,200)}</div>
  <div style="position:absolute;top:8mm;left:18mm">{svg_stripes(30,30,"rgba(255,255,255,0.3)",5,2)}</div>
  <!-- Large dot grid Constructys style -->
  <div style="position:absolute;top:-10mm;right:20mm;opacity:0.15">{svg_dots(80,80,"rgba(255,255,255,0.6)",2.5,7)}</div>
  <div style="display:flex;align-items:center;gap:4mm;margin-top:8mm">
    {ico_badge(ico_users, "#fff", 36, 18)}
    <div class="m" style="font-size:26pt;font-weight:900;color:#fff;text-transform:uppercase">{H(PAGE_02["titre"])}</div>
  </div>
</div>
<div style="padding:5mm 18mm 12mm">
  <div class="bar-og"></div>
  <p class="body mb4"><b>Presentation complete du candidat,</b> de son parcours professionnel, de sa structure d'accueil et de son activite actuelle dans le domaine du BTP et de l'economie de la construction.</p>
  <div class="card">{rows}</div>
  <!-- Stat cards row -->
  <div style="display:flex;gap:3mm;margin-top:4mm;margin-bottom:4mm">
    <div class="stat-card" style="flex:1">{ico_badge(ico_clock, TQ, 30, 15)}<div><div class="sc-num" style="color:{NV}">8 ans</div><div class="sc-lab">Experience BTP</div></div></div>
    <div class="stat-card" style="flex:1">{ico_badge(ico_globe, OG, 30, 15)}<div><div class="sc-num" style="color:{OG}">2 pays</div><div class="sc-lab">Maroc + France</div></div></div>
    <div class="stat-card" style="flex:1">{ico_badge(ico_briefcase, NV, 30, 15)}<div><div class="sc-num" style="color:{NV}">7</div><div class="sc-lab">Marches publics</div></div></div>
  </div>
  <!-- Constructys style: orange capsule footer -->
  <div class="capsule capsule-og" style="justify-content:center">
    {ico_bolt(16,"#fff")}
    <div style="font-size:9pt;font-weight:600;font-style:italic;text-align:center;line-height:1.4">{H(PAGE_02["pied"])}</div>
  </div>
</div>
{ft(2)}
</section>''')

    # ═══ P3 SOMMAIRE ═══
    si="".join(f'<div class="som"><div class="t">{H(t)}</div><div style="font-size:8.5pt;color:#666;margin-bottom:1mm">{H(s)}</div><div class="p">{H(p)}</div></div>' for n,t,s,p in PAGE_03["sections"])
    P.append(f'''<section class="page" style="background:{TQ};overflow:hidden">
<!-- Constructys p2: bold geometric shapes -->
<div style="position:absolute;top:-20mm;right:10mm;opacity:0.18">{svg_dots(100,100,"rgba(255,255,255,0.6)",3,8)}</div>
<div style="position:absolute;top:30%;right:-30mm;opacity:0.1">{svg_half_circle(200,NV,"left")}</div>
<div style="position:absolute;bottom:-15mm;left:-10mm;opacity:0.08">{svg_half_circle(160,OG,"right")}</div>
<div style="position:absolute;top:15%;left:8mm;opacity:0.12">{svg_circle_deco(80,"white",2)}</div>
<div style="position:absolute;inset:0;opacity:0.04">{svg_stripes(800,1200,gap=8)}</div>
<!-- White pill accent left of title -->
<div style="position:relative;z-index:10;padding:25mm 25mm 20mm">
  <div style="display:flex;align-items:center;gap:5mm;margin-bottom:8mm">
    <div style="width:15mm;height:8mm;background:#fff;border-radius:50px"></div>
    <div class="m" style="font-size:34pt;font-weight:900;color:#fff">Sommaire</div>
  </div>
  <div class="card" style="padding:8mm 10mm">{si}</div>
</div>
<div class="ft" style="color:rgba(255,255,255,0.45)"><b style="color:rgba(255,255,255,0.65)">BAHAFID Mohamed</b> | Rapport d'activites U62 | BTS MEC 2026 | <b style="color:rgba(255,255,255,0.65)">3</b></div>
</section>''')

    # ═══ P4 INTRO - Timeline + progress rings + icon stats ═══
    # Timeline (vertical)
    tl_icons=[ico_graduation, ico_building, ico_tools, ico_cube, ico_briefcase]
    tl_colors=[TQ, NV, OG, TQ, OG]
    tl_html=""
    for i,(date,title,desc) in enumerate(PAGE_04["phases"]):
        active = "tl-active" if i==len(PAGE_04["phases"])-1 else ""
        tl_html+=f'<div class="tl-item {active}" style="--tl-c:{tl_colors[i]}"><div style="position:absolute;left:-12mm;top:0">{ico_badge(tl_icons[i], tl_colors[i], 22, 11)}</div><div class="tl-date">{H(date)}</div><div class="tl-title">{H(title)}</div><div class="tl-desc">{H(desc)}</div></div>'

    # Progress ring stats
    ring_data=[(100,"8 ans","Experience",NV),(87,"7","Marches",OG),(95,"+100M","DH investis",TQ),(100,"2","Pays",PR)]
    rings_html=""
    for pct,val,lab,clr in ring_data:
        rings_html+=f'<div style="flex:1;text-align:center">{svg_progress_ring(pct,clr,55,5)}<div class="m" style="font-size:8pt;font-weight:800;color:{NV};margin-top:1mm">{val}</div><div style="font-size:7pt;color:#888">{lab}</div></div>'

    projs="".join(f'<li>{H(p)}</li>' for p in PAGE_04["projets"])
    P.append(f'''<section class="page">
<!-- Decorative shapes -->
<div style="position:absolute;top:-15mm;right:-10mm;opacity:0.06">{svg_half_circle(150,TQ,"left")}</div>
<div style="padding:14mm 18mm 12mm">
  <div style="position:absolute;top:8mm;left:18mm">{svg_stripes(30,30,f"rgba(92,200,192,0.3)",5,2)}</div>
  <div style="display:flex;align-items:center;gap:3mm;margin-top:4mm">
    {ico_badge(ico_target, OG, 34, 17)}
    <div class="title-huge">{H(PAGE_04["titre"])}</div>
  </div>
  <div class="bar-og"></div>
  <p class="body mb3">{H(PAGE_04["intro_texte"])}</p>
  <!-- Two columns: Timeline left, Stats right -->
  <div class="col2">
    <div>
      <div class="fdiv"><span>Mon parcours</span></div>
      <div class="timeline">{tl_html}</div>
    </div>
    <div>
      <div class="fdiv"><span>Chiffres cles</span></div>
      <div style="display:flex;gap:2mm;flex-wrap:wrap;margin-bottom:3mm">{rings_html}</div>
      <div class="capsule capsule-og mb3" style="justify-content:center;padding:3mm 5mm">
        {ico_star(14,"#fff")}
        <div class="m" style="font-size:9pt;font-weight:800;text-align:center">5 situations professionnelles</div>
      </div>
      <ul class="ol">{projs}</ul>
    </div>
  </div>
  <div class="fdiv"><span>Projets analyses</span></div>
  <div class="col2">
    <div class="stat-card" style="border-left:4px solid {OG}">{ico_badge(ico_building, OG, 30, 15)}<div><div class="m" style="font-size:10pt;font-weight:800;color:{NV}">Projet 1</div><div style="font-size:8pt;color:#888">4 communes — 53,5 M DH TTC — VRD</div></div></div>
    <div class="stat-card" style="border-left:4px solid {TQ}">{ico_badge(ico_road, TQ, 30, 15)}<div><div class="m" style="font-size:10pt;font-weight:800;color:{NV}">Projet 2</div><div style="font-size:8pt;color:#888">Lehri-Kerrouchen — 29 M DH — 25 km</div></div></div>
  </div>
</div>
{ft(4)}
</section>''')

    # ═══ P5 SEP ═══
    P.append(sep_page(PAGE_05["numero"], PAGE_05["titre"], PAGE_05["sous_titre"]))

    # ═══ P6 CONSEIL REGIONAL ═══
    facts_html="".join(f'<li>{H(d)}</li>' for _,d in PAGE_06["facts"])
    org=PAGE_06["organigramme"]
    svcs=" | ".join(s.replace("\n"," ") for s in org["services"])
    mis="".join(f'<li>{H(m)}</li>' for m in PAGE_06["missions"])
    P.append(f'''<section class="page">
<div style="background:{TQ};padding:12mm 18mm 10mm;position:relative;overflow:hidden;border-radius:0 0 16px 16px">
  <div style="position:absolute;top:0;right:0;width:120px;height:100%;opacity:0.1">{svg_stripes(120,180)}</div>
  <div style="position:absolute;top:-8mm;right:25mm;opacity:0.12">{svg_dots(60,60,"rgba(255,255,255,0.5)",2.5,7)}</div>
  <div style="display:flex;align-items:center;gap:4mm">
    {ico_badge(ico_building, "#fff", 36, 18)}
    <div>
      <div class="label-caps" style="color:rgba(255,255,255,0.7);margin-bottom:1mm">STRUCTURE D'ACCUEIL</div>
      <div class="m" style="font-size:22pt;font-weight:800;color:#fff;text-transform:uppercase">{H(PAGE_06["titre"])}</div>
    </div>
  </div>
</div>
<!-- Bleed shape right -->
<div style="position:absolute;right:0;top:50%;width:8mm;height:30mm;background:{OG};border-radius:10px 0 0 10px;opacity:0.7"></div>
<div style="padding:4mm 18mm 12mm">
  <div class="bar-og"></div>
  <div class="col2">
    <div>
      <p class="body mb3"><b>Le Conseil Regional de Beni Mellal-Khenifra</b> est une collectivite territoriale marocaine couvrant 5 provinces sur 28 374 km2, soit 2,5 millions d'habitants. Ses competences incluent les routes, l'amenagement urbain et rural, l'AEP et les equipements publics.</p>
      <!-- Icon fact cards -->
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:2mm">
        <div class="stat-card" style="padding:2.5mm 3mm">{ico_badge(ico_map_pin, TQ, 24, 12)}<div><div class="m" style="font-size:9pt;font-weight:800;color:{NV}">5 Provinces</div><div style="font-size:7pt;color:#888">28 374 km2</div></div></div>
        <div class="stat-card" style="padding:2.5mm 3mm">{ico_badge(ico_users, OG, 24, 12)}<div><div class="m" style="font-size:9pt;font-weight:800;color:{NV}">2,5 M</div><div style="font-size:7pt;color:#888">Habitants</div></div></div>
        <div class="stat-card" style="padding:2.5mm 3mm">{ico_badge(ico_road, NV, 24, 12)}<div><div class="m" style="font-size:9pt;font-weight:800;color:{NV}">Routes</div><div style="font-size:7pt;color:#888">VRD, AEP</div></div></div>
        <div class="stat-card" style="padding:2.5mm 3mm">{ico_badge(ico_shield, TQ, 24, 12)}<div><div class="m" style="font-size:9pt;font-weight:800;color:{NV}">Decret</div><div style="font-size:7pt;color:#888">n.2-12-349</div></div></div>
      </div>
    </div>
    <div>
      <div style="margin-bottom:3mm;border-radius:10px;overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,0.1)">{img_tag("conseil","width:100%;height:34mm;object-fit:cover","Conseil Regional")}</div>
      <div class="card-light tc" style="padding:4mm">
        <div style="font-size:8pt;color:#888">{ico_users(12,"#aaa")} {H(org["president"])}</div>
        <div style="width:1px;height:3mm;background:#DDD;margin:1mm auto"></div>
        <div class="m" style="font-size:9pt;font-weight:700;color:{NV}">{ico_briefcase(12,NV)} {H(org["directeur"])}</div>
        <div style="width:1px;height:3mm;background:#DDD;margin:1mm auto"></div>
        <div style="font-size:7.5pt;color:#888">{H(svcs)}</div>
        <div style="width:1px;height:3mm;background:#DDD;margin:1mm auto"></div>
        <span class="badge badge-og" style="font-size:7.5pt;padding:2mm 4mm">{ico_star(10,"#fff")} {H(org["mon_poste"]).replace(chr(10)," | ")}</span>
      </div>
    </div>
  </div>
  <div class="fdiv"><span>Missions principales</span></div>
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:2mm">
    <div class="ico-row">{ico_badge(ico_chart, OG, 24, 12)}<div class="ico-text"><p>{H(PAGE_06["missions"][0])}</p></div></div>
    <div class="ico-row">{ico_badge(ico_briefcase, TQ, 24, 12)}<div class="ico-text"><p>{H(PAGE_06["missions"][1])}</p></div></div>
    <div class="ico-row">{ico_badge(ico_users, NV, 24, 12)}<div class="ico-text"><p>{H(PAGE_06["missions"][2])}</p></div></div>
    <div class="ico-row">{ico_badge(ico_chart, OG, 24, 12)}<div class="ico-text"><p>{H(PAGE_06["missions"][3])}</p></div></div>
    <div class="ico-row">{ico_badge(ico_shield, TQ, 24, 12)}<div class="ico-text"><p>{H(PAGE_06["missions"][4])}</p></div></div>
  </div>
</div>
{ft(6)}
</section>''')

    # ═══ P7 BIMCO ═══
    outils="".join(hbar(n,p,TQ if i%2==0 else OG,f"{p}%") for i,(n,p) in enumerate(PAGE_07["outils_principaux"]))
    doms="".join(f'<li><b>{H(d)}</b> : {H(s)}</li>' for d,s in PAGE_07["domaines"])
    # Radar chart for skills
    radar_labels=[n for n,_ in PAGE_07["outils_principaux"]]
    radar_vals=[p for _,p in PAGE_07["outils_principaux"]]
    radar_svg=svg_radar(radar_labels, radar_vals, TQ, 160)
    P.append(f'''<section class="page">
<!-- Geometric accents -->
<div style="position:absolute;bottom:0;right:0;width:12mm;height:45mm;background:{TQ};border-radius:20px 0 0 0;opacity:0.6"></div>
<div style="position:absolute;bottom:0;right:12mm;width:8mm;height:25mm;background:{OG};border-radius:15px 15px 0 0;opacity:0.5"></div>
<div style="padding:15mm 18mm 12mm">
  <div style="position:absolute;top:8mm;left:18mm">{svg_stripes(30,30,f"rgba(92,200,192,0.3)",5,2)}</div>
  <div style="display:flex;align-items:center;gap:4mm;margin-top:4mm">
    {ico_badge(ico_cube, OG, 34, 17)}
    <div>
      <div style="color:{OG};font-size:13pt;font-weight:800;font-family:Montserrat,sans-serif">BIMCO</div>
      <div class="title-huge" style="margin-top:1mm;font-size:20pt">Expert BIM &amp; Economie de la Construction</div>
    </div>
  </div>
  <div class="bar-og"></div>
  <p class="body mb3"><b>Cree le 9 janvier 2026,</b> BIMCO est une entreprise d'ingenierie specialisee dans le BIM et l'economie de la construction. SIREN : 999580053 | APE : 7112B | Siege : 44 rue de la Republique, 42510 Bussieres.</p>
  <div class="col2">
    <div>
      <div class="fdiv"><span>Expertise</span></div>
      <div class="ico-row">{ico_badge(ico_cube, TQ, 24, 12)}<div class="ico-text"><h4>Projeteur BIM</h4><p>Revit, Navisworks, IFC</p></div></div>
      <div class="ico-row">{ico_badge(ico_chart, OG, 24, 12)}<div class="ico-text"><h4>Metres &amp; Etudes de prix</h4><p>BIM + traditionnel, DPGF</p></div></div>
      <div class="ico-row">{ico_badge(ico_tools, NV, 24, 12)}<div class="ico-text"><h4>Outils numeriques</h4><p>Python, C#, Apps web</p></div></div>
      <div class="ico-row">{ico_badge(ico_briefcase, TQ, 24, 12)}<div class="ico-text"><h4>Suivi economique</h4><p>Budgets, decomptes</p></div></div>
      <hr class="dot-sep">
      <div class="capsule capsule-lt" style="flex-direction:column;gap:1mm;text-align:center">
        {ico_bolt(14, OG)}
        <div class="m" style="font-size:9.5pt;font-weight:700;color:{NV}">{H(PAGE_07["app"]["titre"])}</div>
        <div style="font-size:8.5pt;color:{OG};font-weight:600">{H(PAGE_07["app"]["url"])}</div>
        <div style="font-size:7.5pt;color:#888">{H(PAGE_07["app"]["stack"])}</div>
      </div>
    </div>
    <div>
      <div class="fdiv"><span>Competences</span></div>
      <!-- Radar chart -->
      <div class="tc mb3">{radar_svg}</div>
      {outils}
      <div class="capsule capsule-nv mt3" style="justify-content:center;padding:4mm 5mm">
        {ico_trophy(14,"#fff")}
        <div style="font-size:8.5pt;font-weight:700">TRIPLE COMPETENCE : MOA + Execution + BIM</div>
      </div>
    </div>
  </div>
</div>
{ft(7)}
</section>''')

    # ═══ P8 SEP P1 ═══
    P.append(sep_page(PAGE_08["numero"], PAGE_08["titre"], PAGE_08["sous_titre"], "chantier_urban"))

    # ═══ P9 FICHE P1 ═══
    fiche="".join(f'<div class="fr"><div class="k">{H(k)}</div><div class="v">{H(v)}</div></div>' for k,v in PAGE_09["fiche"])
    P.append(f'''<section class="page">
<div style="background:{TQ};padding:12mm 18mm 10mm;position:relative;overflow:hidden;border-radius:0 0 16px 16px">
  <div style="position:absolute;top:0;right:0;width:120px;height:100%;opacity:0.1">{svg_stripes(120,180)}</div>
  <div style="position:absolute;top:-8mm;right:20mm;opacity:0.12">{svg_dots(60,60,"rgba(255,255,255,0.5)",2.5,7)}</div>
  <div style="display:flex;align-items:center;gap:4mm">
    {ico_badge(ico_building, "#fff", 36, 18)}
    <div>
      <div class="label-caps" style="color:rgba(255,255,255,0.7);margin-bottom:1mm">PROJET 1</div>
      <div class="m" style="font-size:22pt;font-weight:800;color:#fff;text-transform:uppercase">{H(PAGE_09["titre"]).replace(chr(10)," ")}</div>
    </div>
  </div>
</div>
<!-- Orange bleed right -->
<div style="position:absolute;right:0;top:55%;width:10mm;height:40mm;background:{OG};border-radius:15px 0 0 15px;opacity:0.6"></div>
<div style="padding:4mm 18mm 12mm">
  <div class="bar-og"></div>
  <p class="body mb3"><b>Programme de mise a niveau des centres emergents</b> de 4 communes de la province de Khenifra, initie par le Conseil Regional de Beni Mellal-Khenifra. Ce projet d'amenagement urbain et de VRD couvre 8 corps d'etat techniques sur 4 sites simultanement.</p>
  <div class="col2 mb3">
    <div class="card" style="padding:4mm 5mm">{fiche}</div>
    <div>
      <div class="capsule capsule-og" style="flex-direction:column;text-align:center;padding:8mm 6mm;border-radius:16px;margin-bottom:3mm">
        <div style="position:absolute;top:0;right:0;opacity:0.1">{svg_stripes(80,120,"rgba(255,255,255,0.3)",5,1.5)}</div>
        <div class="m" style="font-size:32pt;font-weight:900">{H(PAGE_09["montant"])}</div>
        <div style="margin-top:2mm;font-size:9.5pt;opacity:0.85">{H(PAGE_09["detail_montant"])}</div>
      </div>
      <div style="border-radius:10px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.1)">{img_tag("chantier_urban","width:100%;height:30mm;object-fit:cover","Chantier urbain")}</div>
    </div>
  </div>
  <!-- Capsule footer -->
  <div class="capsule capsule-lt" style="justify-content:center;padding:3mm 5mm">
    <div style="font-size:8.5pt;color:#888">{H(PAGE_09["pied"])}</div>
  </div>
</div>
{ft(9)}
</section>''')

    # ═══ P10 BUDGET ═══
    communes=PAGE_10["communes"]
    vals=[float(c[1].replace(",",".").split()[0]) for c in communes]
    tot=sum(vals)
    c10=[OG,"#E07800",TQ,NV]
    seg10=[(v/tot*100,c) for v,c in zip(vals,c10)]
    d10=svg_donut(seg10,"44,6 M DH","Budget HT",200)
    b10="".join(hbar(n,float(p.replace(",",".").replace("%","")),c10[i],f"{m} ({p})") for i,(n,m,p) in enumerate(communes))
    # Bar chart for communes
    bar10=svg_bar_chart([(n.split()[0],v,c10[i]) for i,(n,v) in enumerate(zip([c[0] for c in communes],vals))], 240, 100, 40)
    cc=[OG,"#E07800","#C06800","#A05800",TQ,"#4AB0A8",NV,"#2D5A8A"]
    corps="".join(f'<div style="background:{cc[i]};border-radius:8px;padding:3mm 2mm;text-align:center;color:#fff"><div class="m" style="font-size:14pt;font-weight:900">{H(n)}</div><div style="font-size:7pt;font-weight:700;margin:0.5mm 0;line-height:1.15">{H(t)}</div><div style="font-size:5.5pt;opacity:0.8;line-height:1.15">{H(d)}</div></div>' for i,(n,t,d) in enumerate(PAGE_10["parties"]))
    P.append(f'''<section class="page">
<div style="background:{OG};padding:8mm 18mm;position:relative;overflow:hidden">
  <div style="position:absolute;top:0;right:0;width:120px;height:100%;opacity:0.08">{svg_stripes(120,160)}</div>
  <div style="position:absolute;top:3mm;right:15mm;opacity:0.12">{svg_dots(55,55,"rgba(255,255,255,0.4)",2.5,8)}</div>
  <div style="display:flex;align-items:center;gap:4mm">
    {ico_badge(ico_chart, "#fff", 32, 16)}
    <div class="m" style="font-size:20pt;font-weight:800;color:#fff;text-transform:uppercase">{H(PAGE_10["titre"])}</div>
  </div>
</div>
<div style="padding:4mm 18mm 10mm">
  <div class="fdiv"><span>Repartition par commune</span></div>
  <div style="display:flex;align-items:center;gap:5mm;margin-bottom:3mm">
    <div style="flex-shrink:0">{d10}</div>
    <div style="flex:1">{b10}</div>
  </div>
  <div class="capsule capsule-og mb3" style="justify-content:center;padding:3.5mm 5mm">{ico_star(13,"#fff")}<div class="m" style="font-size:9pt;font-weight:700">68% du budget sur 2 communes : Ouaoumana + Sebt Ait Rahou</div></div>
  <div class="m" style="font-size:10pt;font-weight:800;color:{NV};text-align:center;margin-bottom:3mm">8 CORPS D'ETAT DU MARCHE</div>
  <div style="display:flex;gap:2mm">{corps}</div>
</div>
{ft(10)}
</section>''')

    # ═══ P11-14 CPAR ═══
    P.append(cpar_page(SITUATION_1,11))
    P.append(cpar_page(SITUATION_2,12))
    P.append(cpar_page(SITUATION_3,13))
    P.append(cpar_page(SITUATION_4,14))

    # ═══ P15 DIFFICULTES - Constructys article style with badges + icons ═══
    badge_cls=["badge-og","badge-tq","badge-pr","badge-nv"]
    defi_icons=[ico_warning, ico_tools, ico_shield, ico_target]
    defs_html=""
    for i,(df,pb,sl) in enumerate(PAGE_15["defis"]):
        bc=badge_cls[i % len(badge_cls)]
        color_accent=[OG,TQ,PR,NV][i % 4]
        defs_html+=f'''<div style="flex:1;padding:3mm;background:#fff;border-radius:10px;box-shadow:0 1px 8px rgba(0,0,0,0.06);border-top:4px solid {color_accent}">
          <div style="display:flex;align-items:center;gap:2mm;margin-bottom:2mm">{ico_badge(defi_icons[i % 4], color_accent, 24, 12)}<span class="badge {bc}">{H(df)}</span></div>
          <div class="title-art" style="font-size:10pt;margin:1mm 0">{H(df)}</div>
          <p class="body" style="font-size:8.5pt;margin-bottom:2mm"><b>{H(pb.split(chr(10))[0])}</b></p>
          <p class="body" style="font-size:8.5pt;line-height:1.5">{H(chr(10).join(pb.split(chr(10))[1:]))}</p>
          <hr class="dot-sep">
          <div style="display:flex;align-items:center;gap:2mm">{ico_check(12, color_accent)}<p class="body" style="font-size:8.5pt;color:{color_accent};font-weight:600">{H(sl)}</p></div>
        </div>'''
    P.append(f'''<section class="page">
<!-- Decorative shapes -->
<div style="position:absolute;bottom:0;right:0;width:15mm;height:35mm;background:{TQ};border-radius:20px 0 0 0;opacity:0.5"></div>
<div style="padding:14mm 15mm 12mm">
  <div class="title-sec">{H(PAGE_15["titre"])}</div>
  <div class="bar-og"></div>
  <p class="body mb4"><b>Les quatre chantiers simultanement menes sur la province de Khenifra</b> ont genere des defis techniques, logistiques et financiers majeurs. Chaque difficulte a fait l'objet d'une reponse methodique.</p>
  <div style="display:flex;gap:3mm;margin-bottom:4mm">{defs_html}</div>
  <div class="capsule capsule-nv" style="justify-content:center;padding:4mm 6mm">
    <div style="font-size:9pt;line-height:1.5;text-align:center">Le suivi technique exige presence terrain, communication transparente et anticipation des risques.</div>
  </div>
</div>
{ft(15)}
</section>''')

    # ═══ P16 SEP P2 ═══
    P.append(sep_page(PAGE_16["numero"], PAGE_16["titre"], PAGE_16["sous_titre"], "montagne"))

    # ═══ P17 FICHE P2 ═══
    fp2="".join(f'<div class="fr"><div class="k">{H(k)}</div><div class="v">{H(v)}</div></div>' for k,v in PAGE_17["fiche"])
    met="".join(f'<div style="text-align:center;padding:4mm 2mm;background:#fff;border-radius:10px;border:2px solid #EEE"><div class="kpi-huge" style="font-size:15pt">{H(v)}</div><div class="kpi-label" style="font-size:7pt">{H(l)}</div></div>' for v,l in PAGE_17["metres"])
    P.append(f'''<section class="page">
<div style="background:{TQ};padding:12mm 18mm 10mm;position:relative;overflow:hidden;border-radius:0 0 16px 16px">
  <div style="position:absolute;top:0;right:0;width:120px;height:100%;opacity:0.1">{svg_stripes(120,180)}</div>
  <div style="position:absolute;top:-5mm;right:20mm;opacity:0.12">{svg_dots(60,60,"rgba(255,255,255,0.5)",2.5,7)}</div>
  <div style="display:flex;align-items:center;gap:4mm">
    {ico_badge(ico_road, "#fff", 36, 18)}
    <div>
      <div class="label-caps" style="color:rgba(255,255,255,0.7);margin-bottom:1mm">PROJET 2</div>
      <div class="m" style="font-size:22pt;font-weight:800;color:#fff;text-transform:uppercase">{H(PAGE_17["titre"])}</div>
    </div>
  </div>
</div>
<div style="position:absolute;right:0;top:48%;width:10mm;height:35mm;background:{OG};border-radius:15px 0 0 15px;opacity:0.6"></div>
<div style="padding:4mm 18mm 12mm">
  <div class="bar-og"></div>
  <div class="col2 mb3">
    <div>
      <p class="body mb3"><b>La route Lehri-Kerrouchen,</b> longue de 25 km en zone montagneuse, s'inscrit dans le Programme National des Routes Rurales (PRR3). Le marche comprend 3 sections et 53 prix au bordereau.</p>
      <div class="card" style="padding:4mm 5mm">{fp2}</div>
    </div>
    <div>
      <div class="capsule capsule-og" style="flex-direction:column;text-align:center;padding:8mm 6mm;border-radius:16px;margin-bottom:3mm">
        <div style="position:absolute;top:0;right:0;opacity:0.1">{svg_stripes(80,120,"rgba(255,255,255,0.3)",5,1.5)}</div>
        <div class="m" style="font-size:32pt;font-weight:900">{H(PAGE_17["montant"])}</div>
        <div style="margin-top:2mm;font-size:9.5pt;opacity:0.85">Programme PRR3</div>
      </div>
      <div style="border-radius:10px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.1)">{img_tag("route","width:100%;height:28mm;object-fit:cover","Route Lehri-Kerrouchen")}</div>
    </div>
  </div>
  <hr class="dot-sep">
  <div class="title-art" style="font-size:10pt;margin-bottom:3mm">Principaux metres</div>
  <div style="display:flex;gap:3mm">{met}</div>
</div>
{ft(17)}
</section>''')

    # ═══ P18 BUDGET ROUTE ═══
    items=PAGE_18["items"]
    v18=[float(it[1].replace(",",".").split()[0]) for it in items]
    t18=sum(v18)
    c18=[OG,"#E07800","#C06800",TQ,"#4AB0A8",NV]
    seg18=[(v/t18*100,c) for v,c in zip(v18,c18)]
    d18=svg_donut(seg18,"24,2 M DH","Budget HT",200)
    b18="".join(hbar(n,float(p.replace(",",".").replace("%","")),c18[i],f"{m} ({p})") for i,(n,m,p) in enumerate(items))
    # Bar chart
    bar18=svg_bar_chart([(n[:8],v,c18[i]) for i,(n,v) in enumerate(zip([it[0] for it in items],v18))], 320, 100, 32)
    P.append(f'''<section class="page">
<div style="background:{OG};padding:8mm 18mm;position:relative;overflow:hidden">
  <div style="position:absolute;top:0;right:0;width:120px;height:100%;opacity:0.08">{svg_stripes(120,160)}</div>
  <div style="position:absolute;top:3mm;right:15mm;opacity:0.12">{svg_dots(50,50,"rgba(255,255,255,0.4)",2.5,8)}</div>
  <div style="display:flex;align-items:center;gap:4mm">
    {ico_badge(ico_chart, "#fff", 32, 16)}
    <div class="m" style="font-size:20pt;font-weight:800;color:#fff;text-transform:uppercase">{H(PAGE_18["titre"])}</div>
  </div>
</div>
<div style="padding:5mm 18mm 12mm">
  <p class="body mb3"><b>Le budget total du projet routier s'eleve a 24,2 M DH HT.</b> La repartition met en evidence la predominance du corps de chaussee et des ouvrages hydrauliques, qui representent ensemble 58% de l'enveloppe globale du marche.</p>
  <div style="display:flex;align-items:center;gap:5mm;margin-bottom:3mm">
    <div style="flex-shrink:0">{d18}</div>
    <div style="flex:1">{b18}</div>
  </div>
  <!-- Bar chart visualization -->
  <div class="tc mb3">{bar18}</div>
  <div class="capsule capsule-og mb3" style="justify-content:center;padding:3.5mm 5mm">{ico_star(13,"#fff")}<div class="m" style="font-size:9pt;font-weight:700">{H(PAGE_18["callout"]).replace(chr(10)," | ")}</div></div>
  <div class="fdiv"><span>Analyse</span></div>
  <div class="col2">
    <div class="stat-card" style="border-left:4px solid {TQ}">{ico_badge(ico_road, TQ, 26, 13)}<div><div class="m" style="font-size:9pt;font-weight:800;color:{NV}">Chaussee 30,1%</div><div style="font-size:7.5pt;color:#888">GNB + GNF2 sur 25 km</div></div></div>
    <div class="stat-card" style="border-left:4px solid {OG}">{ico_badge(ico_tools, OG, 26, 13)}<div><div class="m" style="font-size:9pt;font-weight:800;color:{NV}">Hydraulique 28,1%</div><div style="font-size:7.5pt;color:#888">Buses + dalots + gabions</div></div></div>
  </div>
</div>
{ft(18)}
</section>''')

    # ═══ P19 SITUATION 5 ═══
    P.append(cpar_page(SITUATION_5,19))

    # ═══ P20 DEFIS ROUTE - Same style as P15 with badges + icons ═══
    badge_cls2=["badge-tq","badge-og","badge-pr","badge-nv"]
    defi_icons2=[ico_warning, ico_road, ico_shield, ico_target]
    defis_html=""
    for i,(t,pb,sl) in enumerate(PAGE_20["defis"]):
        bc=badge_cls2[i % len(badge_cls2)]
        color_accent=[TQ,OG,PR,NV][i % 4]
        defis_html+=f'''<div style="flex:1;padding:3mm;background:#fff;border-radius:10px;box-shadow:0 1px 8px rgba(0,0,0,0.06);border-top:4px solid {color_accent}">
          <div style="display:flex;align-items:center;gap:2mm;margin-bottom:2mm">{ico_badge(defi_icons2[i % 4], color_accent, 24, 12)}<span class="badge {bc}">{H(t.split(chr(10))[0][:20])}</span></div>
          <div class="title-art" style="font-size:10pt;margin:1mm 0">{H(t)}</div>
          <div class="bar-og" style="width:12mm;margin:1mm 0 2mm;background:{color_accent}"></div>
          <p class="body" style="font-size:8.5pt;line-height:1.45;color:#666">{H(pb[:220])}</p>
          <p class="body" style="font-size:8pt;line-height:1.4;color:#065f46;margin-top:2mm"><b>✓ </b>{H(sl[:200])}</p>
        </div>'''
    P.append(f'''<section class="page">
<div style="position:absolute;bottom:0;left:0;width:12mm;height:40mm;background:{OG};border-radius:0 20px 0 0;opacity:0.5"></div>
<div style="padding:14mm 15mm 12mm">
  <div class="title-sec">{H(PAGE_20["titre"]).replace(chr(10)," ")}</div>
  <div class="bar-og"></div>
  <p class="body mb4"><b>La construction d'une route de 25 km en zone montagneuse</b> a confronte l'equipe a des aleas geotechniques et climatiques majeurs. Voici les principaux defis rencontres et les reponses apportees sur le terrain.</p>
  <div style="display:flex;gap:3mm;margin-bottom:4mm">{defis_html}</div>
  <div class="capsule capsule-nv" style="justify-content:center;padding:4mm 6mm">
    <div style="font-size:9pt;line-height:1.5;text-align:center">Ces defis ont forge ma capacite a adapter les methodes de suivi aux aleas du terrain et a anticiper les ecarts budgetaires.</div>
  </div>
</div>
{ft(20)}
</section>''')

    # ═══ P21 ACTIVITES COMPLEMENTAIRES ═══
    marc="".join(f'<li><b>Marche {H(r)}</b> : {H(n)} <span style="color:{OG};font-weight:600">({H(t)})</span></li>' for r,n,t in PAGE_21["marches"])
    fr="".join(f'<li><b>{H(d)} - {H(p)}</b> : {H(dt)}</li>' for d,p,dt in PAGE_21["france"])
    P.append(f'''<section class="page">
<div style="position:absolute;top:0;right:0;width:10mm;height:50mm;background:{TQ};border-radius:0 0 0 20px;opacity:0.5"></div>
<div style="padding:14mm 18mm 12mm">
  <div style="display:flex;align-items:center;gap:3mm">
    {ico_badge(ico_globe, OG, 34, 17)}
    <div class="title-sec">{H(PAGE_21["titre"])}</div>
  </div>
  <div class="bar-og"></div>
  <div class="col2">
    <div>
      <div style="display:flex;align-items:center;gap:2mm;margin-bottom:2mm">{ico_badge(ico_map_pin, OG, 22, 11)}<div class="title-art" style="font-size:11pt">{H(PAGE_21["maroc_titre"])}</div></div>
      <div class="bar-og" style="width:20mm;margin:2mm 0 3mm"></div>
      <p class="body mb3"><b>Outre les deux projets principaux analyses en detail,</b> j'ai participe au suivi de 5 autres marches publics au Conseil Regional, couvrant des routes, pistes rurales, AEP et voirie.</p>
      <ul class="ol">{marc}</ul>
    </div>
    <div>
      <div style="display:flex;align-items:center;gap:2mm;margin-bottom:2mm">{ico_badge(ico_building, TQ, 22, 11)}<div class="title-art" style="font-size:11pt">{H(PAGE_21["france_titre"])}</div></div>
      <div class="bar-og" style="width:20mm;margin:2mm 0 3mm"></div>
      <p class="body mb3"><b>Mon experience en France m'a permis</b> de decouvrir le gros oeuvre cote execution, completant ainsi ma vision MOA acquise au Maroc par une connaissance terrain des contraintes de chantier.</p>
      <ul class="ol bl-tq">{fr}</ul>
      <div style="margin-top:3mm;border-radius:10px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.1)">{img_tag("france","width:100%;height:28mm;object-fit:cover","Chantier France")}</div>
    </div>
  </div>
  <!-- Capsule synthesis -->
  <div class="capsule capsule-og mt3" style="justify-content:center;padding:4mm 6mm">
    <div style="font-size:9pt;font-weight:600;text-align:center">7 marches publics suivis au total — plus de 100 M DH d'investissements</div>
  </div>
</div>
{ft(21)}
</section>''')

    # ═══ P22 SEP BILAN (Navy bg) ═══
    P.append(f'''<section class="page" style="background:{NV};overflow:hidden">
<div style="position:absolute;inset:0;opacity:0.04">{svg_stripes(800,1200,"rgba(255,255,255,0.3)",gap=8)}</div>
<div style="position:absolute;top:15mm;right:15mm;opacity:0.08">{svg_circle_deco(90,"#F39200",2)}</div>
<div style="position:absolute;bottom:20mm;left:12mm;opacity:0.06">{svg_dots(50,50,"rgba(255,255,255,0.4)",2,8)}</div>
<!-- Constructys shapes on navy -->
<div style="position:absolute;top:0;right:0;width:25%;height:15mm;background:{TQ};border-radius:0 0 0 50px;opacity:0.3"></div>
<div style="position:absolute;bottom:0;left:0;width:12mm;height:35mm;background:{OG};border-radius:0 20px 0 0;opacity:0.3"></div>
<div style="position:relative;z-index:10;height:100%;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;padding:30mm">
  <div class="m" style="font-size:80pt;font-weight:900;color:rgba(255,255,255,0.06);line-height:1;margin-bottom:5mm">{H(PAGE_22["numero"])}</div>
  <div class="m" style="font-size:30pt;font-weight:800;color:#fff;text-transform:uppercase;line-height:1.1;margin-bottom:4mm">{H(PAGE_22["titre"])}</div>
  <div style="width:25mm;height:3px;background:{OG};border-radius:2px;margin-bottom:4mm"></div>
  <div style="font-size:11pt;color:rgba(255,255,255,0.7);max-width:150mm;line-height:1.6">{H(PAGE_22["sous_titre"])}</div>
</div>
<div class="ft" style="color:rgba(255,255,255,0.35)"><b style="color:rgba(255,255,255,0.55)">BAHAFID Mohamed</b> | Rapport d'activites U62 | BTS MEC 2026</div>
</section>''')

    # ═══ P23 COMPETENCES ═══
    rows23=""
    for act,sc,sit,niv in PAGE_23["tableau"]:
        bc=NV if niv=="Expert" else TQ
        bp=100 if niv=="Expert" else 75
        rows23+=f'<tr><td style="text-align:left">{H(act)}</td><td class="tc">{H(sc)}</td><td class="tc"><span class="badge badge-og">{H(sit)}</span></td><td><div class="pbar"><div class="pbar-t"><div class="pbar-f" style="width:{bp}%;background:{bc}"></div></div><span class="pbar-l" style="color:{bc}">{H(niv)}</span></div></td></tr>'
    # Competence radar
    comp_radar=svg_radar(["Metres","Analyse offres","Suivi financier","Communication","BIM","Controle"],
                         [95,90,88,85,80,85], OG, 150)
    P.append(f'''<section class="page">
<div style="padding:14mm 18mm 12mm">
  <div style="display:flex;align-items:center;gap:3mm">
    {ico_badge(ico_graduation, OG, 34, 17)}
    <div class="title-sec">{H(PAGE_23["titre"])}</div>
  </div>
  <div class="bar-og"></div>
  <div style="display:flex;gap:5mm;margin-bottom:3mm">
    <div style="flex:1"><p class="body"><b>Ce tableau synthetise l'ensemble des activites realisees</b> durant mes 8 annees d'experience dans le BTP, en les reliant aux sous-competences du referentiel BTS MEC.</p></div>
    <div style="flex-shrink:0">{comp_radar}</div>
  </div>
  <table class="tbl"><thead><tr><th>Activite realisee</th><th>Sous-competence</th><th>Situation</th><th>Niveau</th></tr></thead><tbody>{rows23}</tbody></table>
  <hr class="dot-sep">
  <p class="body" style="font-size:8.5pt;color:#888">{H(PAGE_23["legende"])}</p>
</div>
{ft(23)}
</section>''')

    # ═══ P24 COMPARAISON ═══
    cmp=""
    for theme,maroc,france in PAGE_24["comparaison"]:
        cmp+=f'<div class="cmp"><div class="cmp-th"><div class="cmp-c">{H(theme[0])}</div><div class="cmp-tl">{H(theme)}</div></div><div class="cmp-cd"><div class="cmp-bn cmp-br">MAROC</div><div class="cmp-tx">{H(maroc)}</div></div><div class="cmp-cd"><div class="cmp-bn cmp-bb">FRANCE</div><div class="cmp-tx">{H(france)}</div></div></div>'
    P.append(f'''<section class="page">
<div style="padding:14mm 18mm 12mm">
  <div style="display:flex;align-items:center;gap:3mm;margin-bottom:2mm">{ico_badge(ico_globe, OG, 28, 14)}<div class="label-caps">ANALYSE COMPARATIVE</div></div>
  <div class="title-sec">MAROC vs FRANCE</div>
  <div class="bar-og"></div>
  <p class="body mb4"><b>Cette analyse comparative met en evidence</b> les similitudes et differences entre les systemes de marches publics marocain et francais, tels que je les ai pratiques au cours de mon parcours professionnel.</p>
  {cmp}
  <div class="capsule capsule-nv mt3" style="justify-content:center;padding:4mm 5mm"><div style="font-size:9pt;line-height:1.5;text-align:center">{H(PAGE_24["synthese"]).replace(chr(10),"<br>")}</div></div>
</div>
{ft(24)}
</section>''')

    # ═══ P25 BILAN REFLEXIF ═══
    cs=[(OG,"#FFF8F0"),(TQ,"#E8F8F6"),(NV,"#EDF1F5")]
    bilan_icons=[ico_target, ico_trophy, ico_star]
    bil=""
    for i,(t,tx) in enumerate(PAGE_25["blocs"]):
        bc,bg=cs[i]
        bil+=f'<div style="border-radius:12px;padding:5mm 6mm;margin-bottom:4mm;border-left:4px solid {bc};background:{bg}"><div style="display:flex;align-items:center;gap:3mm;margin-bottom:2mm">{ico_badge(bilan_icons[i], bc, 28, 14)}<div class="title-art" style="font-size:12pt;color:{bc}">{H(t)}</div></div><p class="body">{H(tx)}</p></div>'
    # Photo bilan en bandeau haut
    bilan_photo_uri = b64img("bilan", max_w=900, max_h=300)
    bilan_bg = f'background-image:url({bilan_photo_uri});background-size:cover;background-position:center' if bilan_photo_uri else f'background:{NV}'
    P.append(f'''<section class="page">
<div style="position:absolute;top:0;right:0;width:8mm;height:30mm;background:{OG};border-radius:0 0 0 15px;opacity:0.5"></div>
<!-- Bandeau photo haut -->
<div style="height:28mm;position:relative;overflow:hidden;{bilan_bg}">
  <div style="position:absolute;inset:0;background:linear-gradient(to right,rgba(30,58,95,0.75),rgba(30,58,95,0.3))"></div>
  <div style="position:relative;z-index:1;height:100%;display:flex;align-items:center;padding:0 18mm;gap:4mm">
    {ico_badge(ico_target, OG, 36, 18)}
    <div class="m" style="font-size:20pt;font-weight:900;color:#fff;text-transform:uppercase">{H(PAGE_25["titre"])}</div>
  </div>
</div>
<div style="padding:5mm 18mm 10mm">
  <div class="bar-og"></div>
  <p class="body mb3"><b>Ce bilan reflexif analyse mon parcours professionnel</b> en lien avec les competences du BTS MEC. Il porte un regard critique sur mes pratiques, identifie les axes d'amelioration et met en perspective les apports de mon experience pour le diplome vise.</p>
  {bil}
</div>
{ft(25)}
</section>''')

    # ═══ P26 BIM ═══
    conv=PAGE_26["convention"]
    ci="".join(f'<div style="display:flex;margin-bottom:2.5mm"><div style="width:40mm;font-weight:700;font-size:8.5pt;color:{TQ}">{H(k)}</div><div class="body" style="flex:1;font-size:8.5pt">{H(v)}</div></div>' for k,v in conv["items"])
    wfs="".join(f'<li>{H(s)}</li>' for s in PAGE_26["workflow"])
    cas=PAGE_26.get("cas_concret",{})
    P.append(f'''<section class="page">
<div style="position:absolute;bottom:0;left:0;width:10mm;height:40mm;background:{TQ};border-radius:0 20px 0 0;opacity:0.4"></div>
<div style="padding:14mm 18mm 12mm">
  <div style="display:flex;align-items:center;gap:3mm">
    {ico_badge(ico_cube, TQ, 34, 17)}
    <div class="title-sec">{H(PAGE_26["titre"])}</div>
  </div>
  <div class="bar-og"></div>
  <p class="body mb3"><b>Le BIM (Building Information Modeling)</b> permet d'extraire automatiquement les quantites depuis la maquette numerique, reduisant les erreurs de metres manuels et accelerant le processus de chiffrage pour l'economiste de la construction.</p>
  <div class="col2">
    <div>
      <div class="title-art" style="font-size:10pt;margin-bottom:2mm">{H(conv["titre"])}</div>
      <div class="bar-og" style="width:15mm;margin:2mm 0 3mm"></div>
      {ci}
      <hr class="dot-sep">
      <div class="title-art" style="font-size:10pt;margin-bottom:2mm">Workflow BIM</div>
      <div class="bar-og" style="width:15mm;margin:2mm 0 3mm"></div>
      <ul class="ol bl-tq">{wfs}</ul>
    </div>
    <div>
      <div class="card mb3" style="border-left:4px solid {OG};padding:5mm">
        <div class="title-art" style="font-size:10pt">{H(cas["titre"])}</div>
        <div class="bar-og" style="width:15mm;margin:2mm 0 2mm"></div>
        <p class="body">{H(cas["details"])}</p>
      </div>
      <div class="card-light" style="border-left:4px solid {TQ};padding:5mm">
        <div class="title-art" style="font-size:10pt;color:{TQ}">Apport du BIM pour le MEC</div>
        <div style="width:15mm;height:3px;background:{TQ};border-radius:2px;margin:2mm 0"></div>
        <p class="body">{H(PAGE_26["apport_mec"])}</p>
      </div>
    </div>
  </div>
</div>
{ft(26)}
</section>''')

    # ═══ P27 PROJET PRO ═══
    hz_c=[OG,TQ,NV]
    hz_icons=[ico_clock, ico_target, ico_star]
    hz="".join(f'<div style="flex:1;border-radius:12px;border-top:5px solid {hz_c[i]};background:#fff;box-shadow:0 2px 10px rgba(0,0,0,0.06);padding:5mm"><div style="display:flex;align-items:center;gap:2mm;margin-bottom:2mm">{ico_badge(hz_icons[i], hz_c[i], 26, 13)}<div class="m" style="font-size:11pt;font-weight:800;color:{hz_c[i]}">{H(l)}</div></div><div style="font-size:9pt;color:{OG};font-weight:600;margin-bottom:2mm">{H(d)}</div><p class="body" style="white-space:pre-line;font-size:8.5pt">{H(ds)}</p></div>' for i,(l,d,ds) in enumerate(PAGE_27["horizons"]))
    P.append(f'''<section class="page">
<div style="position:absolute;bottom:0;right:0;width:20mm;height:30mm;background:{TQ};border-radius:20px 0 0 0;opacity:0.4"></div>
<div style="position:absolute;bottom:30mm;right:0;width:10mm;height:15mm;background:{NV};border-radius:15px 0 0 15px;opacity:0.3"></div>
<div style="padding:14mm 18mm 12mm">
  <div style="display:flex;align-items:center;gap:3mm">
    {ico_badge(ico_briefcase, OG, 34, 17)}
    <div class="title-sec">{H(PAGE_27["titre"])}</div>
  </div>
  <div class="bar-og"></div>
  <p class="body mb4"><b>Mon projet professionnel s'articule autour de trois horizons temporels,</b> tous axes sur la convergence entre l'economie de la construction et les outils numeriques BIM. BIMCO a vocation a devenir un cabinet d'ingenierie de reference.</p>
  <div style="display:flex;gap:4mm;margin-bottom:5mm">{hz}</div>
  <div class="qt" style="position:relative">{svg_quote_mark(35,OG)}<span style="position:relative;z-index:1">{H(PAGE_27["citation"]).replace(chr(10)," ")}</span></div>
  <hr class="dot-sep">
  <div class="capsule capsule-nv" style="justify-content:center;flex-direction:column;text-align:center;padding:5mm 6mm">
    <div style="font-size:8pt;font-weight:700;opacity:0.7;margin-bottom:1mm">MA DOUBLE COMPETENCE</div>
    <div class="m" style="font-size:11pt;font-weight:800">Economiste terrain + Developpeur BIM</div>
    <div style="font-size:8.5pt;margin-top:1mm;opacity:0.8">Un avantage differenciant rare sur le marche de l'ingenierie</div>
  </div>
</div>
{ft(27)}
</section>''')

    # ═══ P28 CONCLUSION ═══
    pts="".join(f'<div class="mb3" style="border-left:4px solid {OG};padding:3mm 0 3mm 5mm"><div class="title-art" style="font-size:11pt;color:{NV}">{H(t)}</div><p class="body mt2">{H(d)}</p></div>' for t,d in PAGE_28["points"])
    P.append(f'''<section class="page">
<div style="background:{NV};padding:12mm 18mm 10mm;position:relative;overflow:hidden">
  <div style="position:absolute;top:0;right:0;width:120px;height:100%;opacity:0.05">{svg_stripes(120,180)}</div>
  <div style="position:absolute;top:-10mm;right:25mm;opacity:0.08">{svg_dots(60,60,"rgba(255,255,255,0.5)",2.5,7)}</div>
  <div style="display:flex;align-items:center;gap:4mm">
    {ico_badge(ico_check, "#fff", 36, 18)}
    <div class="m" style="font-size:22pt;font-weight:800;color:#fff;text-transform:uppercase">{H(PAGE_28["titre"])}</div>
  </div>
  <div style="width:25mm;height:3px;background:{OG};border-radius:2px;margin:3mm 0 0"></div>
</div>
<div style="padding:5mm 18mm 12mm">
  <p class="body mb4"><b>Ce rapport presente 5 situations professionnelles</b> issues de 2 projets majeurs totalisant 82,5 M DH d'investissements, completees par 5 marches complementaires et une experience terrain en France.</p>
  {pts}
  <div class="qt mt3" style="position:relative">{svg_quote_mark(35,OG)}<span style="position:relative;z-index:1">{H(PAGE_28["citation"]).replace(chr(10)," ")}</span></div>
  <hr class="dot-sep">
  <div class="capsule capsule-nv" style="justify-content:center;padding:3mm 5mm">
    <div style="font-size:8.5pt;text-align:center">{H(PAGE_28["pied"])}</div>
  </div>
</div>
{ft(28)}
</section>''')

    # ═══ P29 ANNEXE 1 ═══
    docs="".join(
        f'<div style="display:flex;flex-direction:column;align-items:center">'
        f'<div style="width:100%;height:70mm;border-radius:10px;overflow:hidden;'
        f'box-shadow:0 3px 12px rgba(0,0,0,0.12);border:1px solid #EEE">'
        f'{img_tag(k,"width:100%;height:70mm;object-fit:contain;background:#F8F8F8",t,max_w=500,max_h=500)}'
        f'</div>'
        f'<div class="m tc" style="margin-top:2mm;font-size:9pt;font-weight:800;color:{OG}">{H(t)}</div>'
        f'<div class="body tc" style="font-size:7.5pt;color:#888">{H(d.replace(chr(10)," · "))}</div>'
        f'</div>'
        for k,t,d in PAGE_29["documents"]
    )
    P.append(f'''<section class="page">
<div style="position:absolute;bottom:0;right:0;width:15mm;height:40mm;background:{TQ};border-radius:20px 0 0 0;opacity:0.4"></div>
<div style="padding:14mm 18mm 12mm">
  <div class="title-sec">{H(PAGE_29["titre"])}</div>
  <div class="bar-og"></div>
  <p class="body mb4"><b>Les documents ci-dessous attestent de la realite des activites professionnelles</b> decrites dans ce rapport. Ils incluent des pieces officielles liees aux procedures de marches publics marocains.</p>
  <div style="display:flex;gap:5mm">{docs}</div>
  <hr class="dot-sep">
  <div class="capsule capsule-lt" style="justify-content:center;padding:3mm 5mm">
    <div style="font-size:8.5pt;color:#888;text-align:center">Ces documents originaux seront presentes lors de l'epreuve orale pour validation par le jury.</div>
  </div>
</div>
{ft(29)}
</section>''')

    # ═══ P30 ANNEXE 2 ═══
    photos="".join(
        f'<div style="border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.1)">'
        f'{img_tag(k,"width:100%;height:40mm;object-fit:cover",c,max_w=500,max_h=400)}'
        f'<div style="background:{NV};color:#fff;font-size:7pt;font-weight:600;'
        f'text-align:center;padding:1.5mm;font-family:Montserrat,sans-serif">{H(c)}</div>'
        f'</div>'
        for k,c in PAGE_30["photos"]
    )
    P.append(f'''<section class="page">
<div style="position:absolute;bottom:0;left:0;width:12mm;height:35mm;background:{OG};border-radius:0 20px 0 0;opacity:0.4"></div>
<div style="padding:14mm 18mm 12mm">
  <div class="title-sec">{H(PAGE_30["titre"])}</div>
  <div class="bar-og"></div>
  <p class="body mb4"><b>Selection de photographies prises sur les chantiers suivis</b> dans le cadre des deux projets principaux. Ces images temoignent de la realite des travaux d'amenagement urbain et de construction routiere.</p>
  <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:4mm">{photos}</div>
</div>
{ft(30)}
</section>''')

    return P

# ━━━━━━ MAIN ━━━━━━
def main():
    print("="*60)
    print("RAPPORT U62 - PDF MAGAZINE STYLE CONSTRUCTYS V2")
    print("="*60)
    pages=build()
    full=f'''<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8">
<title>Rapport U62 - BAHAFID Mohamed - BTS MEC 2026</title>
<style>{CSS}</style></head><body>{"".join(pages)}</body></html>'''
    os.makedirs(os.path.dirname(OUT_PDF),exist_ok=True)
    with open(OUT_HTML,"w",encoding="utf-8") as f: f.write(full)
    print(f"HTML : {OUT_HTML}")
    print("Generation PDF (Chromium)...")
    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        browser=pw.chromium.launch()
        page=browser.new_page()
        page.set_content(full,wait_until="networkidle")
        page.pdf(path=OUT_PDF,format="A4",margin={"top":"0","right":"0","bottom":"0","left":"0"},print_background=True)
        browser.close()
    sz=os.path.getsize(OUT_PDF)/1024
    print(f"PDF  : {OUT_PDF}")
    print(f"Pages: {len(pages)} | Taille: {sz:.0f} Ko")
    print("="*60)

if __name__=="__main__":
    main()
