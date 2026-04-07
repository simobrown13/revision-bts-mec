
import sys, os
sys.stdout.reconfigure(encoding="utf-8")
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from lxml import etree

BG=RGBColor(0xF5,0xF8,0xFA); CARD=RGBColor(0xE8,0xF5,0xF6); CARD2=RGBColor(0xF0,0xF4,0xF8)
TURQ=RGBColor(0x00,0x95,0x9E); TURQD=RGBColor(0x00,0x6E,0x78); TURQL=RGBColor(0xCC,0xEE,0xF0)
NAVY=RGBColor(0x1C,0x33,0x40); GRIS=RGBColor(0x6E,0x8A,0x96); GRISL=RGBColor(0xD4,0xE1,0xE5)
GOLD=RGBColor(0xF5,0xA1,0x18); ROUGE=RGBColor(0xE5,0x4E,0x3C); VERT=RGBColor(0x27,0xAE,0x60)
ORANGE=RGBColor(0xF3,0x9C,0x12); BLANC=RGBColor(0xFF,0xFF,0xFF); DARK=RGBColor(0x0F,0x2A,0x33)
W=Inches(13.333); H=Inches(7.5)
MEDIA="d:/PREPA BTS MEC/08_U62_Rapport_Activites/Soutenance_PowerPoint/extracted_media"
print("imports ok")
