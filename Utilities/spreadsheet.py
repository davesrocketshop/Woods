# ***************************************************************************
# *   Copyright (c) 2025 David Carter <dcarter@davidcarter.ca>              *
# *                                                                         *
# *   This program is free software; you can redistribute it and/or modify  *
# *   it under the terms of the GNU Lesser General Public License (LGPL)    *
# *   as published by the Free Software Foundation; either version 2 of     *
# *   the License, or (at your option) any later version.                   *
# *   for detail see the LICENCE text file.                                 *
# *                                                                         *
# *   This program is distributed in the hope that it will be useful,       *
# *   but WITHOUT ANY WARRANTY; without even the implied warranty of        *
# *   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the         *
# *   GNU Library General Public License for more details.                  *
# *                                                                         *
# *   You should have received a copy of the GNU Library General Public     *
# *   License along with this program; if not, write to the Free Software   *
# *   Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  *
# *   USA                                                                   *
# *                                                                         *
# ***************************************************************************
"""Class for creating material files from a spreadsheet"""

__title__ = "FreeCAD Materials Generation"
__author__ = "David Carter"
__url__ = "https://www.davesrocketshop.com"

from openpyxl import load_workbook
import os
import uuid
from typing import Any
import cv2
from base64 import b64encode
from PIL import Image
from io import BytesIO

FILENAME = "Resources/Data/Wood Properties FC.xlsx"
IMAGES = "Resources/Data/Images"
OUTPUT_DIR = "Resources/Materials"

ROW_MIN = 5
ROW_MAX = 7
COLUMN_MAX = 45

# Column numbers
COLUMN_NAME = 0 # A
COLUMN_SOFTWOOD = 2 # C
COLUMN_STEAM_BEND = 3 # D
COLUMN_HARDNESS = 4 # E
COLUMN_DENSITY = 5 # F
COLUMN_FLEX_MODULUS = 6 # G
COLUMN_POISSON_LONG = 7 # H
COLUMN_POISSON_RAD = 8 # I
COLUMN_FLEX_STRENGTH = 9 # J
COLUMN_COMPRESS = 10 # K
COLUMN_SHRINK_RAD = 11 # L
COLUMN_SHRINK_TAN = 12 # M
COLUMN_SHRINK_VOL = 13 # N
COLUMN_IMAGE = 14 # O
COLUMN_ALT_NAMES = 15 # P
COLUMN_TAGS = 16 # Q
COLUMN_REF1 = 17 # R
COLUMN_REF2 = 18 # S
COLUMN_UUID = 19 # T
COLUMN_UUID2 = 20 # U
COLUMN_RANGE = 21 # V
COLUMN_CITES = 22 # W
COLUMN_IUCN_REDLIST = 23 # X
COLUMN_IUCN_REDLIST_URL = 24 # Y
COLUMN_MACH_CHIP_THICKNESS_EXPONENT = 25 # Z
COLUMN_MACH_SURFACE_SPEED_CARBIDE = 26 # AA
COLUMN_MACH_SURFACE_SPEED_HSS = 27 # AB
COLUMN_MACH_UNIT_CUTTING_FORCE = 28 # AC
COLUMN_MACH_MAX_LOAD = 29 # AD
COLUMN_FLEX_MOD_TANG_LONG = 30 # AE
COLUMN_FLEX_MOD_RAD_LONG = 31 # AF
COLUMN_SHEAR_LONG_RAD = 32 # AG
COLUMN_SHEAR_LONG_TANG = 33 # AH
COLUMN_SHEAR_RAD_TANG = 34 # AI
COLUMN_ULTIMATE_STRENGTH_LONG = 35 # AJ
COLUMN_ULTIMATE_STRENGTH_CROSS = 36 # AK
COLUMN_COMPRESS_STRENGTH_CROSS = 37 # AL
COLUMN_SHEAR_LONG = 38 # AM
COLUMN_POISSON_LONG_TANG = 39 # AN
COLUMN_POISSON_RAD_TANG = 40 # AO
COLUMN_POISSON_TANG_RAD = 41 # AP
COLUMN_POISSON_TANG_LONG = 42 # AQ
COLUMN_THERMAL_CONDUCTIVITY = 43 # AR

# Averaged values
VrlBack = 0.056
VlrBack = 0.376
VrlTop = 0.048
VlrTop = 0.386

def parseURL(cell) -> str:
    if cell.hyperlink:
        return str(cell.hyperlink.target)
    return cell.value

def parseSteam(cell) -> str | None:
    if cell.value is None:
        return None
    steam = cell.value
    if steam == '?':
        return None
    return steam
    # return steam.strip("%")

def parseCell(cell : Any) -> tuple[Any, bool]:
    value = cell.value
    if value is None:
        return '', False
    if isinstance(value, str) and value.startswith("="):
        if value == "=VrlBack":
            return VrlBack, True
        if value == "=VlrBack":
            return VlrBack, True
        if value == "=VrlTop":
            return VrlTop, True
        if value == "=VlrTop":
            return VlrTop, True
    return value, False

# def parseFloatCell(cell : str) -> float:
#     try:
#         return float(parseCell(cell))
#     except ValueError:
#         pass
#     return 0

def parseRow(row : tuple) -> dict:
    result = {}
    result["name"] = str(row[COLUMN_NAME].value).strip().title()
    result["softwood"] = row[COLUMN_SOFTWOOD].value
    result["steam"] = parseSteam(row[COLUMN_STEAM_BEND])
    result["hardness"] = row[COLUMN_HARDNESS].value
    result["density"] = row[COLUMN_DENSITY].value
    result["flex_mod"] = row[COLUMN_FLEX_MODULUS].value
    result["poisson_long"], long_averaged = parseCell(row[COLUMN_POISSON_LONG])
    result["poisson_rad"], rad_averaged = parseCell(row[COLUMN_POISSON_RAD])
    result["long_averaged"] = long_averaged
    result["rad_averaged"] = rad_averaged
    averaged = (long_averaged or rad_averaged)
    result["averaged"] = averaged
    result["flex_strength"] = row[COLUMN_FLEX_STRENGTH].value
    result["compress"] = row[COLUMN_COMPRESS].value
    result["shrink_rad"] = row[COLUMN_SHRINK_RAD].value
    result["shrink_tan"] = row[COLUMN_SHRINK_TAN].value
    result["shrink_vol"] = row[COLUMN_SHRINK_VOL].value
    result["image"] = row[COLUMN_IMAGE].value
    result["alt"] = row[COLUMN_ALT_NAMES].value
    result["tags"] = row[COLUMN_TAGS].value
    result["ref1"] = parseURL(row[COLUMN_REF1])
    result["ref2"] = parseURL(row[COLUMN_REF2])
    if row[COLUMN_UUID].value is None:
        row[COLUMN_UUID].value = str(uuid.uuid4())
    result["UUID"] = row[COLUMN_UUID].value
    if row[COLUMN_UUID2].value is None and averaged:
        row[COLUMN_UUID2].value = str(uuid.uuid4())
    result["UUID2"] = row[COLUMN_UUID2].value
    result["range"] = row[COLUMN_RANGE].value
    result["CITES"] = row[COLUMN_CITES].value
    result["Redlist"] = row[COLUMN_IUCN_REDLIST].value
    result["RedlistURL"] = row[COLUMN_IUCN_REDLIST_URL].value
    result["MachChipThickness"] = row[COLUMN_MACH_CHIP_THICKNESS_EXPONENT].value
    result["MachSpeedCarbide"] = row[COLUMN_MACH_SURFACE_SPEED_CARBIDE].value
    result["MachSpeedHSS"] = row[COLUMN_MACH_SURFACE_SPEED_HSS].value
    result["MachUnitCuttingForce"] = row[COLUMN_MACH_UNIT_CUTTING_FORCE].value
    result["MachMaxLoad"] = row[COLUMN_MACH_MAX_LOAD].value
    result["FlexModulusTangLong"] = row[COLUMN_FLEX_MOD_TANG_LONG].value
    result["FlexModulusRadLong"] = row[COLUMN_FLEX_MOD_RAD_LONG].value
    result["ShearLongRad"] = row[COLUMN_SHEAR_LONG_RAD].value
    result["ShearLongTang"] = row[COLUMN_SHEAR_LONG_TANG].value
    result["ShearRadTang"] = row[COLUMN_SHEAR_RAD_TANG].value
    result["UltimateLong"] = row[COLUMN_ULTIMATE_STRENGTH_LONG].value
    result["UltimateCross"] = row[COLUMN_ULTIMATE_STRENGTH_CROSS].value
    result["CompressCross"] = row[COLUMN_COMPRESS_STRENGTH_CROSS].value
    result["ShearLong"] = row[COLUMN_SHEAR_LONG].value
    result["PoissonLongTang"] = row[COLUMN_POISSON_LONG_TANG].value
    result["PoissonRadTang"] = row[COLUMN_POISSON_RAD_TANG].value
    result["PoissonTangRad"] = row[COLUMN_POISSON_TANG_RAD].value
    result["PoissonTangLong"] = row[COLUMN_POISSON_TANG_LONG].value
    result["ThermalConductivity"] = row[COLUMN_THERMAL_CONDUCTIVITY].value
    return result

def getTags(row : dict) -> list:
    tags = []
    names = row['alt'].split(',')
    for name in names:
        tag = name.strip().lower()
        tags.append(tag)
    if row['tags'] is not None:
        names = row['tags'].split(',')
        for name in names:
            tag = name.strip().lower()
            tags.append(tag)
    return tags

def getRange(row : dict) -> list:
    ranges = []
    names = row['range'].split(',')
    for range in names:
        ranges.append(range.strip().upper())
    return ranges

def createInherits(row : dict) -> str:
    yam = ""
    if row["softwood"]:
        yam =  'Inherits:\n'
        yam += '  Softwood:\n'
        yam += '    UUID: "f9d83964-24ca-44df-a570-d1af36756a99"\n'
    else:
        yam =  'Inherits:\n'
        yam += '  Hardwood:\n'
        yam += '    UUID: "2a78c735-c21b-4bf9-8606-149d74f88fa8"\n'
    return yam

def createBotanical(row : dict) -> str:
    yam =   "  Wood - Botanical:\n"
    yam +=  '    UUID: "1273eaa6-8185-4130-8072-ff61132568d9"\n'
    # yam += f'    Species: "{row["species"]}"\n'
    yam += f'    SpeciesURL: "{row["ref2"]}"\n'
    yam += f'    WoodDatabase: "{row["ref1"]}"\n'
    yam += f'    Softwood: "{row["softwood"]}"\n'
    ranges = getRange(row)
    if ranges:
        yam += f'    Range:\n'
        for range in ranges:
            yam += f'      - "{range}"\n'
    if row["CITES"]:
        yam += f'    CITESAppendix: "{row["CITES"]}"\n'
    yam += f'    IUCNRedList: "{row["Redlist"]}"\n'
    yam += f'    IUCNRedListURL: "{row["RedlistURL"]}"\n'
    return yam

def createMachinability(row : dict) -> str:
    if row["MachSpeedHSS"] or row["MachSpeedCarbide"] or row["MachUnitCuttingForce"] or row["MachChipThickness"]:
        yam =   '  Machinability:\n'
        yam +=  '    UUID: "9d81fcb2-bf81-48e3-bb57-d45ecf380096"\n'
        if row["MachSpeedHSS"]:
            yam += f'    SurfaceSpeedHSS: "{row["MachSpeedHSS"]} mm/min"\n'
        if  row["MachSpeedCarbide"]:
            yam += f'    SurfaceSpeedCarbide: "{row["MachSpeedCarbide"]} mm/min"\n'
        if row["MachUnitCuttingForce"]:
            yam += f'    UnitCuttingForce: "{row["MachUnitCuttingForce"]} N/mm^2"\n'
        if row["MachChipThickness"]:
            yam += f'    ChipThicknessExponent: "{row["MachChipThickness"]}"\n'
        return yam
    else:
        return ""

def createShrinkage(row : dict) -> str:
    yam =   '  Wood - Shrinkage:\n'
    yam +=  '    UUID: "ec84f5bb-99cf-448a-86a5-cac2ebcab31c"\n'
    shrinkRadial = row["shrink_rad"]
    shrinkTangential = row["shrink_tan"]
    shrinkVolume = row["shrink_vol"]
    shrinkLong = max(1 - (1 - shrinkVolume)/((1 - shrinkRadial)*(1 - shrinkTangential)), 0)
    yam += f'    ShrinkRadial: "{shrinkRadial * 100.0}"\n'
    yam += f'    ShrinkTangential: "{shrinkTangential * 100.0}"\n'
    yam += f'    ShrinkVolume: "{shrinkVolume * 100.0}"\n'
    yam += f'    ShrinkLong: "{shrinkLong * 100.0}"\n'
    return yam

def createPhysical(row : dict) -> str:
    yam =   '  Wood:\n'
    yam +=  '    UUID: "901459aa-fd5e-43b8-aad6-71578f76c3f6"\n'
    # yam += f'    MoistureContent: "{row["density"]}"\n'
    yam += f'    SteamBendable: "{row["steam"] * 100.0}"\n'
    yam += f'    Density: "{row["density"]} kg/m^3"\n'
    yam += f'    PoissonRatioLongRad: "{row["poisson_long"]}"\n'
    yam += f'    PoissonRatioLongTan: "{row["PoissonLongTang"]}"\n'
    yam += f'    PoissonRatioRadTan: "{row["PoissonRadTang"]}"\n'
    yam += f'    PoissonRatioTanRad: "{row["PoissonTangRad"]}"\n'
    yam += f'    PoissonRatioRadLong: "{row["poisson_rad"]}"\n'
    yam += f'    PoissonRatioTanLong: "{row["PoissonTangLong"]}"\n'
    yam += f'    ShearModulusLongRad: "{row["ShearLongRad"]} MPa"\n'
    yam += f'    ShearModulusLong: "{row["ShearLong"]} kPa"\n'
    yam += f'    YoungsModulusLong: "{row["flex_mod"]} MPa"\n'
    yam += f'    YoungsModulusRadLong: "{row["FlexModulusRadLong"]}"\n'
    yam += f'    YoungsModulusRad: "{row["FlexModulusRadLong"] * row["flex_mod"]} MPa"\n'
    yam += f'    UltimateStrengthLong: "{row["UltimateLong"]} kPa"\n'
    yam += f'    UltimateStrengthCross: "{row["UltimateCross"]} kPa"\n'
    # yam += f'    CompressiveStrengthLong: "{row["density"]} kPa"\n'
    yam += f'    CompressiveStrengthCross: "{row["CompressCross"]} kPa"\n'
    yam += f'    ModulusOfRuptureLong: "{row["flex_strength"]} kPa"\n'
    return yam

def createAppearance(base : str | None, diffuse : tuple) ->str:
    yam =   "AppearanceModels:\n"
    yam +=  "  Texture Rendering:\n"
    yam +=  '    UUID: "bbdcc65b-67ca-489c-bd5c-a36e33d1c160"\n'
    yam +=  '    AmbientColor: "(0.333333, 0.333333, 0.333333, 1)"\n'
    yam += f'    DiffuseColor: "{diffuse}"\n'
    yam +=  '    EmissiveColor: "(0, 0, 0, 1)"\n'
    yam +=  '    Shininess: "0.9"\n'
    yam +=  '    SpecularColor: "(0.533333, 0.533333, 0.533333, 1)"\n'
    yam +=  '    Transparency: "0"\n'
    if base is not None:
        yam += '    TextureImage:' + base
    return yam

def createYaml(row : dict, base : str | None, diffuse : tuple, averaged : bool = False) -> str:
    yam = "# File created by the Woods workbench\n"
    yam += "General:\n"
    # Add UUIDs
    if averaged:
        yam += f'  UUID: "{row["UUID2"]}"\n'
        yam += f'  Name: "{row["name"]} (Averaged)"\n'
    else:
        yam += f'  UUID: "{row["UUID"]}"\n'
        yam += f'  Name: "{row["name"]}"\n'
    yam += f'  Author: "Gregory Holmberg"\n'
    yam += f'  License: "CDLA-Sharing-1.0"\n'
    yam += f'  SourceURL: "https://research.fs.usda.gov/treesearch/62200"\n'
    yam += f'  ReferenceSource: "USDA FPL Wood Handbook 2021"\n'
    tags = getTags(row)
    if len(tags) > 0:
        yam += f'  Tags:\n'
        for tag in tags:
            yam += f'    - "{tag}"\n'
    yam += f'  Description: >-2\n'
    yam += '    Automatically created by the Woods workbench.\n'
    if averaged:
        yam += '    \n'
        yam += '    \n'
        yam += '    This file includes averaged values in the absence of known values.\n'
        yam += '    Use with caution as the values may produce incorrect results.\n'
    yam += createInherits(row)

    yam += 'Models:\n'
    yam += createBotanical(row)
    yam += createMachinability(row)
    yam += createShrinkage(row)
    yam += createPhysical(row)
    yam += createAppearance(base, diffuse)

    return yam

def createCard(row : dict, base : str | None, diffuse : tuple) -> None:
    name = row["name"]
    if name is not None:
        if row["averaged"]:
            yaml = createYaml(row, base, diffuse, True)
            outputName = f"{OUTPUT_DIR}/{name} (Averaged).FCMat"
            outfile = open(outputName, "w", encoding="utf-8")
            outfile.write(yaml)
            outfile.close()

        yaml = createYaml(row, base, diffuse, False)
        outputName = f"{OUTPUT_DIR}/{name}.FCMat"
        outfile = open(outputName, "w", encoding="utf-8")
        outfile.write(yaml)
        outfile.close()

def imageToPng(imageData : bytes) -> bytes:
    # Create an in-memory binary stream for the input JPG data
    imageBuffer = BytesIO(imageData)

    # Open the image using Pillow
    img = Image.open(imageBuffer)

    # Create an in-memory binary stream for the output PNG data
    pngBuffer = BytesIO()

    # Save the image as PNG to the output buffer
    img.save(pngBuffer, format="PNG")

    # Get the bytes data of the PNG image
    pngData = pngBuffer.getvalue()

    return pngData

def checkImage(data : dict) -> tuple[str | None, Any]:
    base = None
    diffuse = (0.859, 0.780, 0.584, 1)
    if data["image"] is not None:
        image = f"{IMAGES}/{data['image']}"
        if os.path.exists(image):
            im = cv2.imread(image)
            A = cv2.mean(im)

            # BGR to RGB
            diffuse = (A[2] / 255.0, A[1] / 255.0, A[0] / 255.0, 1.0)

            # Convert the image to base64
            with open(image, "rb") as image_file:
                # v1.0 only works with PNG. Use this to maintain compatibility
                png = imageToPng(image_file.read())
                encoded_string = b64encode(png)
                encoded_output = encoded_string.decode('utf-8')

            base = " |-2"
            while len(encoded_output) > 0:
                base += "\n      "
                base += encoded_output[:74]
                encoded_output = encoded_output[74:]
            base += "\n"

    return base, diffuse

# Create the output folder if required
os.makedirs(OUTPUT_DIR, exist_ok=True)

# wb = load_workbook(filename=FILENAME, read_only=True)
wb = load_workbook(filename=FILENAME, read_only=False)
ws = wb['All']
# for row in ws.iter_rows(min_row=5, max_row=247, max_col=22, values_only=True):
for row in ws.iter_rows(min_row=ROW_MIN, max_row=ROW_MAX, max_col=COLUMN_MAX):
    cell = row[20]
    parsed = parseRow(row)
    base, diffuse = checkImage(parsed)
    createCard(parsed, base, diffuse)

wb.save(filename=FILENAME)
