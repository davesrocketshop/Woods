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

FILENAME = "Resources/Data/Wood Properties V2.xlsx"
IMAGES = "Resources/Data/Images"
OUTPUT_DIR = "Resources/Materials/Physical"

# Column numbers
COLUMN_NAME = 0 # A
COLUMN_STEAM_BEND = 3 # D
COLUMN_HARDNESS = 4 # E
COLUMN_DENSITY = 5 # F
COLUMN_FLEX = 6 # G
COLUMN_POISSON_LONG = 7 # I
COLUMN_POISSON_RAD = 8 # J
COLUMN_FLEX = 9 # K
COLUMN_COMPRESS = 10 # L
COLUMN_SHRINK_RAD = 11 # M
COLUMN_SHRINK_TAN = 12 # N
COLUMN_SHRINK_VOL = 13 # O
COLUMN_IMAGE = 14 # Q
COLUMN_ALT_NAMES = 15 # R
COLUMN_TAGS = 16 # S
COLUMN_REF1 = 17 # T
COLUMN_REF2 = 18 # U
COLUMN_UUID = 19 # V
COLUMN_UUID2 = 20 # W

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

def parseFloatCell(cell : str) -> float:
    try:
        return float(parseCell(cell))
    except ValueError:
        pass
    return 0

def parseRow(row : tuple) -> dict:
    result = {}
    result["name"] = str(row[COLUMN_NAME].value).strip().title()
    result["steam"] = parseSteam(row[COLUMN_STEAM_BEND])
    result["hardness"] = row[COLUMN_HARDNESS].value
    result["density"] = row[COLUMN_DENSITY].value
    result["flex"] = row[COLUMN_FLEX].value
    result["poisson_long"], long_averaged = parseCell(row[COLUMN_POISSON_LONG])
    result["poisson_rad"], rad_averaged = parseCell(row[COLUMN_POISSON_RAD])
    result["long_averaged"] = long_averaged
    result["rad_averaged"] = rad_averaged
    averaged = (long_averaged or rad_averaged)
    result["averaged"] = averaged
    result["flex"] = row[COLUMN_FLEX].value
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

def createYaml(row : dict, averaged : bool = False) -> str:
    yam = "# File created by the Woods workbench\n"
    yam += "General:\n"
    # Add UUIDs
    if averaged:
        yam += f'  UUID: "{row["UUID2"]}"\n'
        yam += f'  Name: "{row["name"]} (Averaged)"\n'
    else:
        yam += f'  UUID: "{row["UUID"]}"\n'
        yam += f'  Name: "{row["name"]}"\n'
    yam += f'  Author: "Woods Workbench"\n'
    yam += f'  License: "GPL 3.0"\n'
    tags = getTags(row)
    if len(tags) > 0:
        yam += f'  Tags:\n'
        for tag in tags:
            yam += f'    - "{tag}"\n'
    yam += f'  Description: |2\n'
    yam += '    Automatically created by the Woods workbench\n'
    if averaged:
        yam += '    This file includes averaged values in the absence of known values.\n'
        yam += '    Use with caution as the values may produce incorrect results.\n'

    yam += "AppearanceModels:\n"
    yam += "  Texture Rendering:\n"
    yam += '    UUID: "bbdcc65b-67ca-489c-bd5c-a36e33d1c160"\n'
    yam += '    AmbientColor: "(0.333333, 0.333333, 0.333333, 1)"\n'
    yam += '    DiffuseColor: "(0.859, 0.780, 0.584, 1)"\n'
    yam += '    EmissiveColor: "(0, 0, 0, 1)"\n'
    yam += '    Shininess: "0.9"\n'
    yam += '    SpecularColor: "(0.533333, 0.533333, 0.533333, 1)"\n'
    yam += '    Transparency: "0"\n'

    yam += 'Models:\n'
    yam += '  LinearElastic:\n'
    yam += '    UUID: "7b561d1d-fb9b-44f6-9da9-56a4f74d7536"\n'
    yam += f'    Density: "{row["density"]} kg/m^3"\n'
    # yam += f'    PoissonRatio: "{row[3]}"\n'
    # yam += f'    YoungsModulus: "{row[4]} Pa"\n'
    # yam += f'    YieldStrength: "{row[5]} Pa"\n'
    # yam += f'    CompressiveStrength: "{row[6]} Pa"\n'

    return yam

def createCard(row : dict) -> None:
    name = row["name"]
    if name is not None:
        if row["averaged"]:
            yaml = createYaml(row, True)
            outputName = f"{OUTPUT_DIR}/{name} (Averaged).FCMat"
            outfile = open(outputName, "w", encoding="utf-8")
            outfile.write(yaml)
            outfile.close()

        yaml = createYaml(row, False)
        outputName = f"{OUTPUT_DIR}/{name}.FCMat"
        outfile = open(outputName, "w", encoding="utf-8")
        outfile.write(yaml)
        outfile.close()

def checkImage(data : dict) -> None:
    if data["image"] is not None:
        image = f"{IMAGES}/{data['image']}"
        if not os.path.exists(image):
            print(data['image'])

# wb = load_workbook(filename=FILENAME, read_only=True)
wb = load_workbook(filename=FILENAME, read_only=False)
ws = wb['All']
# for row in ws.iter_rows(min_row=5, max_row=247, max_col=22, values_only=True):
for row in ws.iter_rows(min_row=5, max_row=247, max_col=22):
    cell = row[20]
    # print(type(cell))
    # if cell.hyperlink:
    #     print(f"Cell with hyperlink: {cell.hyperlink.target}")
    #     cell.value = str(cell.hyperlink.target)
    # else:
    #     print(cell.value)
    # print(row)
    parsed = parseRow(row)
    # checkImage(parsed)
    createCard(parsed)
# print(f"VrlBack={wb.defined_names['VrlBack'].attr_text}")
# print(f"VlrBack={wb.defined_names['VlrBack'].attr_text}")
# print(f"VrlTop={wb.defined_names['VrlTop'].attr_text}")
# print(f"VlrTop={wb.defined_names['VlrTop'].attr_text}")

wb.save(filename=FILENAME)
