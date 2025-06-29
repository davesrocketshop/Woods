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

def parseCell(cell : str) -> str:
    value = cell
    if value is None:
        return ''
    if value.startswith("="):
        values = value.split(',')
        value = values[1][:-1].strip('\'"')
    value = value.replace("\n", " ")
    value = value.replace('"', '')
    return value

def parseFloatCell(cell : str) -> float:
    try:
        return float(parseCell(cell))
    except ValueError:
        pass
    return 0

def parseRow(row : tuple) -> dict:
    result = {}
    result["name"] = row[COLUMN_NAME].value
    result["steam"] = parseSteam(row[COLUMN_STEAM_BEND])
    result["hardness"] = row[COLUMN_HARDNESS].value
    result["density"] = row[COLUMN_DENSITY].value
    result["flex"] = row[COLUMN_FLEX].value
    result["poisson_long"] = row[COLUMN_POISSON_LONG].value
    result["poisson_rad"] = row[COLUMN_POISSON_RAD].value
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
    result["UUID"] = row[COLUMN_UUID].value
    result["UUID2"] = row[COLUMN_UUID2].value
    # print(result)
    return result

def createYaml(row : tuple) -> str:
    yam = "# File created by the Woods workbench\n"
    yam += "General:\n"
    # Add UUIDs
    yam += f'  UUID: "{uuid.uuid4()}"\n'
    yam += f'  Name: "{row[1]}"\n'
    yam += f'  Author: "Woods Workbench"\n'
    yam += f'  License: "GPL 3.0"\n'
    yam += f'  Description: "Automatically created by the Woods workbench"\n'

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
    yam += f'    Density: "{row[2]} kg/m^3"\n'
    yam += f'    PoissonRatio: "{row[3]}"\n'
    yam += f'    YoungsModulus: "{row[4]} Pa"\n'
    yam += f'    YieldStrength: "{row[5]} Pa"\n'
    yam += f'    CompressiveStrength: "{row[6]} Pa"\n'

    return yam

def createCard(row : tuple) -> None:
    name = row[1]
    if name is not None:
        yaml = createYaml(row)
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
    checkImage(parsed)
    # createCard(parsed)

wb.save(filename=FILENAME)
