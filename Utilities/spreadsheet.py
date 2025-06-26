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
import uuid

FILENAME = "Y:\\Materials\\wood-database.com\\Wood Properties V1.xlsx"
OUTPUT_DIR = "Resources/Materials/Physical"

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

def parseRow(row : tuple) -> tuple:
    category = row[0]
    if category is None:
        category = ''
    name = parseCell(row[1])
    density = parseFloatCell(row[2])
    poisson = parseFloatCell(row[3])
    young = parseFloatCell(row[4])
    tensileYield = parseFloatCell(row[5])
    compressiveYield = parseFloatCell(row[7])
    result = (category, name, density, poisson, young, tensileYield, compressiveYield)
    print(result)
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

wb = load_workbook(filename=FILENAME, read_only=True)
ws = wb['onshape']
for row in ws.iter_rows(min_row=3, max_row=253, max_col=9, values_only=True):
    parsed = parseRow(row)
    createCard(parsed)
