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

__author__ = "David Carter"
__url__ = "https://www.davesrocketshop.com"


import FreeCAD
from pathlib import PurePath

# Add materials to the user config dir
materials = FreeCAD.ParamGet("User parameter:BaseApp/Preferences/Mod/Material/Resources/Modules/Woods")
matdir = str(PurePath(FreeCAD.getUserAppDataDir(), "Mod/Woods/Resources/Materials"))
materials.SetString("ModuleDir", matdir)
moddir = str(PurePath(FreeCAD.getUserAppDataDir(), "Mod/Woods/Resources/Models"))
materials.SetString("ModuleModelDir", moddir)
materials.SetString("ModuleIcon", str(PurePath(FreeCAD.getUserAppDataDir(), "Mod/Woods/Resources/icons/woods.png")))
