# SPDX-License-Identifier: LGPL-2.1-or-later
# SPDX-FileCopyrightText: 2025 David Carter <dcarter@davidcarter.ca>
# SPDX-FileNotice: Part of the Woods addons.

################################################################################
#                                                                              #
#   Copyright (c) 2025 David Carter <dcarter@davidcarter.ca>                   #
#                                                                              #
#   This addon is free software; you can redistribute it and/or modify it      #
#   under the terms of the GNU Lesser General Public License as published      #
#   by the Free Software Foundation; either version 2.1 of the License, or     #
#   (at your option) any later version.                                        #
#                                                                              #
#   This addon is distributed in the hope that it will be useful,              #
#   but WITHOUT ANY WARRANTY; without even the implied warranty of             #
#   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.                       #
#                                                                              #
#   See the GNU Lesser General Public License for more details.                #
#                                                                              #
#   You should have received a copy of the GNU Lesser General Public License   #
#   along with this addon; if not, write to the Free Software Foundation,      #
#   Inc., 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA           #
#                                                                              #
################################################################################

__author__ = "David Carter"
__url__ = "https://www.davesrocketshop.com"


import FreeCAD
from pathlib import PurePath

# Add materials to the user config dir
materials = FreeCAD.ParamGet("User parameter:BaseApp/Preferences/Mod/Material/Resources/Modules/Woods")
matdir = str(PurePath(FreeCAD.getUserAppDataDir(), "Mod/Woods/freecad/Woods/Resources/Materials"))
materials.SetString("ModuleDir", matdir)
moddir = str(PurePath(FreeCAD.getUserAppDataDir(), "Mod/Woods/freecad/Woods/Resources/Models"))
materials.SetString("ModuleModelDir", moddir)
materials.SetString("ModuleIcon", str(PurePath(FreeCAD.getUserAppDataDir(), "Mod/Woods/freecad/Woods/Resources/Icons/Logo.png")))
