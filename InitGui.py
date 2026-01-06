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

class Woods(Workbench):
    """Woods is not *really* a workbench, so this class is basically empty."""

    Icon = FreeCAD.getUserAppDataDir() + "Mod/Woods/Resources/icons/woods.png"

    def __init__(self):
        super().__init__()
        FreeCAD.Console.PrintMessage("Woods workbench loaded\n")

    def Activated(self):
        """This function is executed when the workbench is activated"""

    def Deactivated(self):
        """This function is executed when the workbench is deactivated"""

    def GetClassName(self):
        """This function is mandatory if this is a full python workbench"""
        return "Gui::PythonWorkbench"
