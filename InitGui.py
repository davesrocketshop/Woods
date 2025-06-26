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
