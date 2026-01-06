# SPDX-License-Identifier: LGPL-2.1-or-later
# SPDX-FileNotice: Part of the Woods addons.

import freecad.Woods as module
from importlib import resources


materials = resources.files(module) / 'Resources/Materials'
models = resources.files(module) / 'Resources/Models'
icons = resources.files(module) / 'Resources/Icons'


def asIcon ( name : str ):

    file = name + '.svg'

    icon = icons / file

    with resources.as_file(icon) as path:
        return str( path )