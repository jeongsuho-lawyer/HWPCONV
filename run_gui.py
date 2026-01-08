#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""PyInstaller entry point script"""

import sys
import os

# Add src directory to path
if getattr(sys, 'frozen', False):
    # Running as PyInstaller bundle
    base_path = sys._MEIPASS
else:
    # Running as normal Python
    base_path = os.path.dirname(os.path.abspath(__file__))
    sys.path.insert(0, os.path.join(base_path, 'src'))

from hwpconv.gui import HwpConverterApp

if __name__ == '__main__':
    app = HwpConverterApp()

    # Close splash screen
    try:
        import pyi_splash
        pyi_splash.close()
    except ImportError:
        pass

    app.mainloop()
