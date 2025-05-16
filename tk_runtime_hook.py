"""
Runtime hook to fix Tkinter issues in bundled macOS apps
"""
import os
import sys
import platform

# Fix for Tcl/Tk resources on macOS
if platform.system() == 'Darwin':
    # Determine the correct paths for the bundled application
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        # Get the base directory of the bundled app
        base_dir = sys._MEIPASS
        
        # Set environment variables to help Tkinter find its resources
        os.environ['TCL_LIBRARY'] = os.path.join(base_dir, 'tcl')
        os.environ['TK_LIBRARY'] = os.path.join(base_dir, 'tk')
        
        # Set this to avoid "NSInternalInconsistencyException" with Tkinter menus
        os.environ['TK_SILENCE_DEPRECATION'] = '1'
        
        # This helps avoid some macOS-specific Tkinter issues
        os.environ['PYTHONUTF8'] = '1'
