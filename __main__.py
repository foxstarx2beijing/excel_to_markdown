#!/usr/bin/env python3
"""
Main entry point for Excel/CSV to Markdown converter
"""

import sys

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--gui":
        # Run the GUI version
        from gui import main
        main()
    else:
        # Run the CLI version
        from excel_to_md import main
        main() 