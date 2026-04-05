import sys
import os

# Add the src directory to the sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), 'src')))

from datatransferringtool.cli import app

if __name__ == "__main__":
    app()
