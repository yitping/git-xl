#!/bin/bash
PYTHON_ARCH=$(python -c "import sys;print('x64' if sys.maxsize > 2**32 else 'x86')")
SCRIPT_DIR="scripts/bash"

echo $PYTHON_ARCH

python $SCRIPT_DIR/update-version-info.py
python -m PyInstaller --onefile git-xltrail-diff.py --name=git-xltrail-diff-$PYTHON_ARCH --version-file $SCRIPT_DIR/git-xltrail-version-info.py --icon $SCRIPT_DIR/git-xltrail-logo.ico
python -m PyInstaller --onefile git-xltrail.py --name=git-xltrail-$PYTHON_ARCH --version-file $SCRIPT_DIR/git-xltrail-version-info.py --icon $SCRIPT_DIR/git-xltrail-logo.ico
