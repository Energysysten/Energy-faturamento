#!/usr/bin/env python3
import os
import sys

PROJECT_DIR = "/Users/leonardocarmo/Documents/Claude/Projects/Faturamento"
os.chdir(PROJECT_DIR)
os.environ.setdefault("HOME", os.path.expanduser("~"))

sys.argv = [
    "streamlit", "run", "dashboard.py",
    "--server.port", "8501",
    "--server.headless", "true",
    "--server.address", "localhost",
]

from streamlit.web.cli import main
main()
