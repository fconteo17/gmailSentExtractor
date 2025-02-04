#!/usr/bin/env python
"""
Launcher script for Gmail Export Tool.
"""
import argparse
from src.main import main as cli_main
from src.gui import main as gui_main

def main():
    parser = argparse.ArgumentParser(description='Gmail Export Tool')
    parser.add_argument('--cli', action='store_true', help='Run in command-line interface mode')
    args = parser.parse_args()

    if args.cli:
        cli_main()
    else:
        gui_main()

if __name__ == "__main__":
    main() 