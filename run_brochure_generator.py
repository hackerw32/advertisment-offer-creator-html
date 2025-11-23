#!/usr/bin/env python3
"""
Launcher script for the Professional Brochure Generator

This script checks for required dependencies and launches the application.
"""

import sys
import subprocess

def check_dependencies():
    """Check if required dependencies are installed"""
    missing = []

    # Check tkinter
    try:
        import tkinter
    except ImportError:
        missing.append("tkinter (included with Python on most systems)")

    # Check Pillow
    try:
        from PIL import Image
    except ImportError:
        missing.append("Pillow")

    # Check reportlab
    try:
        import reportlab
    except ImportError:
        missing.append("reportlab")

    return missing


def install_dependencies():
    """Install missing pip packages"""
    print("Installing required packages...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "Pillow", "reportlab"])
    print("Installation complete!")


def main():
    print("=" * 60)
    print("Professional Brochure Generator with Booklet Imposition")
    print("=" * 60)
    print()

    missing = check_dependencies()

    if missing:
        print("Missing dependencies detected:")
        for dep in missing:
            print(f"  - {dep}")
        print()

        if "tkinter" in str(missing):
            print("tkinter is usually included with Python.")
            print("On Ubuntu/Debian: sudo apt-get install python3-tk")
            print("On Fedora: sudo dnf install python3-tkinter")
            print("On macOS: tkinter should be included with Python from python.org")
            print("On Windows: tkinter should be included with Python installer")
            print()

        # Try to install pip packages
        pip_packages = [d for d in missing if d in ["Pillow", "reportlab"]]
        if pip_packages:
            response = input("Install missing pip packages automatically? (y/n): ")
            if response.lower() == 'y':
                install_dependencies()
            else:
                print("\nPlease install manually with:")
                print("  pip install Pillow reportlab")
                return

    # Check again
    missing = check_dependencies()
    if missing:
        print("\nCannot start: missing dependencies")
        return

    print("Starting Brochure Generator...")
    print()

    # Import and run
    from brochure_generator import BrochureGeneratorApp
    app = BrochureGeneratorApp()
    app.run()


if __name__ == "__main__":
    main()
