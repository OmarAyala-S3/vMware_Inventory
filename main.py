"""
VMware Inventory System - Punto de entrada principal
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ui.app import VMwareInventoryApp

def main():
    app = VMwareInventoryApp()
    app.run()

if __name__ == "__main__":
    main()
