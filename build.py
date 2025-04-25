import os
import subprocess

def build_executable():
    print("Building QLVT Tool executable...")
    
    # Make sure PyInstaller is installed
    try:
        import PyInstaller
    except ImportError:
        print("PyInstaller not found, installing...")
        subprocess.run(["pip", "install", "pyinstaller"], check=True)
    
    # Build the executable
    subprocess.run([
        "pyinstaller",
        "--name=QLVT-Tool-V2",
        "--windowed",  # GUI mode, no console window
        "--onefile",   # Create a single executable file
        "--icon=NONE", # No icon for now, you can add one later
        "--add-data=sample_data.xlsx;.",  # Include sample data
        "main.py"
    ], check=True)
    
    print("\nBuild completed!")
    print(f"Executable can be found at: {os.path.abspath('dist/QLVT-Tool-V2.exe')}")
    print("\nDon't forget to distribute these files with the executable:")
    print("- README.md (Instructions)")
    print("- sample_data.xlsx (Sample data)")

if __name__ == "__main__":
    build_executable()
