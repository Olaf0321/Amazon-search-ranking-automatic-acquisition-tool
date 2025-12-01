"""
Build script to create executable and copy necessary files.
This script handles building the exe and ensuring the .playwright-browsers
folder and all data files are in the correct locations.
"""
import os
import shutil
import subprocess
import sys
from pathlib import Path

def main():
    # Get the project root directory
    project_root = Path(__file__).parent.absolute()
    
    # Define paths
    dist_dir = project_root / "dist"
    build_dir = project_root / "build"
    playwright_browsers_src = project_root / ".playwright-browsers"
    playwright_browsers_dst = dist_dir / ".playwright-browsers"
    
    print("=" * 60)
    print("Building Amazon Ranking Tool Executable")
    print("=" * 60)
    
    # Step 1: Clean previous builds
    print("\n[1/5] Cleaning previous builds...")
    if dist_dir.exists():
        shutil.rmtree(dist_dir)
        print("   ✓ Removed dist directory")
    if build_dir.exists():
        shutil.rmtree(build_dir)
        print("   ✓ Removed build directory")
    
    # Step 2: Run PyInstaller
    print("\n[2/5] Running PyInstaller...")
    spec_file = project_root / "app.spec"
    try:
        result = subprocess.run(
            [sys.executable, "-m", "PyInstaller", str(spec_file), "--clean"],
            cwd=project_root,
            check=True,
            capture_output=False
        )
        print("   ✓ PyInstaller completed successfully")
    except subprocess.CalledProcessError as e:
        print(f"   ✗ PyInstaller failed: {e}")
        return 1
    
    # Step 3: Find the exe directory (onedir mode creates a folder)
    print("\n[3/5] Locating executable directory...")
    exe_dirs = [d for d in dist_dir.iterdir() if d.is_dir() and (d / "AmazonRankingTool.exe").exists()]
    exe_files = list(dist_dir.glob("**/AmazonRankingTool.exe"))
    
    if not exe_files:
        print("   ✗ Executable file not found!")
        return 1
    
    exe_file = exe_files[0]
    exe_dir = exe_file.parent
    
    print(f"   ✓ Found executable directory: {exe_dir.name}")
    print(f"   ✓ Found executable: {exe_file.name}")
    
    # Step 4: Copy .playwright-browsers folder if it exists
    print("\n[4/5] Copying .playwright-browsers folder...")
    playwright_browsers_dst = exe_dir / ".playwright-browsers"
    
    if playwright_browsers_src.exists() and playwright_browsers_src.is_dir():
        if playwright_browsers_dst.exists():
            shutil.rmtree(playwright_browsers_dst)
        
        # Copy the entire folder
        shutil.copytree(playwright_browsers_src, playwright_browsers_dst)
        file_count = len(list(playwright_browsers_dst.rglob('*'))) - len(list(playwright_browsers_dst.rglob('**/')))  # Count files only
        print(f"   ✓ Copied .playwright-browsers folder ({file_count} files)")
    else:
        print("   ⚠ Warning: .playwright-browsers folder not found in project root")
        print("   ⚠ Make sure to run 'playwright install chromium' before building")
        print("   ⚠ The application will fail if this folder is not present at runtime")
    
    # Step 5: Verify data files are included
    print("\n[5/5] Verifying data files...")
    required_files = [
        "asins.csv",
        "keywords.csv",
        "spreadsheetIDs.csv",
        "weighty-vertex-464012-u4-7cd9bab1166b.json"
    ]
    
    all_present = True
    for file_name in required_files:
        file_path = exe_dir / file_name
        if file_path.exists():
            print(f"   ✓ {file_name} found")
        else:
            print(f"   ✗ {file_name} MISSING!")
            all_present = False
    
    # Check for .env file (optional)
    env_file = exe_dir / ".env"
    if env_file.exists():
        print(f"   ✓ .env file found (optional)")
    else:
        print(f"   ⚠ .env file not found (optional - you may need to create it)")
        print(f"      Create it in: {exe_dir}")
    
    print("\n" + "=" * 60)
    if all_present:
        print("✓ Build completed successfully!")
        print(f"\nExecutable location: {exe_file}")
        print(f"\nTo distribute your application:")
        print(f"  1. Copy the entire '{exe_dir.name}' folder to the target location")
        print(f"  2. The folder should contain:")
        print(f"     - AmazonRankingTool.exe")
        print(f"     - .playwright-browsers/ (folder)")
        print(f"     - *.csv files")
        print(f"     - weighty-vertex-464012-u4-7cd9bab1166b.json")
        print(f"     - (Optional) .env file")
        print(f"\n  3. Users can edit the CSV files directly in this folder")
        print(f"  4. All files must remain in the same folder as the exe")
    else:
        print("⚠ Build completed with warnings - some files may be missing!")
        print("Please check the output above and ensure all required files are present.")
    print("=" * 60)
    
    return 0 if all_present else 1

if __name__ == "__main__":
    sys.exit(main())

