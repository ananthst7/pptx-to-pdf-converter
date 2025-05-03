import os
import argparse
import subprocess

def convert_pptx_file(file_path):
    if not file_path.lower().endswith('.pptx'):
        print(f"⚠️ Skipped (not .pptx): {file_path}")
        return

    folder_path = os.path.dirname(file_path)
    print(f"Converting: {os.path.basename(file_path)}")

    try:
        subprocess.run([
            "soffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", folder_path,
            file_path
        ], check=True)
        print(f"✅ Converted: {os.path.basename(file_path)}")
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to convert {file_path}: {e}")

def convert_all_pptx_in_folder(folder_path):
    pptx_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pptx')]

    if not pptx_files:
        print("No .pptx files found in the specified folder.")
        return

    for filename in pptx_files:
        full_path = os.path.join(folder_path, filename)
        convert_pptx_file(full_path)

def main():
    parser = argparse.ArgumentParser(description="Convert a PPTX file or all PPTX files in a folder to PDF using LibreOffice.")
    parser.add_argument("path", help="Path to a .pptx file or folder containing .pptx files")
    args = parser.parse_args()

    if not os.path.exists(args.path):
        print("❌ Error: The specified path does not exist.")
        return

    if os.path.isfile(args.path):
        convert_pptx_file(os.path.abspath(args.path))
    elif os.path.isdir(args.path):
        convert_all_pptx_in_folder(os.path.abspath(args.path))
    else:
        print("❌ Error: Invalid path. Provide a .pptx file or folder containing .pptx files.")

if __name__ == "__main__":
    main()
