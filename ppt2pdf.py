import os
import subprocess
from tqdm import tqdm

def convert_ppt_to_pdf(input_file, libreoffice_path="/Applications/LibreOffice.app/Contents/MacOS/soffice"):
    try:
        cmd = [
            libreoffice_path,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", os.path.dirname(input_file),
            input_file
        ]

        result = subprocess.run(
            cmd,
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,

        )
        return True, result.stdout.decode()
    except subprocess.CalledProcessError as e:
        return False, f"转换失败：{e.stderr.decode()}"
    except FileNotFoundError:
        return False, "LibreOffice 路径错误，无法找到 soffice 命令"

def batch_convert(folder):
    libreoffice_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    ppt_files = [f for f in os.listdir(folder) if f.lower().endswith(('.ppt', '.pptx'))]

    if not ppt_files:
        print("No PPT or PPTX files found in the folder.")
        return

    print(f"Found {len(ppt_files)} files. Starting conversion...\n")

    for filename in tqdm(ppt_files, desc="Converting PPTs"):
        full_path = os.path.join(folder, filename)
        success, error_msg = convert_ppt_to_pdf(full_path, libreoffice_path)
        if success:
            tqdm.write(f"[SUCCESS] {filename} converted.")
        else:
            tqdm.write(f"[ERROR] Failed to convert {filename}:\n{error_msg}")

    print("\nBatch conversion completed.")

if __name__ == "__main__":
    folder_path = '/path/to/your/target'  # to change
    batch_convert(folder_path)

