import os
import subprocess

folder_path = r"C:\Users\joses\Desktop\word-pdf"
output_folder = r"C:\Users\joses\Desktop\word-pdf"

os.makedirs(output_folder, exist_ok=True)

for filename in os.listdir(folder_path):
    if filename.endswith(".docx") or filename.endswith(".doc") or filename.endswith(".odt"):
        doc_path = os.path.join(folder_path, filename)
        try:
            command = [
                r"C:\Program Files\LibreOffice\program\soffice.exe", 
                "--headless",
                "--convert-to", "pdf",
                "--outdir", output_folder,
                doc_path
            ]
            subprocess.run(command, check=True)
            print(f"Convertido com sucesso: {filename}")
        except subprocess.CalledProcessError as e:
            print(f"Erro ao converter {filename}: {e}")
        except FileNotFoundError:
            print("Erro: O LibreOffice não foi encontrado. Verifique o caminho.")
