import os
import win32com.client

input_folder = r"C:\Users\Sebastian Solar\Documents\Conmutador de Python\Word"
output_folder = r"C:\Users\Sebastian Solar\Documents\Conmutador de Python\PDF"

os.makedirs(output_folder, exist_ok=True)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

for file in os.listdir(input_folder):
    if file.endswith(".docx"):
        input_path = os.path.join(input_folder, file)
        output_path = os.path.join(output_folder, file.replace(".docx", ".pdf"))

        print(f"Convirtiendo: {file}")

        doc = word.Documents.Open(input_path)
        doc.SaveAs(output_path, FileFormat=17)
        doc.Close()

word.Quit()

print("✅ Conversión completa")