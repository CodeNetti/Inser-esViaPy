from docx import Document
from docx.shared import Inches
import os
from PIL import Image
import traceback

# Defino o caminho da imagem
caminho_imagem = r"C:/Users/rsimonetti/Pictures/teste.jpeg"
caminho_temp = r"C:/Users/rsimonetti/Pictures/testetemp.jpeg"

# Verifique se o arquivo existe
if os.path.isfile(caminho_imagem):
    try:
        # Verifique se a imagem pode ser aberta com Pillow
        with Image.open(caminho_imagem) as img:
            img.verify()  # Verifica a integridade da imagem
            print("A imagem foi verificada com sucesso.")
        
        # Abror e salvor a imagem temporariamente com Pillow
        with Image.open(caminho_imagem) as img:
            img.save(caminho_temp)
            print("Imagem salva temporariamente.")
        
        documento = Document()
        
        # Insiro a imagem no documento
        print("Inserindo a imagem no documento.")
        documento.add_picture(caminho_temp, width=Inches(4))
        
        # Salve o documento
        documento.save('Carta.docx')
        print("Imagem inserida e documento salvo com sucesso.")
        
        # Remova o arquivo temporário
        os.remove(caminho_temp)

    except Exception as e:
        print(f"Erro ao processar o arquivo de imagem: {e}")
        print(traceback.format_exc())
else:
    print(f"O arquivo {caminho_imagem} não existe.")