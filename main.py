from docx import Document
from docx.shared import Inches
import os
from PIL import Image
import traceback

# Defino o caminho da imagem

caminho_temp = r"C:/Users/rsimonetti/Documents/ProjetoThi/Imagens/testetemp.jpg"
caminho_pasta = r"C:/Users/rsimonetti/Documents/ProjetoThi/Imagens"

documento = Document()

contador = 0
# Verifique se o arquivo existe
try:     
        for fotos in os.listdir(caminho_pasta):
            if fotos.endswith(('.jpg')):
                img = os.path.join(caminho_pasta, fotos)
                with Image.open(img) as image:
                    image.verify()  # Verifica a integridade da imagem
                   # print("A imagem foi verificada com sucesso.")
                with Image.open(img) as image:
                    image.save(caminho_temp)  # Verifica a integridade da imagem
                    
                    
                print("Inserindo a imagem no documento.")
                documento.add_picture(caminho_temp, width=Inches(4))
                contador += 1
                print(contador)
                
                    

        
                

      
        
        # Insiro a imagem no documento
        
    
        # Salve o documento
        documento.save('Carta.docx')
        print("Imagem inseridas e documento salvo com sucesso.")
        
        # Remova o arquivo temporário
        os.remove(caminho_temp)

except Exception as e:
        print(f"Erro ao processar o arquivo de imagem: {e}")
        print(traceback.format_exc())
