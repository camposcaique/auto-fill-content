# Importação das bibliotecas
import openpyxl
from PIL import Image, ImageDraw, ImageFont

# Abrir a planilha
workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
    
    # Cada célula que contém a info que precisamos
    nome_curso = linha[0].value # Nome do curso
    nome_participante = linha[1].value # Nome do participante
    tipo_participacao = linha[2].value # Tipo de participação
    data_inicio = linha[3].value # Data do início
    data_final = linha[4].value # Data final
    carga_horaria = linha[5].value # Carga horária
    data_emissao = linha[6].value # Data de emissão
    
    # Definindo a fonte que iremos usar
    fonte_nome = ImageFont.truetype('./tahomabd.ttf', 90)
    fonte_geral = ImageFont.truetype('./tahoma.ttf', 80)
    fonte_data = ImageFont.truetype('./tahoma.ttf', 60)

    image = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(image)
    
    # Comandos de preenchimento das células
    desenhar.text((1050,830), nome_participante,fill='black',font = fonte_nome)
    desenhar.text((1100,945), nome_curso,fill='black',font = fonte_nome)
    desenhar.text((1430,1060), tipo_participacao,fill='black',font = fonte_nome)
    desenhar.text((1500,1185),str(carga_horaria),fill='black',font= fonte_nome)
    desenhar.text((745, 1785),data_inicio,fill='black',font= fonte_data)
    desenhar.text((745, 1930),data_final,fill='black',font= fonte_data)
    desenhar.text((2220,1930),data_emissao,fill='black',font= fonte_data)

    # Comando que salva o arquivo com o número e o nome do participante
    image.save(f'./{indice} {nome_participante} certificado.png')