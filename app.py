
import openpyxl 
from PIL import Image,ImageDraw, ImageFont

workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2, max_row=2)):
    
    nome_curso = linha[0].value 
    nome_participante = linha[1].value 
    tipo_participacao = linha[2].value 
    data_inicio = linha[3].value 
    data_final = linha[4].value 
    carga_horaria = linha[5].value 
    data_emissao = linha[6].value 
    input('')

   
    fonte_nome = ImageFont.truetype('./tahomabd.ttf', 90)
    fonte_geral = ImageFont.truetype('./tahoma.ttf', 80)
    fonte_data = ImageFont.truetype('./tahoma.ttf', 70)

    image = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(image)

    desenhar.text( (1020, 827), nome_participante, fill='black',font=fonte_nome)
    desenhar.text( (1090, 950), nome_curso, fill='black', font=fonte_geral)
    desenhar.text( (1430,1068), tipo_participacao, fill='black', font=fonte_geral)
    desenhar.text( (1490, 1187), str(carga_horaria), fill= 'black', font=fonte_geral)

    desenhar.text( (714,1755), str(data_inicio), fill='red', font=fonte_data)
    desenhar.text( (714,1920),str(data_final), fill='red', font=fonte_data)
    desenhar.text( (2180,1920),str(data_emissao), fill='red', font=fonte_data)
    
    
    image.save(f'./{indice} {nome_participante} certificado.png')
