# recifeplacas.github.io
# Esse script faz com que uma base em excel seja calculada os indicadores e seja convertida em pdf por empresas
 
 # todos os importe necessarios 
import os 
import locale
import pandas as pd
from docx import Document
from docx2pdf import convert
from datetime import datetime
from num2words import num2words
from datetime import datetime, timedelta
from docx.shared import RGBColor
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

nomes = ['grupo via sul piedade','grupo via euro olinda','jjk Autos'\
         ,'Rivaldo_moto', 'Pedragon rui b', 'Kia imb', 'granvia imb.'\
        , 'tambai av recife','tambai imbiribeira','honda torre novos','honda usados '\
         ,'grupo via sul repasse','honda imb novos','honda prado novos',\
         'honda piedade novos','honda caruaru novos', 'Rivaldo_carro', 'hyundai', 'hyundai_segundo']

codigos= [1000, 1014,  1001, 1002, 1003, 1004, 1005, 1006, 1007, 1008,\
          1009, 1010, 1011, 1012, 1013, 1015, 1016, 'yyyy', 'xxxx'] 

dicionario = dict(zip(nomes, codigos)) # zipando duas lista e criando dict 
for nome, codigo in dicionario.items(): # loop para interagir 
         # impontando base 
        tabela = pd.read_excel('relatorio_geral.xlsx', sheet_name = 'setembro')  
        tabela = tabela.query("MÊS == 'setembro'")       
        
        #calculando a quantidade
        # quantidade indicadores
        tabela['lojas'] = tabela['lojas'].str.strip().str.lower()
        quantidade = tabela.loc[tabela['codigos'] == codigo, 'V. TOTAL - R$'].count()

        # soma indicadores
        soma_loja_codigo = tabela.loc[tabela['codigos'] == codigo, 'V. TOTAL - R$'].sum() 
        
        # seaparando os itens
        itens_loja_especifica = tabela.loc[tabela['codigos'] == codigo, 'PLACAS'] 
        itens_string = ', '.join(itens_loja_especifica)

        data_atual = datetime.now().strftime('%d/%m/%Y') # criando a data 

        # valor por escrito 
        valor_descrito = num2words(soma_loja_codigo, lang= 'pt_BR') 

        locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')

        data_atual = datetime.now()  

        mes_anterior = data_atual - timedelta(days=20)

        nome_mes_anterior = mes_anterior.strftime('%B')  

        mensagem = f'''Segue Relaçao das placas, loja {nome.upper()} Loja Nº({codigo}),este 
relatorio no valor de R$ {soma_loja_codigo:,.2f}, ({valor_descrito} reais) , é referente
aos itens placas do Mercosul confeccionadas para os veículos abaixo relacionados, no
período mes de ({nome_mes_anterior})-2023, solicito conferencia para emissao de nota fiscal.
'''
        texto = f'''Atenciosamente: Rafael Leao da silva, Fone: 81 9 83685747 / 81 9 85436959
Rua Major marcelo menezes - casa, 32 Iputinga - Recife- PE - Fone 3223-5317
data: {data_atual}'''

        texto1 = 'Este documento é produzido de forma integral em python'
        programador = ' programadorJr : Rafael'

        doc = Document() # criando documento 
        doc.add_heading('Recife Placas', 0) #  titulo
        doc.add_heading('RELAÇAO DE PLACAS CNPJ 15.436.220.0001-30') # subtitulo 

        p = doc.add_paragraph('')# paragrafos
        p.alignment = 3

        p.add_run(mensagem).bold= True
        prior_paragraph = p.insert_paragraph_before()# paragrafos
        p.alignment = 3

        doc.add_paragraph(itens_string).alignment = 3# paragrafos

        p = doc.add_paragraph()# paragrafos
        p.add_run(texto)
        p.alignment = 3

        p = doc.add_paragraph()# paragrafos
        p.add_run(texto1 + programador)
        p.alignment = 3

        # Definir a cor vermelha (RGB: 255, 0, 0)
        cor_verde = RGBColor(0, 128, 0)

        run = p.runs[0]
        font = run.font
        font.color.rgb = cor_verde

        temp_file = f'C:/Users/TRANS MASSENA/usuario_rafael/setembro_2023/{nome}.docx' # caminho docx 

        doc.save(temp_file)# salvando caminho 

        output_file = f'C:/Users/TRANS MASSENA/usuario_rafael/setembro_2023/{nome}_{soma_loja_codigo:,.2f}.pdf'# caminho pdf

        convert(temp_file, output_file)# convertendo word em pdf 

        os.remove(temp_file) # removendo o word 

        print(f'A loja {nome}.{codigo} esta concluido') # imprimir
                
             
          
   
