import os
import xlsxwriter
from moviepy.video.io.VideoFileClip import VideoFileClip
from datetime import timedelta



#Pasta para o programa começar a procurar os arquivos
pasta_inicial = 'PastaParaProcurar'

#Onde o arquivo excel será criado
output_file = 'LocalDoArquivoExcel'

#Criar uma pasta excel usando o modulo xlswriter
workbook = xlsxwriter.Workbook(output_file)
#Adicionar uma planinha a pasta exceel
worksheet = workbook.add_worksheet()

#Colocar o cabeçalho
worksheet.write('A1', 'Nome do Arquivo')
worksheet.write('B1', 'Nome da Pasta')
worksheet.write('C1', 'Duração')

#Linha inicial, para não colocar no cabeçalho
linha_excel = 2

#Criar um loop para passar por todos as pastas, arquivos, tendo como base a pasta inicial e usando o modulo 'os'
for root, dirs, files in os.walk(pasta_inicial):
    #Pular as pastas que terminem com '1',printando as pastas que vão ser lidas e as que foram puladas
    if os.path.basename(root).endswith('1'):
        
        print('\033[91m' + 'Pasta Pulada: ' + os.path.basename(root) + '\033[0m')

        continue
    else:
        print('\033[92m' + 'Pasta Lida: ' + os.path.basename(root) + '\033[0m')

        
    
    

    for arquivo in files:
        #Olha por todos os arquivos numa lista de arquivos e salva os que terminam com .mp4
        if arquivo.endswith('.mp4'):
            # Valores que vao ser passados
            nome_arquivo = arquivo
            nome_pasta = os.path.basename(root)
            # Caminho do arquivo para pegar duração
            caminho_arquivo = os.path.join(root, arquivo)
            
            #Pegar duração do arquivo.mp4

            video = VideoFileClip(caminho_arquivo)
            duration = video.duration
            
            #Transformar segundos em horas, pega o valor de segundos dividindo por 3600, o que sobra vira minutos, depois divide os minutos por 60, e o que sobra vira segundos
            horas, remainder = divmod(duration, 3600)
            minutos, segundos = divmod(remainder, 60)
            
            #Se não tiver mais que uma hora, sera salvo somente os minutos e os segundos
            if horas == 0:
                duration_str = f'{int(minutos):02d}m {int(segundos):02d}s'
            else:
                duration_str = f'{int(horas):02d}h {int(minutos):02d}m {int(segundos):02d}s'
            video.close()

            
            
            #Os arquivos que terminam com 2 serão pintadas de vermelho e escritas no excel, depois serão escritas os valores: nome do arquivo, duração do mp4, nome da pasta
            if nome_pasta.endswith('2'):
                
                worksheet.write(f'A{linha_excel}', nome_arquivo, workbook.add_format({'font_color': 'red'}))
                worksheet.write(f'B{linha_excel}', duration_str, workbook.add_format({'font_color': 'red'}))
                worksheet.write(f'C{linha_excel}', nome_pasta, workbook.add_format({'font_color': 'red'}))
            else:
                
                worksheet.write(f'A{linha_excel}', nome_arquivo)
                worksheet.write(f'B{linha_excel}', duration_str)
                worksheet.write(f'C{linha_excel}', nome_pasta)
            
            #Escrever na proxima linha
            linha_excel += 1

# Fechar o excel e salvar
print('Feito')
workbook.close()