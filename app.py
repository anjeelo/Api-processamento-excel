from flask import Flask, request, send_file, render_template
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import os
from copy import copy

app = Flask(__name__)

# Página inicial com o formulário de upload
@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

# Rota para processar o arquivo enviado
@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "Nenhum arquivo enviado", 400
    file = request.files['file']
    if file.filename == '':
        return "Nenhum arquivo selecionado", 400
    if file and file.filename.endswith(('.xlsx', '.xls')):
        # Salvar o arquivo enviado na pasta 'uploads'
        uploads_dir = os.path.join(os.getcwd(), 'uploads')
        if not os.path.exists(uploads_dir):
            os.makedirs(uploads_dir)
        filepath = os.path.join(uploads_dir, file.filename)
        file.save(filepath)

        # Processar o arquivo
        output_filepath = process_file(filepath)

        # Enviar o arquivo processado para o usuário
        return send_file(output_filepath, as_attachment=True)
    else:
        return "Formato de arquivo inválido. Apenas arquivos .xlsx ou .xls são permitidos.", 400

# Função para processar o arquivo Excel
def process_file(filepath):
    # Carregar o arquivo Excel
    df = pd.read_excel(filepath, usecols="A:H")
    df = df.dropna(how='all')  # Remover linhas completamente vazias
    dados_atualizados = [tuple(row) for row in df.itertuples(index=False, name=None)]

    # Processar cada linha de dados
    for or_, ta, obra, localidade, causa, tratativa, endereco, exec_obra in dados_atualizados:
        output_filepath = preencher_planilha(ta, obra, localidade, tratativa, endereco, exec_obra, or_, causa)

    return output_filepath

# Função para preencher a planilha base
def preencher_planilha(ta, obra, localidade, tratativa, endereco, exec_obra, or_, causa, nome_arquivo_base='modelocroqui.xlsx'):
    # Carregar a planilha base
    wb = load_workbook(nome_arquivo_base)
    ws = wb.active  # Selecionar a primeira aba

    # Copiar as imagens da planilha base
    imagens = {}
    for img in ws._images:
        imagens[img.anchor._from] = img

    # Preencher os dados nas células especificadas
    ws['C53'] = obra          # Obra
    ws['H53'] = ta            # TA
    ws['S32'] = localidade    # Localidade
    ws['B56'] = tratativa     # Tratativa
    ws['H31'] = endereco      # Endereço
    ws['L43'] = exec_obra     # Execução de Obra
    ws['C42'] = or_           # OR
    ws['C51'] = causa         # Causa

    # Restaurar as imagens na planilha
    for anchor, img in imagens.items():
        new_img = Image(img.ref)
        new_img.anchor = anchor
        ws.add_image(new_img)

    # Salvar a planilha atualizada
    nome_arquivo_saida = f'{obra}.xlsx'
    output_filepath = os.path.join('uploads', nome_arquivo_saida)
    wb.save(output_filepath)

    return output_filepath

# Iniciar o servidor Flask
if __name__ == "__main__":
    if not os.path.exists('uploads'):
        os.makedirs('uploads')
    app.run(debug=True)