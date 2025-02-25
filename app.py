from flask import Flask, request, send_file, render_template_string
import pandas as pd
from openpyxl import load_workbook
import os
from copy import copy
import logging

logging.basicConfig(level=logging.DEBUG)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # Limite de 16 MB para uploads

# Página inicial com o formulário de upload
@app.route('/', methods=['GET'])
def index():
    return render_template_string('''
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Upload de Arquivo</title>
    </head>
    <body>
        <form action="/upload" method="post" enctype="multipart/form-data">
            <input type="file" name="file" accept=".xlsx, .xls">
            <button type="submit">Enviar</button>
        </form>
    </body>
    </html>
    ''')

# Rota para processar o arquivo enviado
@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "Nenhum arquivo enviado", 400
    file = request.files['file']
    if file.filename == '':
        return "Nenhum arquivo selecionado", 400
    if file and file.filename.endswith(('.xlsx', '.xls')):
        try:
            uploads_dir = os.path.join(os.getcwd(), 'uploads')
            if not os.path.exists(uploads_dir):
                os.makedirs(uploads_dir)
            filepath = os.path.join(uploads_dir, file.filename)
            file.save(filepath)

            output_filepath = process_file(filepath)
            return send_file(output_filepath, as_attachment=True)
        except Exception as e:
            logging.error(f"Erro ao processar o arquivo: {e}")
            return f"Erro ao processar o arquivo: {e}", 500
    else:
        return "Formato de arquivo inválido. Apenas arquivos .xlsx ou .xls são permitidos.", 400

# Função para processar o arquivo Excel
def process_file(filepath):
    df = pd.read_excel(filepath, usecols="A:H")
    df = df.dropna(how='all')
    dados_atualizados = [tuple(row) for row in df.itertuples(index=False, name=None)]

    for or_, ta, obra, localidade, causa, tratativa, endereco, exec_obra in dados_atualizados:
        output_filepath = preencher_planilha(ta, obra, localidade, tratativa, endereco, exec_obra, or_, causa)

    return output_filepath

# Função para preencher a planilha base
def preencher_planilha(ta, obra, localidade, tratativa, endereco, exec_obra, or_, causa, nome_arquivo_base='modelocroqui.xlsx'):
    wb = load_workbook(nome_arquivo_base)
    ws = wb.active

    ws['C53'] = obra
    ws['H53'] = ta
    ws['S32'] = localidade
    ws['B56'] = tratativa
    ws['H31'] = endereco
    ws['L43'] = exec_obra
    ws['C42'] = or_
    ws['C51'] = causa

    nome_arquivo_saida = f'{obra}.xlsx'
    output_filepath = os.path.join('uploads', nome_arquivo_saida)
    wb.save(output_filepath)

    return output_filepath

# Iniciar o servidor Flask com Gunicorn
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))  # Use a porta do Render ou 10000 como fallback
    app.run(host="0.0.0.0", port=port)
