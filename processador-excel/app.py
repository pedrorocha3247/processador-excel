import os
import pandas as pd
from flask import Flask, request, render_template, send_from_directory, flash
from werkzeug.utils import secure_filename
import shutil

app = Flask(__name__)
app.secret_key = 'supersecretkey' # Necessário para usar 'flash'

# Configuração da pasta de upload
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # 1. Obter os dados do formulário
        data_input = request.form.get('data')
        tipo_input = request.form.get('tipo')
        file = request.files.get('arquivo_excel')

        # 2. Validação dos inputs
        if not data_input or not tipo_input or not file or file.filename == '':
            flash('Todos os campos são obrigatórios!')
            return render_template('index.html')

        if not file.filename.endswith(('.xlsx', '.xls')):
            flash('Por favor, envie um arquivo Excel (.xlsx, .xls).')
            return render_template('index.html')

        # 3. Processamento do arquivo
        try:
            # Lê a segunda linha (índice 1) da aba "Insira os dados"
            df = pd.read_excel(file, sheet_name='Insira os dados', header=None, skiprows=1, nrows=1)

            # Transforma a linha do DataFrame em uma lista e depois em uma string
            # Substitui valores nulos (NaN) por strings vazias
            dados = [str(item) if not pd.isna(item) else '' for item in df.iloc[0]]
            
            # Formata a string de acordo com o exemplo fornecido
            # O exemplo usa '#' como separador principal e '##' para campos vazios maiores
            conteudo_txt = "#".join(dados)
            conteudo_txt = conteudo_txt.replace('#nan#', '##') # Limpeza extra para 'nan'

            # 4. Criação da estrutura de pastas e do arquivo .txt
            data_segura = secure_filename(data_input)
            tipo_seguro = secure_filename(tipo_input)

            # Caminho da pasta principal (data) e subpasta (tipo)
            caminho_final = os.path.join('output', data_segura, tipo_seguro)
            if os.path.exists(caminho_final):
                shutil.rmtree(caminho_final) # Remove se já existir para evitar erros
            os.makedirs(caminho_final)

            # Salva o arquivo .txt
            nome_arquivo_txt = f"{tipo_seguro}.txt"
            with open(os.path.join(caminho_final, nome_arquivo_txt), 'w', encoding='utf-8') as f:
                f.write(conteudo_txt)

            flash(f"Arquivo processado com sucesso! Salvo em: output/{data_segura}/{tipo_seguro}/{nome_arquivo_txt}")

        except Exception as e:
            flash(f"Ocorreu um erro ao processar o arquivo: {e}")

        return render_template('index.html')

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)