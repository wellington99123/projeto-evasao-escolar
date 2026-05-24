from flask import Flask, render_template, request, redirect, url_for, send_file
import json
import os
import pandas as pd
from calcular_risco import calcular_risco
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
import io

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/calcular', methods=['POST'])
def calcular():
    df = pd.read_excel('alunos.xlsx', engine='openpyxl')
    df['risco'] = df.apply(calcular_risco, axis=1)
    df['nivel'] = df['risco'].apply(lambda x: 'ALTO' if x > 55 else 'MODERADO' if x > 25 else 'BAIXO')
    alunos = df.to_dict('records')
    if os.path.exists('ocorrencias.json'):
        with open('ocorrencias.json', 'r') as f:
            ocorrencias = json.load(f)
    else:
        ocorrencias = []
    return render_template('index.html', alunos=alunos, ocorrencias=ocorrencias)

@app.route('/registrar', methods=['POST'])
def registrar():
    nome = request.form['nome']
    tipo = request.form['tipo']
    descricao = request.form['descricao']
    if os.path.exists('ocorrencias.json'):
        with open('ocorrencias.json', 'r') as f:
            ocorrencias = json.load(f)
    else:
        ocorrencias = []
    ocorrencias.append({
        'nome': nome,
        'tipo': tipo,
        'descricao': descricao,
        'data': datetime.now().strftime('%d/%m/%Y %H:%M')
    })
    with open('ocorrencias.json', 'w') as f:
        json.dump(ocorrencias, f)
    df = pd.read_excel('alunos.xlsx', engine='openpyxl')
    df['risco'] = df.apply(calcular_risco, axis=1)
    contagem = df[df['nome'] == nome].index
    if len(contagem) > 0:
        idx = contagem[0]
        ocorr_aluno = len([o for o in ocorrencias if o['nome'] == nome])
        df.at[idx, 'risco'] = min(df.at[idx, 'risco'] + (ocorr_aluno * 5), 100)
    df['nivel'] = df['risco'].apply(lambda x: 'ALTO' if x > 55 else 'MODERADO' if x > 25 else 'BAIXO')
    alunos = df.to_dict('records')
    return render_template('index.html', alunos=alunos, ocorrencias=ocorrencias)

@app.route('/importar', methods=['POST'])
def importar():
    arquivo = request.files['planilha']
    arquivo.save('alunos.xlsx')
    df = pd.read_excel('alunos.xlsx', engine='openpyxl')
    df['risco'] = df.apply(calcular_risco, axis=1)
    df['nivel'] = df['risco'].apply(lambda x: 'ALTO' if x > 55 else 'MODERADO' if x > 25 else 'BAIXO')
    alunos = df.to_dict('records')
    if os.path.exists('ocorrencias.json'):
        with open('ocorrencias.json', 'r') as f:
            ocorrencias = json.load(f)
    else:
        ocorrencias = []
    return render_template('index.html', alunos=alunos, ocorrencias=ocorrencias)

@app.route('/aluno/<nome>')
def perfil(nome):
    df = pd.read_excel('alunos.xlsx', engine='openpyxl')
    df['risco'] = df.apply(calcular_risco, axis=1)
    df['nivel'] = df['risco'].apply(lambda x: 'ALTO' if x > 55 else 'MODERADO' if x > 25 else 'BAIXO')
    aluno = df[df['nome'] == nome].to_dict('records')
    if not aluno:
        return redirect(url_for('index'))
    return render_template('perfil.html', aluno=aluno[0])

@app.route('/exportar-pdf')
def exportar_pdf():
    df = pd.read_excel('alunos.xlsx', engine='openpyxl')
    df['risco'] = df.apply(calcular_risco, axis=1)
    df['nivel'] = df['risco'].apply(lambda x: 'ALTO' if x > 55 else 'MODERADO' if x > 25 else 'BAIXO')
    df = df.sort_values('risco', ascending=False)
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    elementos = []
    elementos.append(Paragraph('Relatório de Risco de Evasão Escolar', styles['Title']))
    elementos.append(Paragraph(f'Gerado em: {datetime.now().strftime("%d/%m/%Y %H:%M")}', styles['Normal']))
    elementos.append(Spacer(1, 20))
    dados = [['Nome', 'Turma', 'Faltas', 'Média', 'Risco', 'Nível']]
    for _, row in df.iterrows():
        dados.append([str(row['nome']), str(row['turma']), str(row['faltas_bimestre']), str(row['media_notas']), f"{row['risco']}%", str(row['nivel'])])
    tabela = Table(dados, colWidths=[120, 60, 50, 60, 60, 70])
    tabela.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2E75B6')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 10),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 6),
    ]))
    for i, row in enumerate(df.itertuples(), start=1):
        cor = colors.HexColor('#FFC7CE') if row.nivel == 'ALTO' else colors.HexColor('#FFEB9C') if row.nivel == 'MODERADO' else colors.HexColor('#C6EFCE')
        tabela.setStyle(TableStyle([('BACKGROUND', (0,i), (-1,i), cor)]))
    elementos.append(tabela)
    doc.build(elementos)
    buffer.seek(0)
    return send_file(buffer, download_name='relatorio_evasao.pdf', as_attachment=True, mimetype='application/pdf')

if __name__ == '__main__':
    app.run(debug=True)