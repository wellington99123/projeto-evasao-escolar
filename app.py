from flask import Flask, render_template, request, redirect, url_for, send_file
import os
import pandas as pd
from calcular_risco import calcular_risco
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
import io
import requests

app = Flask(__name__)

# ── Supabase config ──────────────────────────────────────────────
SUPABASE_URL = "https://qwqvnhwkdsxtqufqgufi.supabase.co"
SUPABASE_KEY = "sb_secret_P-rSjvt3yDPQahiGNpeaRw_g-H4BpvA"
HEADERS = {
    "apikey": SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type": "application/json",
    "Prefer": "return=representation"
}

def sb_get(tabela, filtros=""):
    url = f"{SUPABASE_URL}/rest/v1/{tabela}?{filtros}"
    r = requests.get(url, headers=HEADERS)
    return r.json() if r.ok else []

def sb_post(tabela, dados):
    url = f"{SUPABASE_URL}/rest/v1/{tabela}"
    r = requests.post(url, headers=HEADERS, json=dados)
    return r.ok

def sb_delete(tabela, filtro):
    url = f"{SUPABASE_URL}/rest/v1/{tabela}?{filtro}"
    r = requests.delete(url, headers=HEADERS)
    return r.ok

def sb_upsert(tabela, dados):
    url = f"{SUPABASE_URL}/rest/v1/{tabela}"
    headers = {**HEADERS, "Prefer": "resolution=merge-duplicates,return=representation"}
    r = requests.post(url, headers=headers, json=dados)
    return r.ok

# ── Helpers ──────────────────────────────────────────────────────
def calcular_e_salvar_alunos(df):
    df['risco'] = df.apply(calcular_risco, axis=1)
    df['nivel'] = df['risco'].apply(lambda x: 'ALTO' if x > 55 else 'MODERADO' if x > 25 else 'BAIXO')
    registros = df.to_dict('records')
    # Salva cada aluno no Supabase (upsert por nome)
    for r in registros:
        sb_upsert('alunos', {
            'nome': str(r['nome']),
            'turma': str(r.get('turma', '')),
            'faltas_bimestre': int(r.get('faltas_bimestre', 0)),
            'media_notas': float(r.get('media_notas', 0)),
            'renda_salarios': float(r.get('renda_salarios', 0)),
            'reprovacoes': int(r.get('reprovacoes', 0)),
            'trabalha': str(r.get('trabalha', 'Não')),
            'risco': int(r['risco']),
            'nivel': r['nivel']
        })
    return registros

# ── Rotas ────────────────────────────────────────────────────────
@app.route('/')
def index():
    alunos = sb_get('alunos', 'order=risco.desc')
    ocorrencias = sb_get('ocorrencias', 'order=created_at.desc')
    return render_template('index.html', alunos=alunos, ocorrencias=ocorrencias)

@app.route('/calcular', methods=['POST'])
def calcular():
    alunos = sb_get('alunos', 'order=risco.desc')
    ocorrencias = sb_get('ocorrencias', 'order=created_at.desc')
    return render_template('index.html', alunos=alunos, ocorrencias=ocorrencias)

@app.route('/registrar', methods=['POST'])
def registrar():
    nome = request.form['nome']
    tipo = request.form['tipo']
    descricao = request.form['descricao']
    data = datetime.now(tz=__import__('zoneinfo').ZoneInfo('America/Sao_Paulo')).strftime('%d/%m/%Y %H:%M')

    # Salva ocorrência no Supabase
    sb_post('ocorrencias', {
        'nome': nome,
        'tipo': tipo,
        'descricao': descricao,
        'data': data
    })

    # Atualiza risco do aluno com base nas ocorrências
    alunos = sb_get('alunos', f'nome=eq.{requests.utils.quote(nome)}')
    ocorrs = sb_get('ocorrencias', f'nome=eq.{requests.utils.quote(nome)}')
    if alunos:
        aluno = alunos[0]
        novo_risco = min(aluno['risco'] + (len(ocorrs) * 5), 100)
        novo_nivel = 'ALTO' if novo_risco > 55 else 'MODERADO' if novo_risco > 25 else 'BAIXO'
        sb_upsert('alunos', {**aluno, 'risco': novo_risco, 'nivel': novo_nivel})

    return redirect(url_for('index'))

@app.route('/importar', methods=['POST'])
def importar():
    arquivo = request.files['planilha']
    arquivo.save('alunos.xlsx')
    df = pd.read_excel('alunos.xlsx', engine='openpyxl')
    calcular_e_salvar_alunos(df)
    return redirect(url_for('index'))

@app.route('/aluno/<nome>')
def perfil(nome):
    alunos = sb_get('alunos', f'nome=eq.{requests.utils.quote(nome)}')
    if not alunos:
        return redirect(url_for('index'))
    return render_template('perfil.html', aluno=alunos[0])

@app.route('/exportar-pdf')
def exportar_pdf():
    alunos = sb_get('alunos', 'order=risco.desc')
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    elementos = []
    elementos.append(Paragraph('Relatório de Risco de Evasão Escolar', styles['Title']))
    elementos.append(Paragraph(f'Gerado em: {datetime.now().strftime("%d/%m/%Y %H:%M")}', styles['Normal']))
    elementos.append(Spacer(1, 20))
    dados = [['Nome', 'Turma', 'Faltas', 'Média', 'Risco', 'Nível']]
    for row in alunos:
        dados.append([
            str(row['nome']), str(row['turma']),
            str(row['faltas_bimestre']), str(row['media_notas']),
            f"{row['risco']}%", str(row['nivel'])
        ])
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
    for i, row in enumerate(alunos, start=1):
        cor = colors.HexColor('#FFC7CE') if row['nivel'] == 'ALTO' else colors.HexColor('#FFEB9C') if row['nivel'] == 'MODERADO' else colors.HexColor('#C6EFCE')
        tabela.setStyle(TableStyle([('BACKGROUND', (0,i), (-1,i), cor)]))
    elementos.append(tabela)
    doc.build(elementos)
    buffer.seek(0)
    return send_file(buffer, download_name='relatorio_evasao.pdf', as_attachment=True, mimetype='application/pdf')

if __name__ == '__main__':
    app.run(debug=True)
    
