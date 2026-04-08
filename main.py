import os
import pandas as pd
from prettytable import PrettyTable
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.units import cm

DEFAULT_PASTA_HTML = "paginas"
DEFAULT_PASTA_PDF = "pdfs"
LOGO_PDF = os.path.join(DEFAULT_PASTA_HTML, "static", "LogoUnivap.png")
LOGO_HTML_HEADER = "static/LogoUnivap.png"

def carregar_planilha(caminho):
    if not os.path.exists(caminho):
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")
    return pd.read_excel(caminho)

def colunas_disciplinas(df):
    if "MÉDIA AT3" in df.columns:   
        fim = df.columns.get_loc("MÉDIA AT3")
        return list(df.columns[3:fim])
    return list(df.columns[3:-1])

def mostrar_pretty_tables(df):
    disciplinas = colunas_disciplinas(df)
    for i, linha in df.iterrows():
        tabela = PrettyTable()
        tabela.title = f"ALUNO: {linha['Aluno']} (Linha {i+2} na planilha)" 
        tabela.field_names = ["Disciplina", "Nota"]
        
        for coluna in disciplinas:
            nota = linha.get(coluna, "Não informada")
            if pd.isna(nota):
                nota = "Não informada"
            tabela.add_row([coluna, nota])
            
        media = linha.get("MÉDIA AT3", "Não informada")
        tabela.add_row(["MÉDIA AT3", media])
        
        print(tabela)
        print("-" * 50)

HTML_TEMPLATE = f"""
<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="icon" href="./static/LogoUnivapHTML.png" type="image/png">
<title>Boletim de __NOME__</title>
<style>
body {{
    font-family: 'Poppins', 'Segoe UI', Arial, sans-serif;
    margin: 0;
    padding: 0;
    background: linear-gradient(to bottom, #0d47a1, #ffffff);
    color: #1a1a1a;
    min-height: 100vh;
}}
:root {{
    --cor-principal: #0d47a1;
    --cor-secundaria: #1565c0;
    --cor-sucesso: #1976d2;
    --cor-alerta: #d32f2f;
    --cor-nao-info: #757575;
    --cor-fundo-claro: #f5f5f5;
    --sombra: 0 2px 8px rgba(0, 0, 0, 0.1);
    --borda-radius: 8px;
}}
.header-ext {{
    background-color: var(--cor-principal);
    color: #ffffff;
    padding: 30px 20px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2);
    border-bottom: 4px solid var(--cor-secundaria);
    flex-wrap: wrap;
}}
.logo-univap {{
    width: 150px;
    height: auto;
    background: none;
}}
.titulo-principal {{
    font-size: 32px;
    font-weight: 700;
    margin: 0;
}}
.card {{
    background-color: #ffffff;
    padding: 30px;
    border-radius: var(--borda-radius);
    box-shadow: var(--sombra);
    max-width: 900px;
    margin: 30px auto;
    border: 2px solid var(--cor-principal);
}}
h2 {{
    color: var(--cor-principal);
    text-align: center;
    margin-bottom: 30px;
    font-size: 24px;
    font-weight: 600;
}}
.meta {{
    margin-bottom: 30px;
    text-align: center;
    background-color: var(--cor-fundo-claro);
    border: 1px solid var(--cor-principal);
    padding: 20px;
    border-radius: var(--borda-radius);
}}
.meta p {{
    margin: 8px 0;
    font-size: 16px;
    font-weight: 500;
}}
table {{
    width: 100%;
    border-collapse: collapse;
    box-shadow: var(--sombra);
    border-radius: var(--borda-radius);
    overflow: hidden;
    margin-bottom: 30px;
}}
th, td {{
    padding: 15px 10px;
    text-align: center;
    font-size: 14px;
    border: 1px solid #e0e0e0;
    border-bottom: 2px solid var(--cor-principal);
}}
th {{
    background-color: var(--cor-principal);
    color: #ffffff;
    font-weight: 600;
    text-transform: uppercase;
}}
tr:nth-child(even) {{
    background-color: #fafafa;
}}
tr:hover {{
    background-color: #f0f8ff;
}}
.nota-boa {{
    color: var(--cor-sucesso);
    font-weight: bold;
}}
.nota-baixa {{
    color: var(--cor-alerta);
    font-weight: bold;
}}
.nota-nao {{
    color: var(--cor-nao-info);
    font-style: italic;
}}
.btn-container {{
    text-align: center;
    margin-top: 30px;
}}
.btn {{
    background-color: var(--cor-principal);
    color: #ffffff;
    padding: 12px 30px;
    border: none;
    border-radius: 25px;
    text-decoration: none;
    font-weight: 600;
    font-size: 16px;
    transition: all 0.3s ease;
    display: inline-block;
}}
.btn:hover {{
    background-color: var(--cor-secundaria);
    transform: scale(1.05);
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2);
}}
footer {{
    text-align: center;
    color: var(--cor-nao-info);
    font-size: 14px;
    margin: 30px 0;
    padding: 10px;
    background-color: var(--cor-fundo-claro);
    border-top: 1px solid var(--cor-principal);
}}
.linha-media-final {{
    background-color: var(--cor-fundo-claro);
    border-top: 2px solid var(--cor-principal);
    font-weight: bold;
}}
.linha-media-final th {{
    background-color: var(--cor-principal);
    color: #ffffff;
    font-weight: 600;
}}
.linha-media-final td {{
    font-size: 16px;
}}
</style>
</head>
<body>
<div class="header-ext">
    <h1 class="titulo-principal">Boletim de Notas</h1>
    <img src="{LOGO_HTML_HEADER}" alt="Logo Univap" class="logo-univap">
</div>

<div class="card">
    <h2>Notas de <b>__NOME__</b></h2>
    <div class="meta">
        <p><strong>Nome:</strong> <b>__NOME__</b></p>
        <p><strong>Matrícula:</strong> <b>__MATRICULA__</b></p>
        <p><strong>Turma:</strong> <b>__TURMA__</b></p>
    </div>

    <table>
        <tr><th><b>Disciplina</b></th><th><b>Nota</b></th></tr>
__TABLE_ROWS__
__FINAL_AVG_ROW__
    </table>

    <div class="btn-container">
        <a href="../pdfs/__MATRICULA__.pdf" class="btn" download>Baixar PDF</a>
    </div>
</div>

<footer>&copy; 2025 Escola Univap - Todos os direitos reservados</footer>
</body>
</html>
"""

def registrar_fonte_arial():
    for p in ["Arial.ttf", r"C:\Windows\Fonts\Arial.ttf"]:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont("Arial", p))
            pdfmetrics.registerFont(TTFont("Arial-Bold", p)) 
            return "Arial"
    pdfmetrics.registerFont(TTFont("Arial", "Helvetica"))
    pdfmetrics.registerFont(TTFont("Arial-Bold", "Helvetica-Bold"))
    return "Helvetica"

FONT_PADRAO = registrar_fonte_arial()

def gerar_html_aluno(linha, disciplinas, pasta_html=DEFAULT_PASTA_HTML):
    os.makedirs(pasta_html, exist_ok=True)
    matricula = str(linha["Código"])
    nome = linha["Aluno"]
    turma = linha["Turma"]

    rows_html = []
    for d in disciplinas:
        n = linha.get(d, "Não informada")
        if pd.isna(n):
            n = "Não informada"
        
        css_class = "nota-nao"
        try:
            n_float = float(n)
            css_class = "nota-boa" if n_float >= 6.0 else "nota-baixa"
        except:
            pass
            
        rows_html.append(f"<tr><td>{d}</td><td class='{css_class}'>{n}</td></tr>")
    table_rows = "\n".join(rows_html)

    media = linha.get("MÉDIA AT3", "Não informada")
    try:
        media_float = float(media)
        media_class = "nota-boa" if media_float >= 6.0 else "nota-baixa"
    except:
        media_class = "nota-nao"

    final_avg_row = (
        f'<tr class="linha-media-final">'
        f'<th>MÉDIA AT3</th>' 
        f'<td class="{media_class}">{media}</td>' 
        f'</tr>'
    )

    conteudo = (
        HTML_TEMPLATE.replace("__NOME__", nome)
        .replace("__MATRICULA__", matricula)
        .replace("__TURMA__", turma)
        .replace("__TABLE_ROWS__", table_rows)
        .replace("__FINAL_AVG_ROW__", final_avg_row)
    )

    with open(os.path.join(pasta_html, f"{matricula}.html"), "w", encoding="utf-8") as f:
        f.write(conteudo)

def gerar_pdf_aluno(linha, disciplinas, pasta_pdf=DEFAULT_PASTA_PDF):
    os.makedirs(pasta_pdf, exist_ok=True)
    matricula = str(linha["Código"])
    nome = linha["Aluno"]
    turma = linha["Turma"]

    caminho_pdf = os.path.join(pasta_pdf, f"{matricula}.pdf")
    pdf = canvas.Canvas(caminho_pdf, pagesize=A4)
    w, h = A4 
    
    COR_PRINCIPAL = colors.HexColor('#0d47a1')
    COR_APROVADO = colors.HexColor('#0a70c2')
    COR_REPROVADO = colors.HexColor('#c42d3e')
    COR_FUNDO_CLARO = colors.HexColor('#e3f2fd') 

    pdf.setFillColor(COR_PRINCIPAL)
    pdf.rect(0, h - 30, w, 30, fill=1)
    pdf.setFillColor(colors.white)
    pdf.setFont("Arial-Bold", 18)
    pdf.drawString(cm, h - 20, "Boletim de Notas - Colégio UniVap")

    if os.path.exists(LOGO_PDF):
        largura_logo = 90
        altura_logo = 90
        y_posicao = h - 100 
        pdf.drawImage(LOGO_PDF, w - largura_logo - cm, y_posicao, width=largura_logo, height=altura_logo, preserveAspectRatio=True)

    y_aluno = h - 90
    pdf.setFillColor(colors.black)
    pdf.setFont("Arial", 12)
    pdf.drawString(cm, y_aluno, f"Nome: ")
    pdf.setFont("Arial-Bold", 12)
    pdf.drawString(cm + 35, y_aluno, nome)
    pdf.setFont("Arial", 12)
    pdf.drawString(cm, y_aluno - 20, f"Matrícula: ")
    pdf.setFont("Arial-Bold", 12)
    pdf.drawString(cm + 55, y_aluno - 20, matricula)
    pdf.setFont("Arial", 12)
    pdf.drawString(w/2 + cm, y_aluno - 20, f"Turma: ")
    pdf.setFont("Arial-Bold", 12)
    pdf.drawString(w/2 + cm + 40, y_aluno - 20, turma)

    y_tabela = y_aluno - 60
    pdf.setFillColor(COR_PRINCIPAL)
    pdf.setFont("Arial-Bold", 14)
    pdf.drawString(cm, y_tabela + 10, "Notas por Disciplina")

    largura_disciplina = 400
    largura_nota = 100
    altura_linha = 25
    x_disc = cm
    x_nota = x_disc + largura_disciplina
    largura_total = largura_disciplina + largura_nota

    pdf.setStrokeColor(COR_PRINCIPAL)
    pdf.setLineWidth(1)
    
    pdf.setFillColor(COR_PRINCIPAL)
    pdf.rect(x_disc, y_tabela, largura_total, altura_linha, fill=1, stroke=1) 

    pdf.setFillColor(colors.white)
    pdf.setFont("Arial-Bold", 12)
    pdf.drawCentredString(x_disc + largura_disciplina / 2, y_tabela + 8, "DISCIPLINA")
    pdf.drawCentredString(x_nota + largura_nota / 2, y_tabela + 8, "NOTA")

    y_linha = y_tabela - altura_linha
    pdf.setFont("Arial", 12)
    
    for d in disciplinas:
        nota = linha.get(d, "Não informada")
        if pd.isna(nota):
            nota = "Não informada"
        
        pdf.setFillColor(colors.white)
        pdf.rect(x_disc, y_linha, largura_total, altura_linha, fill=1, stroke=1) 
        
        pdf.line(x_nota, y_linha, x_nota, y_linha + altura_linha) 

        pdf.setFillColor(colors.black)
        pdf.drawString(x_disc + 10, y_linha + 8, str(d))
        
        try:
            nota_float = float(nota)
            if nota_float >= 6.0:
                pdf.setFillColor(COR_APROVADO)
            else:
                pdf.setFillColor(COR_REPROVADO)
        except:
            pdf.setFillColor(colors.grey)

        pdf.drawCentredString(x_nota + largura_nota / 2, y_linha + 8, str(nota))
        
        y_linha -= altura_linha

    media = linha.get("MÉDIA AT3", "Não informada")
    y_linha -= 10
    
    pdf.setStrokeColor(COR_PRINCIPAL) 
    pdf.setLineWidth(1)

    pdf.setFillColor(COR_FUNDO_CLARO) 
    pdf.rect(x_disc, y_linha, largura_total, altura_linha + 5, fill=1, stroke=1) 

    pdf.line(x_nota, y_linha, x_nota, y_linha + altura_linha + 5) 

    pdf.setFillColor(COR_PRINCIPAL) 
    pdf.setFont("Arial-Bold", 13)
    pdf.drawString(x_disc + 10, y_linha + 10, "MÉDIA AT3 FINAL")
    
    try:
        media_float = float(media)
        if media_float >= 6.0:
            pdf.setFillColor(COR_APROVADO)
        else:
            pdf.setFillColor(COR_REPROVADO)
    except:
        pdf.setFillColor(colors.grey)

    pdf.drawCentredString(x_nota + largura_nota / 2, y_linha + 10, str(media))

    pdf.setFillColor(colors.grey)
    pdf.setFont("Arial", 9)
    pdf.drawCentredString(w/2, 20, f"Documento gerado em {os.path.basename(caminho_pdf)}. Confirme os dados com a secretaria.")
    
    pdf.showPage()
    pdf.save()

print("==== Sistema de Boletins - Projeto POOI ====")
arquivo = input("Digite o nome do arquivo Excel (ex: alunos.xlsx): ")

try:
    df = carregar_planilha(arquivo)
except Exception as e:
    print("Erro ao carregar planilha:", e)
    exit()

print("\nPlanilha carregada com sucesso!")
print("====================================================================")

disciplinas = colunas_disciplinas(df)
print(f"1. Gerando HTMLs e PDFs para todos os {len(df)} alunos...")
for _, linha in df.iterrows():
    gerar_html_aluno(linha, disciplinas)
    gerar_pdf_aluno(linha, disciplinas)
print(f"Arquivos gerados com sucesso em '{DEFAULT_PASTA_HTML}' e '{DEFAULT_PASTA_PDF}'")
print("====================================================================")

while True:
    print("\nO que deseja fazer agora?")
    print("1. Ver tabelas prettytable no console")
    print("2. Sair")
    opcao = input("Escolha uma opção (1-2): ")
    while opcao != "1" and opcao != "2":
        print("Opção inválida. Escolha uma opção (1-2).")
        opcao = input("Escolha uma opção (1-2): ")
    if opcao == "1":
        mostrar_pretty_tables(df)
    elif opcao == "2":
        print("====================================================================")
        print("Fim do Programa.")
        print("====================================================================")
        break
