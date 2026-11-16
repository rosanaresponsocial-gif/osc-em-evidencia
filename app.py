import streamlit as st
import pandas as pd
import re
import unicodedata
from io import BytesIO

st.set_page_config(page_title="SFA - Família Acolhedora", layout="wide")

st.markdown("""
<style>
.stApp { background-color: #FFFFFF; }
h1, h2, h3 { color: #E85D04 !important; }
h1 { border-bottom: 4px solid #2E9E5B; padding-bottom: 10px; }
div.stButton > button { background-color: #E85D04; color: white; border: none; font-weight: bold; border-radius: 10px; }
div.stButton > button:hover { background-color: #2E9E5B; color: white; }
div[data-testid="stMetric"] { background-color: #FFF3E6; border-left: 5px solid #2E9E5B; padding: 10px; border-radius: 10px; }
</style>
""", unsafe_allow_html=True)

try:
    st.image("logo.png", width=280)
except:
    st.markdown("### 🏠🧒👧 SFA - Família Acolhedora")

st.title("SFA - Família Acolhedora em Evidência")
st.markdown("Sistema de evidências - 36 casos | OSC em Evidência")

def sem_acento(t):
    s = unicodedata.normalize("NFD", str(t))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.lower().strip()

def so_numero(v):
    return re.sub(r"\D", "", str(v))

if "avaliacoes" not in st.session_state:
    st.session_state["avaliacoes"] = []
if "escolha_crianca" not in st.session_state:
    st.session_state["escolha_crianca"] = "Todas"
if "filtro_regiao" not in st.session_state:
    st.session_state["filtro_regiao"] = "Todas"
if "filtro_problema" not in st.session_state:
    st.session_state["filtro_problema"] = "Todos"
if "filtro_prazo" not in st.session_state:
    st.session_state["filtro_prazo"] = "Todos"
if "pdf_bytes" not in st.session_state:
    st.session_state["pdf_bytes"] = None

def limpar_busca():
    st.session_state["escolha_crianca"] = "Todas"
    st.session_state["filtro_regiao"] = "Todas"
    st.session_state["filtro_problema"] = "Todos"
    st.session_state["filtro_prazo"] = "Todos"

def gerar_pdf(total_casos, escolha, avaliacoes):
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.lib import colors
        buf = BytesIO()
        doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=40, bottomMargin=40)
        styles = getSampleStyleSheet()
        story = []
        story.append(Paragraph("<b>SFA - Familia Acolhedora em Evidencia</b>", styles["Title"]))
        story.append(Spacer(1, 12))
        story.append(Paragraph(f"Relatorio de Evidencias - Total de casos: {total_casos}", styles["Normal"]))
        story.append(Paragraph(f"Filtro atual - Crianca: {escolha}", styles["Normal"]))
        story.append(Spacer(1, 12))
        story.append(Paragraph("<b>Validacao - Avaliacoes:</b>", styles["Heading2"]))
        if len(avaliacoes) == 0:
            story.append(Paragraph("Nenhuma avaliacao registrada ainda.", styles["Normal"]))
        else:
            dados = [["Nome", "Decisao", "Comentario"]]
            for a in avaliacoes:
                dados.append([str(a["Nome"]), str(a["Decisão"]), str(a["Comentário"])])
            t = Table(dados, colWidths=[120, 120, 220])
            t.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E85D04")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]))
            story.append(t)
            n_aprov = sum(1 for a in avaliacoes if "Aprovo" in a["Decisão"])
            story.append(Spacer(1, 12))
            story.append(Paragraph(f"{n_aprov} de {len(avaliacoes)} aprovaram.", styles["Normal"]))
            if n_aprov >= 2:
                story.append(Paragraph("<b>RELATORIO APROVADO - 2 de 3 aprovaram</b>", styles["Normal"]))
        doc.build(story)
        buf.seek(0)
        return buf.getvalue()
    except Exception as e:
        return None

st.markdown("### 1️⃣ Subir o Excel Base")
arquivo = st.file_uploader("", type=["xlsx"])
df = None
if arquivo is not None:
    try:
        xls = pd.ExcelFile(arquivo)
        nome_aba = xls.sheet_names[1] if len(xls.sheet_names) > 1 else xls.sheet_names[0]
        df_raw = pd.read_excel(arquivo, sheet_name=nome_aba, header=None)
        linha_cab = None
        for i, row in df_raw.iterrows():
            vals = [str(x).strip() for x in row.values if str(x) != "nan"]
            texto = " ".join(vals)
            if "ID do Caso" in texto or "Criança" in texto:
                linha_cab = i
                break
        if linha_cab is None:
            linha_cab = 3
        df = pd.read_excel(arquivo, sheet_name=nome_aba, header=linha_cab)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.dropna(how="all")
        for c in df.columns:
            if df[c].dtype == object:
                df[c] = df[c].astype(str).str.strip().replace("nan", "")
        st.success(f"✅ Base carregada: {len(df)} casos")
    except Exception as e:
        st.error(f"Erro ao ler Excel: {e}")
        st.stop()

if df is None:
    st.warning("Suba a base para continuar.")
    st.stop()

col_crianca = None
for c in df.columns:
    if "Criança" in c or "Crianca" in c:
        col_crianca = c
        break
if col_crianca is None:
    col_crianca = df.columns[1]

def achar_coluna(palavras):
    for c in df.columns:
        cs = sem_acento(c)
        if all(sem_acento(p) in cs for p in palavras):
            return c
    return None

col_regiao = achar_coluna(["regiao"])
col_problema = achar_coluna(["problema", "inicial"])
col_prazo = achar_coluna(["status", "prazo"])
col_resultado = achar_coluna(["resultado"])
col_bolsa = achar_coluna(["bolsa"])
col_equipe = achar_coluna(["equipe"])

st.markdown("### 2️⃣ 🔍 Buscar por Criança")
st.caption("É só clicar na lista e escolher. Não precisa digitar nem apertar Enter.")
valores_crianca = sorted([str(x).strip() for x in df[col_crianca].dropna().unique() if str(x).strip() != "" and str(x).strip().lower() != "nan"])
lista_criancas = ["Todas"] + valores_crianca
escolha = st.selectbox("Escolha a criança:", lista_criancas, key="escolha_crianca")

c1, c2, c3 = st.columns(3)
with c1:
    regioes = ["Todas"] + sorted([str(x).strip() for x in df[col_regiao].dropna().unique()]) if col_regiao else ["Todas"]
    filtro_regiao = st.selectbox("Filtrar por Região", regioes, key="filtro_regiao")
with c2:
    problemas = ["Todos"] + sorted([str(x).strip() for x in df[col_problema].dropna().unique()]) if col_problema else ["Todos"]
    filtro_problema = st.selectbox("Filtrar por Problema", problemas, key="filtro_problema")
with c3:
    prazos = ["Todos"] + sorted([str(x).strip() for x in df[col_prazo].dropna().unique()]) if col_prazo else ["Todos"]
    filtro_prazo = st.selectbox("Filtrar por Status do Prazo", prazos, key="filtro_prazo")

st.button("🧹 Limpar busca / Mostrar todos", on_click=limpar_busca)

df_f = df.copy()
if escolha != "Todas":
    num_escolha = so_numero(escolha)
    df_f = df_f[df_f[col_crianca].apply(lambda x: so_numero(x) == num_escolha)]

if filtro_regiao != "Todas" and col_regiao:
    df_f = df_f[df_f[col_regiao].astype(str).str.strip() == filtro_regiao.strip()]
if filtro_problema != "Todos" and col_problema:
    df_f = df_f[df_f[col_problema].astype(str).str.strip() == filtro_problema.strip()]
if filtro_prazo != "Todos" and col_prazo:
    df_f = df_f[df_f[col_prazo].astype(str).str.strip() == filtro_prazo.strip()]

st.write(f"Mostrando {len(df_f)} de {len(df)} casos")
st.dataframe(df_f, use_container_width=True)

st.markdown("### 3️⃣ 💰 Custos")
if col_bolsa:
    try:
        total_bolsa = pd.to_numeric(df[col_bolsa], errors="coerce").sum()
        st.metric("Total Bolsa-Auxílio", f"R$ {total_bolsa:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    except: pass
if col_equipe:
    try:
        total_eq = pd.to_numeric(df[col_equipe], errors="coerce").sum()
        st.metric("Total Equipe / Estrutura", f"R$ {total_eq:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    except: pass

st.markdown("### 4️⃣ 📊 Gráficos de Impacto")
g1, g2 = st.columns(2)
with g1:
    if col_problema:
        st.write("**Problema Inicial**")
        st.bar_chart(df[col_problema].value_counts())
with g2:
    if col_prazo:
        st.write("**Status do Prazo**")
        st.bar_chart(df[col_prazo].value_counts())

st.markdown("### 5️⃣ 🔄 Antes x Depois")
if col_problema and col_resultado:
    st.dataframe(df[[col_crianca, col_problema, col_resultado]].head(20), use_container_width=True)

st.markdown("### 6️⃣ 📄 Relatório em Word vira Dado")
def nota_vinculo(txt):
    t = sem_acento(txt)
    if "rompido" in t or "ausente" in t: return 0, "rompido"
    if "fragil" in t or "instavel" in t: return 1, "fragil"
    if "reconstrucao" in t or "melhor" in t: return 2, "reconstrucao"
    if "forte" in t or "consolidado" in t: return 3, "forte"
    return 2, "reconstrucao"
def nota_emocional(txt):
    t = sem_acento(txt)
    if "crise" in t or "grave" in t: return 0, "crise"
    if "ansios" in t or "triste" in t: return 1, "instavel"
    if "estavel" in t or "calmo" in t: return 2, "estavel"
    if "excelente" in t or "seguro" in t: return 3, "excelente"
    return 2, "estavel"
def nota_escola(txt):
    t = sem_acento(txt)
    m = re.search(r"(\d+)\s*dia", t)
    if m:
        d = int(m.group(1))
        if d == 0: return 0, "0 dias"
        if d <= 2: return 1, f"{d} dias"
        if d <= 4: return 2, f"{d} dias"
        return 3, f"{d} dias"
    if "nao foi" in t or "faltou" in t: return 0, "nao foi"
    return 2, "frequente"
def nota_evolucao(txt):
    t = sem_acento(txt)
    if "pior" in t or "regrediu" in t: return 0, "piorou"
    if "pouco" in t or "leve" in t: return 1, "pouca melhora"
    return 2, "melhorou"

texto_word = st.text_area("Cole aqui o texto do Word:", height=150, placeholder="Ex:\nCriança_12\nVínculo em reconstrução\nEstado emocional estável\nFoi 3 dias na escola\nMelhorou um pouco")
if st.button("✨ Transformar em Excel"):
    if texto_word.strip() == "":
        st.warning("Cole um texto primeiro.")
    else:
        linhas = [l.strip() for l in texto_word.splitlines() if l.strip() != ""]
        cod = linhas[0] if linhas else "Sem código"
        txt_full = " ".join(linhas)
        nv, lv = nota_vinculo(txt_full)
        ne, le = nota_emocional(txt_full)
        nesc, lesc = nota_escola(txt_full)
        nev, lev = nota_evolucao(txt_full)
        dados = [{"Cód. Criança": cod, "Vínculo": lv, "Nota Vínculo": nv, "Emocional": le, "Nota Emocional": ne, "Escola": lesc, "Nota Escola": nesc, "Evolução": lev, "Nota Evolução": nev}]
        df_word = pd.DataFrame(dados)
        st.success("✅ Transformado!")
        st.dataframe(df_word, use_container_width=True)
        buf = BytesIO()
        df_word.to_excel(buf, index=False)
        st.download_button("⬇️ Baixar Excel Consolidado", buf.getvalue(), file_name="consolidado_word.xlsx")

st.markdown("### 7️⃣ ✅ Validação - 2 de 3 aprovam")
nome_av = st.text_input("Nome da avaliadora:", key="nome_av")
decisao = st.selectbox("Decisão:", ["Aprovo", "Aprovo com ressalvas", "Não aprovo"], key="decisao_av")
coment = st.text_area("Comentário:", key="coment_av")
if st.button("Registrar Avaliação"):
    if nome_av.strip() == "":
        st.warning("Digite o nome.")
    else:
        st.session_state["avaliacoes"].append({"Nome": nome_av, "Decisão": decisao, "Comentário": coment})
        st.success(f"✅ Avaliação de {nome_av} registrada: {decisao}")
if len(st.session_state["avaliacoes"]) > 0:
    st.dataframe(pd.DataFrame(st.session_state["avaliacoes"]), use_container_width=True)
    n_aprov = sum(1 for a in st.session_state["avaliacoes"] if "Aprovo" in a["Decisão"])
    st.write(f"**{n_aprov} de {len(st.session_state['avaliacoes'])} aprovaram**")
    if n_aprov >= 2:
        st.success(f"✅ RELATÓRIO APROVADO - {n_aprov} de 3 aprovaram")

# ---------- 8 - PDF COM BOTAO ----------
st.markdown("### 8️⃣ 📄 Relatório em PDF")
st.write(f"Total de casos: {len(df)} | Filtro atual: {escolha}")
if len(st.session_state["avaliacoes"]) > 0:
    for a in st.session_state["avaliacoes"]:
        st.write(f"- {a['Nome']} | {a['Decisão']} | {a['Comentário']}")

if st.button("📄 Gerar Relatório em PDF"):
    pdf = gerar_pdf(len(df), escolha, st.session_state["avaliacoes"])
    if pdf is None:
        st.error("Não consegui gerar o PDF. Vou tentar instalar a biblioteca. Aperte Ctrl+P por enquanto e salve como PDF.")
    else:
        st.session_state["pdf_bytes"] = pdf
        st.success("✅ PDF gerado! Agora clica em Baixar.")

if st.session_state["pdf_bytes"] is not None:
    st.download_button("⬇️ Baixar PDF", st.session_state["pdf_bytes"], file_name="relatorio_sfa.pdf", mime="application/pdf")
