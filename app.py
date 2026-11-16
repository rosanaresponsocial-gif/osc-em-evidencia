import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="SFA - Família Acolhedora", layout="wide")

# ---------- LOGO ----------
try:
    st.image("logo.png", width=280)
except:
    st.markdown("### 🏠🧒👧 SFA - Família Acolhedora")

st.title("SFA - Família Acolhedora em Evidência")
st.markdown("Sistema de evidências - 36 casos | OSC em Evidência")

# ---------- ESTADO INICIAL ----------
if "avaliacoes" not in st.session_state:
    st.session_state["avaliacoes"] = []
if "busca_crianca" not in st.session_state:
    st.session_state["busca_crianca"] = ""
if "filtro_regiao" not in st.session_state:
    st.session_state["filtro_regiao"] = "Todas"
if "filtro_problema" not in st.session_state:
    st.session_state["filtro_problema"] = "Todos"
if "filtro_prazo" not in st.session_state:
    st.session_state["filtro_prazo"] = "Todos"

def limpar_busca():
    st.session_state["busca_crianca"] = ""
    st.session_state["filtro_regiao"] = "Todas"
    st.session_state["filtro_problema"] = "Todos"
    st.session_state["filtro_prazo"] = "Todos"

# ---------- 1 - SUBIR EXCEL ----------
st.markdown("### 1️⃣ Subir o Excel Base")
st.markdown("Suba o Excel - Base_Dados_SFA (.xlsx)")
arquivo = st.file_uploader("", type=["xlsx"])

df = None
if arquivo is not None:
    # Lê sem cabeçalho para achar a linha do cabeçalho
    df_raw = pd.read_excel(arquivo, sheet_name=1, header=None)
    linha_cab = None
    for i, row in df_raw.iterrows():
        vals = [str(x).strip() for x in row.values if str(x) != "nan"]
        if "ID do Caso" in vals or "Cód. Criança" in str(vals):
            linha_cab = i
            break
    if linha_cab is None:
        linha_cab = 3
    df = pd.read_excel(arquivo, sheet_name=1, header=linha_cab)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")
    st.success(f"✅ Base carregada: {len(df)} casos | Aba: 2_Base_Casos_Criancas")
    st.info(f"💡 Achei o cabeçalho na linha {linha_cab+1} do Excel")
    st.write("Colunas encontradas:", list(df.columns))

if df is None:
    st.warning("Suba a base para continuar.")
    st.stop()

# Acha coluna da criança
col_crianca = None
for c in df.columns:
    if "Criança" in c or "Crianca" in c:
        col_crianca = c
        break
if col_crianca is None:
    col_crianca = df.columns[1]

# ---------- 2 - BUSCAR ----------
st.markdown("### 2️⃣ 🔍 Buscar por Criança")
busca = st.text_input("Buscar", placeholder="Digite: Criança_12 ou 12", key="busca_crianca")

c1, c2, c3 = st.columns(3)
with c1:
    regioes = ["Todas"] + sorted([str(x) for x in df["Região (SP)"].dropna().unique()]) if "Região (SP)" in df.columns else ["Todas"]
    filtro_regiao = st.selectbox("Filtrar por Região", regioes, key="filtro_regiao")
with c2:
    problemas = ["Todos"] + sorted([str(x) for x in df["Problema Inicial Identificado"].dropna().unique()]) if "Problema Inicial Identificado" in df.columns else ["Todos"]
    filtro_problema = st.selectbox("Filtrar por Problema", problemas, key="filtro_problema")
with c3:
    prazos = ["Todos"] + sorted([str(x) for x in df["Status do Prazo"].dropna().unique()]) if "Status do Prazo" in df.columns else ["Todos"]
    filtro_prazo = st.selectbox("Filtrar por Status do Prazo", prazos, key="filtro_prazo")

# CORREÇÃO: usa on_click, não atribui depois do widget
st.button("🧹 Limpar busca / Mostrar todos", on_click=limpar_busca)

df_f = df.copy()
if busca.strip() != "":
    b = busca.strip().lower()
    # aceita "12" ou "Criança_12"
    num = re.sub(r"\D", "", b)
    def match(v):
        s = str(v).lower()
        if b in s:
            return True
        if num and re.sub(r"\D", "", s) == num:
            return True
        return False
    df_f = df_f[df_f[col_crianca].apply(match)]

if filtro_regiao != "Todas" and "Região (SP)" in df.columns:
    df_f = df_f[df_f["Região (SP)"].astype(str) == filtro_regiao]
if filtro_problema != "Todos" and "Problema Inicial Identificado" in df.columns:
    df_f = df_f[df_f["Problema Inicial Identificado"].astype(str) == filtro_problema]
if filtro_prazo != "Todos" and "Status do Prazo" in df.columns:
    df_f = df_f[df_f["Status do Prazo"].astype(str) == filtro_prazo]

st.write(f"Mostrando **{len(df_f)}** de **{len(df)}** casos")
st.dataframe(df_f, use_container_width=True)

# ---------- 3 - CUSTOS ----------
st.markdown("### 3️⃣ 💰 Custos")
if "Custo Bolsa-Auxílio (R$)" in df.columns:
    try:
        total_bolsa = pd.to_numeric(df["Custo Bolsa-Auxílio (R$)"], errors="coerce").sum()
        st.metric("Total Bolsa-Auxílio", f"R$ {total_bolsa:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    except:
        pass
if "Custo Equipe / Estrutura (R$)" in df.columns:
    try:
        total_eq = pd.to_numeric(df["Custo Equipe / Estrutura (R$)"], errors="coerce").sum()
        st.metric("Total Equipe / Estrutura", f"R$ {total_eq:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    except:
        pass

# ---------- 4 - GRÁFICOS ----------
st.markdown("### 4️⃣ 📊 Gráficos de Impacto")
g1, g2 = st.columns(2)
with g1:
    if "Problema Inicial Identificado" in df.columns:
        vc = df["Problema Inicial Identificado"].value_counts()
        st.write(f"**Problema Inicial**")
        st.write(f"Total: {len(df)} casos | Coluna usada: Problema Inicial Identificado")
        st.bar_chart(vc)
with g2:
    if "Status do Prazo" in df.columns:
        vc2 = df["Status do Prazo"].value_counts()
        st.write(f"**Status do Prazo**")
        st.write(f"Total: {len(df)} casos | Coluna usada: Status do Prazo")
        st.bar_chart(vc2)

# ---------- 5 - COMPARAÇÃO ----------
st.markdown("### 5️⃣ 🔄 Antes x Depois")
st.write("Compare Problema Inicial x Resultado / Transformação")
if "Problema Inicial Identificado" in df.columns and "Resultado / Transformação" in df.columns:
    st.dataframe(df[["Cód. Criança (Anonimizado)" if "Cód. Criança (Anonimizado)" in df.columns else col_crianca, "Problema Inicial Identificado", "Resultado / Transformação"]].head(20), use_container_width=True)

# ---------- 6 - WORD VIRA DADO ----------
st.markdown("### 6️⃣ 📄 Relatório em Word vira Dado")
st.markdown("Cole o texto do relatório. O app transforma em notas 0 a 3.")

def nota_vinculo(txt):
    t = txt.lower()
    if "rompido" in t or "ausente" in t or "nenhum" in t:
        return 0, "rompido"
    if "fragil" in t or "frágil" in t or "instavel" in t or "instável" in t or "inicial" in t:
        return 1, "fragil"
    if "reconstrucao" in t or "reconstrução" in t or "em reconstru" in t or "melhor" in t:
        return 2, "reconstrucao"
    if "forte" in t or "consolidado" in t or "estavel" in t or "estável" in t:
        return 3, "forte"
    return 2, "reconstrucao"

def nota_emocional(txt):
    t = txt.lower()
    if "crise" in t or "grave" in t:
        return 0, "crise"
    if "ansios" in t or "agitad" in t or "triste" in t:
        return 1, "instavel"
    if "estavel" in t or "estável" in t or "calmo" in t:
        return 2, "estavel"
    if "excelente" in t or "muito bem" in t or "seguro" in t:
        return 3, "excelente"
    return 2, "estavel"

def nota_escola(txt):
    t = txt.lower()
    m = re.search(r"(\d+)\s*dia", t)
    if m:
        d = int(m.group(1))
        if d == 0:
            return 0, "0 dias"
        if d <= 2:
            return 1, f"{d} dias"
        if d <= 4:
            return 2, f"{d} dias"
        return 3, f"{d} dias"
    if "nao foi" in t or "não foi" in t or "faltou" in t:
        return 0, "nao foi"
    return 2, "frequente"

def nota_evolucao(txt):
    t = txt.lower()
    if "pior" in t or "regrediu" in t:
        return 0, "piorou"
    if "pouco" in t or "leve" in t:
        return 1, "pouca melhora"
    if "melhorou" in t or "melhora" in t or "avan" in t:
        return 2, "melhorou"
    if "muito" in t and "melhor" in t:
        return 3, "muita melhora"
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
        dados = [{
            "Cód. Criança": cod,
            "Vínculo": lv, "Nota Vínculo": nv,
            "Emocional": le, "Nota Emocional": ne,
            "Escola": lesc, "Nota Escola": nesc,
            "Evolução": lev, "Nota Evolução": nev,
            "Média Vínculo": f"{nv:.2f} / 3",
            "Média Emocional": f"{ne:.2f} / 3",
            "Média Escola": f"{nesc:.2f} / 3",
            "Média Evolução": f"{nev:.2f} / 3",
        }]
        df_word = pd.DataFrame(dados)
        st.success("✅ Transformado!")
        st.dataframe(df_word, use_container_width=True)
        buf = BytesIO()
        df_word.to_excel(buf, index=False)
        st.download_button("⬇️ Baixar Excel Consolidado", buf.getvalue(), file_name="consolidado_word.xlsx")
        st.download_button("⬇️ Baixar também em CSV", df_word.to_csv(index=False).encode("utf-8"), file_name="consolidado_word.csv")
        if st.button("💾 Salvar correções"):
            st.success("Correções salvas!")

# ---------- 7 - VALIDAÇÃO ----------
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
    else:
        st.info("Ainda falta aprovação para chegar a 2 de 3.")

# ---------- 8 - PDF ----------
st.markdown("### 8️⃣ 📄 Relatório em PDF")
st.write("Resumo final para impressão / PDF do navegador (Ctrl+P)")
st.write(f"Total de casos: {len(df)}")
if len(st.session_state["avaliacoes"]) > 0:
    st.write("Avaliadoras:")
    for a in st.session_state["avaliacoes"]:
        st.write(f"- {a['Nome']} | {a['Decisão']} | {a['Comentário']}")
st.info("Para gerar PDF: aperte Ctrl+P no navegador e salve como PDF.")
