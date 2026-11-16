import os
import streamlit as st
import pandas as pd

ARQUIVO_BASE = "dados_base.xlsx"
USUARIO_CORRETO = "sfaevidencia"
SENHA_CORRETA = "Sfa2026"

st.set_page_config(page_title="SFA - Restrito", page_icon="🧒", layout="centered")

# ---------- LOGIN ----------
if "logado" not in st.session_state:
    st.session_state.logado = False

if not st.session_state.logado:
    st.title("SFA - Restrito")
    st.write("Area restrita. Faca login para continuar.")
    usuario = st.text_input("Usuario")
    senha = st.text_input("Senha", type="password")
    if st.button("Entrar"):
        if usuario == USUARIO_CORRETO and senha == SENHA_CORRETA:
            st.session_state.logado = True
            st.rerun()
        else:
            st.error("Usuario ou senha incorretos.")
    st.stop()

# ---------- CARREGAR BASE ----------
def carregar_base():
    if not os.path.exists(ARQUIVO_BASE):
        st.error("Base nao encontrada. Fale com a administradora.")
        st.stop()
    try:
        df = pd.read_excel(ARQUIVO_BASE, dtype=str, header=3, sheet_name=1).fillna("")
    except Exception:
        try:
            df = pd.read_excel(ARQUIVO_BASE, dtype=str, header=3).fillna("")
        except Exception:
            st.error("A base esta com formato invalido. Fale com a administradora para subir a base de novo.")
            st.stop()
    df.columns = [str(c).strip() for c in df.columns]
    if "ID do Caso" in df.columns:
        df = df.rename(columns={"ID do Caso": "Codigo"})
    if "Codigo" not in df.columns:
        st.error("Coluna Codigo nao encontrada na base.")
        st.write("Colunas encontradas:")
        st.write(list(df.columns))
        st.stop()
    df = df[df["Codigo"].astype(str).str.strip() != ""]
    df = df[df["Codigo"].astype(str).str.strip().str.lower() != "nan"]
    return df

df_base = carregar_base()

# ---------- TELA PRINCIPAL ----------
st.title("OSC em Evidencia - SFA")
st.write(f"Total de criancas na base: {len(df_base)}")

if os.path.exists("logo.png"):
    st.image("logo.png", width=200)

codigo = st.selectbox("Escolha o codigo da crianca:", [""] + sorted(df_base["Codigo"].tolist()))

if codigo == "":
    st.info("Escolha um codigo acima para ver a avaliacao.")
    st.stop()

linha = df_base[df_base["Codigo"] == codigo].iloc[0]

st.subheader(f"Avaliacao - {codigo}")
st.write("Dados da crianca:")

# Mostra os dados mais importantes, se existirem na base
for coluna in df_base.columns:
    if coluna == "Codigo":
        continue
    valor = str(linha[coluna]).strip()
    if valor != "" and valor.lower() != "nan":
        st.write(f"**{coluna}:** {valor}")

# Avaliacao simples baseada no Status do Prazo
st.subheader("Avaliacao")
status = ""
for col in df_base.columns:
    if "Status" in col:
        status = str(linha[col])
        break

if "Antes do Prazo" in status:
    st.success("Avaliacao: Caso com desacolhimento antes do prazo. Bom acompanhamento.")
elif "Dentro do Prazo" in status:
    st.success("Avaliacao: Caso com desacolhimento dentro do prazo. Acompanhamento adequado.")
elif "Apos o Prazo" in status:
    st.warning("Avaliacao: Caso com desacolhimento apos o prazo. Verificar os motivos do atraso.")
elif "Em Acolhimento" in status or "Em Curso" in status:
    st.info("Avaliacao: Caso em acolhimento ativo. Continuar acompanhamento.")
else:
    st.info(f"Avaliacao: Status registrado como: {status}")

if st.button("Sair"):
    st.session_state.logado = False
    st.rerun()
