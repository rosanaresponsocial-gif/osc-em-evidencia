import streamlit as st
import pandas as pd
import hashlib
import os
from fpdf import FPDF

st.set_page_config(page_title="SFA - Restrito", layout="wide")

ARQUIVO_BASE = "dados_base.xlsx"
ARQUIVO_USUARIOS = "usuarios.csv"
ARQUIVO_COMPLEMENTAR = "dados_complementares.csv"
ADMIN_LOGIN = "sfaevidencia"
SENHA_ADMIN_NOVA = "Sfa2026"

def criptografar(senha):
    return hashlib.sha256(senha.encode()).hexdigest()

def carregar_usuarios():
    if not os.path.exists(ARQUIVO_USUARIOS):
        df = pd.DataFrame(columns=["nome", "login", "senha_hash", "status"])
        df.to_csv(ARQUIVO_USUARIOS, index=False)
        return df
    return pd.read_csv(ARQUIVO_USUARIOS, dtype=str).fillna("")

def salvar_usuarios(df):
    df.to_csv(ARQUIVO_USUARIOS, index=False)

def carregar_base():
    if not os.path.exists(ARQUIVO_BASE):
        st.error("Base nao encontrada. Fale com a administradora.")
        st.stop()
    return pd.read_excel(ARQUIVO_BASE, dtype=str).fillna("")
def carregar_complementar():
    if not os.path.exists(ARQUIVO_COMPLEMENTAR):
        return pd.DataFrame(columns=["codigo", "preenchido_por", "observacao", "data_avaliacao"])
    return pd.read_csv(ARQUIVO_COMPLEMENTAR, dtype=str).fillna("")

if "logado" not in st.session_state:
    st.session_state.logado = False
    st.session_state.login = ""
    st.session_state.nome = ""

if not st.session_state.logado:
    st.title("SFA - Acesso Restrito")
    st.write("Somente pessoas aprovadas pela administradora.")
    aba = st.tabs(["Entrar", "Cadastrar"])

    with aba[0]:
        login = st.text_input("Login", key="login_entrar")
        senha = st.text_input("Senha", type="password", key="senha_entrar")
        if st.button("Entrar", use_container_width=True):
            usuarios = carregar_usuarios()
            # RESET DA ADMIN - entra com a senha nova
            if login == ADMIN_LOGIN and senha == SENHA_ADMIN_NOVA:
                if login not in usuarios["login"].values:
                    novo = pd.DataFrame([{"nome": "Roseane", "login": ADMIN_LOGIN, "senha_hash": criptografar(SENHA_ADMIN_NOVA), "status": "aprovado"}])
                    usuarios = pd.concat([usuarios, novo], ignore_index=True)
                else:
                    usuarios.loc[usuarios["login"] == ADMIN_LOGIN, "senha_hash"] = criptografar(SENHA_ADMIN_NOVA)
                    usuarios.loc[usuarios["login"] == ADMIN_LOGIN, "status"] = "aprovado"
                salvar_usuarios(usuarios)
                st.session_state.logado = True
                st.session_state.login = login
                st.session_state.nome = "Roseane"
                st.rerun()
            else:
                user = usuarios[usuarios["login"] == login]
                if user.empty:
                    st.error("Login nao encontrado.")
                elif user.iloc[0]["senha_hash"] != criptografar(senha):
                    st.error("Senha errada.")
                elif user.iloc[0]["status"] != "aprovado" and login != ADMIN_LOGIN:
                    st.warning("Seu cadastro esta aguardando aprovacao da administradora.")
                else:
                    st.session_state.logado = True
                    st.session_state.login = login
                    st.session_state.nome = user.iloc[0]["nome"]
                    st.rerun()

    with aba[1]:
        nome_novo = st.text_input("Seu nome completo")
        login_novo = st.text_input("Crie um login (sem espaco)", key="login_novo")
        senha_nova = st.text_input("Crie uma senha", type="password", key="senha_nova")
        if st.button("Enviar cadastro para aprovacao", use_container_width=True):
            if not nome_novo or not login_novo or not senha_nova:
                st.error("Preencha tudo.")
            else:
                usuarios = carregar_usuarios()
                if login_novo in usuarios["login"].values:
                    st.error("Esse login ja existe. Escolha outro.")
                else:
                    status_novo = "aprovado" if login_novo == ADMIN_LOGIN else "pendente"
                    novo = pd.DataFrame([{"nome": nome_novo, "login": login_novo, "senha_hash": criptografar(senha_nova), "status": status_novo}])
                    usuarios = pd.concat([usuarios, novo], ignore_index=True)
                    salvar_usuarios(usuarios)
                    if status_novo == "aprovado":
                        st.success("Cadastro de administradora criado! Agora va para Entrar.")
                    else:
                        st.success("Cadastro enviado! Aguarde aprovacao.")
    st.stop()

st.sidebar.write(f"Ola, {st.session_state.nome}")
st.sidebar.write(f"Login: {st.session_state.login}")
if st.button("Sair", key="sair"):
    st.session_state.logado = False
    st.rerun()

eh_admin = (st.session_state.login == ADMIN_LOGIN)

if eh_admin:
    st.title("Painel da Administradora")
    usuarios = carregar_usuarios()
    pendentes = usuarios[usuarios["status"] == "pendente"]
    st.subheader(f"Aguardando aprovacao: {len(pendentes)}")
    for idx, row in pendentes.iterrows():
        col1, col2, col3 = st.columns([2, 1, 1])
        col1.write(f"{row['nome']} - login: {row['login']}")
        if col2.button("Aprovar", key=f"ap_{idx}"):
            usuarios.loc[idx, "status"] = "aprovado"
            salvar_usuarios(usuarios)
            st.rerun()
        if col3.button("Negar", key=f"ne_{idx}"):
            usuarios.loc[idx, "status"] = "negado"
            salvar_usuarios(usuarios)
            st.rerun()
    st.divider()

st.title("Avaliacao - So Codigo")
df_base = carregar_base()
df_comp = carregar_complementar()
col_codigo = df_base.columns[0]
st.write(f"Total de criancas na base: {len(df_base)}")
codigo_escolhido = st.selectbox("Escolha o codigo:", [""] + sorted(df_base[col_codigo].unique().tolist()))

if codigo_escolhido:
    st.info(f"Voce esta avaliando: {codigo_escolhido}")
    with st.form("form_comp"):
        obs = st.text_area("Observacao / Dados complementares")
        data_av = st.date_input("Data da avaliacao")
        enviar = st.form_submit_button("Salvar avaliacao")
        if enviar:
            novo = pd.DataFrame([{"codigo": codigo_escolhido, "preenchido_por": st.session_state.login, "observacao": obs, "data_avaliacao": str(data_av)}])
            df_comp = pd.concat([df_comp, novo], ignore_index=True)
            df_comp.to_csv(ARQUIVO_COMPLEMENTAR, index=False)
            st.success("Salvo com sucesso!")

    if st.button("Baixar PDF - So Codigo"):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", "B", 16)
        pdf.cell(0, 10, "Relatorio Descritivo - SFA", ln=True, align="C")
        pdf.set_font("Arial", "", 12)
        pdf.cell(0, 10, f"Codigo: {codigo_escolhido}", ln=True)
        pdf.cell(0, 10, f"Preenchido por: {st.session_state.login}", ln=True)
        comp_esse = df_comp[df_comp["codigo"] == codigo_escolhido]
        for _, r in comp_esse.iterrows():
            pdf.multi_cell(0, 10, f"- {r['data_avaliacao']}: {r['observacao']}")
        pdf_bytes = pdf.output()
        st.download_button(label="Baixar PDF agora", data=bytes(pdf_bytes), file_name=f"relatorio_{codigo_escolhido}.pdf", mime="application/pdf")
