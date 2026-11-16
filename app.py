import streamlit as st
import pandas as pd
import hashlib
import os
from fpdf import FPDF

st.set_page_config(page_title="SFA - Restrito", layout="wide")

# --- CONFIG ---
ARQUIVO_BASE = "dados_base.xlsx"
ARQUIVO_USUARIOS = "usuarios.csv"
ARQUIVO_COMPLEMENTAR = "dados_complementares.csv"
ADMIN_LOGIN = "roseane"  # <-- TROQUE AQUI pelo seu login de admin

# --- FUNÇÕES SEGURANÇA ---
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
    # Ninguém vê a planilha, só o app lê
    if not os.path.exists(ARQUIVO_BASE):
        st.error("Base não encontrada. Fale com a administradora.")
        st.stop()
    return pd.read_excel(ARQUIVO_BASE, dtype=str).fillna("")

def carregar_complementar():
    if not os.path.exists(ARQUIVO_COMPLEMENTAR):
        return pd.DataFrame(columns=["codigo", "preenchido_por", "observacao", "data_avaliacao"])
    return pd.read_csv(ARQUIVO_COMPLEMENTAR, dtype=str).fillna("")

# --- SESSÃO ---
if "logado" not in st.session_state:
    st.session_state.logado = False
    st.session_state.login = ""
    st.session_state.nome = ""

# --- TELA DE LOGIN / CADASTRO ---
if not st.session_state.logado:
    st.title("🔒 SFA - Acesso Restrito")
    st.write("Somente pessoas aprovadas pela administradora.")

    aba = st.tabs(["Entrar", "Cadastrar"])
    
    with aba[0]:
        login = st.text_input("Login", key="login_entrar")
        senha = st.text_input("Senha", type="password", key="senha_entrar")
        if st.button("Entrar", use_container_width=True):
            usuarios = carregar_usuarios()
            user = usuarios[usuarios["login"] == login]
            if user.empty:
                st.error("Login não encontrado.")
            elif user.iloc[0]["senha_hash"] != criptografar(senha):
                st.error("Senha errada.")
            elif user.iloc[0]["status"] != "aprovado":
                st.warning("Seu cadastro está aguardando aprovação da Roseane.")
            else:
                st.session_state.logado = True
                st.session_state.login = login
                st.session_state.nome = user.iloc[0]["nome"]
                st.rerun()

    with aba[1]:
        nome_novo = st.text_input("Seu nome completo")
        login_novo = st.text_input("Crie um login (sem espaço)", key="login_novo")
        senha_nova = st.text_input("Crie uma senha", type="password", key="senha_nova")
        if st.button("Enviar cadastro para aprovação", use_container_width=True):
            if not nome_novo or not login_novo or not senha_nova:
                st.error("Preencha tudo.")
            else:
                usuarios = carregar_usuarios()
                if login_novo in usuarios["login"].values:
                    st.error("Esse login já existe. Escolha outro.")
                else:
                    novo = pd.DataFrame([{
                        "nome": nome_novo,
                        "login": login_novo,
                        "senha_hash": criptografar(senha_nova),
                        "status": "pendente"
                    }])
                    usuarios = pd.concat([usuarios, novo], ignore_index=True)
                    salvar_usuarios(usuarios)
                    st.success("Cadastro enviado! Aguarde a Roseane aprovar.")
    st.stop()

# --- LOGADO ---
st.sidebar.write(f"Olá, **{st.session_state.nome}**")
st.sidebar.write(f"Login: {st.session_state.login}")
if st.button("Sair", key="sair"):
    st.session_state.logado = False
    st.rerun()

eh_admin = (st.session_state.login == ADMIN_LOGIN)

# --- ABA ADMIN: APROVAR ---
if eh_admin:
    st.title("👑 Painel da Administradora")
    usuarios = carregar_usuarios()
    pendentes = usuarios[usuarios["status"] == "pendente"]
    
    st.subheader(f"Aguardando aprovação: {len(pendentes)}")
    for idx, row in pendentes.iterrows():
        col1, col2, col3 = st.columns([2, 1, 1])
        col1.write(f"**{row['nome']}** - login: `{row['login']}`")
        if col2.button("✅ Aprovar", key=f"ap_{idx}"):
            usuarios.loc[idx, "status"] = "aprovado"
            salvar_usuarios(usuarios)
            st.rerun()
        if col3.button("❌ Negar", key=f"ne_{idx}"):
            usuarios.loc[idx, "status"] = "negado"
            salvar_usuarios(usuarios)
            st.rerun()
    st.divider()

# --- APP PRINCIPAL ---
st.title("Avaliação - Só Código")
df_base = carregar_base()
df_comp = carregar_complementar()

# Mostra só colunas de código, nunca nome
# Ajuste aqui o nome da coluna de código na sua planilha
col_codigo = df_base.columns[0]  # primeira coluna = código
st.write(f"Total de crianças na base: **{len(df_base)}**")

# Filtro por código
codigo_escolhido = st.selectbox("Escolha o código:", [""] + sorted(df_base[col_codigo].unique().tolist()))

if codigo_escolhido:
    dados_crianca = df_base[df_base[col_codigo] == codigo_escolhido].iloc[0]
    st.info(f"Você está avaliando: **{codigo_escolhido}**")
    
    with st.form("form_comp"):
        obs = st.text_area("Observação / Dados complementares")
        data_av = st.date_input("Data da avaliação")
        enviar = st.form_submit_button("Salvar avaliação")
        if enviar:
            novo = pd.DataFrame([{
                "codigo": codigo_escolhido,
                "preenchido_por": st.session_state.login,
                "observacao": obs,
                "data_avaliacao": str(data_av)
            }])
            df_comp = pd.concat([df_comp, novo], ignore_index=True)
            df_comp.to_csv(ARQUIVO_COMPLEMENTAR, index=False)
            st.success("Salvo com sucesso!")

    # --- GERAR PDF SÓ COM CÓDIGO ---
    if st.button("⬇️ Baixar PDF - Só Código"):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", "B", 16)
        pdf.cell(0, 10, "Relatorio Descritivo - SFA", ln=True, align="C")
        pdf.set_font("Arial", "", 12)
        pdf.cell(0, 10, f"Codigo: {codigo_escolhido}", ln=True)
        pdf.cell(0, 10, f"Preenchido por: {st.session_state.login}", ln=True)
        # Pega os complementares desse código
        comp_esse = df_comp[df_comp["codigo"] == codigo_escolhido]
        for _, r in comp_esse.iterrows():
            pdf.multi_cell(0, 10, f"- {r['data_avaliacao']}: {r['observacao']}")
        
        pdf_bytes = pdf.output()
        st.download_button(
            label="Baixar PDF agora",
            data=bytes(pdf_bytes),
            file_name=f"relatorio_{codigo_escolhido}.pdf",
            mime="application/pdf"
        )
