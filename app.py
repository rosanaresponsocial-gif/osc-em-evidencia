import os
import io
import re
import streamlit as st
import pandas as pd
from docx import Document

ARQUIVO_BASE = "dados_base.xlsx"
USUARIO_CORRETO = "sfaevidencia"
SENHA_CORRETA = "Sfa2026"

st.set_page_config(page_title="SFA - Restrito", page_icon="🧒", layout="centered")

if "logado" not in st.session_state:
    st.session_state.logado = False

if not st.session_state.logado:
    st.title("SFA - Restrito")
    usuario = st.text_input("Usuario")
    senha = st.text_input("Senha", type="password")
    if st.button("Entrar"):
        if usuario == USUARIO_CORRETO and senha == SENHA_CORRETA:
            st.session_state.logado = True
            st.rerun()
        else:
            st.error("Usuario ou senha incorretos.")
    st.stop()

def carregar_base():
    if not os.path.exists(ARQUIVO_BASE):
        st.error("Base nao encontrada.")
        st.stop()
    try:
        df = pd.read_excel(ARQUIVO_BASE, dtype=str, header=3, sheet_name=1).fillna("")
    except Exception:
        df = pd.read_excel(ARQUIVO_BASE, dtype=str, header=3).fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    if "ID do Caso" in df.columns:
        df = df.rename(columns={"ID do Caso": "Codigo"})
    df = df[df["Codigo"].astype(str).str.strip() != ""]
    df = df[df["Codigo"].astype(str).str.strip().str.lower() != "nan"]
    return df

df_base = carregar_base()

st.title("OSC em Evidencia - SFA")
st.write(f"Total de criancas na base: {len(df_base)}")

aba1, aba2 = st.tabs(["Consultar", "Adicionar via Word"])

with aba1:
    codigo = st.selectbox("Escolha o codigo da crianca:", [""] + sorted(df_base["Codigo"].tolist()))
    if codigo != "":
        linha = df_base[df_base["Codigo"] == codigo].iloc[0]
        st.subheader(f"Avaliacao - {codigo}")
        for coluna in df_base.columns:
            if coluna == "Codigo":
                continue
            valor = str(linha[coluna]).strip()
            if valor != "" and valor.lower() != "nan":
                st.write(f"**{coluna}:** {valor}")

with aba2:
    st.subheader("Adicionar relatorio via Word")
    st.write("Suba o Word com a tabela Data / Evolucao. O app vai ler a tabela e contar tudo.")
    arquivo_word = st.file_uploader("Escolha o arquivo Word (.docx)", type=["docx"])

    def ler_tabela_word(arquivo):
        doc = Document(arquivo)
        registros = []
        # Procura todas as tabelas do Word
        for tabela in doc.tables:
            for i, linha in enumerate(tabela.rows):
                celulas = [c.text.strip() for c in linha.cells]
                if len(celulas) < 2:
                    continue
                # Pula o cabecalho Data / Evolucao
                if i == 0 and "data" in celulas[0].lower():
                    continue
                if celulas[0] == "" and celulas[1] == "":
                    continue
                registros.append({"Data": celulas[0], "Evolucao": celulas[1]})
        # Se nao achou tabela, tenta ler os paragrafos
        if len(registros) == 0:
            for p in doc.paragraphs:
                if p.text.strip() != "":
                    registros.append({"Data": "", "Evolucao": p.text.strip()})
        return registros

    def contar(registros):
        texto_tudo = " ".join([r["Evolucao"] for r in registros]).lower()
        total_visitas = texto_tudo.count("visita")
        total_contatos = texto_tudo.count("contato") + texto_tudo.count("whatsapp") + texto_tudo.count("ligacao") + texto_tudo.count("ligação")
        total_articulacoes = texto_tudo.count("articula")
        total_consultas = texto_tudo.count("consulta") + texto_tudo.count("ubs") + texto_tudo.count("pediatra") + texto_tudo.count("saude") + texto_tudo.count("saúde")
        total_atendimentos = texto_tudo.count("atendimento")
        return total_visitas, total_contatos, total_articulacoes, total_consultas, total_atendimentos

    if arquivo_word is not None:
        registros = ler_tabela_word(arquivo_word)
        st.success(f"Achei {len(registros)} registros no Word!")
        st.dataframe(pd.DataFrame(registros))

        visitas, contatos, articulacoes, consultas, atendimentos = contar(registros)
        datas = [r["Data"] for r in registros if r["Data"] != ""]
        primeiro = datas[0] if len(datas) > 0 else ""
        ultimo = datas[-1] if len(datas) > 0 else ""
        resumo = " | ".join([r["Evolucao"][:120] for r in registros[:5]])

        st.write("**O que o app contou sozinho:**")
        st.write(f"- Total de registros: {len(registros)}")
        st.write(f"- Total de Visitas: {visitas}")
        st.write(f"- Total de Contatos: {contatos}")
        st.write(f"- Total de Articulacoes: {articulacoes}")
        st.write(f"- Total de Consultas / Saude: {consultas}")
        st.write(f"- Total de Atendimentos: {atendimentos}")
        st.write(f"- Primeiro contato: {primeiro}")
        st.write(f"- Ultimo contato: {ultimo}")

        novo_codigo = st.text_input("ID do Caso / Codigo da crianca desse relatorio", value="")
        if st.button("Adicionar e gerar Excel novo"):
            if novo_codigo.strip() == "":
                st.error("Preencha o ID do Caso, ex: SFA-SP-202542")
            else:
                nova_linha = {}
                for col in df_base.columns:
                    nova_linha[col] = ""
                nova_linha["Codigo"] = novo_codigo
                # Cria colunas novas se nao existirem
                nova_linha["Total Registros"] = str(len(registros))
                nova_linha["Total Visitas"] = str(visitas)
                nova_linha["Total Contatos"] = str(contatos)
                nova_linha["Total Articulacoes"] = str(articulacoes)
                nova_linha["Total Consultas Saude"] = str(consultas)
                nova_linha["Total Atendimentos"] = str(atendimentos)
                nova_linha["Primeiro Contato"] = primeiro
                nova_linha["Ultimo Contato"] = ultimo
                nova_linha["Resumo Evolucao"] = resumo

                df_novo = pd.concat([df_base, pd.DataFrame([nova_linha])], ignore_index=True)
                # Preenche colunas novas que faltavam na base antiga
                df_novo = df_novo.fillna("")
                st.success(f"Caso {novo_codigo} adicionado! Baixe o Excel novo abaixo.")
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    df_novo.to_excel(writer, index=False, startrow=3, sheet_name="2_Base_Casos_Criancas")
                st.download_button("Baixar Excel atualizado", data=buffer.getvalue(), file_name="dados_base_novo.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.warning("Depois de baixar, suba esse arquivo no GitHub com o nome dados_base.xlsx para virar a base oficial.")

if st.button("Sair"):
    st.session_state.logado = False
    st.rerun()
