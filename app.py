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

aba1, aba2, aba3 = st.tabs(["Consultar", "Adicionar via Word", "Relatorio com Graficos"])

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
        for tabela in doc.tables:
            for i, linha in enumerate(tabela.rows):
                celulas = [c.text.strip() for c in linha.cells]
                if len(celulas) < 2:
                    continue
                if i == 0 and "data" in celulas[0].lower():
                    continue
                if celulas[0] == "" and celulas[1] == "":
                    continue
                registros.append({"Data": celulas[0], "Evolucao": celulas[1]})
        if len(registros) == 0:
            for p in doc.paragraphs:
                if p.text.strip() != "":
                    registros.append({"Data": "", "Evolucao": p.text.strip()})
        return registros

    def contar(registros):
        visitas = 0
        contatos = 0
        articulacoes = 0
        consultas = 0
        atendimentos = 0
        for r in registros:
            texto = r["Evolucao"].lower()
            if "visita" in texto:
                visitas += 1
            if "contato" in texto or "whatsapp" in texto or "ligacao" in texto or "ligação" in texto:
                contatos += 1
            if "articula" in texto:
                articulacoes += 1
            if "atendimento" in texto:
                atendimentos += 1
            if "consulta realizada" in texto or "consulta com" in texto or "consulta na" in texto or "pediatra" in texto or "contato realizado com a ubs" in texto or "atendimento na ubs" in texto:
                consultas += 1
        return visitas, contatos, articulacoes, consultas, atendimentos

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
        # Grafico do Word que acabou de subir
        st.subheader("Grafico desse relatorio")
        df_graf_word = pd.DataFrame({"Tipo": ["Visitas", "Contatos", "Articulacoes", "Consultas Saude", "Atendimentos"], "Quantidade": [visitas, contatos, articulacoes, consultas, atendimentos]})
        st.bar_chart(df_graf_word.set_index("Tipo"))
        novo_codigo = st.text_input("ID do Caso / Codigo da crianca desse relatorio", value="")
        if st.button("Adicionar e gerar Excel novo"):
            if novo_codigo.strip() == "":
                st.error("Preencha o ID do Caso, ex: SFA-SP-202542")
            else:
                nova_linha = {}
                for col in df_base.columns:
                    nova_linha[col] = ""
                nova_linha["Codigo"] = novo_codigo
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
                df_novo = df_novo.fillna("")
                st.success(f"Caso {novo_codigo} adicionado! Baixe o Excel novo abaixo.")
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    df_novo.to_excel(writer, index=False, startrow=3, sheet_name="2_Base_Casos_Criancas")
                st.download_button("Baixar Excel atualizado", data=buffer.getvalue(), file_name="dados_base_novo.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.warning("Depois de baixar, suba esse arquivo no GitHub com o nome dados_base.xlsx para virar a base oficial.")

with aba3:
    st.subheader("Relatorio Geral - Graficos")
    st.write("Aqui voce ve o resumo de todos os casos que ja foram alimentados via Word.")
    # Converte as colunas de contagem para numero
    cols = ["Total Registros", "Total Visitas", "Total Contatos", "Total Articulacoes", "Total Consultas Saude", "Total Atendimentos"]
    for c in cols:
        if c not in df_base.columns:
            df_base[c] = "0"
        df_base[c] = pd.to_numeric(df_base[c], errors="coerce").fillna(0)
    df_com_dados = df_base[df_base["Total Registros"] > 0]
    if len(df_com_dados) == 0:
        st.info("Ainda nao tem nenhum caso com dados do Word. Suba um Word na aba 2 e depois volte aqui.")
    else:
        st.write(f"Casos com relatorio Word: {len(df_com_dados)} de {len(df_base)}")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Visitas", int(df_com_dados["Total Visitas"].sum()))
        col2.metric("Total Contatos", int(df_com_dados["Total Contatos"].sum()))
        col3.metric("Total Articulacoes", int(df_com_dados["Total Articulacoes"].sum()))
        col1b, col2b, col3b = st.columns(3)
        col1b.metric("Total Consultas Saude", int(df_com_dados["Total Consultas Saude"].sum()))
        col2b.metric("Total Atendimentos", int(df_com_dados["Total Atendimentos"].sum()))
        col3b.metric("Media de Registros por caso", round(float(df_com_dados["Total Registros"].mean()), 1))

        st.subheader("1. Total de acoes realizadas")
        df_total = pd.DataFrame({"Tipo": ["Visitas", "Contatos", "Articulacoes", "Consultas Saude", "Atendimentos"], "Quantidade": [int(df_com_dados["Total Visitas"].sum()), int(df_com_dados["Total Contatos"].sum()), int(df_com_dados["Total Articulacoes"].sum()), int(df_com_dados["Total Consultas Saude"].sum()), int(df_com_dados["Total Atendimentos"].sum())]})
        st.bar_chart(df_total.set_index("Tipo"))

        st.subheader("2. Visitas por crianca")
        df_vis = df_com_dados[["Codigo", "Total Visitas"]].sort_values("Total Visitas", ascending=False)
        st.bar_chart(df_vis.set_index("Codigo"))

        st.subheader("3. Contatos por crianca")
        df_cont = df_com_dados[["Codigo", "Total Contatos"]].sort_values("Total Contatos", ascending=False)
        st.bar_chart(df_cont.set_index("Codigo"))

        st.subheader("4. Tabela resumo para copiar")
        st.dataframe(df_com_dados[["Codigo", "Total Registros", "Total Visitas", "Total Contatos", "Total Articulacoes", "Total Consultas Saude", "Total Atendimentos", "Primeiro Contato", "Ultimo Contato"]])

if st.button("Sair"):
    st.session_state.logado = False
    st.rerun()
