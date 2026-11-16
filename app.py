import os
import io
import streamlit as st
import pandas as pd
from docx import Document
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

ARQUIVO_BASE = "dados_base.xlsx"
USUARIO_CORRETO = "sfaevidencia"
SENHA_CORRETA = "Sfa2026"

st.set_page_config(page_title="SFA - OSC em Evidencia", page_icon="🧒", layout="centered")

# --- INTERFACE BONITA - VERDE + LARANJA ---
st.markdown("""
<style>
.stApp { background: linear-gradient(180deg, #FFF8E7 0%, #E8F5E9 100%); }
h1, h2, h3 { color: #2E7D32 !important; }
div[data-testid="stMetric"] { background: white; border: 2px solid #FF8C00; border-radius: 15px; padding: 10px; box-shadow: 2px 2px 8px #FFE0B2; }
div[data-testid="stMetric"] label { color: #E65100 !important; font-weight: bold; }
div[data-testid="stMetric"] div { color: #2E7D32 !important; }
.stButton>button { background-color: #FF8C00 !important; color: white !important; border-radius: 20px !important; border: none !important; font-weight: bold; }
.stDownloadButton>button { background-color: #2E7D32 !important; color: white !important; border-radius: 20px !important; font-weight: bold; }
.topo { background: linear-gradient(90deg, #2E7D32, #66BB6A, #FF8C00, #FFB74D); padding: 18px; border-radius: 20px; text-align: center; color: white; margin-bottom: 10px; }
.topo h1 { color: white !important; margin: 0; font-size: 28px; }
.topo p { color: white; margin: 5px 0 0 0; font-size: 15px; }
.card-alegre { background: white; border-left: 6px solid #FF8C00; border-radius: 12px; padding: 12px; margin: 8px 0; }
</style>
""", unsafe_allow_html=True)

if "logado" not in st.session_state:
    st.session_state.logado = False
if "novos_casos" not in st.session_state:
    st.session_state.novos_casos = []

if not st.session_state.logado:
    st.markdown('<div class="topo"><h1>🧒👧🏽👦🏻 OSC em Evidencia 👧🏾🧑🏽‍🤝‍🧑🏻</h1><p>👨‍👩‍👧‍👦👩‍👧🏽👨‍👧🏻👩‍👩‍👧 Famílias diversas, infâncias felizes 💚🧡</p></div>', unsafe_allow_html=True)
    st.markdown('<div class="card-alegre">🌈 <b>Um lugar alegre para cuidar de cada criança!</b><br>👶 Crianças &nbsp; 🧒 Adolescentes &nbsp; 👨‍👩‍👧 Famílias &nbsp; 💚 Acolhimento</div>', unsafe_allow_html=True)
    usuario = st.text_input("Usuario")
    senha = st.text_input("Senha", type="password")
    if st.button("Entrar 💚"):
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
if len(st.session_state.novos_casos) > 0:
    df_base = pd.concat([df_base, pd.DataFrame(st.session_state.novos_casos)], ignore_index=True).fillna("")

st.markdown('<div class="topo"><h1>🧒👧🏽 OSC em Evidencia - SFA 👦🏻👧🏾</h1><p>👨‍👩‍👧‍👦 👩‍👧🏽 👨‍👧🏻 👩‍👩‍👧‍👦 👨‍👨‍👧 Famílias diversas • Crianças • Adolescentes • Acolhimento 💚🧡</p><p>🌈 Cada criança importa! Cada família importa!</p></div>', unsafe_allow_html=True)
st.markdown('<div class="card-alegre">👶🧒👧 <b>Bem-vinda!</b> Aqui a gente cuida com amor, alegria e respeito a diversidade. 💚🧡🌈</div>', unsafe_allow_html=True)
st.write(f"👧🏽 Total de criancas na base: **{len(df_base)}**")
if len(st.session_state.novos_casos) > 0:
    st.info(f"✨ Voce ja adicionou {len(st.session_state.novos_casos)} caso(s) via Word nessa sessao.")

aba1, aba2, aba3 = st.tabs(["🔍 Consultar", "📄 Adicionar via Word", "📊 Relatorio com Graficos + PDF"])

with aba1:
    st.markdown("### 🔍 Consultar uma criança")
    codigo = st.selectbox("Escolha o codigo da crianca:", [""] + sorted(df_base["Codigo"].tolist()))
    if codigo != "":
        linha = df_base[df_base["Codigo"] == codigo].iloc[0]
        st.subheader(f"💚 Avaliacao - {codigo} 🧒")
        for coluna in df_base.columns:
            if coluna == "Codigo":
                continue
            valor = str(linha[coluna]).strip()
            if valor != "" and valor.lower() != "nan":
                st.write(f"**{coluna}:** {valor}")

with aba2:
    st.markdown("### 📄 Adicionar relatorio via Word")
    st.write("Suba o Word com a tabela Data / Evolucao. O app vai ler e contar tudo sozinho! ✨")
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
        visitas = 0; contatos = 0; articulacoes = 0; consultas = 0; atendimentos = 0
        for r in registros:
            texto = r["Evolucao"].lower()
            if "visita" in texto: visitas += 1
            if "contato" in texto or "whatsapp" in texto or "ligacao" in texto or "ligação" in texto: contatos += 1
            if "articula" in texto: articulacoes += 1
            if "atendimento" in texto: atendimentos += 1
            if "consulta realizada" in texto or "consulta com" in texto or "consulta na" in texto or "pediatra" in texto or "contato realizado com a ubs" in texto or "atendimento na ubs" in texto: consultas += 1
        return visitas, contatos, articulacoes, consultas, atendimentos

    if arquivo_word is not None:
        registros = ler_tabela_word(arquivo_word)
        st.success(f"🎉 Achei {len(registros)} registros no Word!")
        st.dataframe(pd.DataFrame(registros))
        visitas, contatos, articulacoes, consultas, atendimentos = contar(registros)
        datas = [r["Data"] for r in registros if r["Data"] != ""]
        primeiro = datas[0] if len(datas) > 0 else ""
        ultimo = datas[-1] if len(datas) > 0 else ""
        resumo = " | ".join([r["Evolucao"][:120] for r in registros[:5]])
        st.write("**O que o app contou sozinho:**")
        st.write(f"- 🧡 Total de registros: {len(registros)}")
        st.write(f"- 🏠 Total de Visitas: {visitas}")
        st.write(f"- 💬 Total de Contatos: {contatos}")
        st.write(f"- 🤝 Total de Articulacoes: {articulacoes}")
        st.write(f"- 🩺 Total de Consultas / Saude: {consultas}")
        st.write(f"- 💚 Total de Atendimentos: {atendimentos}")
        st.subheader("📊 Grafico desse relatorio")
        df_graf_word = pd.DataFrame({"Tipo": ["Visitas", "Contatos", "Articulacoes", "Consultas Saude", "Atendimentos"], "Quantidade": [visitas, contatos, articulacoes, consultas, atendimentos]})
        st.bar_chart(df_graf_word.set_index("Tipo"))
        novo_codigo = st.text_input("ID do Caso / Codigo da crianca desse relatorio", value="")
        if st.button("Adicionar e gerar Excel novo ✨"):
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
                st.session_state.novos_casos.append(nova_linha)
                df_novo = pd.concat([df_base, pd.DataFrame([nova_linha])], ignore_index=True).fillna("")
                st.success(f"Caso {novo_codigo} adicionado! Agora ele ja aparece no Relatorio. 💚")
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    df_novo.to_excel(writer, index=False, startrow=3, sheet_name="2_Base_Casos_Criancas")
                st.download_button("Baixar Excel atualizado 📥", data=buffer.getvalue(), file_name="dados_base_novo.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.warning("Para salvar para sempre: baixe e suba esse arquivo no GitHub com o nome dados_base.xlsx")

with aba3:
    st.markdown("### 📊 Relatorio Geral - Graficos")
    st.markdown('<div class="card-alegre">🌈 <b>Relatório alegre da OSC!</b> 💚🧡 Aqui está tudo que fizemos com amor pelas crianças e famílias!</div>', unsafe_allow_html=True)
    cols = ["Total Registros", "Total Visitas", "Total Contatos", "Total Articulacoes", "Total Consultas Saude", "Total Atendimentos"]
    for c in cols:
        if c not in df_base.columns:
            df_base[c] = "0"
        df_base[c] = pd.to_numeric(df_base[c], errors="coerce").fillna(0)
    df_com_dados = df_base[df_base["Total Registros"] > 0]
    if len(df_com_dados) == 0:
        st.info("Ainda nao tem nenhum caso com dados do Word. Suba um Word na aba 2, digite o Codigo e clique em Adicionar. Depois volte aqui. 💚")
    else:
        st.write(f"👨‍👩‍👧‍👦 Casos com relatorio Word: **{len(df_com_dados)}** de {len(df_base)}")
        col1, col2, col3 = st.columns(3)
        col1.metric("🏠 Total Visitas", int(df_com_dados["Total Visitas"].sum()))
        col2.metric("💬 Total Contatos", int(df_com_dados["Total Contatos"].sum()))
        col3.metric("🤝 Total Articulacoes", int(df_com_dados["Total Articulacoes"].sum()))
        col1b, col2b, col3b = st.columns(3)
        col1b.metric("🩺 Total Consultas Saude", int(df_com_dados["Total Consultas Saude"].sum()))
        col2b.metric("💚 Total Atendimentos", int(df_com_dados["Total Atendimentos"].sum()))
        col3b.metric("📝 Media de Registros", round(float(df_com_dados["Total Registros"].mean()), 1))

        st.subheader("1. 🧡 Total de acoes realizadas 💚")
        df_total = pd.DataFrame({"Tipo": ["Visitas", "Contatos", "Articulacoes", "Consultas Saude", "Atendimentos"], "Quantidade": [int(df_com_dados["Total Visitas"].sum()), int(df_com_dados["Total Contatos"].sum()), int(df_com_dados["Total Articulacoes"].sum()), int(df_com_dados["Total Consultas Saude"].sum()), int(df_com_dados["Total Atendimentos"].sum())]})
        st.bar_chart(df_total.set_index("Tipo"))

        st.subheader("2. 🏠 Visitas por crianca")
        df_vis = df_com_dados[["Codigo", "Total Visitas"]].sort_values("Total Visitas", ascending=False)
        st.bar_chart(df_vis.set_index("Codigo"))

        st.subheader("3. 💬 Contatos por crianca")
        df_cont = df_com_dados[["Codigo", "Total Contatos"]].sort_values("Total Contatos", ascending=False)
        st.bar_chart(df_cont.set_index("Codigo"))

        st.subheader("4. 📋 Tabela resumo")
        st.dataframe(df_com_dados[["Codigo", "Total Registros", "Total Visitas", "Total Contatos", "Total Articulacoes", "Total Consultas Saude", "Total Atendimentos", "Primeiro Contato", "Ultimo Contato"]])

        # --- GERAR PDF ---
        st.subheader("5. 📄 Baixar Relatorio em PDF")
        st.write("Clique para baixar um PDF lindo, verde e laranja, pronto para enviar para a coordenacao! 💚🧡")
        def gerar_pdf():
            buf = io.BytesIO()
            with PdfPages(buf) as pdf:
                # Pagina 1 - Capa + Grafico
                fig, ax = plt.subplots(figsize=(8, 6))
                fig.patch.set_facecolor("#FFF8E7")
                ax.axis("off")
                ax.text(0.5, 0.95, "OSC em Evidencia - SFA", ha="center", va="top", fontsize=20, weight="bold", color="#2E7D32", transform=ax.transAxes)
                ax.text(0.5, 0.88, "Relatorio Alegre - Criancas, Adolescentes e Familias", ha="center", va="top", fontsize=12, color="#E65100", transform=ax.transAxes)
                ax.text(0.5, 0.82, "Familias diversas, infancias felizes!", ha="center", va="top", fontsize=11, color="#2E7D32", transform=ax.transAxes)
                ax.text(0.1, 0.72, f"Casos com relatorio: {len(df_com_dados)}", fontsize=12, color="#333333", transform=ax.transAxes)
                ax.text(0.1, 0.66, f"Total Visitas: {int(df_com_dados['Total Visitas'].sum())}", fontsize=12, color="#333333", transform=ax.transAxes)
                ax.text(0.1, 0.60, f"Total Contatos: {int(df_com_dados['Total Contatos'].sum())}", fontsize=12, color="#333333", transform=ax.transAxes)
                ax.text(0.1, 0.54, f"Total Articulacoes: {int(df_com_dados['Total Articulacoes'].sum())}", fontsize=12, color="#333333", transform=ax.transAxes)
                ax.text(0.1, 0.48, f"Total Consultas Saude: {int(df_com_dados['Total Consultas Saude'].sum())}", fontsize=12, color="#333333", transform=ax.transAxes)
                ax.text(0.1, 0.42, f"Total Atendimentos: {int(df_com_dados['Total Atendimentos'].sum())}", fontsize=12, color="#333333", transform=ax.transAxes)
                ax.text(0.5, 0.15, "Feito com amor, verde e laranja!", ha="center", fontsize=11, color="#FF8C00", style="italic", transform=ax.transAxes)
                pdf.savefig(fig)
                plt.close(fig)
                # Pagina 2 - Grafico de barras
                fig2, ax2 = plt.subplots(figsize=(8, 4))
                fig2.patch.set_facecolor("white")
                cores = ["#2E7D32", "#FF8C00", "#66BB6A", "#FFB74D", "#E65100"]
                ax2.bar(df_total["Tipo"], df_total["Quantidade"], color=cores)
                ax2.set_title("Total de Acoes Realizadas", color="#2E7D32", weight="bold")
                ax2.set_ylabel("Quantidade")
                plt.xticks(rotation=15)
                for i, v in enumerate(df_total["Quantidade"]):
                    ax2.text(i, v + 0.5, str(v), ha="center", weight="bold", color="#2E7D32")
                pdf.savefig(fig2)
                plt.close(fig2)
                # Pagina 3 - Tabela por crianca
                fig3, ax3 = plt.subplots(figsize=(8, 5))
                ax3.axis("off")
                ax3.set_title("Resumo por Crianca", color="#2E7D32", weight="bold", fontsize=14)
                dados_tabela = df_com_dados[["Codigo", "Total Visitas", "Total Contatos", "Total Articulacoes", "Total Consultas Saude"]].values.tolist()
                col_labels = ["Codigo", "Visitas", "Contatos", "Articulacoes", "Consultas"]
                tabela = ax3.table(cellText=dados_tabela, colLabels=col_labels, loc="center", cellLoc="center")
                tabela.auto_set_font_size(False)
                tabela.set_fontsize(8)
                tabela.scale(1, 1.5)
                pdf.savefig(fig3)
                plt.close(fig3)
            buf.seek(0)
            return buf.getvalue()

        pdf_bytes = gerar_pdf()
        st.download_button("📄 Baixar Relatorio em PDF 💚🧡", data=pdf_bytes, file_name="relatorio_SFA.pdf", mime="application/pdf")

if st.button("Sair"):
    st.session_state.logado = False
    st.rerun()
