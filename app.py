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

def achar_coluna_codigo(df):
    for col in df.columns:
        vals = df[col].astype(str).str.strip().tolist()
        if any("SFA" in v for v in vals):
            return col
    for col in df.columns:
        nome = col.lower()
        if "id" in nome and "caso" in nome:
            return col
        if "codigo" in nome or "código" in nome:
            return col
    return df.columns[0]

def carregar_base():
    if not os.path.exists(ARQUIVO_BASE):
        st.error("Base nao encontrada.")
        st.stop()
    try:
        df = pd.read_excel(ARQUIVO_BASE, dtype=str, header=3, sheet_name=1).fillna("")
    except Exception:
        df = pd.read_excel(ARQUIVO_BASE, dtype=str, header=3).fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    col_codigo = achar_coluna_codigo(df)
    if col_codigo != "Codigo":
        df = df.rename(columns={col_codigo: "Codigo"})
    df["Codigo"] = df["Codigo"].astype(str).str.strip()
    df = df[df["Codigo"] != ""]
    df = df[df["Codigo"].str.lower() != "nan"]
    return df

df_base = carregar_base()
if len(st.session_state.novos_casos) > 0:
    df_base = pd.concat([df_base, pd.DataFrame(st.session_state.novos_casos)], ignore_index=True).fillna("")
    df_base["Codigo"] = df_base["Codigo"].astype(str).str.strip()

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
        novo_codigo = st.text_input("ID do Caso / Codigo da crianca desse relatorio - ex: SFA-SP-202542", value="")
        st.caption("⚠️ Digite o codigo completo com SFA, nao so o numero 1, 2, 3. Ex: SFA-SP-202542")
        if st.button("Adicionar e gerar Excel novo ✨"):
            if novo_codigo.strip() == "":
                st.error("Preencha o ID do Caso, ex: SFA-SP-202542")
            else:
                nova_linha = {}
                for col in df_base.columns:
                    nova_linha[col] = ""
                nova_linha["Codigo"] = str(novo_codigo).strip()
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
        df_base[c] = pd.to_numeric(df_base[c], errors="coerce").fillna(0).astype(int)
    df_com_dados = df_base[df_base["Total Registros"] > 0].copy()
    df_com_dados["Codigo"] = df_com_dados["Codigo"].astype(str).str.strip()
    df_todas = df_base.copy()
    df_todas["Codigo"] = df_todas["Codigo"].astype(str).str.strip()

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

        st.subheader("5. 📄 Baixar Relatorio em PDF - So quem tem Word")
        st.write("PDF resumido, so com as criancas que ja tem relatorio Word. 💚🧡")
        def gerar_pdf():
            buf = io.BytesIO()
            with PdfPages(buf) as pdf:
                fig = plt.figure(figsize=(8.5, 11))
                fig.patch.set_facecolor("#FFF8E7")
                ax = fig.add_axes([0, 0, 1, 1])
                ax.set_xlim(0, 10); ax.set_ylim(0, 11); ax.axis("off")
                ax.add_patch(plt.Rectangle((0, 9.8), 10, 1.2, facecolor="#2E7D32", edgecolor="none"))
                ax.text(5, 10.6, "OSC em Evidencia - SFA", ha="center", va="center", fontsize=20, weight="bold", color="white")
                ax.text(5, 10.1, "Relatorio Alegre - Criancas, Adolescentes e Familias", ha="center", va="center", fontsize=11, color="white")
                ax.text(5, 9.3, "Familias diversas, infancias felizes!", ha="center", va="center", fontsize=13, color="#2E7D32", weight="bold")
                ax.text(5, 8.9, "Cada crianca importa! Cada familia importa!", ha="center", va="center", fontsize=11, color="#E65100")
                ax.text(1.0, 8.0, "Casos com relatorio: " + str(len(df_com_dados)), ha="left", va="center", fontsize=13, color="#222222")
                ax.text(1.0, 7.5, "Total de Visitas: " + str(int(df_com_dados["Total Visitas"].sum())), ha="left", va="center", fontsize=13, color="#222222")
                ax.text(1.0, 7.0, "Total de Contatos: " + str(int(df_com_dados["Total Contatos"].sum())), ha="left", va="center", fontsize=13, color="#222222")
                ax.text(1.0, 6.5, "Total de Articulacoes: " + str(int(df_com_dados["Total Articulacoes"].sum())), ha="left", va="center", fontsize=13, color="#222222")
                ax.text(1.0, 6.0, "Total de Consultas Saude: " + str(int(df_com_dados["Total Consultas Saude"].sum())), ha="left", va="center", fontsize=13, color="#222222")
                ax.text(1.0, 5.5, "Total de Atendimentos: " + str(int(df_com_dados["Total Atendimentos"].sum())), ha="left", va="center", fontsize=13, color="#222222")
                ax.text(5, 4.7, "Feito com amor, verde e laranja!", ha="center", va="center", fontsize=12, color="#FF8C00", style="italic", weight="bold")
                ax.text(1.0, 4.0, "Codigos incluidos neste relatorio:", ha="left", va="center", fontsize=11, color="#2E7D32", weight="bold")
                y = 3.5
                for cod in df_com_dados["Codigo"].tolist():
                    ax.text(1.3, y, "- " + str(cod), ha="left", va="center", fontsize=11, color="#333333")
                    y -= 0.45
                    if y < 0.5: break
                pdf.savefig(fig); plt.close(fig)
                fig2, ax2 = plt.subplots(figsize=(8.5, 5))
                fig2.patch.set_facecolor("white")
                fig2.subplots_adjust(left=0.30, right=0.92, top=0.85, bottom=0.15)
                df_plot = df_total.sort_values("Quantidade", ascending=True)
                cores = ["#2E7D32", "#66BB6A", "#FF8C00", "#FFB74D", "#E65100"]
                ax2.barh(df_plot["Tipo"], df_plot["Quantidade"], color=cores[:len(df_plot)], height=0.55)
                ax2.set_title("Total de Acoes Realizadas", color="#2E7D32", weight="bold", fontsize=14, pad=15)
                ax2.set_xlabel("Quantidade", fontsize=11); ax2.set_ylabel("")
                ax2.grid(axis="x", linestyle="--", alpha=0.4)
                for i, v in enumerate(df_plot["Quantidade"]):
                    ax2.text(int(v) + 0.2, i, str(int(v)), va="center", ha="left", weight="bold", color="#2E7D32", fontsize=12)
                ax2.set_xlim(0, max(int(df_plot["Quantidade"].max()) + 2, 5))
                pdf.savefig(fig2); plt.close(fig2)
                fig3, ax3 = plt.subplots(figsize=(8.5, 6))
                ax3.axis("off")
                ax3.set_title("Resumo por Crianca", color="#2E7D32", weight="bold", fontsize=14, pad=20)
                dados_tabela = []
                for _, row in df_com_dados.iterrows():
                    dados_tabela.append([str(row["Codigo"]), str(int(row["Total Visitas"])), str(int(row["Total Contatos"])), str(int(row["Total Articulacoes"])), str(int(row["Total Consultas Saude"])), str(int(row["Total Atendimentos"]))])
                col_labels = ["Codigo", "Visitas", "Contatos", "Articul.", "Consultas", "Atendim."]
                tabela = ax3.table(cellText=dados_tabela, colLabels=col_labels, loc="center", cellLoc="center", colWidths=[0.35, 0.13, 0.13, 0.13, 0.13, 0.13])
                tabela.auto_set_font_size(False); tabela.set_fontsize(9); tabela.scale(1, 1.8)
                for j in range(len(col_labels)):
                    tabela[0, j].set_facecolor("#2E7D32"); tabela[0, j].set_text_props(color="white", weight="bold")
                pdf.savefig(fig3); plt.close(fig3)
            buf.seek(0)
            return buf.getvalue()
        pdf_bytes = gerar_pdf()
        st.download_button("📄 Baixar PDF - So quem tem Word 💚🧡", data=pdf_bytes, file_name="relatorio_SFA.pdf", mime="application/pdf")

    # --- NOVO: PDF DE TODAS AS CRIANCAS ---
    st.divider()
    st.subheader("6. 👨‍👩‍👧‍👦 Baixar Relatorio em PDF - TODAS as Criancas")
    st.markdown('<div class="card-alegre">🌈 <b>Relatório completo!</b> Uma página para cada criança, com todas as informações do Excel. Ideal para arquivar e enviar para a coordenação!</div>', unsafe_allow_html=True)
    st.write(f"Total de crianças que vão para o PDF: **{len(df_todas)}**")

    def gerar_pdf_todas():
        buf = io.BytesIO()
        with PdfPages(buf) as pdf:
            # CAPA
            fig = plt.figure(figsize=(8.5, 11))
            fig.patch.set_facecolor("#FFF8E7")
            ax = fig.add_axes([0, 0, 1, 1])
            ax.set_xlim(0, 10); ax.set_ylim(0, 11); ax.axis("off")
            ax.add_patch(plt.Rectangle((0, 9.8), 10, 1.2, facecolor="#2E7D32", edgecolor="none"))
            ax.text(5, 10.6, "OSC em Evidencia - SFA", ha="center", va="center", fontsize=20, weight="bold", color="white")
            ax.text(5, 10.1, "Relatorio COMPLETO - Todas as Criancas", ha="center", va="center", fontsize=11, color="white")
            ax.text(5, 9.3, "Familias diversas, infancias felizes!", ha="center", va="center", fontsize=13, color="#2E7D32", weight="bold")
            ax.text(5, 8.9, "Cada crianca importa! Cada familia importa!", ha="center", va="center", fontsize=11, color="#E65100")
            ax.text(1.0, 8.0, "Total de criancas na base: " + str(len(df_todas)), ha="left", va="center", fontsize=13, color="#222222")
            ax.text(1.0, 7.5, "Com relatorio Word: " + str(len(df_com_dados)), ha="left", va="center", fontsize=13, color="#222222")
            ax.text(1.0, 7.0, "Sem relatorio Word ainda: " + str(len(df_todas) - len(df_com_dados)), ha="left", va="center", fontsize=13, color="#222222")
            ax.text(5, 6.0, "Neste PDF, cada crianca tem", ha="center", va="center", fontsize=12, color="#333333")
            ax.text(5, 5.6, "uma pagina completa com todas", ha="center", va="center", fontsize=12, color="#333333")
            ax.text(5, 5.2, "as suas informacoes!", ha="center", va="center", fontsize=12, color="#333333")
            ax.text(5, 4.3, "Feito com amor, verde e laranja!", ha="center", va="center", fontsize=12, color="#FF8C00", style="italic", weight="bold")
            pdf.savefig(fig); plt.close(fig)

            # INDICE / LISTA DE TODAS
            fig_idx, ax_idx = plt.subplots(figsize=(8.5, 11))
            fig_idx.patch.set_facecolor("white")
            ax_idx.axis("off")
            ax_idx.set_title("Indice - Todas as Criancas", color="#2E7D32", weight="bold", fontsize=14, pad=20)
            dados_idx = []
            for i, (_, row) in enumerate(df_todas.iterrows(), start=1):
                dados_idx.append([str(i), str(row["Codigo"]), str(int(row["Total Registros"])) if "Total Registros" in row else "0", str(int(row["Total Visitas"])) if "Total Visitas" in row else "0"])
            col_labels_idx = ["N.", "Codigo", "Registros", "Visitas"]
            tabela_idx = ax_idx.table(cellText=dados_idx, colLabels=col_labels_idx, loc="center", cellLoc="center", colWidths=[0.1, 0.5, 0.2, 0.2])
            tabela_idx.auto_set_font_size(False); tabela_idx.set_fontsize(9); tabela_idx.scale(1, 1.4)
            for j in range(len(col_labels_idx)):
                tabela_idx[0, j].set_facecolor("#FF8C00"); tabela_idx[0, j].set_text_props(color="white", weight="bold")
            pdf.savefig(fig_idx); plt.close(fig_idx)

            # UMA PAGINA POR CRIANCA
            for idx, (_, linha) in enumerate(df_todas.iterrows(), start=1):
                fig_c = plt.figure(figsize=(8.5, 11))
                fig_c.patch.set_facecolor("#FFF8E7")
                ax_c = fig_c.add_axes([0, 0, 1, 1])
                ax_c.set_xlim(0, 10); ax_c.set_ylim(0, 11); ax_c.axis("off")
                # Faixa topo verde/laranja
                ax_c.add_patch(plt.Rectangle((0, 10.0), 10, 1.0, facecolor="#2E7D32", edgecolor="none"))
                ax_c.text(5, 10.6, f"Crianca {idx} de {len(df_todas)}", ha="center", va="center", fontsize=12, weight="bold", color="white")
                ax_c.text(5, 10.2, "Codigo: " + str(linha["Codigo"]), ha="center", va="center", fontsize=14, weight="bold", color="white")
                y = 9.5
                # Mostra todos os campos que tem valor
                for coluna in df_todas.columns:
                    if coluna == "Codigo":
                        continue
                    valor = str(linha[coluna]).strip()
                    if valor == "" or valor.lower() == "nan":
                        continue
                    # Quebra texto longo em 2 linhas se precisar
                    titulo = str(coluna) + ": "
                    # Se o valor for muito longo, corta para caber e continua na linha de baixo
                    if len(valor) > 90:
                        ax_c.text(0.6, y, titulo, ha="left", va="top", fontsize=10, weight="bold", color="#2E7D32")
                        y -= 0.35
                        # Quebra em pedacos de 90 caracteres
                        for k in range(0, len(valor), 90):
                            ax_c.text(0.6, y, valor[k:k+90], ha="left", va="top", fontsize=10, color="#222222")
                            y -= 0.30
                            if y < 0.5:
                                break
                    else:
                        ax_c.text(0.6, y, titulo + valor, ha="left", va="top", fontsize=10, color="#222222", weight="bold" if "Total" in str(coluna) else "normal")
                        y -= 0.35
                    if y < 0.5:
                        break
                ax_c.text(5, 0.3, "OSC em Evidencia - SFA  |  Feito com amor 💚🧡", ha="center", va="center", fontsize=9, color="#FF8C00", style="italic")
                pdf.savefig(fig_c); plt.close(fig_c)
        buf.seek(0)
        return buf.getvalue()

    pdf_todas_bytes = gerar_pdf_todas()
    st.download_button("👨‍👩‍👧‍👦 Baixar PDF - TODAS as Criancas 💚🧡", data=pdf_todas_bytes, file_name="relatorio_TODAS_criancas_SFA.pdf", mime="application/pdf")
    st.caption("Esse PDF pode ficar grande se tiver muitas crianças, é normal demorar uns segundinhos para gerar.")

if st.button("Sair"):
    st.session_state.logado = False
    st.rerun()
