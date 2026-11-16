import streamlit as st
import pandas as pd
import re
import io
import os
import unicodedata
from datetime import date

st.set_page_config(page_title="SFA - Família Acolhedora", layout="wide", page_icon="🏠")

# --- INTERFACE NOVA: VERDE, LARANJA, BRANCO ---
st.markdown("""
<style>
    .stApp { background-color: #FFFFFF; }
    h1 { color: #1E7A34 !important; font-size: 36px !important; }
    h2 { color: #1E7A34 !important; border-left: 6px solid #FF7A00; padding-left: 12px; background: #FFF4E6; padding-top: 8px; padding-bottom: 8px; border-radius: 0 12px 12px 0; }
    h3 { color: #E86A00 !important; }
    .stMetric { background-color: #F0F9F1; padding: 15px; border-radius: 14px; border: 2px solid #C8E6C9; }
    .stMetric label { color: #1E7A34 !important; font-weight: bold; }
    .stButton > button { background-color: #FF7A00 !important; color: white !important; border-radius: 12px; font-weight: bold; border: none; padding: 10px 20px; }
    .stButton > button:hover { background-color: #1E7A34 !important; }
    .stDownloadButton > button { background-color: #1E7A34 !important; color: white !important; border-radius: 12px; font-weight: bold; }
    .stTextInput > div > div > input, .stTextArea > div > div > textarea { border-radius: 12px; border: 2px solid #FFB366; }
    div[data-testid="stExpander"] { border: 2px solid #FFB366; border-radius: 12px; }
</style>
""", unsafe_allow_html=True)

# --- LOGO ---
if os.path.exists("logo.png"):
    c_logo1, c_logo2, c_logo3 = st.columns([1, 2, 1])
    with c_logo2:
        st.image("logo.png", width=320)
else:
    st.markdown("<div style='text-align:center; font-size:70px;'>🏠 🧒 👧 🌱</div>", unsafe_allow_html=True)
    st.caption("Dica: Suba a imagem do logo como logo.png no GitHub para ela aparecer aqui no topo.")

st.markdown("<div style='text-align:center;'><h1>🏠 SFA - Família Acolhedora</h1><h3 style='color:#FF7A00;'>OSC em Evidência | Relatório de Impacto Social</h3><p style='font-size:18px;'>🧒 <b>Toda criança merece uma família.</b> 👨‍👩‍👧 <b>Toda família merece apoio.</b> 🌱 <b>Todo vínculo pode ser reconstruído.</b></p></div>", unsafe_allow_html=True)
st.markdown("<div style='background:#1E7A34; color:white; padding:12px; border-radius:12px; text-align:center; font-size:16px;'>💚 <b>A equipe escreve em Word.</b> 🧡 <b>O app transforma em dado.</b> 🤍 <b>Nada se perde.</b></div>", unsafe_allow_html=True)
st.markdown("---")

def divisao_segura(num, den):
    try:
        if den is None or den == 0 or pd.isna(den):
            return 0
        return float(num) / float(den)
    except:
        return 0

def formata_moeda(v):
    try:
        return f"R$ {float(v):,.2f}"
    except:
        return "R$ 0,00"

def normalizar(texto):
    if not texto:
        return ""
    t = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    return t.lower()

def ler_docx_sem_biblioteca(arquivo):
    try:
        import zipfile
        arquivo.seek(0)
        with zipfile.ZipFile(arquivo) as z:
            xml_content = z.read('word/document.xml')
            textos = re.findall(r'<w:t[^>]*>(.*?)</w:t>', xml_content.decode())
            texto2 = " ".join(textos)
            if texto2.strip():
                return texto2, "Lido ✅"
    except Exception as e:
        return "", f"Não consegui ler: {e}"
    return "", "Vazio"

MAPA_VINCULO = {"forte": 3, "bom": 3, "otimo": 3, "excelente": 3, "muito bom": 3, "reconstrucao": 2, "reconstruindo": 2, "melhorando": 2, "razoavel": 2, "regular": 2, "fragil": 1, "fraco": 1, "delicado": 1, "dificil": 1, "distante": 1, "rompendo": 0, "rompido": 0, "ausente": 0, "nenhum": 0, "sem vinculo": 0}
MAPA_EMOCIONAL = {"tranquila": 3, "tranquilo": 3, "calma": 3, "bem": 3, "feliz": 3, "animada": 3, "estavel": 2, "melhor": 2, "mais calma": 2, "segura": 2, "instavel": 1, "agitada": 1, "ansiosa": 1, "chorando": 1, "triste": 1, "nervosa": 1, "crise": 0, "em crise": 0, "muito mal": 0}
MAPA_ESCOLA = {"todos os dias": 3, "todos dias": 3, "frequentando": 3, "3 dias": 2, "3-4 dias": 2, "quase todos os dias": 2, "1-2 dias": 1, "1 dia": 1, "2 dias": 1, "pouco": 1, "nao foi": 0, "nao vai": 0, "faltou": 0, "sem escola": 0}
MAPA_EVOLUCAO = {"melhorou muito": 3, "melhorou bastante": 3, "muito melhor": 3, "melhorou um pouco": 2, "melhorou": 2, "evoluindo": 2, "progresso": 2, "avanco": 2, "estagnou": 1, "mesma coisa": 1, "sem mudanca": 1, "regrediu": 0, "piorou": 0, "pior": 0}

def achar_nota(texto_norm, mapa):
    for palavra in sorted(mapa.keys(), key=len, reverse=True):
        if palavra in texto_norm:
            return palavra, mapa[palavra]
    return "Não identificado", None

def extrair_codigo(texto):
    m = re.search(r"crian[cç]a[_\s\-]*(\d+)", texto.lower())
    if m:
        return f"Criança_{m.group(1)}"
    return "Não identificado"

def mostrar_grafico(df, coluna, titulo, emoji):
    st.subheader(f"{emoji} {titulo}")
    if coluna is None:
        st.warning(f"⚠️ Coluna não encontrada no Excel para: {titulo}")
        st.caption(f"Colunas que o app achou: {list(df.columns)}")
        return
    try:
        contagem = df[coluna].fillna("Sem informação").astype(str).str.strip().value_counts()
        if len(contagem) == 0:
            st.info("Sem dados para mostrar aqui.")
            return
        st.bar_chart(contagem, color="#FF7A00")
        st.caption(f"Total: {int(contagem.sum())} casos | Coluna usada: {coluna}")
        st.dataframe(contagem.reset_index().rename(columns={"index": titulo, coluna: "Quantidade"}), use_container_width=True)
    except Exception as e:
        st.error(f"Não consegui desenhar o gráfico: {e}")
        st.dataframe(df[coluna].value_counts(), use_container_width=True)

if "base" not in st.session_state:
    st.session_state.base = None
if "avaliacoes" not in st.session_state:
    st.session_state.avaliacoes = []
if "registros_word" not in st.session_state:
    st.session_state.registros_word = []

st.header("1️⃣ Subir Excel Base")
uploaded = st.file_uploader("Suba o Excel - Base_Dados_SFA (.xlsx)", type=["xlsx", "xls"], key="excel_base")
if uploaded:
    try:
        xls = pd.ExcelFile(uploaded)
        aba = None
        for n in xls.sheet_names:
            if any(p in n.lower() for p in ["base", "dados", "casos", "sfa"]):
                aba = n
                break
        if aba is None:
            aba = xls.sheet_names[1] if len(xls.sheet_names) > 1 else xls.sheet_names[0]
        # CORRECAO GRAFICO: Acha sozinho a linha do cabecalho de verdade - pula titulo
        df_teste = pd.read_excel(uploaded, sheet_name=aba, header=None)
        linha_cabecalho = 0
        for i, linha in df_teste.iterrows():
            texto_linha = " ".join([str(x).lower() for x in linha.values if pd.notna(x)])
            if "crian" in texto_linha or "regi" in texto_linha or "problema" in texto_linha or "desfecho" in texto_linha or "objetivo" in texto_linha:
                linha_cabecalho = i
                break
        df = pd.read_excel(uploaded, sheet_name=aba, header=linha_cabecalho).dropna(how="all")
        # Tira colunas vazias Unnamed
        df = df.loc[:, ~df.columns.astype(str).str.contains("Unnamed")]
        df = df.dropna(axis=1, how="all")
        # Limpa nomes das colunas - tira espaco extra
        df.columns = [str(c).strip() for c in df.columns]
        st.session_state.base = df
        st.success(f"✅ Base carregada: {len(df)} casos | Aba: {aba}")
        st.info(f"💡 Achei o cabeçalho na linha {linha_cabecalho + 1} do Excel")
        st.caption(f"Colunas encontradas: {list(df.columns)}")
    except Exception as e:
        st.error(f"Erro: {e}")

if st.session_state.base is None:
    st.info("👆 Suba o Excel primeiro.")
    st.stop()

df_total = st.session_state.base.copy()
def achar(palavras):
    for c in df_total.columns:
        if any(p in str(c).lower() for p in palavras):
            return c
    return None

col_codigo = achar(["cód", "cod", "criança", "crianca"])
col_regiao = achar(["região", "regiao"])
col_problema = achar(["problema"])
col_desfecho = achar(["desfecho"])
col_objetivo = achar(["objetivo"])
col_status = achar(["status", "prazo"])
col_interv = achar(["total", "interven"])
col_atendidos = achar(["atendidos", "núcleo", "nucleo"])
col_custo = achar(["custo total", "custo"])

st.header("2️⃣ 🔍 Buscar por Criança")
busca = st.text_input("Buscar", placeholder="Digite: Criança_12 ou 12", key="busca_crianca")
c2, c3, c4 = st.columns(3)
with c2:
    op_regiao = sorted(df_total[col_regiao].dropna().astype(str).unique().tolist()) if col_regiao else []
    filtro_regiao = st.selectbox("Filtrar por Região", ["Todas"] + op_regiao)
with c3:
    op_problema = sorted(df_total[col_problema].dropna().astype(str).unique().tolist()) if col_problema else []
    filtro_problema = st.selectbox("Filtrar por Problema", ["Todos"] + op_problema)
with c4:
    op_status = sorted(df_total[col_status].dropna().astype(str).unique().tolist()) if col_status else []
    filtro_status = st.selectbox("Filtrar por Status do Prazo", ["Todos"] + op_status)
if st.button("🧹 Limpar busca / Mostrar todos"):
    st.session_state["busca_crianca"] = ""
    st.rerun()

df = df_total.copy()
if busca:
    bl = str(busca).strip().lower()
    mask = pd.Series([False]*len(df))
    for col in df.columns:
        try:
            mask = mask | df[col].astype(str).str.lower().str.contains(bl, na=False)
        except:
            pass
    df = df[mask]
    if len(df) > 0:
        st.success(f"✅ Encontrei {len(df)} caso(s) com '{busca}'")
    else:
        st.warning(f"⚠️ Nenhum caso com '{busca}'")
if filtro_regiao != "Todas" and col_regiao:
    df = df[df[col_regiao].astype(str) == filtro_regiao]
if filtro_problema != "Todos" and col_problema:
    df = df[df[col_problema].astype(str) == filtro_problema]
if filtro_status != "Todos" and col_status:
    df = df[df[col_status].astype(str) == filtro_status]
st.write(f"Mostrando **{len(df)}** de **{len(df_total)}** casos.")
st.dataframe(df, use_container_width=True)
if len(df) == 0:
    st.warning("⚠️ Nenhum caso. Clique em Limpar busca.")
    st.stop()

st.header("3️⃣ Painel de Impacto")
concluidos = int((df[col_objetivo].astype(str).str.lower().str.strip() == "sim").sum()) if col_objetivo else 0
s = df[col_status].astype(str) if col_status else pd.Series([""]*len(df))
antes = int(s.str.contains("Antes", case=False, na=False).sum()) if col_status else 0
dentro = int(s.str.contains("Dentro", case=False, na=False).sum()) if col_status else 0
apos = int(s.str.contains("Após", case=False, na=False).sum()) if col_status else 0
em_curso = int(s.str.contains("Em Curso|Em acolhimento", case=False, na=False).sum()) if col_status else 0
no_prazo_perc = divisao_segura((antes+dentro), concluidos)*100
total_interv = int(pd.to_numeric(df[col_interv], errors="coerce").fillna(0).sum()) if col_interv else 0
total_atend = int(pd.to_numeric(df[col_atendidos], errors="coerce").fillna(0).sum()) if col_atendidos else 0
custo_tot = float(pd.to_numeric(df[col_custo], errors="coerce").fillna(0).sum()) if col_custo else 0
media_custo = divisao_segura(custo_tot, len(df))
m1, m2, m3, m4 = st.columns(4)
with m1: st.metric("🧒 Casos Filtrados", len(df))
with m2: st.metric("✅ Concluídos", int(concluidos))
with m3: st.metric("⏰ % No Prazo", f"{no_prazo_perc:.1f}%", f"{antes} antes + {dentro} dentro")
with m4: st.metric("💬 Intervenções", total_interv)
m5, m6, m7 = st.columns(3)
with m5: st.metric("👨‍👩‍👧 Pessoas Atendidas", total_atend)
with m6: st.metric("💰 Custo Total", formata_moeda(custo_tot))
with m7: st.metric("💰 Custo Médio / Criança", formata_moeda(media_custo))

st.header("💛 A História Por Trás dos Números")
st.markdown("**😢 Antes:** A criança chega em acolhimento, afastada da família, com medo, sem vínculo. **🤝 Durante:** A equipe faz conversas, relatórios, encaminhamentos. **🏠 Depois:** A criança volta para casa com o vínculo reconstruído, dentro do prazo. **🌱 Isso é impacto social.**")

st.header("4️⃣ Gráficos de Impacto")
st.markdown("Se algum gráfico aparecer com aviso, me manda print do aviso que eu te digo qual é o nome da coluna no seu Excel.")
g1, g2 = st.columns(2)
with g1:
    mostrar_grafico(df, col_problema, "Por Problema Inicial", "📊")
    mostrar_grafico(df, col_desfecho, "Por Desfecho", "🏠")
with g2:
    mostrar_grafico(df, col_status, "Por Status do Prazo", "⏰")
    mostrar_grafico(df, col_regiao, "Por Região", "📍")
if st.session_state.registros_word:
    st.subheader("📈 Evolução dos Relatórios em Word")
    df_gw = pd.DataFrame(st.session_state.registros_word)
    for c in ["Vínculo (nota 0-3)", "Emocional (nota 0-3)", "Escola (nota 0-3)", "Evolução (nota 0-3)"]:
        df_gw[c] = pd.to_numeric(df_gw[c], errors="coerce")
    st.bar_chart(df_gw[["Vínculo (nota 0-3)", "Emocional (nota 0-3)", "Escola (nota 0-3)", "Evolução (nota 0-3)"]].mean(), color="#1E7A34")
    st.caption("Média das notas 0-3 dos relatórios em Word já transformados.")

st.header("5️⃣ Alertas Automáticos")
alertas = []
if col_objetivo:
    sem_obj = df[df[col_objetivo].astype(str).str.strip() == ""]
    if len(sem_obj) > 0:
        alertas.append(f"⚠️ {len(sem_obj)} casos sem Objetivo Atingido.")
if em_curso > 0:
    alertas.append(f"🔔 {em_curso} casos Em Acolhimento.")
if apos > 0:
    alertas.append(f"🔔 {apos} casos Após o Prazo.")
if not alertas:
    st.success("✅ Nenhum alerta.")
else:
    for a in alertas:
        st.warning(a)

st.header("6️⃣ 📄 Automação: Relatório em Word vira Dado em Excel")
st.markdown("**Equipe escreve em Word como sempre. App lê e transforma em nota. Se ele não entender, você corrige na hora.**")
st.code("""Modelo para copiar no Word:
Criança: Criança_12
Vínculo com família de origem: em reconstrução
Estado emocional: estável
Frequência escolar: foi 3 dias na escola
Evolução: melhorou um pouco""")

col_w1, col_w2 = st.columns(2)
with col_w1:
    arquivos_word = st.file_uploader("Suba 1 ou vários relatórios em Word (.docx)", type=["docx"], accept_multiple_files=True, key="word_upload_v2")
    quem_word = st.text_input("Quem está subindo?", placeholder="Ex: Roseane", key="quem_word_v2")
with col_w2:
    data_word = st.date_input("Data dos atendimentos", value=date.today(), key="data_word_v2")
    st.write("Se o Word não for lido, use a Opção B abaixo:")

st.markdown("**Opção B - Se o Word não subir, cole o texto aqui (sempre funciona):**")
texto_colado = st.text_area("Cole aqui o texto do relatório Word", placeholder="Cole o relatório inteiro aqui... Ex: Criança_12, vínculo em reconstrução, criança estável...", height=120)
nome_colado = st.text_input("Nome para esse relatório colado", placeholder="Ex: relatorio_Crianca_12", key="nome_colado")

if st.button("🔄 Transformar em Excel", type="primary", use_container_width=True):
    itens = []
    if arquivos_word:
        for arq in arquivos_word:
            texto, status = ler_docx_sem_biblioteca(arq)
            st.info(f"📄 {arq.name} - {status}")
            if texto.strip():
                st.expander(f"👁️ Ver o que o app leu de {arq.name}").write(texto[:2000])
                itens.append((arq.name, texto))
            else:
                st.error(f"❌ {arq.name} veio vazio. Cole o texto na Opção B.")
    if texto_colado.strip():
        nome_arq = nome_colado.strip() if nome_colado.strip() else "texto_colado.docx"
        itens.append((nome_arq, texto_colado))
        st.info(f"📄 Texto colado - lido ✅ - {len(texto_colado)} caracteres")
    if not itens:
        st.error("Suba um Word ou cole o texto. Nenhum dos dois veio.")
    else:
        for nome_arq, texto in itens:
            texto_norm = normalizar(texto)
            codigo = extrair_codigo(texto)
            if codigo == "Não identificado":
                codigo = extrair_codigo(nome_arq)
            pv, nv = achar_nota(texto_norm, MAPA_VINCULO)
            pe, ne = achar_nota(texto_norm, MAPA_EMOCIONAL)
            psc, nsc = achar_nota(texto_norm, MAPA_ESCOLA)
            pev, nev = achar_nota(texto_norm, MAPA_EVOLUCAO)
            notas = [x for x in [nv, ne, nsc, nev] if x is not None]
            media = round(sum(notas)/len(notas), 2) if notas else 0
            registro = {
                "Data": str(data_word),
                "Quem subiu": quem_word,
                "Arquivo Word": nome_arq,
                "Cód. Criança": codigo,
                "Texto completo do Word": texto[:1000],
                "Vínculo (texto encontrado)": pv,
                "Vínculo (nota 0-3)": nv if nv is not None else "",
                "Emocional (texto)": pe,
                "Emocional (nota 0-3)": ne if ne is not None else "",
                "Escola (texto)": psc,
                "Escola (nota 0-3)": nsc if nsc is not None else "",
                "Evolução (texto)": pev,
                "Evolução (nota 0-3)": nev if nev is not None else "",
                "Média do atendimento (0-3)": media,
            }
            st.session_state.registros_word.append(registro)
        st.success(f"✅ {len(itens)} relatório(s) transformados!")
        st.rerun()

if st.session_state.registros_word:
    st.subheader(f"📊 Planilha Consolidada - {len(st.session_state.registros_word)} relatórios")
    df_word = pd.DataFrame(st.session_state.registros_word)
    st.dataframe(df_word, use_container_width=True)
    for col_nota in ["Vínculo (nota 0-3)", "Emocional (nota 0-3)", "Escola (nota 0-3)", "Evolução (nota 0-3)"]:
        df_word[col_nota] = pd.to_numeric(df_word[col_nota], errors="coerce")
    cm1, cm2, cm3, cm4 = st.columns(4)
    with cm1: st.metric("Média Vínculo", f"{df_word['Vínculo (nota 0-3)'].mean():.2f} / 3")
    with cm2: st.metric("Média Emocional", f"{df_word['Emocional (nota 0-3)'].mean():.2f} / 3")
    with cm3: st.metric("Média Escola", f"{df_word['Escola (nota 0-3)'].mean():.2f} / 3")
    with cm4: st.metric("Média Evolução", f"{df_word['Evolução (nota 0-3)'].mean():.2f} / 3")
    st.markdown("**✏️ Correção manual - se o app entendeu errado, corrija aqui:**")
    df_editado = st.data_editor(df_word, use_container_width=True, num_rows="dynamic", key="editor_word_v2")
    if st.button("💾 Salvar correções"):
        st.session_state.registros_word = df_editado.to_dict("records")
        st.success("✅ Correções salvas!")
        st.rerun()
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        pd.DataFrame(st.session_state.registros_word).to_excel(writer, index=False, sheet_name="Relatorios_Word")
        ws = writer.sheets["Relatorios_Word"]
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = min(max_len + 2, 50)
    buffer.seek(0)
    st.download_button("⬇️ Baixar Excel Consolidado - Estruturado (.xlsx)", buffer, "relatorios_word_para_excel_SFA.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("⬇️ Baixar tambem em CSV (para Google Sheets)", pd.DataFrame(st.session_state.registros_word).to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig"), "relatorios_word_para_excel_SFA.csv", "text/csv")
    st.info("⚠️ Para não perder: Baixe o Excel e guarde no Drive. Na próxima vez, suba ele de novo para acumular.")
    csv_antigo = st.file_uploader("Suba o CSV antigo para acumular", type=["csv"], key="csv_antigo_v2")
    if csv_antigo:
        try:
            df_ant = pd.read_csv(csv_antigo, sep=None, engine="python")
            nomes_exist = set(pd.DataFrame(st.session_state.registros_word)["Arquivo Word"].astype(str)) if len(st.session_state.registros_word) > 0 else set()
            novos_ant = df_ant[~df_ant["Arquivo Word"].astype(str).isin(nomes_exist)]
            if len(novos_ant) > 0:
                st.session_state.registros_word = st.session_state.registros_word + novos_ant.to_dict("records")
                st.success(f"✅ Acumulado +{len(novos_ant)}")
                st.rerun()
            else:
                st.info("Já estava acumulado.")
        except Exception as e:
            st.error(f"Erro: {e}")
    if st.button("🗑️ Limpar todos"):
        st.session_state.registros_word = []
        st.rerun()
else:
    st.info("Nenhum Word transformado ainda. Suba 1 Word de teste para ver a mágica.")

st.header("7️⃣ Validação Humana - 3 Avaliadoras")
nome = st.text_input("Nome da avaliadora")
dec = st.radio("Decisão", ["Aprovo", "Reprovo - precisa ajustar"], horizontal=True)
coment = st.text_area("Comentário / O que ajustar?")
if st.button("Registrar Avaliação"):
    if nome:
        st.session_state.avaliacoes.append({"Avaliadora": nome, "Decisão": dec, "Comentário": coment})
        st.success(f"✅ Avaliação de {nome} registrada: {dec}")
    else:
        st.error("Coloque o nome da avaliadora.")
if st.session_state.avaliacoes:
    df_av = pd.DataFrame(st.session_state.avaliacoes)
    st.dataframe(df_av, use_container_width=True)
    aprovos = int((df_av["Decisão"] == "Aprovo").sum())
    st.metric("Aprovações", f"{aprovos} de {len(df_av)}")
    if len(df_av) >= 3:
        if aprovos >= 2:
            st.success(f"✅ RELATÓRIO APROVADO - {aprovos} de 3 aprovaram")
        else:
            st.error(f"❌ REPROVADO - só {aprovos} de 3 aprovaram.")
    else:
        st.info(f"Faltam {3 - len(df_av)} avaliadoras.")
    if st.button("Limpar avaliações - Rodada 2"):
        st.session_state.avaliacoes = []
        st.rerun()

st.header("8️⃣ Relatório em PDF")
texto_validacao = ""
if st.session_state.avaliacoes:
    for av in st.session_state.avaliacoes:
        texto_validacao += f"- {av['Avaliadora']}: {av['Decisão']} | {av['Comentário']}\n"
else:
    texto_validacao = "Ainda sem avaliações."
texto_alertas = "\n".join(alertas) if alertas else "Nenhum alerta."
n_word = len(st.session_state.registros_word)
media_word_txt = "Sem relatórios Word ainda"
if n_word > 0:
    df_w = pd.DataFrame(st.session_state.registros_word)
    df_w["Média do atendimento (0-3)"] = pd.to_numeric(df_w["Média do atendimento (0-3)"], errors="coerce")
    media_word_txt = f"{df_w['Média do atendimento (0-3)'].mean():.2f} / 3 em {n_word} relatórios"
texto_relatorio = f"""RELATÓRIO DE IMPACTO SOCIAL - SFA
RESUMO:
- Casos analisados: {len(df)} | Concluídos: {int(concluidos)}
- % No Prazo: {no_prazo_perc:.1f}% ({antes} antes + {dentro} dentro)
- Total Intervenções: {total_interv} | Pessoas Atendidas: {total_atend}
- Custo Total: {formata_moeda(custo_tot)} | Médio: {formata_moeda(media_custo)}
AUTOMAÇÃO WORD -> EXCEL:
- Relatórios em Word transformados: {n_word}
- Média geral: {media_word_txt}
ALERTAS:
{texto_alertas}
VALIDAÇÃO:
{texto_validacao}
Fonte: Base real do Excel + relatórios Word transformados. Dados anonimizados - LGPD.
"""
st.text_area("Prévia do Relatório", texto_relatorio, height=300)
st.download_button("⬇️ Baixar Relatório em TXT (para virar PDF)", texto_relatorio.encode("utf-8"), "relatorio_impacto_sfa.txt", "text/plain")
st.download_button("⬇️ Baixar Base Filtrada em CSV", df.to_csv(index=False).encode("utf-8"), "base_filtrada_sfa.csv", "text/csv")
st.markdown("<div style='background:#FFF4E6; padding:12px; border-radius:12px; text-align:center;'>💚 <b>Família Acolhedora</b> 🧡 <b>OSC em Evidência</b> 🤍 <b>Toda criança merece uma família</b></div>", unsafe_allow_html=True)
