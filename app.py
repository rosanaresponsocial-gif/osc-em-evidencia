import streamlit as st
import pandas as pd
import re
from datetime import date
import io

st.set_page_config(page_title="SFA - Relatório de Impacto", layout="wide")

st.markdown("""
<style>
    .main { background-color: #FFF8F0; }
    h1 { color: #6B3A2A; font-size: 32px; }
    h2 { color: #8B5A2B; border-left: 5px solid #E8A87C; padding-left: 10px; }
    .stMetric { background-color: white; padding: 15px; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
    .stButton > button { border-radius: 10px; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

st.title("🏠 SFA - Relatório de Impacto Social")
st.markdown("### OSC em Evidência | Serviço de Acolhimento em Família Acolhedora")
st.markdown("> **A equipe escreve em Word. O app transforma em dado. Nada se perde.**")
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

# --- Dicionário qualitativo -> nota ---
# O app procura essas palavras no Word
def nota_por_palavras(texto, mapa):
    t = texto.lower()
    for palavra, nota in mapa.items():
        if palavra in t:
            return palavra, nota
    return "Não identificado", None

MAPA_VINCULO = {"forte": 3, "em reconstrução": 2, "em reconstrucao": 2, "frágil": 1, "fragil": 1, "rompendo": 0, "rompido": 0, "ausente": 0}
MAPA_EMOCIONAL = {"tranquila": 3, "tranquilo": 3, "calma": 3, "estável": 2, "estavel": 2, "instável": 1, "instavel": 1, "em crise": 0, "crise": 0, "agitada": 1}
MAPA_ESCOLA = {"todos os dias": 3, "todos os dias na escola": 3, "3-4 dias": 2, "3 dias": 2, "1-2 dias": 1, "não foi": 0, "nao foi": 0, "faltou": 0}
MAPA_EVOLUCAO = {"melhorou muito": 3, "melhorou bastante": 3, "melhorou um pouco": 2, "melhorou": 2, "estagnou": 1, "estável": 1, "regrediu": 0, "piorou": 0}

def extrair_codigo(texto):
    m = re.search(r"crian[cç]a[_\s\-]*(\d+)", texto.lower())
    if m:
        return f"Criança_{m.group(1)}"
    return "Não identificado"

def ler_docx(arquivo):
    # Tenta ler com python-docx, se não tiver, lê como texto
    try:
        import docx
        doc = docx.Document(arquivo)
        textos = [p.text for p in doc.paragraphs]
        # Também pega tabelas
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    textos.append(cell.text)
        return "\n".join(textos)
    except Exception as e:
        return ""

if "base" not in st.session_state:
    st.session_state.base = None
if "avaliacoes" not in st.session_state:
    st.session_state.avaliacoes = []
if "registros_word" not in st.session_state:
    st.session_state.registros_word = []
if "passo" not in st.session_state:
    st.session_state.passo = 0
    st.session_state.dados = {}

# --- 1. SUBIR EXCEL BASE ---
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
        df = pd.read_excel(uploaded, sheet_name=aba, header=0).dropna(how="all")
        st.session_state.base = df
        st.success(f"✅ Base carregada: {len(df)} casos | Aba: {aba}")
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

# --- 2. BUSCA + FILTROS ---
st.header("2️⃣ 🔍 Buscar por Criança")
cb1, cb2 = st.columns([3, 1])
with cb1:
    busca = st.text_input("Buscar", placeholder="Digite: Criança_12 ou 12", key="busca_crianca", label_visibility="collapsed")
with cb2:
    st.button("🔍 BUSCAR", type="primary", use_container_width=True)
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

# --- 3. PAINEL ---
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
media_interv = divisao_segura(total_interv, concluidos)
media_custo = divisao_segura(custo_tot, len(df))
m1, m2, m3, m4 = st.columns(4)
with m1: st.metric("Casos Filtrados", len(df))
with m2: st.metric("Concluídos", int(concluidos))
with m3: st.metric("% No Prazo", f"{no_prazo_perc:.1f}%", f"{antes} antes + {dentro} dentro")
with m4: st.metric("Intervenções", total_interv)
m5, m6, m7 = st.columns(3)
with m5: st.metric("Pessoas Atendidas", total_atend)
with m6: st.metric("Custo Total", formata_moeda(custo_tot))
with m7: st.metric("Custo Médio / Criança", formata_moeda(media_custo))

# --- HISTÓRIA ---
st.header("💛 A História Por Trás dos Números")
st.markdown("""
**Antes:** A criança chega em acolhimento, afastada da família, com medo, sem vínculo.
**Durante:** A equipe faz conversas com a família de origem, com a família acolhedora, com a Vara, relatórios, encaminhamentos.
**Depois:** A criança volta para casa com o vínculo reconstruído, dentro do prazo.
""")
if col_objetivo:
    cs = df[df[col_objetivo].astype(str).str.lower().str.strip() == "sim"]
    if len(cs) > 0:
        ex = cs.iloc[0]
        st.success(f"**Exemplo real:** {ex[col_codigo] if col_codigo else 'Caso'} | Problema: {ex[col_problema] if col_problema else ''} | Desfecho: {ex[col_desfecho] if col_desfecho else ''} | Status: {ex[col_status] if col_status else ''}")

# --- 4. GRÁFICOS ---
st.header("4️⃣ Gráficos de Impacto")
g1, g2 = st.columns(2)
with g1:
    if col_problema:
        st.subheader("Por Problema Inicial")
        st.bar_chart(df[col_problema].value_counts())
    if col_desfecho:
        st.subheader("Por Desfecho")
        st.bar_chart(df[col_desfecho].value_counts())
with g2:
    if col_status:
        st.subheader("Por Status do Prazo")
        st.bar_chart(df[col_status].value_counts())
    if col_regiao:
        st.subheader("Por Região")
        st.bar_chart(df[col_regiao].value_counts())

# --- 5. ALERTAS ---
st.header("5️⃣ Alertas Automáticos")
alertas = []
if col_objetivo:
    sem_obj = df[df[col_objetivo].astype(str).str.strip() == ""]
    if len(sem_obj) > 0:
        alertas.append(f"⚠️ {len(sem_obj)} casos sem Objetivo Atingido.")
if col_interv:
    sem_interv = df[pd.to_numeric(df[col_interv], errors="coerce").fillna(0) == 0]
    if len(sem_interv) > 0:
        alertas.append(f"⚠️ {len(sem_interv)} casos com 0 intervenções.")
if col_custo:
    sem_custo = df[pd.to_numeric(df[col_custo], errors="coerce").fillna(0) == 0]
    if len(sem_custo) > 0:
        alertas.append(f"⚠️ {len(sem_custo)} casos sem custo.")
if em_curso > 0:
    alertas.append(f"🔔 {em_curso} casos Em Acolhimento.")
if apos > 0:
    alertas.append(f"🔔 {apos} casos Após o Prazo.")
if not alertas:
    st.success("✅ Nenhum alerta.")
else:
    for a in alertas:
        st.warning(a)

# --- 6. AUTOMAÇÃO WORD -> EXCEL - A PARTE NOVA ---
st.header("6️⃣ 📄 Automação: Relatório em Word vira Dado em Excel")
st.markdown("""
**Como a equipe usa no dia a dia:**
1. Escreve o relatório em Word como sempre, com texto livre
2. Sobe o Word aqui embaixo
3. O app lê, acha a criança, acha as palavras-chave e já transforma em nota 0-3
4. Você baixa o Excel consolidado - ele vira sua base oficial, não some mais
""")
st.info("💡 Dica para a equipe: No Word, escreva sempre frases como: 'Vínculo forte', 'Criança estável', 'Foi todos os dias na escola', 'Melhorou muito'. Assim o app entende 100%.")

# Modelo de Word para copiar
st.markdown("**Modelo para copiar e colar no Word da equipe:**")
st.code("""Relatório de Atendimento - SFA
Criança: Criança_12
Data: 16/11/2026
Técnica: Roseane

Vínculo com família de origem: em reconstrução
Vínculo com família acolhedora: forte
Estado emocional: estável
Frequência escolar: foi 3 dias na escola
Evolução: melhorou um pouco

Observação: Mãe participou da visita, criança mais calma, combinamos retorno na sexta.""")

col_w1, col_w2 = st.columns(2)
with col_w1:
    arquivos_word = st.file_uploader("Suba 1 ou vários relatórios em Word (.docx)", type=["docx"], accept_multiple_files=True, key="word_upload")
    quem_word = st.text_input("Quem está subindo os relatórios?", placeholder="Ex: Roseane", key="quem_word")
with col_w2:
    data_word = st.date_input("Data dos atendimentos", value=date.today(), key="data_word")
    st.write("O app vai procurar no texto:")
    st.write("- Criança_XX")
    st.write("- Vínculo: forte / em reconstrução / frágil / rompendo")
    st.write("- Emocional: tranquila / estável / instável / em crise")
    st.write("- Escola: todos os dias / 3 dias / 1-2 dias / não foi")
    st.write("- Evolução: melhorou muito / melhorou um pouco / estagnou / regrediu")

if st.button("🔄 Transformar Word em Excel", type="primary", use_container_width=True):
    if not arquivos_word:
        st.error("Suba pelo menos 1 arquivo Word.")
    else:
        novos = 0
        for arq in arquivos_word:
            texto = ler_docx(arq)
            if not texto.strip():
                st.warning(f"⚠️ {arq.name} veio vazio ou não consegui ler. Tente salvar como .docx de novo.")
                continue
            codigo = extrair_codigo(texto)
            # Se não achou código no texto, tenta no nome do arquivo
            if codigo == "Não identificado":
                codigo = extrair_codigo(arq.name)
            pv, nv = nota_por_palavras(texto, MAPA_VINCULO)
            pe, ne = nota_por_palavras(texto, MAPA_EMOCIONAL)
            psc, nsc = nota_por_palavras(texto, MAPA_ESCOLA)
            pev, nev = nota_por_palavras(texto, MAPA_EVOLUCAO)
            notas = [x for x in [nv, ne, nsc, nev] if x is not None]
            media = round(sum(notas)/len(notas), 2) if notas else 0
            registro = {
                "Data": str(data_word),
                "Quem subiu": quem_word,
                "Arquivo Word": arq.name,
                "Cód. Criança": codigo,
                "Texto completo do Word": texto[:800],
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
            novos += 1
        st.success(f"✅ {novos} relatório(s) transformados em dados!")
        st.rerun()

# Mostra tabela consolidada + correção manual + download
if st.session_state.registros_word:
    st.subheader(f"📊 Planilha Consolidada - {len(st.session_state.registros_word)} relatórios transformados")
    df_word = pd.DataFrame(st.session_state.registros_word)
    st.dataframe(df_word, use_container_width=True)
    # Médias
    for col_nota in ["Vínculo (nota 0-3)", "Emocional (nota 0-3)", "Escola (nota 0-3)", "Evolução (nota 0-3)"]:
        df_word[col_nota] = pd.to_numeric(df_word[col_nota], errors="coerce")
    cm1, cm2, cm3, cm4 = st.columns(4)
    with cm1: st.metric("Média Vínculo", f"{df_word['Vínculo (nota 0-3)'].mean():.2f} / 3")
    with cm2: st.metric("Média Emocional", f"{df_word['Emocional (nota 0-3)'].mean():.2f} / 3")
    with cm3: st.metric("Média Escola", f"{df_word['Escola (nota 0-3)'].mean():.2f} / 3")
    with cm4: st.metric("Média Evolução", f"{df_word['Evolução (nota 0-3)'].mean():.2f} / 3")
    st.bar_chart(df_word.groupby("Cód. Criança")["Média do atendimento (0-3)"].mean())
    st.markdown("**Correção manual:** Se o app entendeu alguma palavra errado, corrija aqui antes de baixar. O Excel final já sai corrigido.")
    df_editado = st.data_editor(df_word, use_container_width=True, num_rows="dynamic", key="editor_word")
    # Salva edição
    if st.button("💾 Salvar correções"):
        st.session_state.registros_word = df_editado.to_dict("records")
        st.success("✅ Correções salvas!")
        st.rerun()
    # Download - ESSA É A PERSISTÊNCIA: baixa e vira base oficial
    csv_bytes = pd.DataFrame(st.session_state.registros_word).to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Baixar Excel Consolidado - Word virou Excel (CSV)", csv_bytes, "relatorios_word_para_excel_sfa.csv", "text/csv")
    st.info("⚠️ IMPORTANTE - Para não perder: Baixe esse CSV e guarde no Drive. Na próxima vez, suba ele junto com os novos Words. Assim ele acumula e nunca some.")
    # Upload de CSV antigo para acumular
    st.markdown("**Acumular com planilha antiga:**")
    csv_antigo = st.file_uploader("Suba o CSV que você baixou da última vez para acumular", type=["csv"], key="csv_antigo")
    if csv_antigo:
        try:
            df_ant = pd.read_csv(csv_antigo)
            # Evita duplicar pelo nome do arquivo
            nomes_existentes = set(pd.DataFrame(st.session_state.registros_word)["Arquivo Word"].astype(str)) if len(st.session_state.registros_word) > 0 else set()
            novos_ant = df_ant[~df_ant["Arquivo Word"].astype(str).isin(nomes_existentes)]
            if len(novos_ant) > 0:
                st.session_state.registros_word = st.session_state.registros_word + novos_ant.to_dict("records")
                st.success(f"✅ Acumulado! +{len(novos_ant)} registros antigos. Total agora: {len(st.session_state.registros_word)}")
                st.rerun()
            else:
                st.info("Já estava tudo acumulado, nenhum registro novo.")
        except Exception as e:
            st.error(f"Erro ao acumular: {e}")
    if st.button("🗑️ Limpar todos os registros Word"):
        st.session_state.registros_word = []
        st.rerun()
else:
    st.info("Nenhum Word transformado ainda. Suba 1 Word de teste para ver a mágica.")

# --- 7. VALIDAÇÃO ---
st.header("7️⃣ Validação Humana - 3 Avaliadoras")
st.write("Passa com **2 de 3 aprovações**. Você organiza, não avalia.")
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

# --- 8. RELATÓRIO ---
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
- Relatórios em Word transformados em dados: {n_word}
- Média geral dos atendimentos: {media_word_txt}
- Equipe escreve em Word, app transforma em nota 0-3. Nada se perde.

ALERTAS:
{texto_alertas}

VALIDAÇÃO:
{texto_validacao}

Fonte: Base real do Excel + relatórios Word transformados. Dados anonimizados - LGPD.
"""
st.text_area("Prévia do Relatório", texto_relatorio, height=300)
st.download_button("⬇️ Baixar Relatório em TXT (para virar PDF)", texto_relatorio.encode("utf-8"), "relatorio_impacto_sfa.txt", "text/plain")
st.download_button("⬇️ Baixar Base Filtrada em CSV", df.to_csv(index=False).encode("utf-8"), "base_filtrada_sfa.csv", "text/csv")
st.info("Para virar PDF: Baixa o TXT > Abre no Word / Google Docs > Arquivo > Baixar como PDF. Ou Ctrl+P > Salvar como PDF.")
