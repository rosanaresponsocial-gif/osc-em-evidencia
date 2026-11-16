import streamlit as st
import pandas as pd
from datetime import date

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
st.markdown("> **Cada número aqui é uma criança que voltou para casa. Nenhum número foi inventado.**")
st.markdown("---")

def divisao_segura(num, den):
    try:
        if den is None or den == 0 or pd.isna(den):
            return 0
        return float(num) / float(den)
    except:
        return 0

def formata_moeda(valor):
    try:
        return f"R$ {float(valor):,.2f}"
    except:
        return "R$ 0,00"

# Mapeamento qualitativo -> quantitativo
MAPA = {
    "Rompendo": 0, "Frágil": 1, "Em reconstrução": 2, "Forte": 3,
    "Em crise": 0, "Instável": 1, "Estável": 2, "Tranquila": 3,
    "Não foi": 0, "Foi 1-2 dias": 1, "Foi 3-4 dias": 2, "Foi todos os dias": 3,
    "Regrediu": 0, "Estagnou": 1, "Melhorou um pouco": 2, "Melhorou muito": 3,
    "Nenhum": 0, "1 contato": 1, "2-3 contatos": 2, "Contato diário": 3,
}

if "base" not in st.session_state:
    st.session_state.base = None
if "avaliacoes" not in st.session_state:
    st.session_state.avaliacoes = []
if "passo" not in st.session_state:
    st.session_state.passo = 0
    st.session_state.dados = {}
if "registros_diarios" not in st.session_state:
    st.session_state.registros_diarios = []

# --- 1. SUBIR EXCEL ---
st.header("1️⃣ Subir Excel")
uploaded = st.file_uploader("Suba o Excel - Base_Dados_SFA (.xlsx)", type=["xlsx", "xls"])
if uploaded:
    try:
        xls = pd.ExcelFile(uploaded)
        aba_base = None
        for n in xls.sheet_names:
            if any(p in n.lower() for p in ["base", "dados", "casos", "sfa"]):
                aba_base = n
                break
        if aba_base is None:
            aba_base = xls.sheet_names[1] if len(xls.sheet_names) > 1 else xls.sheet_names[0]
        df = pd.read_excel(uploaded, sheet_name=aba_base, header=0).dropna(how="all")
        st.session_state.base = df
        st.success(f"✅ Carregado! Aba: {aba_base} | {len(df)} casos")
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

# --- 2. BUSCA ---
st.header("2️⃣ 🔍 Buscar por Criança")
c_b1, c_b2 = st.columns([3, 1])
with c_b1:
    busca = st.text_input("Buscar", placeholder="Digite: Criança_12 ou 12", key="busca_crianca", label_visibility="collapsed")
with c_b2:
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
    busca_limpa = str(busca).strip().lower()
    mask = pd.Series([False] * len(df))
    for col in df.columns:
        try:
            mask = mask | df[col].astype(str).str.lower().str.contains(busca_limpa, na=False)
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
    caso_sucesso = df[df[col_objetivo].astype(str).str.lower().str.strip() == "sim"]
    if len(caso_sucesso) > 0:
        ex = caso_sucesso.iloc[0]
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
    alertas.append(f"🔔 {em_curso} casos Em Acolhimento - acompanhar prazo.")
if apos > 0:
    alertas.append(f"🔔 {apos} casos Após o Prazo - verificar motivo.")
if not alertas:
    st.success("✅ Nenhum alerta.")
else:
    for a in alertas:
        st.warning(a)

# --- 6B. NOVA PARTE - AUTOMAÇÃO QUALITATIVO -> QUANTITATIVO ---
st.header("6️⃣ 📋 Registro Diário da Equipe - Qualitativo vira Número")
st.markdown("**A equipe atende, registra em 1 minuto no celular, e o dado já cai na planilha. Nada se perde mais.**")
st.info("Como usar: Depois de cada atendimento, a técnica abre o app no celular, escolhe a criança e clica nas opções. O sistema já calcula a nota e salva.")

with st.form("form_diario", clear_on_submit=True):
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        data_atend = st.date_input("Data do atendimento", value=date.today())
        quem = st.text_input("Quem atendeu? (nome da técnica)", placeholder="Ex: Roseane")
        cod_crianca = st.text_input("Cód. Criança", placeholder="Ex: Criança_12")
        tipo_atend = st.selectbox("Tipo de atendimento", ["Visita domiciliar", "Atendimento na sede", "Ligação / WhatsApp", "Reunião com Vara / Rede", "Acompanhamento escolar / saúde"])
    with col_d2:
        vinculo_origem = st.selectbox("1. Vínculo com família de origem", ["Rompendo", "Frágil", "Em reconstrução", "Forte"])
        vinculo_acolh = st.selectbox("2. Vínculo com família acolhedora", ["Rompendo", "Frágil", "Em reconstrução", "Forte"])
        emocional = st.selectbox("3. Estado emocional da criança", ["Em crise", "Instável", "Estável", "Tranquila"])
        escola = st.selectbox("4. Frequência escolar na semana", ["Não foi", "Foi 1-2 dias", "Foi 3-4 dias", "Foi todos os dias"])
        contato_familia = st.selectbox("5. Contato com família de origem na semana", ["Nenhum", "1 contato", "2-3 contatos", "Contato diário"])
        evolucao = st.selectbox("6. Evolução geral desde o último atendimento", ["Regrediu", "Estagnou", "Melhorou um pouco", "Melhorou muito"])
    
    obs = st.text_area("Observação qualitativa - o que aconteceu? (1 frase)", placeholder="Ex: Mãe participou da visita, criança mais calma, combinamos retorno na sexta.")
    salvar = st.form_submit_button("💾 Salvar atendimento e transformar em número", type="primary", use_container_width=True)

    if salvar:
        if not cod_crianca or not quem:
            st.error("Preencha Cód. Criança e Quem atendeu.")
        else:
            notas = [MAPA[vinculo_origem], MAPA[vinculo_acolh], MAPA[emocional], MAPA[escola], MAPA[contato_familia], MAPA[evolucao]]
            media = sum(notas) / len(notas)
            registro = {
                "Data": str(data_atend),
                "Quem atendeu": quem,
                "Cód. Criança": cod_crianca,
                "Tipo": tipo_atend,
                "Vínculo Origem (texto)": vinculo_origem,
                "Vínculo Origem (nota 0-3)": MAPA[vinculo_origem],
                "Vínculo Acolhedora (texto)": vinculo_acolh,
                "Vínculo Acolhedora (nota 0-3)": MAPA[vinculo_acolh],
                "Emocional (texto)": emocional,
                "Emocional (nota 0-3)": MAPA[emocional],
                "Escola (texto)": escola,
                "Escola (nota 0-3)": MAPA[escola],
                "Contato Família (texto)": contato_familia,
                "Contato Família (nota 0-3)": MAPA[contato_familia],
                "Evolução (texto)": evolucao,
                "Evolução (nota 0-3)": MAPA[evolucao],
                "Média do dia (0-3)": round(media, 2),
                "Observação": obs,
            }
            st.session_state.registros_diarios.append(registro)
            st.success(f"✅ Salvo! {cod_crianca} - Média do dia: {media:.2f} / 3.00 - {evolucao}")
            st.rerun()

# Mostra registros diários + média + download
if st.session_state.registros_diarios:
    st.subheader(f"📊 Registros Diários Salvos - {len(st.session_state.registros_diarios)} atendimentos")
    df_diario = pd.DataFrame(st.session_state.registros_diarios)
    st.dataframe(df_diario, use_container_width=True)
    c_m1, c_m2, c_m3 = st.columns(3)
    with c_m1: st.metric("Média Geral Vínculo Origem", f"{df_diario['Vínculo Origem (nota 0-3)'].mean():.2f} / 3")
    with c_m2: st.metric("Média Geral Emocional", f"{df_diario['Emocional (nota 0-3)'].mean():.2f} / 3")
    with c_m3: st.metric("Média Geral Evolução", f"{df_diario['Evolução (nota 0-3)'].mean():.2f} / 3")
    st.bar_chart(df_diario.groupby("Cód. Criança")["Média do dia (0-3)"].mean())
    st.download_button("⬇️ Baixar Planilha Diária em Excel / CSV - Já com notas", df_diario.to_csv(index=False).encode("utf-8"), "registros_diarios_sfa.csv", "text/csv")
    if st.button("🗑️ Limpar registros diários"):
        st.session_state.registros_diarios = []
        st.rerun()
else:
    st.info("Nenhum registro diário ainda. Faça o primeiro atendimento acima para testar.")

# --- 6. NOVO CASO ---
st.header("6️⃣ Registrar Novo Caso - Interface que Pergunta")
if st.button("➕ Começar novo caso"):
    st.session_state.passo = 1
    st.session_state.dados = {}
    st.rerun()
if st.session_state.passo >= 1:
    st.subheader(f"Passo {st.session_state.passo} de 5")
    if st.session_state.passo == 1:
        st.session_state.dados["Cód. Criança"] = st.text_input("Cód. Criança (anonimizado) Ex: Criança_36", value=st.session_state.dados.get("Cód. Criança", ""))
        st.session_state.dados["Idade (anos)"] = st.number_input("Idade", 0, 18, 5)
        if st.button("Próxima ➡️ Problema"):
            st.session_state.passo = 2
            st.rerun()
    elif st.session_state.passo == 2:
        lista_prob = sorted(df_total[col_problema].dropna().astype(str).unique().tolist()) if col_problema else ["Negligência"]
        st.session_state.dados["Problema Inicial Identificado"] = st.selectbox("Problema", lista_prob)
        st.session_state.dados["Solução Aplicada"] = st.text_area("Solução Aplicada")
        if st.button("Próxima ➡️ Intervenções"):
            st.session_state.passo = 3
            st.rerun()
    elif st.session_state.passo == 3:
        tot = 0
        for campo in ["Conversas Família Origem", "Conversas Família Acolhedora", "Intervenções Vara", "Relatórios Técnicos", "Encaminhamentos Saúde (só fluxo)"]:
            v = st.number_input(campo, min_value=0, step=1, value=0, key=campo)
            st.session_state.dados[campo] = v
            tot += v
        st.session_state.dados["Total de Intervenções Realizadas"] = tot
        st.success(f"Total: {tot}")
        if st.button("Próxima ➡️ Impacto"):
            st.session_state.passo = 4
            st.rerun()
    elif st.session_state.passo == 4:
        st.session_state.dados["Resultado/Transformação Alcançada"] = st.text_area("Resultado / Transformação")
        st.session_state.dados["Desfecho do Caso"] = st.selectbox("Desfecho", ["Reintegração Família de Origem", "Reintegração Família Extensa", "Encaminhamento Família Substituta", "Em Acolhimento (Ativo)"])
        st.session_state.dados["Objetivo Atingido (Sim / Em Andamento)"] = st.selectbox("Objetivo?", ["Sim", "Em Andamento"])
        st.session_state.dados["Status do Prazo"] = st.selectbox("Status", ["Desacolhida Antes do Prazo", "Desacolhida Dentro do Prazo", "Desacolhida Após o Prazo", "Em Acolhimento (Em Curso)"])
        if st.button("Próxima ➡️ Custo"):
            st.session_state.passo = 5
            st.rerun()
    elif st.session_state.passo == 5:
        bolsa = st.number_input("Bolsa-Auxílio R$", min_value=0.0, step=100.0)
        equipe = st.number_input("Equipe / Estrutura R$", min_value=0.0, step=100.0)
        st.session_state.dados["Custo Total por Criança (R$)"] = bolsa + equipe
        if st.button("🎉 Salvar na Base"):
            nova = pd.DataFrame([st.session_state.dados])
            st.session_state.base = pd.concat([st.session_state.base, nova], ignore_index=True)
            st.success("✅ Salvo!")
            st.session_state.passo = 0
            st.rerun()

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
n_diario = len(st.session_state.registros_diarios)
media_diario_txt = f"{pd.DataFrame(st.session_state.registros_diarios)['Média do dia (0-3)'].mean():.2f} / 3" if n_diario > 0 else "Sem registros diários ainda"

texto_relatorio = f"""RELATÓRIO DE IMPACTO SOCIAL - SFA

RESUMO:
- Casos analisados: {len(df)} | Concluídos: {int(concluidos)}
- % No Prazo: {no_prazo_perc:.1f}% ({antes} antes + {dentro} dentro)
- Total Intervenções: {total_interv} | Pessoas Atendidas: {total_atend}
- Custo Total: {formata_moeda(custo_tot)} | Médio: {formata_moeda(media_custo)}

HISTÓRIA:
Antes: Criança chega afastada. Durante: Equipe faz conversas, relatórios, encaminhamentos. Depois: Volta para casa com vínculo reconstruído.

REGISTROS DIÁRIOS DA EQUIPE - AUTOMAÇÃO QUALITATIVO -> QUANTITATIVO:
- Total de atendimentos registrados no app: {n_diario}
- Média geral de evolução nos atendimentos: {media_diario_txt}
- Cada atendimento vira nota 0-3 em vínculo, emocional, escola, contato, evolução. Nada se perde.

ALERTAS:
{texto_alertas}

VALIDAÇÃO:
{texto_validacao}

Fonte: Base real do Excel + registros diários. Dados anonimizados - LGPD.
"""
st.text_area("Prévia do Relatório", texto_relatorio, height=300)
st.download_button("⬇️ Baixar Relatório em TXT (para virar PDF)", texto_relatorio.encode("utf-8"), "relatorio_impacto_sfa.txt", "text/plain")
st.download_button("⬇️ Baixar Base Filtrada em CSV", df.to_csv(index=False).encode("utf-8"), "base_filtrada_sfa.csv", "text/csv")
st.info("Para virar PDF: Baixa o TXT > Abre no Word / Google Docs > Arquivo > Baixar como PDF. Ou Ctrl+P > Salvar como PDF.")
