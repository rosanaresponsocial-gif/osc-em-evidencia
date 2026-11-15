import streamlit as st
import pandas as pd

st.set_page_config(page_title="SFA - Relatório de Impacto", layout="wide")
st.title("🏠 SFA - Relatório de Impacto Social")
st.write("Suba o Excel. Filtre, busque, veja gráficos, receba alertas, valide e baixe o relatório.")

# --- FUNÇÃO SEGURA - NUNCA DIVIDE POR ZERO ---
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

# --- SESSÃO ---
if "base" not in st.session_state:
    st.session_state.base = None
if "avaliacoes" not in st.session_state:
    st.session_state.avaliacoes = []
if "passo" not in st.session_state:
    st.session_state.passo = 0
    st.session_state.dados = {}

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
        st.success(f"✅ Carregado! Aba: {aba_base} | {len(df)} casos | {len(df.columns)} colunas")
    except Exception as e:
        st.error(f"Erro ao ler Excel: {e}")

if st.session_state.base is None:
    st.info("👆 Suba o Excel primeiro.")
    st.stop()

df_total = st.session_state.base.copy()

def achar(palavras):
    for c in df_total.columns:
        cl = str(c).lower()
        if any(p in cl for p in palavras):
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

# --- 2. FILTROS + BUSCA ---
st.header("2️⃣ Filtros + Busca por Criança")
c1, c2, c3, c4 = st.columns(4)
with c1:
    busca = st.text_input("🔍 Buscar por Cód. Criança", placeholder="Ex: Criança_12")
with c2:
    op_regiao = sorted(df_total[col_regiao].dropna().astype(str).unique().tolist()) if col_regiao else []
    filtro_regiao = st.selectbox("Filtrar por Região", ["Todas"] + op_regiao)
with c3:
    op_problema = sorted(df_total[col_problema].dropna().astype(str).unique().tolist()) if col_problema else []
    filtro_problema = st.selectbox("Filtrar por Problema", ["Todos"] + op_problema)
with c4:
    op_status = sorted(df_total[col_status].dropna().astype(str).unique().tolist()) if col_status else []
    filtro_status = st.selectbox("Filtrar por Status do Prazo", ["Todos"] + op_status)

df = df_total.copy()
if busca and col_codigo:
    df = df[df[col_codigo].astype(str).str.contains(busca, case=False, na=False)]
if filtro_regiao != "Todas" and col_regiao:
    df = df[df[col_regiao].astype(str) == filtro_regiao]
if filtro_problema != "Todos" and col_problema:
    df = df[df[col_problema].astype(str) == filtro_problema]
if filtro_status != "Todos" and col_status:
    df = df[df[col_status].astype(str) == filtro_status]

st.write(f"Mostrando **{len(df)}** de **{len(df_total)}** casos após filtros.")
st.dataframe(df, use_container_width=True)

if len(df) == 0:
    st.warning("⚠️ Nenhum caso com esse filtro / busca. Limpe a busca ou mude o filtro para ver os dados.")
    st.stop()

# --- 3. PAINEL - COM DIVISÃO SEGURA ---
st.header("3️⃣ Painel de Impacto - Da base filtrada")

# Concluídos - conta Sim, ignorando maiúsculo/minúsculo e espaços
concluidos = 0
if col_objetivo:
    concluidos = int((df[col_objetivo].astype(str).str.lower().str.strip() == "sim").sum())

s = df[col_status].astype(str) if col_status else pd.Series([""] * len(df))
antes = int(s.str.contains("Antes", case=False, na=False).sum()) if col_status else 0
dentro = int(s.str.contains("Dentro", case=False, na=False).sum()) if col_status else 0
apos = int(s.str.contains("Após", case=False, na=False).sum()) if col_status else 0
em_curso = int(s.str.contains("Em Curso|Em acolhimento", case=False, na=False).sum()) if col_status else 0

# TODAS AS DIVISÕES AGORA USAM divisao_segura - NUNCA QUEBRA
no_prazo_perc = divisao_segura((antes + dentro), concluidos) * 100
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
        alertas.append(f"⚠️ {len(sem_obj)} casos sem Objetivo Atingido preenchido.")
if col_interv:
    sem_interv = df[pd.to_numeric(df[col_interv], errors="coerce").fillna(0) == 0]
    if len(sem_interv) > 0:
        alertas.append(f"⚠️ {len(sem_interv)} casos com 0 intervenções - verificar se falta alimentar.")
if col_custo:
    sem_custo = df[pd.to_numeric(df[col_custo], errors="coerce").fillna(0) == 0]
    if len(sem_custo) > 0:
        alertas.append(f"⚠️ {len(sem_custo)} casos sem custo - relatório ficará sem custo médio.")
if em_curso > 0:
    alertas.append(f"🔔 {em_curso} casos Em Acolhimento - acompanhar prazo.")
if apos > 0:
    alertas.append(f"🔔 {apos} casos Desacolhidos Após o Prazo - verificar motivo do atraso.")
if not alertas:
    st.success("✅ Nenhum alerta. Base completa nos filtros atuais.")
else:
    for a in alertas:
        st.warning(a)

# --- 6. NOVO CASO - INTERFACE QUE PERGUNTA ---
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
            st.success("✅ Salvo! Painel atualizado.")
            st.session_state.passo = 0
            st.rerun()

# --- 7. VALIDAÇÃO ---
st.header("7️⃣ Validação Humana - 3 Avaliadoras (além da organizadora)")
st.write("Cada avaliadora avalia o relatório filtrado. Passa com **2 de 3 aprovações**. Você organiza, não avalia.")
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
            st.success(f"✅ RELATÓRIO APROVADO - {aprovos} de 3 aprovaram (regra 2 de 3)")
        else:
            st.error(f"❌ REPROVADO - só {aprovos} de 3 aprovaram. Ajustar e fazer Rodada 2.")
    else:
        st.info(f"Faltam {3 - len(df_av)} avaliadoras para completar 3. Faltam 3 pessoas além de você.")
    if st.button("Limpar avaliações - Começar Rodada 2"):
        st.session_state.avaliacoes = []
        st.rerun()

# --- 8. RELATÓRIO PDF / TXT ---
st.header("8️⃣ Relatório em PDF")
st.write("Baixa o relatório em texto. No navegador, você imprime como PDF: Ctrl+P > Salvar como PDF.")

texto_validacao = ""
if st.session_state.avaliacoes:
    for av in st.session_state.avaliacoes:
        texto_validacao += f"- {av['Avaliadora']}: {av['Decisão']} | {av['Comentário']}\n"
else:
    texto_validacao = "Ainda sem avaliações. Precisa de 3 avaliadoras além da organizadora."

texto_alertas = "\n".join(alertas) if alertas else "Nenhum alerta."

texto_relatorio = f"""RELATÓRIO DE IMPACTO SOCIAL - SFA
Serviço de Acolhimento em Família Acolhedora

RESUMO:
- Casos analisados (filtro atual): {len(df)}
- Concluídos com objetivo atingido: {int(concluidos)}
- % No Prazo (Antes + Dentro): {no_prazo_perc:.1f}% ({antes} antes + {dentro} dentro)
- Total de Intervenções: {total_interv}
- Pessoas Atendidas: {total_atend}
- Custo Total: {formata_moeda(custo_tot)} | Custo Médio por Criança: {formata_moeda(media_custo)}

PERGUNTAS DO GUIA:
1. Intervenções até objetivo: {total_interv} total, média {media_interv:.1f} por caso concluído
2. Antes do prazo: {antes} | Dentro do prazo: {dentro} | Após: {apos} | Em curso: {em_curso}
3. Transformações: {int(concluidos)} casos com objetivo atingido
4. Custo por criança: {formata_moeda(media_custo)} em média

ALERTAS:
{texto_alertas}

VALIDAÇÃO:
{texto_validacao}

Fonte: Base de dados real do Excel, sem número inventado. Dados anonimizados - LGPD.
"""

st.text_area("Prévia do Relatório para PDF - copie ou baixe", texto_relatorio, height=300)
st.download_button("⬇️ Baixar Relatório em TXT (para virar PDF)", texto_relatorio.encode("utf-8"), "relatorio_impacto_sfa.txt", "text/plain")
st.download_button("⬇️ Baixar Base Filtrada em CSV", df.to_csv(index=False).encode("utf-8"), "base_filtrada_sfa.csv", "text/csv")
st.info("Para virar PDF de verdade: Baixa o TXT > Abre no Word / Google Docs > Arquivo > Baixar como PDF. Ou no navegador: Ctrl+P > Salvar como PDF.")
