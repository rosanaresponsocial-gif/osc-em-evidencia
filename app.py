import streamlit as st
import pandas as pd

st.set_page_config(page_title="SFA - Relatório de Impacto", layout="wide")
st.title("🏠 SFA - Relatório de Impacto Social")
st.write("Suba o Excel. A equipe responde. A IA + automação geram o relatório. Sem inventar número.")

# --- 1. SUBIR EXCEL ---
st.header("1️⃣ Subir Excel da Base SFA")
uploaded = st.file_uploader("Suba o arquivo Excel (.xlsx) - Base_Dados_SFA", type=["xlsx", "xls"])

# Mantém a base na sessão
if "base" not in st.session_state:
    st.session_state.base = None
if "passo" not in st.session_state:
    st.session_state.passo = 0
    st.session_state.dados = {}

if uploaded:
    try:
        # Lê todas as abas
        xls = pd.ExcelFile(uploaded)
        st.write(f"Abas encontradas: {xls.sheet_names}")

        # Tenta achar a aba de Base de Dados - a sua é a Aba 2
        # Procura por nome que contenha Base / Dados / Casos
        aba_base = None
        for nome in xls.sheet_names:
            if any(p in nome.lower() for p in ["base", "dados", "casos", "sfa"]):
                aba_base = nome
                break
        if aba_base is None:
            aba_base = xls.sheet_names[1] if len(xls.sheet_names) > 1 else xls.sheet_names[0]

        df = pd.read_excel(uploaded, sheet_name=aba_base, header=0)
        # Remove linhas totalmente vazias
        df = df.dropna(how="all")
        st.session_state.base = df
        st.success(f"✅ Excel carregado! Aba usada: {aba_base} | {len(df)} linhas | {len(df.columns)} colunas")

        with st.expander("Ver colunas que o app encontrou"):
            st.write(list(df.columns))
        with st.expander("Ver os primeiros casos"):
            st.dataframe(df.head(10), use_container_width=True)

    except Exception as e:
        st.error(f"❌ Erro ao ler Excel: {e}")
        st.write("Dica: verifique se o arquivo não está aberto no Excel na hora de subir.")

# --- 2. PAINEL DE IMPACTO - SÓ SE TEM BASE ---
if st.session_state.base is not None:
    df = st.session_state.base
    st.header("2️⃣ Painel de Impacto - Calculado da sua base real")

    # Tenta achar as colunas pelo nome, sem depender de nome exato
    def achar_coluna(palavras):
        for c in df.columns:
            cl = str(c).lower()
            if any(p in cl for p in palavras):
                return c
        return None

    col_objetivo = achar_coluna(["objetivo", "atingido"])
    col_status = achar_coluna(["status", "prazo"])
    col_total_interv = achar_coluna(["total", "interven"])
    col_atendidos = achar_coluna(["atendidos", "núcleo", "nucleo", "pessoas"])
    col_custo = achar_coluna(["custo total", "custo"])

    total_casos = len(df)
    # Concluídos = Objetivo == Sim
    concluidos = 0
    if col_objetivo:
        concluidos = (df[col_objetivo].astype(str).str.lower().str.strip() == "sim").sum()

    # Prazo - conta Antes + Dentro
    antes = dentro = apos = em_curso = 0
    if col_status:
        s = df[col_status].astype(str)
        antes = s.str.contains("Antes", case=False, na=False).sum()
        dentro = s.str.contains("Dentro", case=False, na=False).sum()
        apos = s.str.contains("Após|Após", case=False, na=False).sum()
        em_curso = s.str.contains("Em Curso|Em acolhimento", case=False, na=False).sum()

    no_prazo_perc = round((antes + dentro) / concluidos * 100, 1) if concluidos > 0 else 0
    total_interv = int(pd.to_numeric(df[col_total_interv], errors="coerce").sum()) if col_total_interv else 0
    total_atendidos = int(pd.to_numeric(df[col_atendidos], errors="coerce").sum()) if col_atendidos else 0
    custo_total = float(pd.to_numeric(df[col_custo], errors="coerce").sum()) if col_custo else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1: st.metric("Total de Casos", total_casos)
    with c2: st.metric("Concluídos (Objetivo Sim)", int(concluidos))
    with c3: st.metric("% No Prazo (Antes+Dentro)", f"{no_prazo_perc}%", f"{antes} antes + {dentro} dentro")
    with c4: st.metric("Total Intervenções", total_interv)
    c5, c6, c7 = st.columns(3)
    with c5: st.metric("Pessoas Atendidas", total_atendidos)
    with c6: st.metric("Custo Total", f"R$ {custo_total:,.2f}")
    with c7: st.metric("Custo Médio / Criança", f"R$ {custo_total/total_casos:,.2f}" if total_casos else "R$ 0")

    st.subheader("Respostas às 5 perguntas do Guia do Agente - calculadas da base")
    st.write(f"**1. Quantas intervenções até o objetivo?** Total: {total_interv} | Média por caso concluído: {total_interv/concluidos:.1f}" if concluidos else "Sem dado")
    st.write(f"**2. Desacolhidas antes do prazo?** {antes} casos")
    st.write(f"**3. Desacolhidas dentro do prazo?** {dentro} casos | Após: {apos} | Em curso: {em_curso}")
    st.write(f"**4. Problemas solucionados / transformações:** {concluidos} casos com objetivo atingido, de {total_casos} totais. Ver tabela abaixo por desfecho.")
    st.write(f"**5. Custo por criança?** Total R$ {custo_total:,.2f} | Médio R$ {custo_total/total_casos:,.2f}" if total_casos else "Sem dado")

    if col_status:
        st.write("**Distribuição por Status do Prazo:**")
        st.bar_chart(df[col_status].value_counts())

    st.download_button("⬇️ Baixar Relatório em CSV", df.to_csv(index=False).encode("utf-8"), "relatorio_sfa.csv", "text/csv")

    # --- 3. INTERFACE QUE PERGUNTA - NOVO CASO ---
    st.header("3️⃣ Registrar Novo Caso - Interface que vai perguntando")
    st.write("A equipe responde passo a passo. O novo caso entra na base e o painel atualiza.")

    if st.button("➕ Começar novo caso - Passo 1"):
        st.session_state.passo = 1
        st.session_state.dados = {}
        st.rerun()

    PROBLEMAS = [
        "Negligência e insegurança alimentar grave",
        "Conflito familiar agudo e fragilidade de vínculos",
        "Ausência de rede de apoio e vulnerabilidade habitacional",
        "Evasão escolar e violação de direitos básicos",
        "Dependência química parental com exposição a risco",
        "Dificuldade de acesso a acompanhamento especializado (Saúde/CAPS)",
        "Abandono temporário por internação/situação de crise do responsável",
    ]

    if st.session_state.passo >= 1:
        st.subheader(f"Passo {st.session_state.passo} de 6")
        if st.session_state.passo == 1:
            st.write("**O que foi feito? Quem é o caso?**")
            st.session_state.dados["Cód. Criança"] = st.text_input("Cód. Criança (anonimizado) - Ex: Criança_36", value=st.session_state.dados.get("Cód. Criança",""))
            st.session_state.dados["Idade (anos)"] = st.number_input("Idade", min_value=0, max_value=18, value=0)
            st.session_state.dados["Nº Atendidos (Núcleo)"] = st.number_input("Nº Atendidos (núcleo)", min_value=1, value=6)
            if st.button("Próxima ➡️ Problema"):
                st.session_state.passo = 2; st.rerun()
        elif st.session_state.passo == 2:
            st.write("**Qual era o problema inicial?**")
            st.session_state.dados["Problema Inicial Identificado"] = st.selectbox("Problema", PROBLEMAS)
            st.session_state.dados["Solução Aplicada"] = st.text_area("Solução Aplicada")
            if st.button("Próxima ➡️ Intervenções"):
                st.session_state.passo = 3; st.rerun()
        elif st.session_state.passo == 3:
            st.write("**O que a equipe fez? Digite só os números reais.**")
            campos = ["Conversas com Família de Origem","Conversas com Família Acolhedora","Intervenções com Vara","Intervenções na Rede","Relatórios Técnicos","Reuniões de Discussão de Caso","Conversas Individuais","Encaminhamentos de Saúde (só fluxo, sem diagnóstico)"]
            total = 0
            for campo in campos:
                v = st.number_input(campo, min_value=0, step=1, value=0, key=campo)
                st.session_state.dados[campo] = v
                total += v
            st.session_state.dados["Total de Intervenções Realizadas"] = total
            st.success(f"Total de Intervenções: {total}")
            if st.button("Próxima ➡️ Impacto"):
                st.session_state.passo = 4; st.rerun()
        elif st.session_state.passo == 4:
            st.write("**Qual foi o impacto gerado?**")
            st.session_state.dados["Resultado/Transformação Alcançada"] = st.text_area("Resultado / Transformação")
            st.session_state.dados["Desfecho do Caso"] = st.selectbox("Desfecho", ["Reintegração Família de Origem","Reintegração Família Extensa","Encaminhamento Família Substituta","Em Acolhimento (Ativo)"])
            st.session_state.dados["Objetivo Atingido (Sim / Em Andamento)"] = st.selectbox("Objetivo Atingido?", ["Sim","Em Andamento"])
            st.session_state.dados["Status do Prazo"] = st.selectbox("Status do Prazo", ["Desacolhida Antes do Prazo","Desacolhida Dentro do Prazo","Desacolhida Após o Prazo","Em Acolhimento (Em Curso)"])
            if st.button("Próxima ➡️ Custo"):
                st.session_state.passo = 5; st.rerun()
        elif st.session_state.passo == 5:
            st.write("**Quanto custou?**")
            bolsa = st.number_input("Custo Bolsa-Auxílio (R$)", min_value=0.0, step=100.0)
            equipe = st.number_input("Custo Equipe / Estrutura (R$)", min_value=0.0, step=100.0)
            st.session_state.dados["Custo Bolsa-Auxílio (R$)"] = bolsa
            st.session_state.dados["Custo Equipe / Estrutura (R$)"] = equipe
            st.session_state.dados["Custo Total por Criança (R$)"] = bolsa + equipe
            st.success(f"Custo Total: R$ {bolsa+equipe:,.2f}")
            if st.button("🎉 Gerar Relatório + Salvar na Base"):
                st.session_state.passo = 6; st.rerun()
        elif st.session_state.passo == 6:
            d = st.session_state.dados
            st.success("✅ Novo caso gerado! Relatório pronto, sem número inventado.")
            st.write(f"**Caso:** {d.get('Cód. Criança')} | **Problema:** {d.get('Problema Inicial Identificado')} | **Transformação:** {d.get('Resultado/Transformação Alcançada')}")
            st.write(f"**Intervenções:** {d.get('Total de Intervenções Realizadas')} | **Desfecho:** {d.get('Desfecho do Caso')} | **Prazo:** {d.get('Status do Prazo')} | **Custo:** R$ {d.get('Custo Total por Criança (R$)',0):,.2f}")
            # Adiciona na base
            nova = pd.DataFrame([d])
            st.session_state.base = pd.concat([st.session_state.base, nova], ignore_index=True)
            st.write(f"Base agora com {len(st.session_state.base)} casos. Recarregue o painel acima para ver o número atualizado.")
            if st.button("🔄 Registrar outro caso"):
                st.session_state.passo = 0
                st.session_state.dados = {}
                st.rerun()
else:
    st.info("👆 Suba o Excel primeiro para ver o painel. Depois registre o novo caso.")

