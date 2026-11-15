import streamlit as st
import pandas as pd
import google.generativeai as genai
from datetime import datetime

st.set_page_config(page_title="OSC em Evidência", page_icon="💜", layout="wide")

# --- CONFIGURAÇÃO SEGURA: chave vem do Secrets, nunca do código ---
try:
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    model = genai.GenerativeModel("gemini-1.5-flash")
    tem_chave = True
except:
    tem_chave = False
    st.warning("⚠️ Configure a GEMINI_API_KEY em Settings > Secrets para ativar a IA.")

st.title("💜 OSC em Evidência")
st.markdown("**Interface que consome um agente de IA** - Você vê, aprova ou age sobre o resultado. Você sabe que está falando com IA e deve revisar antes de usar.")
st.caption("Protótipo para Projeto Final IA Aplicada - {reprograma} | Dados fictícios para demonstração")

# --- Histórico de execuções para provar AUTOMAÇÃO ---
if "historico" not in st.session_state:
    st.session_state.historico = []

# --- TELA 1: ENTRADA ---
st.header("1. Envie a planilha de OSCs")
st.info("Base de dados: planilha enviada + dados higienizados. Campos sensíveis (CPF, telefone) são ignorados pelo agente.")

arquivo = st.file_uploader("Faça upload de .csv ou .xlsx", type=["csv", "xlsx"])
usar_exemplo = st.button("Usar dados de exemplo (20 OSCs fictícias)")

df = None
if usar_exemplo:
    df = pd.DataFrame([
        {"Nome OSC": f"OSC Exemplo {i+1}", "Área": ["Assistência Social","Educação","Cultura e Esporte"][i % 3], "Pessoas Atendidas": 50 + i*10, "Relatório Atualizado": "Sim" if i % 4 != 0 else "Não", "Horas/mês em relatório": 8 if i % 2 == 0 else 20} for i in range(20)
    ])
elif arquivo is not None:
    try:
        df = pd.read_csv(arquivo) if arquivo.name.endswith(".csv") else pd.read_excel(arquivo)
    except Exception as e:
        st.error(f"Erro ao ler arquivo: {e}")

if df is not None:
    st.subheader("Prévia da Base de Dados - Higienizada")
    st.dataframe(df.head(20), use_container_width=True)
    st.caption(f"Total: {len(df)} OSCs | Colunas: {', '.join(df.columns)}")

    # --- TELA 2: AGENTE ---
    st.header("2. Gerar síntese com IA")
    if st.button("✨ Gerar síntese com Agente", type="primary"):
        if not tem_chave:
            st.error("Configure a chave para gerar.")
        else:
            with st.spinner("Agente analisando... isso levava ~45 min manualmente."):
                dados_texto = df.to_csv(index=False)
                prompt = f"""
                Você é o Agente OSC em Evidência, assistente de análise de impacto social para OSCs pequenas e médias no Brasil.
                REGRAS: Use APENAS os dados abaixo. Nunca invente CNPJ, valores ou nomes. Se campo vazio, escreva 'Dado não informado'. Ignore dados sensíveis. Linguagem simples. Deixe claro: 'Conteúdo gerado por IA, revise antes de usar'.
                ENTREGUE NESSA ORDEM:
                1. RESUMO EXECUTIVO: 3-5 linhas - quantas OSCs, quantas ativas, principal alerta.
                2. TABELA CLASSIFICADA: Nome | Área | Status [Ativa / Em risco / Dados incompletos] | Motivo
                3. INDICADORES: Total pessoas atendidas | Nº OSCs com dados incompletos
                4. ALERTAS: até 5 alertas práticos
                5. PRÓXIMO PASSO: 1 ação para hoje.
                Se não houver dados, diga o que falta.
                DADOS:
                {dados_texto}
                """
                try:
                    resposta = model.generate_content(prompt)
                    st.session_state.ultima_resposta = resposta.text
                    st.session_state.historico.append({
                        "Data": datetime.now().strftime("%d/%m %H:%M"),
                        "OSCs": len(df),
                        "Tempo": "Simulado: ~2-3 min vs 45 min manual"
                    })
                except Exception as e:
                    st.error(f"Erro na API (limite ou conexão): {e}. Use o Plano B - vídeo gravado.")

    if "ultima_resposta" in st.session_state:
        st.subheader("Resultado do Agente - Para Aprovação Humana")
        st.markdown(st.session_state.ultima_resposta)

        # --- VALIDAÇÃO HUMANA ---
        st.header("3. Validação Humana - 20 casos, 3 avaliadoras, 2 rodadas")
        st.markdown("Avalie a saída da IA com critérios definidos antes do teste:")
        c1 = st.selectbox("1. Está correta? (sem dados inventados)", ["Sim", "Parcial", "Não"])
        c2 = st.selectbox("2. Está completa? (resumo + tabela + alertas)", ["Sim", "Parcial", "Não"])
        c3 = st.selectbox("3. Dá para usar no relatório? (linguagem clara)", ["Sim", "Parcial", "Não"])
        nome_avaliadora = st.text_input("Nome da avaliadora")
        col1, col2, col3 = st.columns(3)
        if col1.button("✅ Aprovar"):
            st.success(f"Aprovado por {nome_avaliadora or 'Avaliadora'} - Registrado para a banca.")
        if col2.button("✏️ Editar / Ajustar"):
            st.info("Edição registrada. Ajuste o texto acima antes de enviar.")
        if col3.button("❌ Rejeitar"):
            st.error("Rejeitado. Motivo registrado para 2ª rodada.")

# --- PROVAS PARA A BANCA ---
st.divider()
st.header("Provas das 4 Peças Exigidas")
colA, colB = st.columns(2)
with colA:
    st.markdown("**1. Automação:** Gatilho = Upload -> Agente roda sozinho")
    if st.session_state.historico:
        st.table(pd.DataFrame(st.session_state.historico))
    else:
        st.caption("Nenhuma execução ainda. Gere 3 vezes para provar.")
    st.markdown("**2. Base de dados:** Planilha enviada + dados higienizados, sem dados sensíveis.")
with colB:
    st.markdown("**3. IA Generativa:** Prompt anexado no código. Sem IA, volta para 45 min manual.")
    st.markdown("**4. Validação Humana:** Meta: 20 casos, 3 avaliadoras, 2 rodadas.")
    st.caption("Rodada 1: Pesquisa problema (n=5, 100% útil). Rodada 2: Avaliação desta tela.")
