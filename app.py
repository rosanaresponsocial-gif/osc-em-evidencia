import streamlit as st
import pandas as pd
from datetime import date

st.set_page_config(page_title="SFA - Relatório de Impacto", layout="wide")
st.title("🏠 SFA - Assistente de Impacto Social")
st.write("A equipe responde passo a passo. A IA + automação geram o relatório. Sem inventar número.")

if "passo" not in st.session_state:
    st.session_state.passo = 1
    st.session_state.dados = {}
if "base" not in st.session_state:
    # Carrega sua base real se você subir o Excel, ou começa vazia
    st.session_state.base = pd.DataFrame()

PROBLEMAS = [
    "Negligência e insegurança alimentar grave",
    "Conflito familiar agudo e fragilidade de vínculos",
    "Ausência de rede de apoio e vulnerabilidade habitacional",
    "Evasão escolar e violação de direitos básicos",
    "Dependência química parental com exposição a risco",
    "Dificuldade de acesso a acompanhamento especializado (Saúde/CAPS)",
    "Abandono temporário por internação/situação de crise do responsável",
]
SOLUCOES = {
    PROBLEMAS[0]: "Articulação com CRAS/CREAS, inclusão em benefícios e mediação familiar",
    PROBLEMAS[1]: "Escuta qualificada, plano de acompanhamento familiar e oficinas de parentalidade",
    PROBLEMAS[2]: "Encaminhamento para programa habitacional e fortalecimento da família extensa",
    PROBLEMAS[3]: "Busca ativa escolar, articulação com DRE/CEU e acompanhamento pedagógico",
    PROBLEMAS[4]: "Articulação com CAPS AD/UBS e construção de rede protetiva comunitária",
    PROBLEMAS[5]: "Encaminhamento e monitoramento de fluxo junto à UBS/NIR (Regulação de Saúde)",
    PROBLEMAS[6]: "Acolhimento protetivo e busca de família extensa apta ao cuidado",
}

def avancar():
    st.session_state.passo += 1
    st.rerun()
def voltar():
    st.session_state.passo -= 1
    st.rerun()

st.progress(st.session_state.passo / 6)

# PASSO 1
if st.session_state.passo == 1:
    st.header("Passo 1 de 6 - O que foi feito? Quem é o caso?")
    st.session_state.dados["codigo"] = st.text_input("Cód. Criança (anonimizado) *", value=st.session_state.dados.get("codigo",""), placeholder="Ex: Criança_36")
    c1, c2, c3 = st.columns(3)
    with c1: st.session_state.dados["idade"] = st.number_input("Idade (anos)", min_value=0, max_value=18, value=st.session_state.dados.get("idade", 5))
    with c2: st.session_state.dados["regiao"] = st.selectbox("Região (SP)", ["Zona Norte","Zona Sul","Zona Leste","Zona Oeste","Centro"])
    with c3: st.session_state.dados["nucleo"] = st.number_input("Nº Atendidos (núcleo familiar)", min_value=1, value=st.session_state.dados.get("nucleo", 6))
    st.session_state.dados["data_acolh"] = str(st.date_input("Data Acolhimento", value=date(2025,1,1)))
    st.session_state.dados["prazo_dias"] = st.number_input("Prazo Estimado (dias)", min_value=30, value=st.session_state.dados.get("prazo_dias", 365))
    if st.button("Próxima ➡️ O problema"): avancar()

# PASSO 2
elif st.session_state.passo == 2:
    st.header("Passo 2 de 6 - Qual era o problema inicial?")
    prob = st.selectbox("Problema Inicial Identificado *", PROBLEMAS)
    st.session_state.dados["problema"] = prob
    st.info(f"Solução sugerida pela base: {SOLUCOES[prob]}")
    st.session_state.dados["solucao"] = st.text_area("Solução Aplicada - confirme ou ajuste", value=SOLUCOES[prob])
    col1, col2 = st.columns(2)
    with col1:
        if st.button("⬅️ Voltar"): voltar()
    with col2:
        if st.button("Próxima ➡️ O que a equipe fez"): avancar()

# PASSO 3
elif st.session_state.passo == 3:
    st.header("Passo 3 de 6 - O que a equipe fez? Intervenções")
    st.write("Digite só os números reais. O app soma o total sozinho.")
    campos = [
        ("conv_origem","Conversas com Família de Origem"),
        ("conv_acolhedora","Conversas com Família Acolhedora"),
        ("conv_substituta","Conversas com Família Substituta"),
        ("vara","Intervenções com Vara da Infância"),
        ("rede","Intervenções na Rede Socioeducativa"),
        ("relatorios","Relatórios Técnicos Realizados"),
        ("formacoes","Formações com os Grupos"),
        ("reunioes","Reuniões de Discussão de Caso"),
        ("individuais","Conversas Individuais"),
        ("pesquisas","Pesquisas / Estudos Diagnósticos"),
        ("saude","Encaminhamentos / Contatos de Saúde (só fluxo, sem diagnóstico - LGPD)"),
        ("outras","Outras Intervenções Técnicas"),
    ]
    c1, c2 = st.columns(2)
    for i, (chave, rotulo) in enumerate(campos):
        col = c1 if i % 2 == 0 else c2
        with col:
            st.session_state.dados[chave] = st.number_input(rotulo, min_value=0, step=1, value=st.session_state.dados.get(chave, 0))
    total = sum(st.session_state.dados.get(k, 0) for k,_ in campos)
    st.session_state.dados["total_interv"] = total
    st.success(f"Total de Intervenções calculado: **{total}**")
    st.session_state.dados["tempo_visitas"] = st.number_input("Tempo Total de Visitas (horas)", min_value=0.0, step=0.5)
    col1, col2 = st.columns(2)
    with col1:
        if st.button("⬅️ Voltar"): voltar()
    with col2:
        if st.button("Próxima ➡️ Impacto gerado"): avancar()

# PASSO 4
elif st.session_state.passo == 4:
    st.header("Passo 4 de 6 - Qual foi o impacto gerado?")
    st.session_state.dados["resultado"] = st.text_area("Resultado / Transformação Alcançada *", placeholder="Ex: Reintegração à família de origem com segurança alimentar e autonomia")
    st.session_state.dados["desfecho"] = st.selectbox("Desfecho do Caso", ["Reintegração Família de Origem","Reintegração Família Extensa","Encaminhamento Família Substituta","Em Acolhimento (Ativo)"])
    st.session_state.dados["objetivo"] = st.selectbox("Objetivo Atingido?", ["Sim","Em Andamento"])
    st.session_state.dados["data_desacolh"] = str(st.date_input("Data Desacolhimento Real (deixe em branco se em curso)", value=None))
    st.session_state.dados["status_prazo"] = st.selectbox("Status do Prazo", ["Desacolhida Antes do Prazo","Desacolhida Dentro do Prazo","Desacolhida Após o Prazo","Em Acolhimento (Em Curso)"])
    col1, col2 = st.columns(2)
    with col1:
        if st.button("⬅️ Voltar"): voltar()
    with col2:
        if st.button("Próxima ➡️ Custo"): avancar()

# PASSO 5
elif st.session_state.passo == 5:
    st.header("Passo 5 de 6 - Quanto custou?")
    bolsa = st.number_input("Custo Bolsa-Auxílio (R$)", min_value=0.0, step=100.0)
    equipe = st.number_input("Custo Equipe / Estrutura (R$)", min_value=0.0, step=100.0)
    st.session_state.dados["bolsa"] = bolsa
    st.session_state.dados["equipe_custo"] = equipe
    st.session_state.dados["custo_total"] = bolsa + equipe
    st.success(f"Custo Total por Criança: **R$ {bolsa + equipe:,.2f}** | Fórmula: Bolsa + Equipe")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("⬅️ Voltar"): voltar()
    with col2:
        if st.button("🎉 Gerar Relatório de Impacto"): avancar()

# PASSO 6 - RELATÓRIO
else:
    d = st.session_state.dados
    st.header("Passo 6 de 6 - Relatório de Impacto Gerado")
    st.success("✅ Caso salvo. Relatório gerado pela automação + IA, sem número inventado.")
    st.write(f"**Criança:** {d.get('codigo')} | **Problema:** {d.get('problema')}")
    st.write(f"**Solução:** {d.get('solucao')}")
    st.write(f"**Transformação:** {d.get('resultado')}")
    st.write(f"**Intervenções:** {d.get('total_interv')} | **Tempo de visitas:** {d.get('tempo_visitas')}h | **Atendidos:** {d.get('nucleo')} pessoas")
    st.write(f"**Desfecho:** {d.get('desfecho')} | **Objetivo:** {d.get('objetivo')} | **Prazo:** {d.get('status_prazo')}")
    st.write(f"**Custo:** R$ {d.get('custo_total',0):,.2f} (Bolsa R$ {d.get('bolsa',0):,.2f} + Equipe R$ {d.get('equipe_custo',0):,.2f})")
    st.write(f"**Frase técnica pronta:** A partir de {d.get('total_interv')} intervenções técnicas, o caso {d.get('codigo')} alcançou {d.get('resultado','').lower()} com desfecho {d.get('desfecho','').lower()}, com custo total de R$ {d.get('custo_total',0):,.2f}. Evidência: base de casos + relatórios técnicos.")
    # Salva
    nova = pd.DataFrame([d])
    st.session_state.base = pd.concat([st.session_state.base, nova], ignore_index=True)
    st.write(f"Total de casos na base nesta sessão: {len(st.session_state.base)}")
    if st.button("🔄 Registrar novo caso"):
        st.session_state.passo = 1
        st.session_state.dados = {}
        st.rerun()
    st.download_button("⬇️ Baixar este caso em CSV", pd.DataFrame([d]).to_csv(index=False).encode("utf-8"), "caso_sfa.csv", "text/csv")
