import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime

# Configuração da página para formato Mobile
st.set_page_config(page_title="Finanças Bajaj", page_icon="🏍️", layout="centered")

st.title("🏍️ Gestão Financeira - Bajaj Dominar 200")
st.caption("Cenário 41 Meses e 20 Dias | Retenção: 40,5%")

# Banco de dados local em formato texto estruturado permanente
ARQUIVO_BD = "banco_entregas_bajaj.json"

def carregar_banco():
    if os.path.exists(ARQUIVO_BD):
        with open(ARQUIVO_BD, "r") as f:
            return json.load(f)
    return []

def salvar_banco(dados):
    with open(ARQUIVO_BD, "w") as f:
        json.dump(dados, f, indent=4)

# Carrega os lançamentos anteriores
historico = carregar_banco()

# --- FORMULÁRIO DE ENTRADA ---
st.subheader("📥 Lançar Nova Corrida/Turno")
with st.form("nova_entrega", clear_on_submit=True):
    valor_bruto = st.number_input("Valor Bruto Recebido (R$):", min_value=0.0, step=5.0, format="%.2f")
    data_entrega = st.date_input("Data:", datetime.now())
    submetido = st.form_submit_button("Calcular e Salvar Permanentemente")

    if submetido and valor_bruto > 0:
        # Lógica das porcentagens acordadas
        retencao_total = valor_bruto * 0.405
        lucro_livre = valor_bruto * 0.595
        
        # Divisão interna dos 40,5%
        pai = retencao_total * 0.515
        gasolina = retencao_total * 0.220
        manutencao = retencao_total * 0.108
        seguro_reserva = retencao_total * 0.157
        
        novo_registro = {
            "Data": data_entrega.strftime("%d/%m/%Y"),
            "Bruto (R$)": round(valor_bruto, 2),
            "Lucro Livre (R$)": round(lucro_livre, 2),
            "Pai (R$)": round(pai, 2),
            "Gasolina (R$)": round(gasolina, 2),
            "Manutenção (R$)": round(manutencao, 2),
            "Seguro/Reserva (R$)": round(seguro_reserva, 2)
        }
        
        historico.append(novo_registro)
        salvar_banco(historico)
        st.success(f"Lançamento de R$ {valor_bruto:.2f} gravado com sucesso!")
        st.rerun()

# --- DASHBOARD DE ACUMULADOS ---
if historico:
    df = pd.DataFrame(historico)
    
    st.markdown("---")
    st.subheader("📊 Painel Geral Acumulado")
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Faturamento Bruto Total", f"R$ {df['Bruto (R$)'].sum():.2f}")
    with col2:
        st.metric("🟢 Seu Lucro Livre Total (59,5%)", f"R$ {df['Lucro Livre (R$)'].sum():.2f}")
        
    st.markdown("### 🔒 Divisão de Caixinhas Reservadas (40,5%)")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("👴🏼 Pai", f"R$ {df['Pai (R$)'].sum():.2f}")
    c2.metric("⛽ Gasolina", f"R$ {df['Gasolina (R$)'].sum():.2f}")
    c3.metric("🛠️ Mecânica", f"R$ {df['Manutenção (R$)'].sum():.2f}")
    c4.metric("🛡️ Seguro", f"R$ {df['Seguro/Reserva (R$)'].sum():.2f}")
    
    # Meta semanal de custos fixos (R$ 242,00)
    custo_acumulado = df['Pai (R$)'].sum() + df['Gasolina (R$)'].sum() + df['Manutenção (R$)'].sum() + df['Seguro/Reserva (R$)'].sum()
    st.markdown("### 📈 Progresso de Custos Fixos (Meta Semanal: R$ 242,00)")
    progresso = min(float(custo_acumulado / 242.0), 1.0) if custo_acumulado > 0 else 0.0
    st.progress(progresso)
    st.caption(f"Você já garantiu R$ {custo_acumulado:.2f} para manter a moto e o pagamento do seu pai em dia.")

    # --- HISTÓRICO VISUAL ---
    st.markdown("---")
    st.subheader("📋 Histórico de Lançamentos Permanentes")
    st.dataframe(df, use_container_width=True)
else:
    st.info("Nenhuma entrega lançada ainda. Digite um valor acima para ver o painel automático!")
