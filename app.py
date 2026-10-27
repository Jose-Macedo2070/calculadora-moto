import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from datetime import datetime

# Configuração da página para formato Mobile
st.set_page_config(page_title="Finanças Bajaj", page_icon="🏍️", layout="centered")

st.title("🏍️ Gestão Financeira - Bajaj Dominar 200")
st.caption("Banco de Dados na Nuvem (Google Sheets) | Retenção: 40,5%")

# Estabelece a conexão com o Google Sheets utilizando os Secrets do Streamlit
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    # Lê os dados existentes na planilha
    df_historico = conn.read(ttl="0d") # ttl="0d" força o app a buscar os dados em tempo real
except Exception as e:
    st.error("Erro ao conectar com o Google Sheets. Verifique as configurações de credenciais.")
    df_historico = pd.DataFrame()

# --- FORMULÁRIO DE ENTRADA ---
st.subheader("📥 Lançar Nova Corrida/Turno")
with st.form("nova_entrega", clear_on_submit=True):
    valor_bruto = st.number_input("Valor Bruto Recebido (R$):", min_value=0.0, step=5.0, format="%.2f")
    data_entrega = st.date_input("Data:", datetime.now())
    submetido = st.form_submit_button("Calcular e Salvar na Nuvem")

    if submetido and valor_bruto > 0:
        # Lógica exata das porcentagens acordadas
        retencao_total = valor_bruto * 0.405
        lucro_livre = valor_bruto * 0.595
        
        # Divisão interna dos 40,5%
        pai = retencao_total * 0.515
        gasolina = retencao_total * 0.220
        manutencao = retencao_total * 0.108
        seguro_reserva = retencao_total * 0.157
        
        # Nova linha formatada para a planilha
        nova_linha = pd.DataFrame([{
            "Data": data_entrega.strftime("%d/%m/%Y"),
            "Bruto": round(valor_bruto, 2),
            "Lucro_Livre": round(lucro_livre, 2),
            "Pai": round(pai, 2),
            "Gasolina": round(gasolina, 2),
            "Manutencao": round(manutencao, 2),
            "Seguro_Reserva": round(seguro_reserva, 2)
        }])
        
        # Junta os dados antigos com a nova linha
        df_atualizado = pd.concat([df_historico, nova_linha], ignore_index=True)
        
        # Salva de volta na planilha do Google Drive
        conn.update(data=df_atualizado)
        st.success(f"R$ {valor_bruto:.2f} gravado diretamente no Google Sheets!")
        st.rerun()

# --- DASHBOARD DE ACUMULADOS ---
if df_historico is not None and not df_historico.empty:
    st.markdown("---")
    st.subheader("📊 Painel Geral Acumulado (Nuvem)")
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Faturamento Bruto Total", f"R$ {pd.to_numeric(df_historico['Bruto']).sum():.2f}")
    with col2:
        st.metric("🟢 Seu Lucro Livre Total (59,5%)", f"R$ {pd.to_numeric(df_historico['Lucro_Livre']).sum():.2f}")
        
    st.markdown("### 🔒 Divisão de Caixinhas Reservadas (40,5%)")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("👴🏼 Pai", f"R$ {pd.to_numeric(df_historico['Pai']).sum():.2f}")
    c2.metric("⛽ Gasolina", f"R$ {pd.to_numeric(df_historico['Gasolina']).sum():.2f}")
    c3.metric("🛠️ Mecânica", f"R$ {pd.to_numeric(df_historico['Manutencao']).sum():.2f}")
    c4.metric("🛡️ Seguro", f"R$ {pd.to_numeric(df_historico['Seguro_Reserva']).sum():.2f}")
    
    # Cálculo da meta semanal
    custo_acumulado = (pd.to_numeric(df_historico['Pai']).sum() + 
                       pd.to_numeric(df_historico['Gasolina']).sum() + 
                       pd.to_numeric(df_historico['Manutencao']).sum() + 
                       pd.to_numeric(df_historico['Seguro_Reserva']).sum())
    
    st.markdown("### 📈 Progresso de Custos Fixos (Meta Semanal: R$ 242,00)")
    progresso = min(custo_acumulado / 242.0, 1.0) if custo_acumulado > 0 else 0.0
    st.progress(progresso)
    st.caption(f"Você já garantiu R$ {custo_acumulado:.2f} arrecadados na nuvem.")

    # --- HISTÓRICO VISUAL ---
    st.markdown("---")
    st.subheader("📋 Histórico de Lançamentos Sincronizados")
    st.dataframe(df_historico, use_container_width=True)
else:
    st.info("A planilha do Google Sheets está vazia ou carregando dados...")

