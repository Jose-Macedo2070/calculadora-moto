import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime
from io import BytesIO

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
        retencao_total = valor_bruto * 0.405
        lucro_livre = valor_bruto * 0.595
        
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
    
    custo_acumulado = df['Pai (R$)'].sum() + df['Gasolina (R$)'].sum() + df['Manutenção (R$)'].sum() + df['Seguro/Reserva (R$)'].sum()
    st.markdown("### 📈 Progresso de Custos Fixos (Meta Semanal: R$ 242,00)")
    progresso = min(float(custo_acumulado / 242.0), 1.0) if custo_acumulado > 0 else 0.0
    st.progress(progresso)
    st.caption(f"Você já garantiu R$ {custo_acumulado:.2f} para as despesas.")

    # --- BOTÃO DE EXPORTAÇÃO CORRIGIDO ---
    st.markdown("### 📥 Exportar Dados formatados")
    
    # Criar arquivo do Excel em memória de forma limpa e nativa (.xlsx)
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Entregas')
    buffer.seek(0)
    
    st.download_button(
        label="📊 Baixar Planilha para o Excel",
        data=buffer,
        file_name=f"financeiro_bajaj_{datetime.now().strftime('%m_%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # --- HISTÓRICO VISUAL ---
    st.markdown("---")
    st.subheader("📋 Histórico de Lançamentos Permanentes")
    st.dataframe(df, use_container_width=True)

    # --- GERENCIAMENTO (EDITAR / EXCLUIR) ---
    st.markdown("---")
    st.subheader("⚙️ Gerenciar Entregas Lançadas")
    
    opcoes_registro = [f"ID {idx} | {reg['Data']} | R$ {reg['Bruto (R$)']:.2f}" for idx, reg in enumerate(historico)]
    registro_selecionado = st.selectbox("Selecione o registro que deseja alterar:", options=opcoes_registro)
    
    id_selecionado = int(registro_selecionado.split(" ")[1])
    
    col_ed, col_ex = st.columns(2)
    
    with col_ed:
        st.markdown("**Editar Registro**")
        novo_valor = st.number_input("Corrigir Valor Bruto (R$):", min_value=0.0, value=float(historico[id_selecionado]["Bruto (R$)"]), key="edit_val")
        if st.button("✏️ Confirmar Alteração"):
            ret_t = novo_valor * 0.405
            historico[id_selecionado]["Bruto (R$)"] = round(novo_valor, 2)
            historico[id_selecionado]["Lucro Livre (R$)"] = round(novo_valor * 0.595, 2)
            historico[id_selecionado]["Pai (R$)"] = round(ret_t * 0.515, 2)
            historico[id_selecionado]["Gasolina (R$)"] = round(ret_t * 0.220, 2)
            historico[id_selecionado]["Manutenção (R$)"] = round(ret_t * 0.108, 2)
            historico[id_selecionado]["Seguro/Reserva (R$)"] = round(ret_t * 0.157, 2)
            
            salvar_banco(historico)
            st.success("Registro atualizado com sucesso!")
            st.rerun()
            
    with col_ex:
        st.markdown("**Excluir Registro**")
        st.warning("Esta ação não pode ser desfeita.")
        if st.button("🗑️ Excluir permanentemente"):
            historico.pop(id_selecionado)
            salvar_banco(historico)
            st.success("Registro apagado do banco de dados!")
            st.rerun()
else:
    st.info("Nenhuma entrega lançada ainda. Digite um valor acima para criar os primeiros registros permanentes!")
