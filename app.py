import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Configuração da página para formato Mobile
st.set_page_config(page_title="Finanças Bajaj", page_icon="🏍️", layout="centered")

st.title("🏍️ Gestão Financeira - Bajaj Dominar 200")
st.caption("Meta Semanal Agressiva: R$ 1.000,00 Brutos | CLT + 40h Moto | Retenção: 40,5%")

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

# --- FUNÇÃO PARA GERAR EXCEL ESTÉTICO ---
def gerar_excel_estetico(dataframe):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        dataframe.to_excel(writer, index=False, sheet_name='Entregas')
        
        workbook = writer.book
        worksheet = writer.sheets['Entregas']
        
        cor_cabecalho = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
        fonte_cabecalho = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        fonte_dados = Font(name="Arial", size=10, color="000000")
        alinhamento_centro = Alignment(horizontal="center", vertical="center")
        
        borda_fina = Side(border_style="thin", color="D3D3D3")
        estilo_borda = Border(left=borda_fina, right=borda_fina, top=borda_fina, bottom=borda_fina)
        
        for col_idx in range(1, worksheet.max_column + 1):
            cell = worksheet.cell(row=1, column=col_idx)
            cell.fill = cor_cabecalho
            cell.font = fonte_cabecalho
            cell.alignment = alinhamento_centro
            cell.border = estilo_borda
            
        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
            for cell in row:
                cell.font = fonte_dados
                cell.alignment = alinhamento_centro
                cell.border = estilo_borda
                if cell.column > 1:
                    cell.number_format = 'R$ #,##0.00'
        
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 12)
            
    buffer.seek(0)
    return buffer

# --- FORMULÁRIO DE ENTRADA ---
st.subheader("📥 Lançar Nova Corrida/Turno")
with st.form("nova_entrega", clear_on_submit=True):
    valor_bruto = st.number_input("Valor Bruto Recebido (R$):", min_value=0.0, step=5.0, format="%.2f")
    data_entrega = st.date_input("Data:", datetime.now())
    submetido = st.form_submit_button("Calcular e Salvar Permanentemente")

    if submetido and valor_bruto > 0:
        retencao_total = valor_bruto * 0.405
        lucro_livre = valor_bruto * 0.595
        
        pai = retencao_total * 0.3079
        gasolina = retencao_total * 0.2222
        despesas_gerais = retencao_total * 0.4699
        
        novo_registro = {
            "Data": data_entrega.strftime("%d/%m/%Y"),
            "Bruto (R$)": round(valor_bruto, 2),
            "Lucro Livre (R$)": round(lucro_livre, 2),
            "Pai (R$)": round(pai, 2),
            "Gasolina (R$)": round(gasolina, 2),
            "Despesas Gerais (R$)": round(despesas_gerais, 2)
        }
        
        historico.append(novo_registro)
        salvar_banco(historico)
        st.success(f"Lançamento de R$ {valor_bruto:.2f} gravado com sucesso!")
        st.rerun()

# --- DASHBOARD DE ACUMULADOS ---
if historico:
    df = pd.DataFrame(historico)
    
    # Garantir compatibilidade de colunas caso existam registros antigos no JSON local
    if "Despesas Gerais (R$)" not in df.columns:
        df["Despesas Gerais (R$)"] = 0.0
        if "Manutenção (R$)" in df.columns and "Seguro/Reserva (R$)" in df.columns:
            df["Despesas Gerais (R$)"] = df["Manutenção (R$)"] + df["Seguro/Reserva (R$)"]

    st.markdown("---")
    st.subheader("📊 Painel Geral Acumulado")
    
    bruto_total = df['Bruto (R$)'].sum()
    lucro_total = df['Lucro Livre (R$)'].sum()
    custo_acumulado = df['Pai (R$)'].sum() + df['Gasolina (R$)'].sum() + df['Despesas Gerais (R$)'].sum()

    col1, col2 = st.columns(2)
    with col1:
        st.metric("Faturamento Bruto Total", f"R$ {bruto_total:.2f}")
    with col2:
        st.metric("🟢 Seu Lucro Livre Total (59,5%)", f"R$ {lucro_total:.2f}")
        
    st.markdown("### 🔒 Divisão de Caixinhas Reservadas (40,5%)")
    c1, c2, c3 = st.columns(3)
    c1.metric("👴🏼 Pai", f"R$ {df['Pai (R$)'].sum():.2f}")
    c2.metric("⛽ Gasolina", f"R$ {df['Gasolina (R$)'].sum():.2f}")
    c3.metric("🛠️ Despesas Gerais", f"R$ {df['Despesas Gerais (R$)'].sum():.2f}")
    
    # --- NOVA BARRA 1: FATURAMENTO BRUTO TOTAL ---
    st.markdown("### 🚀 Progresso do Faturamento Bruto (Meta Semanal: R$ 1.000,00)")
    progresso_bruto = min(float(bruto_total / 1000.0), 1.0) if bruto_total > 0 else 0.0
    st.progress(progresso_bruto)
    st.caption(f"Você já faturou R$ {bruto_total:.2f} de R$ 1.000,00 planejados.")

    # --- BARRA 2 ATUALIZADA: CUSTOS FIXOS COM REAIS QUE FALTAM ---
    st.markdown("### 📈 Progresso de Custos Fixos (Meta Semanal: R$ 405,00)")
    progresso_custos = min(float(custo_acumulado / 405.0), 1.0) if custo_acumulado > 0 else 0.0
    st.progress(progresso_custos)
    
    # Cálculo de quanto falta para atingir os R$ 405,00 de retenção
    falta_custos = max(405.0 - float(custo_acumulado), 0.0)
    if falta_custos > 0:
        st.caption(f"Você já garantiu R$ {custo_acumulado:.2f}. **Falta R$ {falta_custos:.2f}** para bater a meta de retenção.")
    else:
        st.caption(f"🎉 **Meta de despesas 100% batida!** Todos os R$ 405,00 de custos e obrigações já estão totalmente garantidos.")

    # --- GERAR EXCEL DESIGNED ---
    excel_estetico = gerar_excel_estetico(df)

    # --- BOTÃO DE EXPORTAÇÃO ---
    st.markdown("### 📥 Exportar Dados Formatados")
    st.download_button(
        label="📊 Baixar Planilha para o Excel",
        data=excel_estetico,
        file_name=f"financeiro_bajaj_{datetime.now().strftime('%m_%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # --- HISTÓRICO VISUAL ---
    st.markdown("---")
    st.subheader("📋 Histórico de Lançamentos Permanentes")
    st.dataframe(df[["Data", "Bruto (R$)", "Lucro Livre (R$)", "Pai (R$)", "Gasolina (R$)", "Despesas Gerais (R$)"]], use_container_width=True)

    # --- GERENCIAMENTO (EDITAR / EXCLUIR / EXCLUIR TUDO) ---
    st.markdown("---")
    st.subheader("⚙️ Gerenciar Entregas Lançadas")
    
    aba_unitario, aba_massa = st.tabs(["✏️ Alterar Item Único", "🚨 Limpar Todo o Banco"])
    
    with aba_unitario:
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
                historico[id_selecionado]["Pai (R$)"] = round(ret_t * 0.3079, 2)
                historico[id_selecionado]["Gasolina (R$)"] = round(ret_t * 0.2222, 2)
                historico[id_selecionado]["Despesas Gerais (R$)"] = round(ret_t * 0.4699, 2)
                
                historico[id_selecionado].pop("Manutenção (R$)", None)
                historico[id_selecionado].pop("Seguro/Reserva (R$)", None)
                
                salvar_banco(historico)
                st.success("Registro atualizado com sucesso!")
                st.rerun()
                
        with col_ex:
            st.markdown("**Excluir Registro**")
            st.warning("Esta ação removerá apenas o ID selecionado.")
            if st.button("🗑️ Excluir Item"):
                historico.pop(id_selecionado)
                salvar_banco(historico)
                st.success("Registro apagado com sucesso!")
                st.rerun()
                
    with aba_massa:
        st.markdown("### 🚨 Exclusão Geral do Histórico")
        st.write("Clique abaixo para abrir as opções de exclusão completa dos registros.")
        
        if st.checkbox("⚠️ Quero apagar todos os dados acumulados"):
            st.error("❓ Deseja excluir tudo permanentemente?")
            if st.button("💥 SIM, APAGAR HISTÓRICO"):
                salvar_banco([])
                st.success("O banco de dados foi completamente resetado!")
                st.rerun()
