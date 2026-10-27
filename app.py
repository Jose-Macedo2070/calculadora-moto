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
        
        for row in worksheet.iter_rows(min_row=1, max_row=1, min_col=1, max_col=worksheet.max_column):
            for cell in row:
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
            col_letter = get_column_letter(col.column)
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
    st.dataframe(df, use_container_width=True)

    # --- GERENCIAMENTO (EDITAR / EXCLUIR / EXCLUIR TUDO) ---
    st.markdown("---")
    st.subheader("⚙️ Gerenciar Entregas Lançadas")
    
    aba_unitario, aba_massa = st.tabs(["✏️ Alterar Item Único", "🚨 Limpar Todo o Banco"])
    
    with aba_unitario:
        opcoes_registro = [f"ID {idx} | {reg['Data']} | R$ {reg['Bruto (R$)']:.2f}" for idx, reg in enumerate(historico)]
        registro_selecionado = st.selectbox("Selecione o registro que deseja alterar:", options=opcoes_registro)
        
        # --- LINHA CORRIGIDA AQUI ---
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
                st.success("Registro atualizado!")
                st.rerun()
                
        with col_ex:
            st.markdown("**Excluir Registro**")
            st.warning("Esta ação removerá apenas o ID selecionado.")
            if st.button("🗑️ Excluir Item"):
                historico.pop(id_selecionado)
                salvar_banco(historico)
                st.success("Registro apagado!")
                st.rerun()
                
    with aba_massa:
        st.markdown("### 🚨 Exclusão Geral do Histórico")
        st.write("Clique abaixo para abrir as opções de exclusão completa dos registros.")
        
        if st.checkbox("⚠️ Quero apagar todos os dados acumulados"):
            st.error("❓ Deseja excluir tudo permanentemente?")
            
            col_b1, col_b2 = st.columns(2)
            
            with col_b1:
                st.markdown("""
                    <style>
                    div.stButton > button:first-child {
                        background-color: #d9534f;
                        color: white;
                    }
                    </style>""", unsafe_allow_html=True)
                if st.button("🔴 SIM (Apagar Tudo)"):
                    if os.path.exists(ARQUIVO_BD):
                        os.remove(ARQUIVO_BD)
                    st.success("Todos os registros foram excluídos!")
                    st.rerun()
                    
            with col_b2:
                st.download_button(
                    label="🟢 SALVAR (Baixar Planilha Backup)",
                    data=excel_estetico,
                    file_name=f"backup_entregas_bajaj_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_salvar_massa"
                )
else:
    st.info("Nenhuma entrega lançada ainda. Digite um valor acima para criar os primeiros registros permanentes!")
