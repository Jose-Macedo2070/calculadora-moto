import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime, date, timedelta
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =========================================================
# CONFIGURAÇÃO
# =========================================================

st.set_page_config(
    page_title="Finanças Bajaj",
    page_icon="🏍️",
    layout="centered",
    initial_sidebar_state="collapsed"
)

ARQUIVO_DADOS = "banco_entregas_bajaj.json"

META_BRUTA_SEMANAL = 1000.00
META_RETENCAO_SEMANAL = 405.00

PERCENTUAL_RETENCAO = 0.405
PERCENTUAL_PAI = 0.3079
PERCENTUAL_GASOLINA = 0.2222

APLICATIVOS = ["iFood", "Uber", "99", "Shopee", "Rappi", "Mercado Livre", "Outros"]

# =========================================================
# DESIGN
# =========================================================

st.markdown(
    """
    <style>
        .stApp {
            background-color: #101418;
            color: #f1f5f9;
        }

        [data-testid="stHeader"] {
            background: rgba(16, 20, 24, 0.95);
        }

        .block-container {
            max-width: 1000px;
            padding-top: 1.5rem;
            padding-bottom: 2rem;
        }

        h1, h2, h3 {
            color: #f8fafc !important;
        }

        [data-testid="stMetric"] {
            background: #1b2229;
            border: 1px solid #303943;
            padding: 16px;
            border-radius: 12px;
        }

        [data-testid="stMetricLabel"] {
            color: #cbd5e1 !important;
        }

        [data-testid="stMetricValue"] {
            color: #22c55e !important;
        }

        div.stButton > button,
        div.stDownloadButton > button {
            width: 100%;
            min-height: 44px;
            border-radius: 9px;
            border: 1px solid #374151;
            font-weight: 600;
        }

        div.stButton > button[kind="primary"] {
            background-color: #16a34a;
            color: white;
            border: none;
        }

        div.stButton > button[kind="primary"]:hover {
            background-color: #15803d;
            color: white;
        }

        [data-testid="stDataFrame"] {
            border: 1px solid #303943;
            border-radius: 10px;
            overflow: hidden;
        }

        @media (max-width: 640px) {
            .block-container {
                padding-left: 1rem;
                padding-right: 1rem;
                padding-top: 1rem;
            }

            h1 {
                font-size: 1.65rem !important;
            }

            [data-testid="stMetric"] {
                padding: 12px;
            }
        }
    </style>
    """,
    unsafe_allow_html=True
)

st.title("🏍️ Gestão Financeira - Bajaj Dominar 200")
st.caption("Meta Semanal: R$ 1.000,00 Brutos | CLT + 40h Moto | Retenção: 40,5%")

# =========================================================
# DADOS / FUNÇÕES
# =========================================================

def carregar_dados():
    if not os.path.exists(ARQUIVO_DADOS):
        return []
    try:
        with open(ARQUIVO_DADOS, "r", encoding="utf-8") as arquivo:
            dados = json.load(arquivo)
        if not isinstance(dados, list):
            return []
        
        # Compatibilidade com registros antigos
        for registro in dados:
            registro.setdefault("Hora", "")
            registro.setdefault("Aplicativo", "")
            registro.setdefault("KM percorridos", None)
            
            # Garante que despesas gerais existam se vier de planilhas antigas
            if "Despesas Gerais (R$)" not in registro:
                manut = registro.pop("Manutenção (R$)", 0.0)
                seg = registro.pop("Seguro/Reserva (R$)", 0.0)
                registro["Despesas Gerais (R$)"] = round(manut + seg, 2)

            for campo in ["Bruto (R$)", "Lucro Livre (R$)", "Pai (R$)", "Gasolina (R$)", "Despesas Gerais (R$)"]:
                registro.setdefault(campo, 0.0)
        return dados
    except (json.JSONDecodeError, OSError):
        return []

def salvar_dados(dados):
    try:
        with open(ARQUIVO_DADOS, "w", encoding="utf-8") as arquivo:
            json.dump(dados, arquivo, ensure_ascii=False, indent=4)
        return True
    except OSError:
        return False

def calcular_corrida(bruto):
    bruto = max(float(bruto), 0.0)
    retencao = bruto * PERCENTUAL_RETENCAO
    lucro_livre = bruto - retencao
    pai = retencao * PERCENTUAL_PAI
    gasolina = retencao * PERCENTUAL_GASOLINA
    despesas = retencao - pai - gasolina
    return {
        "Bruto (R$)": round(bruto, 2),
        "Lucro Livre (R$)": round(lucro_livre, 2),
        "Pai (R$)": round(pai, 2),
        "Gasolina (R$)": round(gasolina, 2),
        "Despesas Gerais (R$)": round(despesas, 2)
    }

def obter_inicio_semana():
    hoje = date.today()
    return hoje - timedelta(days=hoje.weekday())

def obter_resumo_semana(df_dados):
    if df_dados.empty:
        return {"bruto": 0.0, "lucro": 0.0, "retencao": 0.0}
    inicio = pd.to_datetime(obter_inicio_semana())
    fim = inicio + timedelta(days=6, hours=23, minutes=59)
    
    semana = df_dados[(df_dados["Data"] >= inicio) & (df_dados["Data"] <= fim)]
    retencao = semana[["Pai (R$)", "Gasolina (R$)", "Despesas Gerais (R$)"]].sum().sum()
    return {
        "bruto": float(semana["Bruto (R$)"].sum()),
        "lucro": float(semana["Lucro Livre (R$)"].sum()),
        "retencao": float(retencao)
    }

def gerar_excel_estetico(dataframe):
    buffer = BytesIO()
    df_excel = dataframe.copy()
    if "Data" in df_excel.columns:
        df_excel["Data"] = df_excel["Data"].dt.strftime("%d/%m/%Y")
        
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_excel.to_excel(writer, index=False, sheet_name='Entregas')
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
                if cell.column > 4:
                    cell.number_format = 'R$ #,##0.00'
        
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 12)
            
    buffer.seek(0)
    return buffer

# Carrega o histórico
historico = carregar_dados()

# =========================================================
# FORMULÁRIO DE ENTRADA
# =========================================================
st.subheader("📥 Lançar Nova Corrida/Turno")
with st.form("nova_entrega", clear_on_submit=True):
    col_v, col_d, col_h = st.columns([2, 1, 1])
    with col_v:
        valor_bruto = st.number_input("Valor Bruto (R$):", min_value=0.0, step=5.0, format="%.2f")
    with col_d:
        data_corrida = st.date_input("Data:", date.today())
    with col_h:
        hora_corrida = st.time_input("Hora:", datetime.now().time())
        
    col_app, col_km = st.columns(2)
    with col_app:
        app_selecionado = st.selectbox("Aplicativo:", options=APLICATIVOS)
    with col_km:
        km_rodados = st.number_input("KM Rodados (Opcional):", min_value=0.0, step=1.0, value=0.0)

    submetido = st.form_submit_button("Salvar Permanentemente", kind="primary")

    if submetido and valor_bruto > 0:
        valores = calcular_corrida(valor_bruto)
        novo_registro = {
            "Data": data_corrida.strftime("%Y-%m-%d"),
            "Hora": hora_corrida.strftime("%H:%M"),
            "Aplicativo": app_selecionado,
            "KM percorridos": km_rodados if km_rodados > 0 else None,
            **valores
        }
        historico.append(novo_registro)
        salvar_dados(historico)
        st.success(f"Lançamento de R$ {valor_bruto:.2f} gravado com sucesso!")
        st.rerun()

# =========================================================
# DASHBOARD / VISUALIZAÇÃO
# =========================================================
if historico:
    # Cria o dataframe formatado para os cálculos
    df = pd.DataFrame(historico)
    df["Data"] = pd.to_datetime(df["Data"])
    df = df.sort_values("Data", ascending=False)
    
    resumo = obter_resumo_semana(df)
    
    st.markdown("---")
    st.subheader("📊 Painel da Semana Atual")
    
    col_b, col_l = st.columns(2)
    with col_b:
        st.metric("Faturamento Bruto da Semana", f"R$ {resumo['bruto']:.2f}")
    with col_l:
        st.metric("🟢 Seu Lucro Livre Semanal (59,5%)", f"R$ {resumo['lucro']:.2f}")
        
    st.markdown("### 🔒 Divisão Acumulada das Caixinhas (40,5%)")
    c1, c2, c3 = st.columns(3)
    c1.metric("👴🏼 Pai", f"R$ {df['Pai (R$)'].sum():.2f}")
    c2.metric("⛽ Gasolina", f"R$ {df['Gasolina (R$)'].sum():.2f}")
    c3.metric("🛠️ Despesas Gerais", f"R$ {df['Despesas Gerais (R$)'].sum():.2f}")
    
    # --- BARRA 1: PROGRESSO DO FATURAMENTO BRUTO ---
