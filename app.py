
import json
import os
from datetime import date, datetime, timedelta
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# =========================================================
# CONFIGURAÇÕES
# =========================================================

st.set_page_config(
    page_title="Finanças Bajaj",
    page_icon="🏍️",
    layout="centered",
)

st.title("🏍️ Gestão Financeira - Bajaj Dominar 200")
st.caption("Meta semanal: R$ 1.000,00 brutos | Retenção: 40,5%")

ARQUIVO_BD = "banco_entregas_bajaj.json"

META_BRUTA_SEMANAL = 1000.00
META_CUSTOS_SEMANAL = 405.00

TAXA_RETENCAO = 0.405
TAXA_LUCRO = 0.595

# Distribuição da retenção
PERCENTUAL_PAI = 0.3079
PERCENTUAL_GASOLINA = 0.2222
PERCENTUAL_DESPESAS = 0.4699

COLUNAS_FINANCEIRAS = [
    "Bruto (R$)",
    "Lucro Livre (R$)",
    "Pai (R$)",
    "Gasolina (R$)",
    "Despesas Gerais (R$)",
]

COLUNAS_HISTORICO = ["Data"] + COLUNAS_FINANCEIRAS


# =========================================================
# BANCO DE DADOS JSON
# =========================================================

def carregar_banco():
    """Carrega o histórico salvo no arquivo JSON."""

    if not os.path.exists(ARQUIVO_BD):
        return []

    try:
        with open(ARQUIVO_BD, "r", encoding="utf-8") as arquivo:
            dados = json.load(arquivo)

        if isinstance(dados, list):
            return dados

        return []

    except (json.JSONDecodeError, OSError):
        st.error(
            "Não foi possível ler o banco de dados. "
            "Verifique o arquivo JSON antes de continuar."
        )
        return []


def salvar_banco(dados):
    """Salva os lançamentos no arquivo JSON."""

    arquivo_temporario = ARQUIVO_BD + ".tmp"

    with open(arquivo_temporario, "w", encoding="utf-8") as arquivo:
        json.dump(
            dados,
            arquivo,
            indent=4,
            ensure_ascii=False,
        )

    os.replace(arquivo_temporario, ARQUIVO_BD)


# =========================================================
# CÁLCULOS FINANCEIROS
# =========================================================

def calcular_registro(valor_bruto, data_entrega):
    """Calcula automaticamente a distribuição de uma entrega."""

    valor_bruto = round(float(valor_bruto), 2)
    retencao = valor_bruto * TAXA_RETENCAO

    return {
        "Data": data_entrega.strftime("%d/%m/%Y"),
        "Bruto (R$)": valor_bruto,
        "Lucro Livre (R$)": round(valor_bruto * TAXA_LUCRO, 2),
        "Pai (R$)": round(retencao * PERCENTUAL_PAI, 2),
        "Gasolina (R$)": round(retencao * PERCENTUAL_GASOLINA, 2),
        "Despesas Gerais (R$)": round(
            retencao * PERCENTUAL_DESPESAS, 2
        ),
    }


def normalizar_historico(historico):
    """Adapta registros antigos ao formato atual."""

    registros_normalizados = []

    for registro in historico:
        bruto = float(registro.get("Bruto (R$)", 0) or 0)

        # Mantém a data original dos registros existentes.
        data_registro = registro.get(
            "Data",
            date.today().strftime("%d/%m/%Y"),
        )

        novo = calcular_registro(
            bruto,
            datetime.strptime(data_registro, "%d/%m/%Y").date(),
        )

        # Preserva os valores antigos de categorias quando existirem.
        # Converte os registros antigos de manutenção e seguro.
        if "Despesas Gerais (R$)" in registro:
            novo["Despesas Gerais (R$)"] = float(
                registro["Despesas Gerais (R$)"] or 0
            )
        elif (
            "Manutenção (R$)" in registro
            or "Seguro/Reserva (R$)" in registro
        ):
            novo["Despesas Gerais (R$)"] = round(
                float(registro.get("Manutenção (R$)", 0) or 0)
                + float(registro.get("Seguro/Reserva (R$)", 0) or 0),
                2,
            )

        # Recalcula as categorias que não existiam no formato antigo.
        for coluna in COLUNAS_FINANCEIRAS:
            if coluna in registro and coluna != "Bruto (R$)":
                if coluna != "Despesas Gerais (R$)" or (
                    "Despesas Gerais (R$)" in registro
                ):
                    novo[coluna] = float(registro[coluna] or 0)

        registros_normalizados.append(novo)

    return registros_normalizados


def moeda(valor):
    """Formata valores para o padrão brasileiro."""

    valor_formatado = f"{valor:,.2f}"
    valor_formatado = (
        valor_formatado.replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )

    return f"R$ {valor_formatado}"


# =========================================================
# EXPORTAÇÃO PARA EXCEL
# =========================================================

def gerar_excel_estetico(dataframe):
    """Gera uma planilha formatada para Excel."""

    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        dataframe.to_excel(
            writer,
            index=False,
            sheet_name="Entregas",
        )

        workbook = writer.book
        worksheet = writer.sheets["Entregas"]

        preenchimento_cabecalho = PatternFill(
            start_color="1B365D",
            end_color="1B365D",
            fill_type="solid",
        )

        fonte_cabecalho = Font(
            name="Arial",
            size=11,
            bold=True,
            color="FFFFFF",
        )

        fonte_dados = Font(
            name="Arial",
            size=10,
            color="000000",
        )

        alinhamento = Alignment(
            horizontal="center",
            vertical="center",
        )

        borda = Side(
            border_style="thin",
            color="D3D3D3",
        )

        estilo_borda = Border(
            left=borda,
            right=borda,
            top=borda,
            bottom=borda,
        )

        # Cabeçalho
        for celula in worksheet[1]:
            celula.fill = preenchimento_cabecalho
            celula.font = fonte_cabecalho
            celula.alignment = alinhamento
            celula.border = estilo_borda

        # Conteúdo
        for linha in worksheet.iter_rows(
            min_row=2,
            max_row=worksheet.max_row,
        ):
            for celula in linha:
                celula.font = fonte_dados
                celula.alignment = alinhamento
                celula.border = estilo_borda

                nome_coluna = worksheet.cell(
                    row=1,
                    column=celula.column,
                ).value

                if nome_coluna in COLUNAS_FINANCEIRAS:
                    celula.number_format = '"R$" #,##0.00'

        # Largura automática
        for coluna in worksheet.columns:
            maior_tamanho = max(
                len(str(celula.value or ""))
                for celula in coluna
            )

            letra = get_column_letter(coluna[0].column)
            worksheet.column_dimensions[letra].width = max(
                maior_tamanho + 4,
                12,
            )

        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        worksheet.row_dimensions[1].height = 25

    buffer.seek(0)
    return buffer


# =========================================================
# CARREGAMENTO INICIAL
# =========================================================

historico = carregar_banco()

# Impede que um JSON corrompido ou com estrutura inesperada
# seja sobrescrito silenciosamente.
if historico:
    try:
        historico = normalizar_historico(historico)
    except (ValueError, TypeError, KeyError) as erro:
        st.error(
            "Há registros incompatíveis no arquivo JSON. "
            "Corrija ou restaure uma cópia do arquivo antes de "
            "salvar novas alterações."
        )
        st.stop()


# =========================================================
# FORMULÁRIO DE NOVA ENTREGA
# =========================================================

st.subheader("📥 Lançar Nova Corrida/Turno")

with st.form("nova_entrega", clear_on_submit=True):
    valor_bruto = st.number_input(
        "Valor bruto recebido (R$)",
        min_value=0.0,
        step=5.0,
        format="%.2f",
    )

    data_entrega = st.date_input(
        "Data da entrega",
        value=date.today(),
        format="DD/MM/YYYY",
    )

    submetido = st.form_submit_button(
        "💾 Calcular e Salvar",
        use_container_width=True,
    )

    if submetido:
        if valor_bruto <= 0:
            st.warning("Informe um valor maior que zero.")
        else:
            novo_registro = calcular_registro(
                valor_bruto,
                data_entrega,
            )

            historico.append(novo_registro)
            salvar_banco(historico)

            st.success("Lançamento salvo com sucesso!")
            st.rerun()


# =========================================================
# DASHBOARD FINANCEIRO
# =========================================================

if historico:
    df = pd.DataFrame(historico, columns=COLUNAS_HISTORICO)

    df["_Data"] = pd.to_datetime(
        df["Data"],
        format="%d/%m/%Y",
        errors="coerce",
    )

    # Totais de todo o histórico
    bruto_total = df["Bruto (R$)"].sum()
    lucro_total = df["Lucro Livre (R$)"].sum()
    pai_total = df["Pai (R$)"].sum()
    gasolina_total = df["Gasolina (R$)"].sum()
    despesas_total = df["Despesas Gerais (R$)"].sum()

    custo_acumulado = pai_total + gasolina_total + despesas_total

    # Período semanal: segunda-feira até domingo.
    hoje = date.today()
    inicio_semana = hoje - timedelta(days=hoje.weekday())
    fim_semana = inicio_semana + timedelta(days=6)

    df_semana = df[
        (df["_Data"].dt.date >= inicio_semana)
        & (df["_Data"].dt.date <= fim_semana)
    ]

    bruto_semana = df_semana["Bruto (R$)"].sum()

    custos_semana = (
        df_semana["Pai (R$)"].sum()
        + df_semana["Gasolina (R$)"].sum()
        + df_semana["Despesas Gerais (R$)"].sum()
    )

    lucro_semana = df_semana["Lucro Livre (R$)"].sum()

    st.markdown("---")
    st.subheader("📊 Painel Geral Acumulado")

    col1, col2 = st.columns(2)

    with col1:
        st.metric(
            "Faturamento bruto total",
            moeda(bruto_total),
        )

    with col2:
        st.metric(
            "Lucro livre total",
            moeda(lucro_total),
        )

    st.markdown("### 🔒 Caixinhas Reservadas")

    c1, c2, c3 = st.columns(3)

    c1.metric("👴 Pai", moeda(pai_total))
    c2.metric("⛽ Gasolina", moeda(gasolina_total))
    c3.metric("🛠️ Despesas", moeda(despesas_total))

    # -----------------------------------------------------
    # PROGRESSO SEMANAL
    # -----------------------------------------------------

    st.markdown("---")
    st.subheader("📅 Metas da Semana")

    st.caption(
        f"Período: {inicio_semana.strftime('%d/%m/%Y')} "
        f"a {fim_semana.strftime('%d/%m/%Y')}"
    )

    # Faturamento
    st.markdown("### 🚀 Faturamento Bruto")

    progresso_bruto = min(
        bruto_semana / META_BRUTA_SEMANAL,
        1.0,
    )

    falta_bruto = max(
        META_BRUTA_SEMANAL - bruto_semana,
        0.0,
    )

    st.progress(progresso_bruto)

    col_atual, col_meta = st.columns(2)

    with col_atual:
        st.metric(
            "Faturado na semana",
            moeda(bruto_semana),
        )

    with col_meta:
        st.metric(
            "Meta semanal",
            moeda(META_BRUTA_SEMANAL),
        )

    st.caption(
        f"Progresso: {progresso_bruto:.1%} "
        f"| Falta: {moeda(falta_bruto)}"
    )

    # Custos
    st.markdown("### 📈 Retenção e Custos Fixos")

    progresso_custos = min(
        custos_semana / META_CUSTOS_SEMANAL,
        1.0,
    )

    falta_custos = max(
        META_CUSTOS_SEMANAL - custos_semana,
        0.0,
    )

    st.progress(progresso_custos)

    col_garantido, col_falta = st.columns(2)

    with col_garantido:
        st.metric(
            "Garantido na semana",
            moeda(custos_semana),
        )

    with col_falta:
        st.metric(
            "Falta para a meta",
            moeda(falta_custos),
        )

    st.caption(
        f"Progresso: {progresso_custos:.1%} "
        f"| Meta: {moeda(META_CUSTOS_SEMANAL)}"
    )

    if custos_semana >= META_CUSTOS_SEMANAL:
        st.success("Meta semanal de custos atingida!")

    st.metric(
        "Lucro livre nesta semana",
        moeda(lucro_semana),
    )

    # -----------------------------------------------------
    # EXPORTAÇÃO
    # -----------------------------------------------------

    st.markdown("---")
    st.subheader("📥 Exportar Dados")

    df_exportacao = df[COLUNAS_HISTORICO].copy()
    excel_estetico = gerar_excel_estetico(df_exportacao)

    st.download_button(
        label="📊 Baixar Planilha Excel",
        data=excel_estetico,
        file_name=(
            f"financeiro_bajaj_{datetime.now().strftime('%m_%Y')}.xlsx"
        ),
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True,
    )

    # -----------------------------------------------------
    # HISTÓRICO
    # -----------------------------------------------------

    st.markdown("---")
    st.subheader("📋 Histórico de Lançamentos")

    st.dataframe(
        df_exportacao,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Data": st.column_config.TextColumn("Data"),
            "Bruto (R$)": st.column_config.NumberColumn(
                "Bruto (R$)", format="R$ %.2f"
            ),
            "Lucro Livre (R$)": st.column_config.NumberColumn(
                "Lucro Livre (R$)", format="R$ %.2f"
            ),
            "Pai (R$)": st.column_config.NumberColumn(
                "Pai (R$)", format="R$ %.2f"
            ),
            "Gasolina (R$)": st.column_config.NumberColumn(
                "Gasolina (R$)", format="R$ %.2f"
            ),
            "Despesas Gerais (R$)": st.column_config.NumberColumn(
                "Despesas Gerais (R$)", format="R$ %.2f"
            ),
        },
    )

    # -----------------------------------------------------
    # GERENCIAMENTO DE REGISTROS
    # -----------------------------------------------------

    st.markdown("---")
    st.subheader("⚙️ Gerenciar Lançamentos")

    aba_editar, aba_excluir = st.tabs(
        ["✏️ Editar ou Excluir", "🚨 Limpar Histórico"]
    )

    with aba_editar:
        opcoes_registro = [
            (
                f"ID {idx} | {registro['Data']} | "
                f"{moeda(registro['Bruto (R$)'])}"
            )
            for idx, registro in enumerate(historico)
        ]

        registro_selecionado = st.selectbox(
            "Selecione o lançamento",
            options=opcoes_registro,
        )

        id_selecionado = int(
            registro_selecionado.split("|")[0].replace("ID", "").strip()
        )

        registro_atual = historico[id_selecionado]

        col_editar, col_excluir_item = st.columns(2)

        with col_editar:
            st.markdown("**Editar registro**")

            with st.form("form_editar"):
                data_atual = datetime.strptime(
                    registro_atual["Data"],
                    "%d/%m/%Y",
                ).date()

                data_editada = st.date_input(
                    "Corrigir data",
                    value=data_atual,
                    key="data_editada",
                    format="DD/MM/YYYY",
                )

                valor_editado = st.number_input(
                    "Corrigir valor bruto (R$)",
                    min_value=0.0,
                    value=float(registro_atual["Bruto (R$)"]),
                    step=5.0,
                    format="%.2f",
                    key="valor_editado",
                )

                confirmar_edicao = st.form_submit_button(
                    "Salvar Alterações",
                    use_container_width=True,
                )

                if confirmar_edicao:
                    if valor_editado <= 0:
                        st.warning("O valor deve ser maior que zero.")
                    else:
                        historico[id_selecionado] = calcular_registro(
                            valor_editado,
                            data_editada,
                        )

                        salvar_banco(historico)
                        st.success("Registro atualizado!")
                        st.rerun()

        with col_excluir_item:
            st.markdown("**Excluir registro**")

            st.warning(
                "Esta ação remove permanentemente "
                "apenas o lançamento selecionado."
            )

            confirmar_exclusao = st.checkbox(
                "Confirmo a exclusão deste lançamento",
                key="confirmar_exclusao_item",
            )

            if st.button(
                "🗑️ Excluir Item",
                disabled=not confirmar_exclusao,
                use_container_width=True,
            ):
                historico.pop(id_selecionado)
                salvar_banco(historico)

                st.success("Registro excluído!")
                st.rerun()

    with aba_excluir:
        st.warning(
            "A exclusão completa remove todos os lançamentos "
            "salvos no arquivo JSON."
        )

        confirmar_tudo = st.checkbox(
            "Quero apagar permanentemente todo o histórico",
            key="confirmar_exclusao_total",
        )

        if st.button(
            "💥 Apagar Todo o Histórico",
            disabled=not confirmar_tudo,
            use_container_width=True,
        ):
            salvar_banco([])
            st.success("Histórico apagado!")
            st.rerun()

else:
    st.info(
        "Nenhum lançamento cadastrado. "
        "Registre sua primeira corrida para visualizar o painel."
    )
