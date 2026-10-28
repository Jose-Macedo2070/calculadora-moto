import streamlit as st
import pandas as pd
import json
import os

from datetime import datetime, date, time, timedelta
from io import BytesIO

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

# Percentuais de distribuição do valor reservado
PERCENTUAL_RETENCAO = 0.405
PERCENTUAL_PAI = 0.3079
PERCENTUAL_GASOLINA = 0.2222
PERCENTUAL_DESPESAS = 0.4699


# =========================================================
# ESTILO VISUAL E RESPONSIVIDADE
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

        div[data-baseweb="input"] input,
        div[data-baseweb="select"] > div {
            border-radius: 8px;
        }

        .subtitulo {
            color: #94a3b8;
            font-size: 0.95rem;
            margin-top: -10px;
            margin-bottom: 20px;
        }

        .aviso {
            background: #1b2229;
            border: 1px solid #303943;
            padding: 14px;
            border-radius: 10px;
            margin-bottom: 16px;
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


# =========================================================
# FUNÇÕES DE DADOS
# =========================================================

def carregar_dados():
    """Carrega o histórico e mantém compatibilidade com registros antigos."""

    if not os.path.exists(ARQUIVO_DADOS):
        return []

    try:
        with open(ARQUIVO_DADOS, "r", encoding="utf-8") as arquivo:
            dados = json.load(arquivo)

        if not isinstance(dados, list):
            return []

        # Registros antigos podem não ter o campo Hora.
        for registro in dados:
            registro.setdefault("Hora", "")

            for campo in [
                "Bruto (R$)",
                "Lucro Livre (R$)",
                "Pai (R$)",
                "Gasolina (R$)",
                "Despesas Gerais (R$)"
            ]:
                registro.setdefault(campo, 0.0)

        return dados

    except (json.JSONDecodeError, OSError):
        st.error(
            "Não foi possível ler o arquivo de dados. "
            "Verifique se o JSON está válido antes de continuar."
        )
        return []


def salvar_dados(dados):
    """Salva o histórico no arquivo JSON."""

    try:
        with open(ARQUIVO_DADOS, "w", encoding="utf-8") as arquivo:
            json.dump(
                dados,
                arquivo,
                ensure_ascii=False,
                indent=4
            )
        return True

    except OSError as erro:
        st.error(f"Não foi possível salvar os dados: {erro}")
        return False


def converter_data(valor):
    """Converte datas armazenadas para objetos date."""

    if isinstance(valor, date) and not isinstance(valor, datetime):
        return valor

    if isinstance(valor, datetime):
        return valor.date()

    try:
        return datetime.strptime(str(valor), "%Y-%m-%d").date()
    except (ValueError, TypeError):
        try:
            return datetime.fromisoformat(str(valor)).date()
        except (ValueError, TypeError):
            return date.today()


def formatar_moeda(valor):
    """Formata um número como moeda brasileira."""

    valor = float(valor or 0)
    texto = f"{valor:,.2f}"
    texto = texto.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {texto}"


def calcular_corrida(bruto):
    """
    Calcula a retenção e a distribuição dos valores.

    A retenção corresponde a 40,5% do valor bruto.
    O restante é considerado lucro livre.
    """

    bruto = max(float(bruto), 0.0)

    retencao = bruto * PERCENTUAL_RETENCAO
    lucro_livre = bruto - retencao

    pai = retencao * PERCENTUAL_PAI
    gasolina = retencao * PERCENTUAL_GASOLINA

    # Usa a diferença para evitar problemas de arredondamento.
    despesas = retencao - pai - gasolina

    return {
        "Bruto (R$)": round(bruto, 2),
        "Lucro Livre (R$)": round(lucro_livre, 2),
        "Pai (R$)": round(pai, 2),
        "Gasolina (R$)": round(gasolina, 2),
        "Despesas Gerais (R$)": round(despesas, 2)
    }


def criar_dataframe(dados):
    """Prepara os registros para exibição e cálculos."""

    colunas = [
        "Data",
        "Hora",
        "Bruto (R$)",
        "Lucro Livre (R$)",
        "Pai (R$)",
        "Gasolina (R$)",
        "Despesas Gerais (R$)"
    ]

    if not dados:
        return pd.DataFrame(columns=colunas)

    df = pd.DataFrame(dados)

    for coluna in colunas:
        if coluna not in df.columns:
            df[coluna] = "" if coluna in ["Data", "Hora"] else 0.0

    df = df[colunas].copy()

    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")

    for coluna in colunas:
        if coluna.endswith("(R$)"):
            df[coluna] = pd.to_numeric(df[coluna], errors="coerce").fillna(0.0)

    df = df.sort_values("Data", ascending=False, na_position="last")
    return df


def obter_inicio_semana():
    """Retorna a segunda-feira da semana atual."""

    hoje = date.today()
    return hoje - timedelta(days=hoje.weekday())


def obter_resumo_semana(df):
    """Retorna os totais da semana atual, de segunda a domingo."""

    if df.empty:
        return {
            "bruto": 0.0,
            "lucro": 0.0,
            "retencao": 0.0,
            "quantidade": 0
        }

    inicio = obter_inicio_semana()
    fim = inicio + timedelta(days=6)

    datas = df["Data"].dt.date
    semana = df[(datas >= inicio) & (datas <= fim)]

    return {
        "bruto": float(semana["Bruto (R$)"].sum()),
        "lucro": float(semana["Lucro Livre (R$)"].sum()),
        "retencao": float(
            semana[
                [
                    "Pai (R$)",
                    "Gasolina (R$)",
                    "Despesas Gerais (R$)"
                ]
            ].sum().sum()
        ),
        "quantidade": len(semana)
    }


def gerar_excel(df):
    """Gera uma planilha Excel em memória para download."""

    planilha = df.copy()

    if "Data" in planilha.columns:
        planilha["Data"] = planilha["Data"].dt.strftime("%d/%m/%Y")

    colunas_valores = [
        "Bruto (R$)",
        "Lucro Livre (R$)",
        "Pai (R$)",
        "Gasolina (R$)",
        "Despesas Gerais (R$)"
    ]

    for coluna in colunas_valores:
        if coluna in planilha.columns:
            planilha[coluna] = pd.to_numeric(
                planilha[coluna],
                errors="coerce"
            ).fillna(0.0)

    saida = BytesIO()

    with pd.ExcelWriter(saida, engine="openpyxl") as escritor:
        planilha.to_excel(
            escritor,
            index=False,
            sheet_name="Histórico"
        )

        aba = escritor.sheets["Histórico"]
        aba.freeze_panes = "A2"
        aba.auto_filter.ref = aba.dimensions

        for coluna in aba.columns:
            letra = coluna[0].column_letter
            maior = max(
                len(str(celula.value or "")) for celula in coluna
            )
            aba.column_dimensions[letra].width = min(maior + 3, 30)

        for coluna in colunas_valores:
            if coluna in planilha.columns:
                indice = planilha.columns.get_loc(coluna) + 1
                letra = aba.cell(row=1, column=indice).column_letter

                for celula in aba[letra][1:]:
                    celula.number_format = '"R$" #,##0.00'

    saida.seek(0)
    return saida.getvalue()


# =========================================================
# ESTADO DA APLICAÇÃO
# =========================================================

if "dados" not in st.session_state:
    st.session_state.dados = carregar_dados()

if "pagina" not in st.session_state:
    st.session_state.pagina = "Dashboard"

if "confirmar_limpeza" not in st.session_state:
    st.session_state.confirmar_limpeza = False

if "registro_edicao" not in st.session_state:
    st.session_state.registro_edicao = None


def navegar(pagina):
    st.session_state.pagina = pagina


def atualizar_registro(indice, registro):
    dados = st.session_state.dados.copy()
    dados[indice] = registro

    if salvar_dados(dados):
        st.session_state.dados = dados
        st.session_state.registro_edicao = None
        st.success("Registro atualizado com sucesso.")
        st.rerun()


# =========================================================
# CABEÇALHO E NAVEGAÇÃO
# =========================================================

st.title("🏍️ Finanças Bajaj")
st.markdown(
    '<div class="subtitulo">Controle de corridas, ganhos e despesas</div>',
    unsafe_allow_html=True
)

col1, col2 = st.columns(2)

with col1:
    if st.button(
        "＋ Nova corrida",
        type="primary",
        use_container_width=True
    ):
        navegar("Nova corrida")
        st.rerun()

with col2:
    if st.button("▦ Dashboard", use_container_width=True):
        navegar("Dashboard")
        st.rerun()

col3, col4 = st.columns(2)

with col3:
    if st.button("▤ Histórico", use_container_width=True):
        navegar("Histórico")
        st.rerun()

with col4:
    if st.button("⚙ Gerenciar", use_container_width=True):
        navegar("Gerenciar")
        st.rerun()

st.divider()


# =========================================================
# ATUALIZAÇÃO DOS DADOS PARA A TELA
# =========================================================

dados = st.session_state.dados
df = criar_dataframe(dados)


# =========================================================
# TELA: NOVA CORRIDA
# =========================================================

if st.session_state.pagina == "Nova corrida":

    st.header("Registrar nova corrida")
    st.caption("Informe o valor recebido e quando a corrida aconteceu.")

    with st.form("form_nova_corrida", clear_on_submit=True):

        col1, col2 = st.columns(2)

        with col1:
            data_corrida = st.date_input(
                "Data da corrida",
                value=date.today(),
                format="DD/MM/YYYY"
            )

        with col2:
            # IMPORTANTE: use format="24h", não "HH:mm".
            # O parâmetro step foi omitido para maior compatibilidade.
            hora_corrida = st.time_input(
                "Horário da corrida",
                value=datetime.now().time().replace(
                    second=0,
                    microsecond=0
                ),
                format="24h"
            )

        bruto = st.number_input(
            "Valor bruto recebido (R$)",
            min_value=0.0,
            value=0.0,
            step=5.0,
            format="%.2f"
        )

        if bruto > 0:
            previa = calcular_corrida(bruto)

            st.markdown("#### Prévia da distribuição")

            a, b = st.columns(2)

            with a:
                st.metric(
                    "Lucro livre",
                    formatar_moeda(previa["Lucro Livre (R$)"])
                )
                st.metric(
                    "Pai",
                    formatar_moeda(previa["Pai (R$)"])
                )

            with b:
                st.metric(
                    "Gasolina",
                    formatar_moeda(previa["Gasolina (R$)"])
                )
                st.metric(
                    "Despesas gerais",
                    formatar_moeda(previa["Despesas Gerais (R$)"])
                )

        salvar_corrida = st.form_submit_button(
            "Salvar corrida",
            type="primary",
            use_container_width=True
        )

        if salvar_corrida:
            if bruto <= 0:
                st.error("Informe um valor maior que zero.")
            else:
                valores = calcular_corrida(bruto)

                novo_registro = {
                    "Data": data_corrida.isoformat(),
                    "Hora": hora_corrida.strftime("%H:%M"),
                    **valores
                }

                novos_dados = st.session_state.dados.copy()
                novos_dados.append(novo_registro)

                if salvar_dados(novos_dados):
                    st.session_state.dados = novos_dados
                    st.success("Corrida registrada com sucesso.")
                    st.rerun()


# =========================================================
# TELA: DASHBOARD
# =========================================================

elif st.session_state.pagina == "Dashboard":

    st.header("Dashboard")

    resumo = obter_resumo_semana(df)

    bruto_total = (
        float(df["Bruto (R$)"].sum()) if not df.empty else 0.0
    )

    lucro_total = (
        float(df["Lucro Livre (R$)"].sum()) if not df.empty else 0.0
    )

    retencao_total = (
        float(
            df[
                [
                    "Pai (R$)",
                    "Gasolina (R$)",
                    "Despesas Gerais (R$)"
                ]
            ].sum().sum()
        ) if not df.empty else 0.0
    )

    st.subheader("Resumo geral")

    c1, c2 = st.columns(2)

    with c1:
        st.metric("Total bruto", formatar_moeda(bruto_total))

    with c2:
        st.metric("Lucro livre", formatar_moeda(lucro_total))

    c3, c4 = st.columns(2)

    with c3:
        st.metric("Total reservado", formatar_moeda(retencao_total))

    with c4:
        st.metric("Corridas registradas", len(df))

    st.divider()

    st.subheader("Semana atual")

    inicio_semana = obter_inicio_semana()
    fim_semana = inicio_semana + timedelta(days=6)

    st.caption(
        f"Período: {inicio_semana.strftime('%d/%m/%Y')} "
        f"a {fim_semana.strftime('%d/%m/%Y')}"
    )

    c1, c2 = st.columns(2)

    with c1:
        st.metric(
            "Bruto da semana",
            formatar_moeda(resumo["bruto"]),
            delta=f"{resumo['bruto'] - META_BRUTA_SEMANAL:.2f}".replace(
                ".", ","
            ) + " vs. meta"
        )

    with c2:
        st.metric(
            "Lucro livre da semana",
            formatar_moeda(resumo["lucro"])
        )

    progresso_bruto = min(
        resumo["bruto"] / META_BRUTA_SEMANAL, 1.0
    )

    st.caption("Progresso da meta bruta semanal")
    st.progress(progresso_bruto)

    st.caption(
        f"{formatar_moeda(resumo['bruto'])} de "
        f"{formatar_moeda(META_BRUTA_SEMANAL)}"
    )

    st.caption("Progresso da meta de retenção semanal")
    progresso_retencao = min(
        resumo["retencao"] / META_RETENCAO_SEMANAL, 1.0
    )
    st.progress(progresso_retencao)

    st.caption(
        f"{formatar_moeda(resumo['retencao'])} de "
        f"{formatar_moeda(META_RETENCAO_SEMANAL)}"
    )

    st.divider()

    st.subheader("Distribuição dos valores")

    c1, c2 = st.columns(2)

    with c1:
        st.metric(
            "Para o pai",
            formatar_moeda(
                float(df["Pai (R$)"].sum()) if not df.empty else 0.0
            )
        )

        st.metric(
            "Gasolina",
            formatar_moeda(
                float(df["Gasolina (R$)"].sum()) if not df.empty else 0.0
            )
        )

    with c2:
        st.metric(
            "Despesas gerais",
            formatar_moeda(
                float(df["Despesas Gerais (R$)"].sum())
                if not df.empty else 0.0
            )
        )

        st.metric(
            "Retenção total",
            formatar_moeda(retencao_total)
        )

    if not df.empty:
        st.divider()
        st.subheader("Ganhos por dia")

        df_grafico = df.dropna(subset=["Data"]).copy()

        if not df_grafico.empty:
            df_grafico["Dia"] = df_grafico["Data"].dt.strftime("%d/%m")
            ganhos_diarios = (
                df_grafico.groupby("Dia")["Bruto (R$)"]
                .sum()
                .sort_index()
            )

            st.bar_chart(ganhos_diarios)

        st.subheader("Últimas corridas")

        ultimas = df.head(5).copy()
        ultimas["Data"] = ultimas["Data"].dt.strftime("%d/%m/%Y")

        st.dataframe(
            ultimas,
            use_container_width=True,
            hide_index=True
        )

    else:
        st.info(
            "Nenhuma corrida registrada ainda. "
            "Acesse 'Nova corrida' para começar."
        )


# =========================================================
# TELA: HISTÓRICO
# =========================================================

elif st.session_state.pagina == "Histórico":

    st.header("Histórico de corridas")

    if df.empty:
        st.info("Ainda não existem corridas registradas.")

    else:
        col1, col2 = st.columns(2)

        with col1:
            data_inicio = st.date_input(
                "Data inicial",
                value=date.today() - timedelta(days=30),
                key="filtro_inicio",
                format="DD/MM/YYYY"
            )

        with col2:
            data_fim = st.date_input(
                "Data final",
                value=date.today(),
                key="filtro_fim",
                format="DD/MM/YYYY"
            )

        if data_inicio > data_fim:
            st.error("A data inicial não pode ser posterior à data final.")

        else:
            datas = df["Data"].dt.date

            df_filtrado = df[
                (datas >= data_inicio) & (datas <= data_fim)
            ].copy()

            df_exibicao = df_filtrado.copy()
            df_exibicao["Data"] = df_exibicao["Data"].dt.strftime(
                "%d/%m/%Y"
            )

            st.caption(f"{len(df_filtrado)} corrida(s) no período selecionado.")

            st.dataframe(
                df_exibicao,
                use_container_width=True,
                hide_index=True
            )

            st.subheader("Totais do período")

            c1, c2 = st.columns(2)

            with c1:
                st.metric(
                    "Bruto",
                    formatar_moeda(df_filtrado["Bruto (R$)"].sum())
                )

            with c2:
                st.metric(
                    "Lucro livre",
                    formatar_moeda(df_filtrado["Lucro Livre (R$)"].sum())
                )

            # Exportação posicionada logo abaixo da tabela.
            st.download_button(
                label="⬇ Exportar histórico para Excel",
                data=gerar_excel(df_filtrado),
                file_name=(
                    f"historico_bajaj_"
                    f"{data_inicio.strftime('%Y%m%d')}_"
                    f"{data_fim.strftime('%Y%m%d')}.xlsx"
                ),
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True
            )


# =========================================================
# TELA: GERENCIAR
# =========================================================

elif st.session_state.pagina == "Gerenciar":

    st.header("Gerenciar registros")

    if not dados:
        st.info("Não existem registros para gerenciar.")

    else:
        st.subheader("Editar ou excluir uma corrida")

        # Cada opção inclui a posição original para alterar o registro correto.
        opcoes = []

        for indice, registro in enumerate(dados):
            data_registro = converter_data(registro.get("Data", ""))
            hora_registro = registro.get("Hora", "--:--") or "--:--"
            bruto_registro = float(registro.get("Bruto (R$)", 0.0) or 0.0)

            rotulo = (
                f"{indice + 1}. "
                f"{data_registro.strftime('%d/%m/%Y')} às {hora_registro} "
                f"— {formatar_moeda(bruto_registro)}"
            )

            opcoes.append((rotulo, indice))

        escolha = st.selectbox(
            "Selecione o registro",
            options=opcoes,
            format_func=lambda item: item[0],
            key="registro_selecionado"
        )

        indice_selecionado = escolha[1]
        registro_selecionado = dados[indice_selecionado]

        st.markdown("#### Dados da corrida")

        with st.form("form_editar_corrida"):

            col1, col2 = st.columns(2)

            with col1:
                data_atual = converter_data(
                    registro_selecionado.get("Data", "")
                )

                data_editada = st.date_input(
                    "Data",
                    value=data_atual,
                    key=f"data_editar_{indice_selecionado}",
                    format="DD/MM/YYYY"
                )

            with col2:
                hora_texto = registro_selecionado.get("Hora", "") or "00:00"

                try:
                    hora_atual = datetime.strptime(
                        hora_texto, "%H:%M"
                    ).time()
                except ValueError:
                    hora_atual = time(0, 0)

                hora_editada = st.time_input(
                    "Horário",
                    value=hora_atual,
                    format="24h",
                    key=f"hora_editar_{indice_selecionado}"
                )

            bruto_atual = float(
                registro_selecionado.get("Bruto (R$)", 0.0) or 0.0
            )

            bruto_editado = st.number_input(
                "Valor bruto (R$)",
                min_value=0.0,
                value=bruto_atual,
                step=5.0,
                format="%.2f",
                key=f"bruto_editar_{indice_selecionado}"
            )

            salvar_edicao = st.form_submit_button(
                "Salvar alterações",
                type="primary",
                use_container_width=True
            )

            if salvar_edicao:
                if bruto_editado <= 0:
                    st.error("O valor bruto precisa ser maior que zero.")

                else:
                    novos_valores = calcular_corrida(bruto_editado)

                    registro_atualizado = {
                        "Data": data_editada.isoformat(),
                        "Hora": hora_editada.strftime("%H:%M"),
                        **novos_valores
                    }

                    atualizar_registro(
                        indice_selecionado,
                        registro_atualizado
                    )

        st.divider()

        st.subheader("Excluir uma corrida")

        st.warning(
            "A exclusão é permanente no arquivo de dados. "
            "Confira a data e o valor antes de continuar."
        )

        if st.button(
            "Excluir corrida selecionada",
            use_container_width=True
        ):
            st.session_state.confirmar_exclusao = True

        if st.session_state.get("confirmar_exclusao", False):
            st.error("Deseja realmente excluir a corrida selecionada?")

            c1, c2 = st.columns(2)

            with c1:
                if st.button("Sim, excluir", type="primary"):
                    novos_dados = dados.copy()
                    novos_dados.pop(indice_selecionado)

                    if salvar_dados(novos_dados):
                        st.session_state.dados = novos_dados
                        st.session_state.confirmar_exclusao = False
                        st.success("Corrida excluída.")
                        st.rerun()

            with c2:
                if st.button("Cancelar exclusão"):
                    st.session_state.confirmar_exclusao = False
                    st.rerun()

        st.divider()

        st.subheader("Limpar todo o histórico")

        st.warning(
            "Esta ação remove todos os registros do arquivo JSON. "
            "Se precisar deles, exporte o histórico para Excel antes."
        )

        if st.button("Limpar todo o histórico", use_container_width=True):
            st.session_state.confirmar_limpeza = True

        if st.session_state.confirmar_limpeza:
            st.error(
                "Confirma a exclusão de todas as corridas? "
                "Essa ação não poderá ser desfeita."
            )

            c1, c2 = st.columns(2)

            with c1:
                if st.button("Sim, apagar tudo", type="primary"):
                    if salvar_dados([]):
                        st.session_state.dados = []
                        st.session_state.confirmar_limpeza = False
                        st.session_state.confirmar_exclusao = False
                        st.success("Histórico apagado.")
                        st.rerun()

            with c2:
                if st.button("Cancelar limpeza"):
                    st.session_state.confirmar_limpeza = False
                    st.rerun()


# =========================================================
# RODAPÉ
# =========================================================

st.divider()

st.caption(
    "Finanças Bajaj • Controle pessoal de corridas e despesas"
)
