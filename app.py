
import json
import os
from datetime import date, datetime, time, timedelta
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# =========================================================
# CONFIGURAÇÕES
# =========================================================

st.set_page_config(
    page_title="Finanças Bajaj",
    page_icon="🏍️",
    layout="centered",
    initial_sidebar_state="collapsed",
)

ARQUIVO_DADOS = "banco_entregas_bajaj.json"

META_BRUTA_SEMANAL = 1000.00
META_RETENCAO_SEMANAL = 405.00

PERCENTUAL_RETENCAO = 0.405
PERCENTUAL_PAI_RETENCAO = 0.3079
PERCENTUAL_GASOLINA_RETENCAO = 0.2222

APLICATIVOS = [
    "Uber",
    "99",
    "Shopee",
    "iFood",
    "Rappi",
    "Outros",
]


# =========================================================
# ESTILO VISUAL
# =========================================================

st.markdown(
    """
    <style>
    .stApp {
        background-color: #0e1117;
        color: #f1f5f9;
    }

    [data-testid="stHeader"] {
        background-color: #0e1117;
    }

    [data-testid="stSidebar"] {
        background-color: #151a23;
    }

    h1, h2, h3 {
        color: #f8fafc;
    }

    p, label, span {
        color: #dbe2ea;
    }

    [data-testid="stMetric"] {
        background-color: #171d27;
        border: 1px solid #2a3342;
        padding: 14px;
        border-radius: 12px;
    }

    [data-testid="stMetricLabel"] {
        color: #aab7c8;
    }

    [data-testid="stMetricValue"] {
        color: #ffffff;
    }

    div.stButton > button,
    div.stDownloadButton > button {
        border-radius: 10px;
        min-height: 42px;
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

    div[data-baseweb="input"] input,
    div[data-baseweb="select"] > div,
    div[data-baseweb="textarea"] textarea {
        background-color: #171d27;
    }

    [data-testid="stDataFrame"] {
        border: 1px solid #2a3342;
        border-radius: 10px;
        overflow: hidden;
    }

    hr {
        border-color: #2a3342;
    }

    .subtitulo {
        color: #94a3b8;
        font-size: 0.92rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# FUNÇÕES DE DADOS
# =========================================================

def carregar_dados():
    """Carrega os registros do JSON e mantém compatibilidade com dados antigos."""

    if not os.path.exists(ARQUIVO_DADOS):
        return []

    try:
        with open(ARQUIVO_DADOS, "r", encoding="utf-8") as arquivo:
            dados = json.load(arquivo)

        if not isinstance(dados, list):
            return []

        for registro in dados:
            registro.setdefault("Data", "")
            registro.setdefault("Hora", "")
            registro.setdefault("Aplicativo", "")
            registro.setdefault("KM percorridos", None)

            registro.setdefault("Bruto (R$)", 0.0)
            registro.setdefault("Lucro Livre (R$)", 0.0)
            registro.setdefault("Pai (R$)", 0.0)
            registro.setdefault("Gasolina (R$)", 0.0)
            registro.setdefault("Despesas Gerais (R$)", 0.0)

        return dados

    except (json.JSONDecodeError, OSError, TypeError):
        st.error(
            "Não foi possível ler o arquivo de dados. "
            "Verifique se o JSON existe e está válido."
        )
        return []


def salvar_dados(dados):
    """Salva os registros no arquivo JSON."""

    arquivo_temporario = ARQUIVO_DADOS + ".tmp"

    with open(arquivo_temporario, "w", encoding="utf-8") as arquivo:
        json.dump(
            dados,
            arquivo,
            ensure_ascii=False,
            indent=4,
        )

    os.replace(arquivo_temporario, ARQUIVO_DADOS)


def converter_data(valor):
    """Converte datas armazenadas em texto para objetos date."""

    if isinstance(valor, datetime):
        return valor.date()

    if isinstance(valor, date):
        return valor

    if not valor:
        return None

    try:
        return datetime.strptime(str(valor), "%Y-%m-%d").date()
    except ValueError:
        pass

    try:
        return datetime.strptime(str(valor), "%d/%m/%Y").date()
    except ValueError:
        return None


def converter_hora(valor):
    """Converte uma hora salva em texto para objeto time."""

    if isinstance(valor, time):
        return valor

    if not valor:
        return time(0, 0)

    for formato in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(str(valor), formato).time()
        except ValueError:
            continue

    return time(0, 0)


def formatar_moeda(valor):
    """Formata um valor no padrão brasileiro."""

    valor = float(valor or 0)
    texto = f"{valor:,.2f}"
    texto = texto.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {texto}"


def valor_numerico(valor, padrao=0.0):
    """Converte valores para número sem interromper a aplicação."""

    try:
        if valor is None or valor == "":
            return padrao
        return float(valor)
    except (ValueError, TypeError):
        return padrao


def calcular_corrida(bruto):
    """
    Calcula a distribuição financeira de uma corrida.

    Retenção total: 40,5% do bruto.
    Lucro livre: 59,5% do bruto.
    A retenção é dividida entre pai, gasolina e despesas gerais.
    """

    bruto = max(0.0, valor_numerico(bruto))

    retencao = bruto * PERCENTUAL_RETENCAO
    lucro_livre = bruto - retencao

    valor_pai = retencao * PERCENTUAL_PAI_RETENCAO
    valor_gasolina = retencao * PERCENTUAL_GASOLINA_RETENCAO

    despesas_gerais = retencao - valor_pai - valor_gasolina

    return {
        "Lucro Livre (R$)": round(lucro_livre, 2),
        "Pai (R$)": round(valor_pai, 2),
        "Gasolina (R$)": round(valor_gasolina, 2),
        "Despesas Gerais (R$)": round(despesas_gerais, 2),
    }


def criar_dataframe(dados):
    """Cria um DataFrame com colunas padronizadas."""

    colunas = [
        "Data",
        "Hora",
        "Aplicativo",
        "KM percorridos",
        "Bruto (R$)",
        "Lucro Livre (R$)",
        "Pai (R$)",
        "Gasolina (R$)",
        "Despesas Gerais (R$)",
    ]

    if not dados:
        return pd.DataFrame(columns=colunas)

    df = pd.DataFrame(dados)

    for coluna in colunas:
        if coluna not in df.columns:
            df[coluna] = None

    df = df[colunas].copy()

    df["Data"] = df["Data"].apply(converter_data)

    for coluna in [
        "Bruto (R$)",
        "Lucro Livre (R$)",
        "Pai (R$)",
        "Gasolina (R$)",
        "Despesas Gerais (R$)",
        "KM percorridos",
    ]:
        df[coluna] = pd.to_numeric(df[coluna], errors="coerce")

    df["Aplicativo"] = df["Aplicativo"].fillna("").astype(str)
    df["Hora"] = df["Hora"].fillna("").astype(str)

    return df


def obter_inicio_semana(data_referencia=None):
    """Retorna a segunda-feira da semana da data informada."""

    if data_referencia is None:
        data_referencia = date.today()

    return data_referencia - timedelta(days=data_referencia.weekday())


def obter_resumo_semana(dados, data_referencia=None):
    """Calcula os totais da semana atual ou da semana selecionada."""

    if data_referencia is None:
        data_referencia = date.today()

    inicio = obter_inicio_semana(data_referencia)
    fim = inicio + timedelta(days=6)

    df = criar_dataframe(dados)

    if df.empty:
        df_semana = df.copy()
    else:
        df_semana = df[
            df["Data"].apply(
                lambda d: d is not None and inicio <= d <= fim
            )
        ]

    bruto = df_semana["Bruto (R$)"].sum()
    lucro_livre = df_semana["Lucro Livre (R$)"].sum()
    pai = df_semana["Pai (R$)"].sum()
    gasolina = df_semana["Gasolina (R$)"].sum()
    despesas = df_semana["Despesas Gerais (R$)"].sum()

    km = df_semana["KM percorridos"].sum(min_count=1)
    km = 0 if pd.isna(km) else float(km)

    return {
        "inicio": inicio,
        "fim": fim,
        "bruto": float(bruto),
        "lucro_livre": float(lucro_livre),
        "pai": float(pai),
        "gasolina": float(gasolina),
        "despesas": float(despesas),
        "retencao": float(pai + gasolina + despesas),
        "km": km,
        "corridas": len(df_semana),
    }


def gerar_excel(dados):
    """Gera uma planilha Excel formatada."""

    df = criar_dataframe(dados)

    if not df.empty:
        df["Data"] = df["Data"].apply(
            lambda d: d.strftime("%d/%m/%Y") if d else ""
        )

    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as escritor:
        df.to_excel(
            escritor,
            index=False,
            sheet_name="Corridas",
        )

        planilha = escritor.sheets["Corridas"]
        planilha.freeze_panes = "A2"
        planilha.auto_filter.ref = planilha.dimensions

        preenchimento = PatternFill(
            fill_type="solid",
            fgColor="172033",
        )

        for celula in planilha[1]:
            celula.font = Font(bold=True, color="FFFFFF")
            celula.fill = preenchimento
            celula.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )

        for coluna in planilha.columns:
            indice = coluna[0].column
            letra = get_column_letter(indice)

            maior = max(
                len(str(celula.value or ""))
                for celula in coluna
            )

            planilha.column_dimensions[letra].width = min(
                max(maior + 2, 12),
                28,
            )

        colunas_monetarias = [
            "Bruto (R$)",
            "Lucro Livre (R$)",
            "Pai (R$)",
            "Gasolina (R$)",
            "Despesas Gerais (R$)",
        ]

        mapa_colunas = {
            celula.value: celula.column
            for celula in planilha[1]
        }

        for nome_coluna in colunas_monetarias:
            if nome_coluna in mapa_colunas:
                indice = mapa_colunas[nome_coluna]

                for linha in range(2, planilha.max_row + 1):
                    planilha.cell(
                        row=linha,
                        column=indice,
                    ).number_format = '"R$" #,##0.00'

        if "KM percorridos" in mapa_colunas:
            indice = mapa_colunas["KM percorridos"]

            for linha in range(2, planilha.max_row + 1):
                planilha.cell(
                    row=linha,
                    column=indice,
                ).number_format = '#,##0.0'

    buffer.seek(0)
    return buffer


def nome_aplicativo_selecionado(aplicativo, outro_nome):
    """Retorna o nome final do aplicativo."""

    if aplicativo == "Outros":
        return outro_nome.strip() if outro_nome.strip() else "Outros"

    return aplicativo


# =========================================================
# ESTADO DA APLICAÇÃO
# =========================================================

if "dados" not in st.session_state:
    st.session_state.dados = carregar_dados()

if "pagina" not in st.session_state:
    st.session_state.pagina = "Nova corrida"

if "nova_corrida_version" not in st.session_state:
    st.session_state.nova_corrida_version = 0

if "mensagem_sucesso_corrida" not in st.session_state:
    st.session_state.mensagem_sucesso_corrida = None

if "mensagem_geral" not in st.session_state:
    st.session_state.mensagem_geral = None


# =========================================================
# CABEÇALHO E NAVEGAÇÃO
# =========================================================

st.title("🏍️ Finanças Bajaj")
st.markdown(
    '<p class="subtitulo">Controle financeiro das suas corridas</p>',
    unsafe_allow_html=True,
)

col1, col2 = st.columns(2)

with col1:
    if st.button(
        "➕ Nova corrida",
        use_container_width=True,
        type="primary" if st.session_state.pagina == "Nova corrida" else "secondary",
    ):
        st.session_state.pagina = "Nova corrida"
        st.rerun()

with col2:
    if st.button(
        "📊 Dashboard",
        use_container_width=True,
        type="primary" if st.session_state.pagina == "Dashboard" else "secondary",
    ):
        st.session_state.pagina = "Dashboard"
        st.rerun()

col3, col4 = st.columns(2)

with col3:
    if st.button(
        "📋 Histórico",
        use_container_width=True,
        type="primary" if st.session_state.pagina == "Histórico" else "secondary",
    ):
        st.session_state.pagina = "Histórico"
        st.rerun()

with col4:
    if st.button(
        "⚙️ Gerenciar",
        use_container_width=True,
        type="primary" if st.session_state.pagina == "Gerenciar" else "secondary",
    ):
        st.session_state.pagina = "Gerenciar"
        st.rerun()

st.divider()

if st.session_state.mensagem_geral:
    st.success(st.session_state.mensagem_geral)
    st.session_state.mensagem_geral = None

if st.session_state.mensagem_sucesso_corrida:
    st.success(st.session_state.mensagem_sucesso_corrida)
    st.session_state.mensagem_sucesso_corrida = None


# =========================================================
# PÁGINA: NOVA CORRIDA
# =========================================================

if st.session_state.pagina == "Nova corrida":
    st.header("Nova corrida")
    st.caption("Informe os dados da corrida para registrar seus ganhos.")

    versao = st.session_state.nova_corrida_version

    with st.container(border=True):
        data_corrida = st.date_input(
            "Data da corrida",
            value=date.today(),
            key=f"nova_data_{versao}",
            format="DD/MM/YYYY",
        )

        hora_corrida = st.time_input(
            "Horário",
            value=datetime.now().time().replace(second=0, microsecond=0),
            key=f"nova_hora_{versao}",
            format="24h",
        )

        aplicativo = st.selectbox(
            "Aplicativo",
            options=APLICATIVOS,
            key=f"novo_aplicativo_{versao}",
        )

        outro_nome = ""

        if aplicativo == "Outros":
            outro_nome = st.text_input(
                "Nome do aplicativo",
                placeholder="Ex.: entrega particular",
                key=f"novo_outro_nome_{versao}",
            )

        km_informado = st.number_input(
            "KM percorridos nesta corrida",
            min_value=0.0,
            value=0.0,
            step=0.1,
            format="%.1f",
            key=f"novo_km_{versao}",
            help="Informe a distância percorrida nesta corrida.",
        )

        bruto = st.number_input(
            "Valor bruto recebido (R$)",
            min_value=0.0,
            value=0.0,
            step=1.0,
            format="%.2f",
            key=f"novo_bruto_{versao}",
        )

        if bruto > 0:
            calculo = calcular_corrida(bruto)

            st.markdown("#### Prévia da distribuição")

            c1, c2 = st.columns(2)

            with c1:
                st.metric(
                    "Lucro livre",
                    formatar_moeda(calculo["Lucro Livre (R$)"]),
                )
                st.metric(
                    "Para seu pai",
                    formatar_moeda(calculo["Pai (R$)"]),
                )

            with c2:
                st.metric(
                    "Gasolina",
                    formatar_moeda(calculo["Gasolina (R$)"]),
                )
                st.metric(
                    "Despesas gerais",
                    formatar_moeda(calculo["Despesas Gerais (R$)"]),
                )

        salvar = st.button(
            "💾 Salvar corrida",
            type="primary",
            use_container_width=True,
            key=f"salvar_corrida_{versao}",
        )

        if salvar:
            if bruto <= 0:
                st.error("Informe um valor bruto maior que zero.")
            elif aplicativo == "Outros" and not outro_nome.strip():
                st.error("Informe o nome do aplicativo selecionado em Outros.")
            else:
                distribuicao = calcular_corrida(bruto)

                novo_registro = {
                    "Data": data_corrida.strftime("%Y-%m-%d"),
                    "Hora": hora_corrida.strftime("%H:%M"),
                    "Aplicativo": nome_aplicativo_selecionado(
                        aplicativo,
                        outro_nome,
                    ),
                    "KM percorridos": round(float(km_informado), 1),
                    "Bruto (R$)": round(float(bruto), 2),
                    "Lucro Livre (R$)": distribuicao["Lucro Livre (R$)"],
                    "Pai (R$)": distribuicao["Pai (R$)"],
                    "Gasolina (R$)": distribuicao["Gasolina (R$)"],
                    "Despesas Gerais (R$)": distribuicao["Despesas Gerais (R$)"],
                }

                novos_dados = list(st.session_state.dados)
                novos_dados.append(novo_registro)

                try:
                    salvar_dados(novos_dados)

                    st.session_state.dados = novos_dados
                    st.session_state.nova_corrida_version += 1
                    st.session_state.mensagem_sucesso_corrida = (
                        "Corrida salva com sucesso! "
                        "O formulário foi limpo e está pronto para a próxima."
                    )
                    st.rerun()

                except OSError as erro:
                    st.error(f"Não foi possível salvar os dados: {erro}")


# =========================================================
# PÁGINA: DASHBOARD
# =========================================================

elif st.session_state.pagina == "Dashboard":
    st.header("Dashboard")

    dados = st.session_state.dados
    df = criar_dataframe(dados)

    resumo = obter_resumo_semana(dados)

    st.subheader("Resumo da semana")
    st.caption(
        f"{resumo['inicio'].strftime('%d/%m/%Y')} até "
        f"{resumo['fim'].strftime('%d/%m/%Y')}"
    )

    c1, c2 = st.columns(2)

    with c1:
        st.metric(
            "Bruto semanal",
            formatar_moeda(resumo["bruto"]),
            delta=f"Meta: {formatar_moeda(META_BRUTA_SEMANAL)}",
        )

    with c2:
        st.metric(
            "Lucro livre",
            formatar_moeda(resumo["lucro_livre"]),
        )

    c3, c4 = st.columns(2)

    with c3:
        st.metric(
            "Retenção acumulada",
            formatar_moeda(resumo["retencao"]),
            delta=f"Meta: {formatar_moeda(META_RETENCAO_SEMANAL)}",
        )

    with c4:
        st.metric(
            "KM na semana",
            f"{resumo['km']:,.1f} km".replace(",", "X").replace(".", ",").replace("X", "."),
        )

    percentual_bruto = min(
        resumo["bruto"] / META_BRUTA_SEMANAL,
        1.0,
    )

    percentual_retencao = min(
        resumo["retencao"] / META_RETENCAO_SEMANAL,
        1.0,
    )

    st.markdown("#### Progresso das metas")

    st.write(
        f"Meta de bruto: {percentual_bruto * 100:.1f}% concluída"
    )
    st.progress(percentual_bruto)

    st.write(
        f"Meta de retenção: {percentual_retencao * 100:.1f}% concluída"
    )
    st.progress(percentual_retencao)

    st.divider()

    st.subheader("Distribuição do dinheiro na semana")

    distribuicao = pd.DataFrame(
        {
            "Categoria": [
                "Lucro livre",
                "Pai",
                "Gasolina",
                "Despesas gerais",
            ],
            "Valor (R$)": [
                resumo["lucro_livre"],
                resumo["pai"],
                resumo["gasolina"],
                resumo["despesas"],
            ],
        }
    )

    st.bar_chart(
        distribuicao.set_index("Categoria"),
        y="Valor (R$)",
    )

    st.divider()

    st.subheader("Ganhos por aplicativo")

    if df.empty or df["Aplicativo"].replace("", pd.NA).dropna().empty:
        st.info("Cadastre corridas para visualizar os ganhos por aplicativo.")
    else:
        df_apps = df.copy()
        df_apps["Aplicativo"] = df_apps["Aplicativo"].replace(
            "",
            "Não informado",
        )

        ganhos_apps = (
            df_apps.groupby("Aplicativo", as_index=False)["Bruto (R$)"]
            .sum()
            .sort_values("Bruto (R$)", ascending=False)
            .set_index("Aplicativo")
        )

        st.bar_chart(ganhos_apps, y="Bruto (R$)")

        st.dataframe(
            ganhos_apps.rename(
                columns={"Bruto (R$)": "Bruto total (R$)"}
            ),
            use_container_width=True,
        )

    st.divider()

    st.subheader("Totais gerais")

    c5, c6 = st.columns(2)

    with c5:
        bruto_total = df["Bruto (R$)"].sum() if not df.empty else 0
        st.metric("Bruto total", formatar_moeda(bruto_total))

    with c6:
        lucro_total = df["Lucro Livre (R$)"].sum() if not df.empty else 0
        st.metric("Lucro livre total", formatar_moeda(lucro_total))

    c7, c8 = st.columns(2)

    with c7:
        st.metric("Corridas registradas", str(len(df)))

    with c8:
        km_total = df["KM percorridos"].sum(min_count=1) if not df.empty else 0
        km_total = 0 if pd.isna(km_total) else float(km_total)
        km_formatado = f"{km_total:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
        st.metric("KM registrados", f"{km_formatado} km")


# =========================================================
# PÁGINA: HISTÓRICO
# =========================================================

elif st.session_state.pagina == "Histórico":
    st.header("Histórico de corridas")

    dados = st.session_state.dados
    df = criar_dataframe(dados)

    if df.empty:
        st.info("Ainda não existem corridas cadastradas.")
    else:
        datas_validas = df["Data"].dropna()

        if datas_validas.empty:
            st.info("Não há datas válidas para filtrar os registros.")
            df_filtrado = df.copy()
        else:
            menor_data = min(datas_validas)
            maior_data = max(datas_validas)

            st.subheader("Filtrar período")

            col1, col2 = st.columns(2)

            with col1:
                data_inicio = st.date_input(
                    "Data inicial",
                    value=menor_data,
                    min_value=menor_data,
                    max_value=maior_data,
                    key="historico_data_inicio",
                    format="DD/MM/YYYY",
                )

            with col2:
                data_fim = st.date_input(
                    "Data final",
                    value=maior_data,
                    min_value=menor_data,
                    max_value=maior_data,
                    key="historico_data_fim",
                    format="DD/MM/YYYY",
                )

            if data_inicio > data_fim:
                st.error("A data inicial não pode ser posterior à data final.")
                df_filtrado = df.iloc[0:0].copy()
            else:
                df_filtrado = df[
                    df["Data"].apply(
                        lambda d: d is not None and data_inicio <= d <= data_fim
                    )
                ].copy()

        st.divider()

        if not df_filtrado.empty:
            df_exibicao = df_filtrado.copy()

            df_exibicao["Data"] = df_exibicao["Data"].apply(
                lambda d: d.strftime("%d/%m/%Y") if d else ""
            )

            df_exibicao = df_exibicao.sort_values(
                by=["Data", "Hora"],
                ascending=False,
            )

            st.dataframe(
                df_exibicao,
                use_container_width=True,
                hide_index=True,
            )

            bruto_periodo = df_filtrado["Bruto (R$)"].sum()
            lucro_periodo = df_filtrado["Lucro Livre (R$)"].sum()
            km_periodo = df_filtrado["KM percorridos"].sum(min_count=1)
            km_periodo = 0 if pd.isna(km_periodo) else float(km_periodo)

            c1, c2 = st.columns(2)

            with c1:
                st.metric(
                    "Bruto no período",
                    formatar_moeda(bruto_periodo),
                )

            with c2:
                st.metric(
                    "Lucro livre no período",
                    formatar_moeda(lucro_periodo),
                )

            st.metric("Corridas no período", str(len(df_filtrado)))
            st.metric("KM no período", f"{km_periodo:.1f} km")

            dados_excel = df_filtrado.to_dict(orient="records")
            excel = gerar_excel(dados_excel)

            st.download_button(
                label="📥 Exportar período para Excel",
                data=excel,
                file_name=(
                    f"corridas_{data_inicio.strftime('%Y%m%d')}_"
                    f"{data_fim.strftime('%Y%m%d')}.xlsx"
                ),
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True,
            )
        else:
            st.info("Nenhuma corrida encontrada nesse período.")


# =========================================================
# PÁGINA: GERENCIAR
# =========================================================

elif st.session_state.pagina == "Gerenciar":
    st.header("Gerenciar corridas")

    dados = st.session_state.dados

    if not dados:
        st.info("Não existem corridas para editar ou excluir.")
    else:
        opcoes_registros = []

        for indice, registro in enumerate(dados):
            data_registro = converter_data(registro.get("Data"))
            data_texto = (
                data_registro.strftime("%d/%m/%Y")
                if data_registro
                else "Data inválida"
            )

            hora_texto = registro.get("Hora", "")
            aplicativo_texto = registro.get("Aplicativo", "Não informado")
            bruto_texto = formatar_moeda(
                registro.get("Bruto (R$)", 0)
            )

            opcoes_registros.append(
                f"{indice} | {data_texto} {hora_texto} | "
                f"{aplicativo_texto} | {bruto_texto}"
            )

        registro_escolhido = st.selectbox(
            "Selecione uma corrida",
            options=range(len(opcoes_registros)),
            format_func=lambda i: opcoes_registros[i],
            key="gerenciar_registro_selecionado",
        )

        registro = dados[registro_escolhido]

        st.subheader("Editar corrida")

        data_original = converter_data(registro.get("Data")) or date.today()
        hora_original = converter_hora(registro.get("Hora"))

        aplicativo_original = registro.get("Aplicativo", "")
        if aplicativo_original in APLICATIVOS:
            aplicativo_padrao = aplicativo_original
            outro_original = ""
        else:
            aplicativo_padrao = (
                "Outros" if aplicativo_original else "Uber"
            )
            outro_original = aplicativo_original

        try:
            km_original = float(registro.get("KM percorridos") or 0)
        except (TypeError, ValueError):
            km_original = 0.0

        bruto_original = valor_numerico(
            registro.get("Bruto (R$)"),
            0.0,
        )

        chave_edicao = f"{registro_escolhido}_{len(dados)}"

        with st.form(key=f"form_editar_{chave_edicao}"):
            data_editada = st.date_input(
                "Data",
                value=data_original,
                format="DD/MM/YYYY",
            )

            hora_editada = st.time_input(
                "Horário",
                value=hora_original,
                format="24h",
            )

            aplicativo_editado = st.selectbox(
                "Aplicativo",
                options=APLICATIVOS,
                index=APLICATIVOS.index(aplicativo_padrao),
            )

            outro_nome_editado = ""

            if aplicativo_editado == "Outros":
                outro_nome_editado = st.text_input(
                    "Nome do aplicativo",
                    value=outro_original,
                )

            km_editado = st.number_input(
                "KM percorridos",
                min_value=0.0,
                value=max(0.0, km_original),
                step=0.1,
                format="%.1f",
            )

            bruto_editado = st.number_input(
                "Valor bruto (R$)",
                min_value=0.0,
                value=max(0.0, bruto_original),
                step=1.0,
                format="%.2f",
            )

            salvar_edicao = st.form_submit_button(
                "Salvar alterações",
                type="primary",
                use_container_width=True,
            )

            if salvar_edicao:
                if bruto_editado <= 0:
                    st.error("O valor bruto deve ser maior que zero.")
                elif aplicativo_editado == "Outros" and not outro_nome_editado.strip():
                    st.error("Informe o nome do aplicativo.")
                else:
                    distribuicao = calcular_corrida(bruto_editado)

                    dados_atualizados = list(st.session_state.dados)

                    dados_atualizados[registro_escolhido] = {
                        "Data": data_editada.strftime("%Y-%m-%d"),
                        "Hora": hora_editada.strftime("%H:%M"),
                        "Aplicativo": nome_aplicativo_selecionado(
                            aplicativo_editado,
                            outro_nome_editado,
                        ),
                        "KM percorridos": round(float(km_editado), 1),
                        "Bruto (R$)": round(float(bruto_editado), 2),
                        "Lucro Livre (R$)": distribuicao["Lucro Livre (R$)"],
                        "Pai (R$)": distribuicao["Pai (R$)"],
                        "Gasolina (R$)": distribuicao["Gasolina (R$)"],
                        "Despesas Gerais (R$)": distribuicao["Despesas Gerais (R$)"],
                    }

                    try:
                        salvar_dados(dados_atualizados)
                        st.session_state.dados = dados_atualizados
                        st.session_state.mensagem_geral = (
                            "Corrida atualizada com sucesso."
                        )
                        st.rerun()
                    except OSError as erro:
                        st.error(f"Erro ao salvar alterações: {erro}")

        st.divider()

        st.subheader("Excluir corrida")

        st.warning(
            "A exclusão remove permanentemente a corrida selecionada do arquivo de dados."
        )

        confirmar_exclusao = st.checkbox(
            "Confirmo que desejo excluir esta corrida.",
            key=f"confirmar_exclusao_{chave_edicao}",
        )

        if st.button(
            "🗑️ Excluir corrida selecionada",
            disabled=not confirmar_exclusao,
            use_container_width=True,
        ):
            dados_atualizados = list(st.session_state.dados)
            dados_atualizados.pop(registro_escolhido)

            try:
                salvar_dados(dados_atualizados)
                st.session_state.dados = dados_atualizados
                st.session_state.mensagem_geral = (
                    "Corrida excluída com sucesso."
                )
                st.rerun()
            except OSError as erro:
                st.error(f"Não foi possível excluir a corrida: {erro}")

        st.divider()

        st.subheader("Apagar todo o histórico")

        st.warning(
            "Esta ação apaga todas as corridas registradas. "
            "Ela não pode ser desfeita."
        )

        confirmar_limpeza = st.checkbox(
            "Confirmo que desejo apagar todo o histórico.",
            key="confirmar_limpeza_total",
        )

        if st.button(
            "🧹 Apagar todo o histórico",
            disabled=not confirmar_limpeza,
            use_container_width=True,
        ):
            try:
                salvar_dados([])
                st.session_state.dados = []
                st.session_state.mensagem_geral = (
                    "Todo o histórico foi apagado."
                )
                st.rerun()
            except OSError as erro:
                st.error(f"Não foi possível apagar o histórico: {erro}")
