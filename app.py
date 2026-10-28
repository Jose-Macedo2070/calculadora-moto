
import streamlit as st
import pandas as pd
import json
import os

from datetime import datetime, date, time, timedelta
from io import BytesIO
from streamlit_javascript import st_javascript

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

APLICATIVOS = ["Uber", "99", "Shopee", "iFood", "Rappi", "Outros"]


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

        .subtitulo {
            color: #94a3b8;
            font-size: 0.95rem;
            margin-top: -10px;
            margin-bottom: 20px;
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
# HORÁRIO LOCAL DO APARELHO
# =========================================================

def obter_data_hora_aparelho():
    """Obtém a data/hora do navegador; usa o horário do servidor como reserva."""
    try:
        valor = st_javascript("new Date().toLocaleString('sv-SE')")
        if isinstance(valor, str) and valor.strip():
            return datetime.strptime(valor.strip(), "%Y-%m-%d %H:%M:%S")
    except Exception:
        pass

    return datetime.now()


# =========================================================
# DADOS
# =========================================================

def carregar_dados():
    if not os.path.exists(ARQUIVO_DADOS):
        return []

    try:
        with open(ARQUIVO_DADOS, "r", encoding="utf-8") as arquivo:
            dados = json.load(arquivo)

        if not isinstance(dados, list):
            return []

        # Compatibilidade com registros antigos.
        for registro in dados:
            registro.setdefault("Hora", "")
            registro.setdefault("Aplicativo", "")
            registro.setdefault("KM percorridos", None)

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
            "Verifique se o arquivo JSON está válido."
        )
        return []


def salvar_dados(dados):
    try:
        with open(ARQUIVO_DADOS, "w", encoding="utf-8") as arquivo:
            json.dump(dados, arquivo, ensure_ascii=False, indent=4)
        return True

    except OSError as erro:
        st.error(f"Não foi possível salvar os dados: {erro}")
        return False


def converter_data(valor):
    if isinstance(valor, datetime):
        return valor.date()

    if isinstance(valor, date):
        return valor

    try:
        return datetime.strptime(str(valor), "%Y-%m-%d").date()
    except (ValueError, TypeError):
        try:
            return datetime.fromisoformat(str(valor)).date()
        except (ValueError, TypeError):
            return date.today()


def formatar_moeda(valor):
    valor = float(valor or 0)
    texto = f"{valor:,.2f}"
    texto = texto.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {texto}"


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


def criar_dataframe(dados):
    colunas = [
        "Data",
        "Hora",
        "Aplicativo",
        "KM percorridos",
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
            df[coluna] = None if coluna == "KM percorridos" else (
                "" if coluna in ["Data", "Hora", "Aplicativo"] else 0.0
            )

    df = df[colunas].copy()
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")

    # KM antigos permanecem vazios; não são convertidos em zero.
    df["KM percorridos"] = pd.to_numeric(
        df["KM percorridos"], errors="coerce"
    )

    for coluna in [
        "Bruto (R$)",
        "Lucro Livre (R$)",
        "Pai (R$)",
        "Gasolina (R$)",
        "Despesas Gerais (R$)"
    ]:
        df[coluna] = pd.to_numeric(
            df[coluna], errors="coerce"
        ).fillna(0.0)

    return df.sort_values("Data", ascending=False, na_position="last")


def obter_inicio_semana():
    hoje = date.today()
    return hoje - timedelta(days=hoje.weekday())


def obter_resumo_semana(df):
    if df.empty:
        return {"bruto": 0.0, "lucro": 0.0, "retencao": 0.0}

    inicio = obter_inicio_semana()
    fim = inicio + timedelta(days=6)

    datas = df["Data"].dt.date
    semana = df[(datas >= inicio) & (datas <= fim)]

    retencao = semana[
        ["Pai (R$)", "Gasolina (R$)", "Despesas Gerais (R$)"]
    ].sum().sum()

    return {
        "bruto": float(semana["Bruto (R$)"].sum()),
        "lucro": float(semana["Lucro Livre (R$)"].sum()),
        "retencao": float(retencao)
    }


def gerar_excel(df):
    planilha = df.copy()

    if "Data" in planilha.columns:
        planilha["Data"] = planilha["Data"].dt.strftime("%d/%m/%Y")

    for coluna in [
        "Bruto (R$)",
        "Lucro Livre (R$)",
        "Pai (R$)",
        "Gasolina (R$)",
        "Despesas Gerais (R$)"
    ]:
        planilha[coluna] = pd.to_numeric(
            planilha[coluna], errors="coerce"
        ).fillna(0.0)

    planilha["KM percorridos"] = pd.to_numeric(
        planilha["KM percorridos"], errors="coerce"
    )

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
            maior = max(len(str(c.value or "")) for c in coluna)
            aba.column_dimensions[letra].width = min(maior + 3, 30)

        for coluna in [
            "Bruto (R$)",
            "Lucro Livre (R$)",
            "Pai (R$)",
            "Gasolina (R$)",
            "Despesas Gerais (R$)"
        ]:
            indice = planilha.columns.get_loc(coluna) + 1
            letra = aba.cell(row=1, column=indice).column_letter

            for celula in aba[letra][1:]:
                celula.number_format = '"R$" #,##0.00'

        indice_km = planilha.columns.get_loc("KM percorridos") + 1
        letra_km = aba.cell(row=1, column=indice_km).column_letter

        for celula in aba[letra_km][1:]:
            celula.number_format = "0.00"

    saida.seek(0)
    return saida.getvalue()


def nome_aplicativo_selecionado(opcao, outro_nome):
    if opcao == "Outros":
        return outro_nome.strip()
    return opcao


# =========================================================
# ESTADO DA APLICAÇÃO
# =========================================================

if "dados" not in st.session_state:
    st.session_state.dados = carregar_dados()

if "pagina" not in st.session_state:
    st.session_state.pagina = "Dashboard"

if "confirmar_limpeza" not in st.session_state:
    st.session_state.confirmar_limpeza = False

if "confirmar_exclusao" not in st.session_state:
    st.session_state.confirmar_exclusao = False

if "nova_corrida_version" not in st.session_state:
    st.session_state.nova_corrida_version = 0

if "mensagem_sucesso_corrida" not in st.session_state:
    st.session_state.mensagem_sucesso_corrida = None


def navegar(pagina):
    st.session_state.pagina = pagina


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
    if st.button("＋ Nova corrida", type="primary", use_container_width=True):
        # Cada clique inicia um formulário novo, sem reaproveitar campos anteriores.
        st.session_state.nova_corrida_version += 1
        st.session_state.pagina = "Nova corrida"
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

if st.session_state.mensagem_sucesso_corrida:
    st.success(st.session_state.mensagem_sucesso_corrida)
    st.session_state.mensagem_sucesso_corrida = None

dados = st.session_state.dados
df = criar_dataframe(dados)


# =========================================================
# NOVA CORRIDA
# =========================================================

if st.session_state.pagina == "Nova corrida":

    st.header("Registrar nova corrida")
    st.caption("Registre os ganhos, o aplicativo e a distância percorrida.")

    versao_corrida = st.session_state.nova_corrida_version
    data_hora_aparelho = obter_data_hora_aparelho()

    col1, col2 = st.columns(2)

    with col1:
        data_corrida = st.date_input(
            "Data da corrida",
            value=data_hora_aparelho.date(),
            format="DD/MM/YYYY",
            key=f"nova_data_{versao_corrida}"
        )

    with col2:
        hora_corrida = st.time_input(
            "Horário da corrida",
            value=data_hora_aparelho.time().replace(second=0, microsecond=0),
            format="24h",
            key=f"nova_hora_{versao_corrida}"
        )

    aplicativo_opcao = st.selectbox(
        "Aplicativo utilizado",
        APLICATIVOS,
        key=f"novo_aplicativo_{versao_corrida}"
    )

    outro_aplicativo = ""

    if aplicativo_opcao == "Outros":
        outro_aplicativo = st.text_input(
            "Digite o nome do aplicativo",
            placeholder="Ex.: Loggi",
            key=f"novo_outro_aplicativo_{versao_corrida}"
        )

    col1, col2 = st.columns(2)

    with col1:
        bruto = st.number_input(
            "Valor bruto recebido (R$)",
            min_value=0.0,
            value=0.0,
            step=5.0,
            format="%.2f",
            key=f"novo_bruto_{versao_corrida}"
        )

    with col2:
        km_percorridos = st.number_input(
            "KM percorridos na corrida",
            min_value=0.0,
            value=0.0,
            step=0.5,
            format="%.2f",
            help="Informe a distância percorrida nessa corrida.",
            key=f"novo_km_{versao_corrida}"
        )

    if bruto > 0:
        previa = calcular_corrida(bruto)

        st.subheader("Prévia da distribuição")
        c1, c2 = st.columns(2)

        with c1:
            st.metric("Lucro livre", formatar_moeda(previa["Lucro Livre (R$)"]))
            st.metric("Pai", formatar_moeda(previa["Pai (R$)"]))

        with c2:
            st.metric("Gasolina", formatar_moeda(previa["Gasolina (R$)"]))
            st.metric("Despesas gerais", formatar_moeda(previa["Despesas Gerais (R$)"]))

        if km_percorridos > 0:
            st.metric(
                "Faturamento bruto por KM",
                formatar_moeda(bruto / km_percorridos) + "/km"
            )

    if st.button(
        "Salvar corrida",
        type="primary",
        use_container_width=True,
        key=f"salvar_corrida_{versao_corrida}"
    ):

        aplicativo = nome_aplicativo_selecionado(
            aplicativo_opcao,
            outro_aplicativo
        )

        if bruto <= 0:
            st.error("Informe um valor bruto maior que zero.")

        elif not aplicativo:
            st.error("Digite o nome do aplicativo.")

        else:
            valores = calcular_corrida(bruto)

            novo_registro = {
                "Data": data_corrida.isoformat(),
                "Hora": hora_corrida.strftime("%H:%M"),
                "Aplicativo": aplicativo,
                "KM percorridos": (
                    round(km_percorridos, 2)
                    if km_percorridos > 0 else None
                ),
                **valores
            }

            novos_dados = st.session_state.dados.copy()
            novos_dados.append(novo_registro)

            if salvar_dados(novos_dados):
                st.session_state.dados = novos_dados
                # O próximo formulário recebe novas chaves e volta aos valores padrão.
                st.session_state.nova_corrida_version += 1
                st.session_state.mensagem_sucesso_corrida = (
                    "Corrida registrada com sucesso. O formulário foi limpo "
                    "e está pronto para a próxima corrida."
                )
                st.session_state.pagina = "Nova corrida"
                st.rerun()


# =========================================================
# DASHBOARD
# =========================================================

elif st.session_state.pagina == "Dashboard":

    st.header("Dashboard")
    resumo = obter_resumo_semana(df)

    bruto_total = float(df["Bruto (R$)"].sum()) if not df.empty else 0.0
    lucro_total = float(df["Lucro Livre (R$)"].sum()) if not df.empty else 0.0

    retencao_total = (
        float(df[["Pai (R$)", "Gasolina (R$)", "Despesas Gerais (R$)"]].sum().sum())
        if not df.empty else 0.0
    )

    km_total = (
        float(df["KM percorridos"].sum(min_count=1))
        if not df.empty and df["KM percorridos"].notna().any()
        else 0.0
    )

    st.subheader("Resumo geral")
    c1, c2 = st.columns(2)

    with c1:
        st.metric("Total bruto", formatar_moeda(bruto_total))

    with c2:
        st.metric("Lucro livre", formatar_moeda(lucro_total))

    c1, c2 = st.columns(2)

    with c1:
        st.metric("Total reservado", formatar_moeda(retencao_total))

    with c2:
        st.metric("KM registrados", f"{km_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    st.divider()
    st.subheader("Semana atual")

    inicio = obter_inicio_semana()
    fim = inicio + timedelta(days=6)

    st.caption(f"Período: {inicio.strftime('%d/%m/%Y')} a {fim.strftime('%d/%m/%Y')}")

    c1, c2 = st.columns(2)

    with c1:
        st.metric("Bruto da semana", formatar_moeda(resumo["bruto"]))

    with c2:
        st.metric("Lucro livre da semana", formatar_moeda(resumo["lucro"]))

    st.caption("Progresso da meta bruta semanal")
    st.progress(min(resumo["bruto"] / META_BRUTA_SEMANAL, 1.0))
    st.caption(f"{formatar_moeda(resumo['bruto'])} de {formatar_moeda(META_BRUTA_SEMANAL)}")

    st.caption("Progresso da meta de retenção semanal")
    st.progress(min(resumo["retencao"] / META_RETENCAO_SEMANAL, 1.0))
    st.caption(f"{formatar_moeda(resumo['retencao'])} de {formatar_moeda(META_RETENCAO_SEMANAL)}")

    if not df.empty:
        st.divider()
        st.subheader("Distribuição dos valores")

        c1, c2 = st.columns(2)

        with c1:
            st.metric("Para o pai", formatar_moeda(df["Pai (R$)"].sum()))
            st.metric("Gasolina", formatar_moeda(df["Gasolina (R$)"].sum()))

        with c2:
            st.metric("Despesas gerais", formatar_moeda(df["Despesas Gerais (R$)"].sum()))
            st.metric("Retenção total", formatar_moeda(retencao_total))

        st.divider()
        st.subheader("Ganhos por aplicativo")

        por_app = (
            df[df["Aplicativo"].fillna("").astype(str).str.strip() != ""]
            .groupby("Aplicativo")["Bruto (R$)"]
            .sum()
            .sort_values(ascending=False)
        )

        if not por_app.empty:
            st.bar_chart(por_app)
        else:
            st.info("Os registros antigos ainda não possuem aplicativo informado.")

        st.subheader("Últimas corridas")
        ultimas = df.head(5).copy()
        ultimas["Data"] = ultimas["Data"].dt.strftime("%d/%m/%Y")

        st.dataframe(ultimas, use_container_width=True, hide_index=True)

    else:
        st.info("Nenhuma corrida registrada ainda. Acesse 'Nova corrida' para começar.")


# =========================================================
# HISTÓRICO
# =========================================================

elif st.session_state.pagina == "Histórico":

    st.header("Histórico de corridas")

    if df.empty:
        st.info("Ainda não existem corridas registradas.")

    else:
        c1, c2 = st.columns(2)

        with c1:
            data_inicio = st.date_input(
                "Data inicial",
                value=date.today() - timedelta(days=30),
                key="filtro_inicio",
                format="DD/MM/YYYY"
            )

        with c2:
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
            df_filtrado = df[(datas >= data_inicio) & (datas <= data_fim)].copy()

            st.caption(f"{len(df_filtrado)} corrida(s) no período selecionado.")

            df_exibicao = df_filtrado.copy()
            df_exibicao["Data"] = df_exibicao["Data"].dt.strftime("%d/%m/%Y")

            st.dataframe(
                df_exibicao,
                use_container_width=True,
                hide_index=True
            )

            st.subheader("Totais do período")
            c1, c2 = st.columns(2)

            with c1:
                st.metric("Bruto", formatar_moeda(df_filtrado["Bruto (R$)"].sum()))

            with c2:
                st.metric("Lucro livre", formatar_moeda(df_filtrado["Lucro Livre (R$)"].sum()))

            # O Excel inclui aplicativo e KM percorridos.
            st.download_button(
                label="⬇ Exportar histórico para Excel",
                data=gerar_excel(df_filtrado),
                file_name=(
                    f"historico_bajaj_"
                    f"{data_inicio.strftime('%Y%m%d')}_"
                    f"{data_fim.strftime('%Y%m%d')}.xlsx"
                ),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )


# =========================================================
# GERENCIAR REGISTROS
# =========================================================

elif st.session_state.pagina == "Gerenciar":

    st.header("Gerenciar registros")

    if not dados:
        st.info("Não existem registros para gerenciar.")

    else:
        opcoes = []

        for indice, registro in enumerate(dados):
            data_registro = converter_data(registro.get("Data", ""))
            hora = registro.get("Hora", "") or "--:--"
            aplicativo = registro.get("Aplicativo", "") or "App não informado"
            bruto_registro = float(registro.get("Bruto (R$)", 0.0) or 0.0)

            rotulo = (
                f"{indice + 1}. {data_registro.strftime('%d/%m/%Y')} "
                f"{hora} — {aplicativo} — {formatar_moeda(bruto_registro)}"
            )

            opcoes.append((rotulo, indice))

        escolha = st.selectbox(
            "Selecione o registro",
            options=opcoes,
            format_func=lambda item: item[0],
            key="registro_selecionado"
        )

        indice = escolha[1]
        registro = dados[indice]

        st.subheader("Editar corrida")

        app_atual = registro.get("Aplicativo", "") or ""
        app_opcao_atual = app_atual if app_atual in APLICATIVOS[:-1] else "Outros"

        app_opcao_editada = st.selectbox(
            "Aplicativo utilizado",
            APLICATIVOS,
            index=APLICATIVOS.index(app_opcao_atual),
            key=f"app_editar_{indice}"
        )

        outro_atual = app_atual if app_opcao_atual == "Outros" else ""

        if app_opcao_editada == "Outros":
            outro_editado = st.text_input(
                "Digite o nome do aplicativo",
                value=outro_atual,
                key=f"outro_editar_{indice}"
            )
        else:
            outro_editado = ""

        col1, col2 = st.columns(2)

        with col1:
            data_editada = st.date_input(
                "Data",
                value=converter_data(registro.get("Data", "")),
                key=f"data_editar_{indice}",
                format="DD/MM/YYYY"
            )

        with col2:
            try:
                hora_atual = datetime.strptime(
                    registro.get("Hora", "") or "00:00",
                    "%H:%M"
                ).time()
            except ValueError:
                hora_atual = time(0, 0)

            hora_editada = st.time_input(
                "Horário",
                value=hora_atual,
                format="24h",
                key=f"hora_editar_{indice}"
            )

        col1, col2 = st.columns(2)

        with col1:
            bruto_editado = st.number_input(
                "Valor bruto (R$)",
                min_value=0.0,
                value=float(registro.get("Bruto (R$)", 0.0) or 0.0),
                step=5.0,
                format="%.2f",
                key=f"bruto_editar_{indice}"
            )

        with col2:
            km_atual = registro.get("KM percorridos")
            if km_atual is None or pd.isna(km_atual):
                km_atual = 0.0

            km_editado = st.number_input(
                "KM percorridos",
                min_value=0.0,
                value=float(km_atual),
                step=0.5,
                format="%.2f",
                key=f"km_editar_{indice}"
            )

        if st.button("Salvar alterações", type="primary", use_container_width=True):

            aplicativo_editado = nome_aplicativo_selecionado(
                app_opcao_editada,
                outro_editado
            )

            if bruto_editado <= 0:
                st.error("O valor bruto precisa ser maior que zero.")

            elif not aplicativo_editado:
                st.error("Digite o nome do aplicativo.")

            else:
                valores = calcular_corrida(bruto_editado)

                registro_atualizado = {
                    "Data": data_editada.isoformat(),
                    "Hora": hora_editada.strftime("%H:%M"),
                    "Aplicativo": aplicativo_editado,
                    "KM percorridos": round(km_editado, 2) if km_editado > 0 else None,
                    **valores
                }

                novos_dados = dados.copy()
                novos_dados[indice] = registro_atualizado

                if salvar_dados(novos_dados):
                    st.session_state.dados = novos_dados
                    st.success("Registro atualizado com sucesso.")
                    st.rerun()

        st.divider()
        st.subheader("Excluir uma corrida")

        if st.button("Excluir corrida selecionada", use_container_width=True):
            st.session_state.confirmar_exclusao = True

        if st.session_state.confirmar_exclusao:
            st.warning("Deseja realmente excluir a corrida selecionada?")

            c1, c2 = st.columns(2)

            with c1:
                if st.button("Sim, excluir", type="primary"):
                    novos_dados = dados.copy()
                    novos_dados.pop(indice)

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
            "Essa ação remove todos os registros. "
            "Exporte o histórico para Excel antes de continuar, se precisar de uma cópia."
        )

        if st.button("Limpar todo o histórico", use_container_width=True):
            st.session_state.confirmar_limpeza = True

        if st.session_state.confirmar_limpeza:
            st.error("Confirma a exclusão de todas as corridas?")

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
st.caption("Finanças Bajaj • Controle pessoal de corridas e despesas")
