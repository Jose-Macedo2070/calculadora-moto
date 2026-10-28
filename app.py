
# =========================================================
# PÁGINA: NOVA CORRIDA
# =========================================================

if st.session_state.pagina == "Nova corrida":
    st.header("Nova corrida")
    st.caption("Informe os dados da corrida para registrar seus ganhos.")

    versao = st.session_state.nova_corrida_version

    # Data e hora atuais no ambiente em que o Python está executando.
    agora = datetime.now().astimezone()

    with st.container(border=True):
        data_corrida = st.date_input(
            "Data da corrida",
            value=agora.date(),
            key=f"nova_data_{versao}",
            format="DD/MM/YYYY",
        )

        hora_corrida = st.time_input(
            "Horário da corrida",
            value=agora.time().replace(
                second=0,
                microsecond=0,
                tzinfo=None,
            ),
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
                st.error(
                    "Informe o nome do aplicativo selecionado em Outros."
                )

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
                    "Despesas Gerais (R$)": distribuicao[
                        "Despesas Gerais (R$)"
                    ],
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
                    st.error(
                        f"Não foi possível salvar os dados: {erro}"
                    )
