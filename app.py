            aplicativo_padrao = "Outros" if aplicativo_original else "Uber"
            outro_original = aplicativo_original

        km_original = valor_numerico(registro.get("KM percorridos"), 0.0)
        bruto_original = valor_numerico(registro.get("Bruto (R$)"), 0.0)
        chave_edicao = f"{registro_escolhido}_{len(dados)}"

        with st.form(key=f"form_editar_{chave_edicao}"):
            data_editada = st.date_input("Data", value=data_original, format="DD/MM/YYYY")
            hora_editada = st.time_input("Horario", value=hora_original, format="24h")
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
                "Salvar alteracoes",
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
                            aplicativo_editado, outro_nome_editado
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
                        st.session_state.mensagem_geral = "Corrida atualizada com sucesso."
                        st.rerun()
                    except OSError as erro:
                        st.error(f"Erro ao salvar alteracoes: {erro}")

        st.divider()
        st.subheader("Excluir corrida")
        st.warning("A exclusao remove permanentemente a corrida selecionada.")

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
                st.session_state.mensagem_geral = "Corrida excluida com sucesso."
                st.rerun()
            except OSError as erro:
                st.error(f"Nao foi possivel excluir a corrida: {erro}")

        st.divider()
        st.subheader("Apagar todo o historico")
        st.warning("Esta acao apaga todas as corridas registradas e nao pode ser desfeita.")

        confirmar_limpeza = st.checkbox(
            "Confirmo que desejo apagar todo o historico.",
            key="confirmar_limpeza_total",
        )

        if st.button(
            "🧹 Apagar todo o historico",
            disabled=not confirmar_limpeza,
            use_container_width=True,
        ):
            try:
                salvar_dados([])
                st.session_state.dados = []
                st.session_state.mensagem_geral = "Todo o historico foi apagado."
                st.rerun()
            except OSError as erro:
                st.error(f"Nao foi possivel apagar o historico: {erro}")
