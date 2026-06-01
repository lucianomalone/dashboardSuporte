import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.io as pio
import tempfile
import os

pio.templates.default = "plotly_white"

def gerar_cores(paleta, tamanho):
    return [paleta[i % len(paleta)] for i in range(tamanho)]

st.set_page_config(
    page_title="Dashboard de Demandas do Suporte",
    layout="wide"
)

st.title("📊 Dashboard de Demandas do Suporte")

# ===============================
# LEITURA DIRETA DO ARQUIVO
# ===============================
# Defina aqui o nome exato do seu arquivo Excel que estará na mesma pasta
nome_do_arquivo = "Ticket.xlsx" 

if os.path.exists(nome_do_arquivo):
    df_total = pd.read_excel(nome_do_arquivo)
    df_total.columns = df_total.columns.str.strip()

    # ===============================
    # Conversão da data
    # ===============================
    df_total["Aberto em"] = pd.to_datetime(
        df_total["Aberto em"],
        errors="coerce"
    )

    total_chamados_base = len(df_total)

    # ===============================
    # FILTRO POR DATA
    # ===============================
    st.sidebar.header("📅 Filtro por Período")

    data_min = df_total["Aberto em"].min()
    data_max = df_total["Aberto em"].max()

    data_inicio = st.sidebar.date_input(
        "Data inicial",
        value=data_min,
        min_value=data_min,
        max_value=data_max
    )

    data_fim = st.sidebar.date_input(
        "Data final",
        value=data_max,
        min_value=data_min,
        max_value=data_max
    )

    df_total = df_total[
        (df_total["Aberto em"] >= pd.to_datetime(data_inicio)) &
        (df_total["Aberto em"] <= pd.to_datetime(data_fim))
    ]

    # ===============================
    # Base analisada
    # ===============================
    df = df_total.copy()

    df = df[~df["Status"].isin([
        "Fora do horário",
        "DEV - Em desenvolvimento"
    ])]

    responsaveis_permitidos = [
        "Luiz Cafarate",
        "Leonardo W.",
        "Natanael Cardoso",
        "Luciano Boteleiro",
        "Thiago O. P."
    ]

    df = df[df["Responsável"].isin(responsaveis_permitidos)]

    total_chamados_analisados = len(df)

    # ===============================
    # Top serviços
    # ===============================
    top_servicos = (
        df["Serviço (Completo)"]
        .value_counts()
        .head(3)
        .reset_index()
    )

    top_servicos.columns = [
        "Serviço (Completo)",
        "Quantidade de Chamados"
    ]

    total_top3 = top_servicos["Quantidade de Chamados"].sum()

    percentual_top3 = (
        (total_top3 / total_chamados_analisados) * 100
        if total_chamados_analisados > 0 else 0
    )

    # ===============================
    # VISÃO GERAL
    # ===============================
    st.markdown("### 📌 Visão Geral")

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "📞 Total de Chamados (Base)",
        total_chamados_base
    )

    col2.metric(
        "📂 Chamados Analisados",
        total_chamados_analisados
    )

    col3.metric(
        "🏆 % Top 3 Serviços",
        f"{percentual_top3:.2f}%"
    )

    # ===============================
    # Gráfico Top Serviços
    # ===============================
    st.subheader("🏆 Top 3 Serviços com Maior Demanda")

    fig_top = px.bar(
        top_servicos,
        x="Serviço (Completo)",
        y="Quantidade de Chamados",
        text="Quantidade de Chamados"
    )

    fig_top.update_traces(
        textposition="outside",
        marker_color=gerar_cores(
            px.colors.qualitative.Set2,
            len(top_servicos)
        )
    )

    st.plotly_chart(fig_top, use_container_width=True)

    # ===============================
    # Categorias por serviço
    # ===============================
    st.subheader("📌 Top 5 Categorias por Serviço")

    figuras_categoria = []

    for servico in top_servicos["Serviço (Completo)"]:

        df_categoria = (
            df[df["Serviço (Completo)"] == servico]
            .groupby("Categoria")
            .size()
            .reset_index(name="Quantidade de Chamados")
            .sort_values("Quantidade de Chamados", ascending=False)
            .head(5)
        )

        fig = px.bar(
            df_categoria,
            x="Categoria",
            y="Quantidade de Chamados",
            text="Quantidade de Chamados",
            title=servico
        )

        fig.update_traces(
            textposition="outside",
            marker_color=gerar_cores(
                px.colors.qualitative.Pastel,
                len(df_categoria)
            )
        )

        figuras_categoria.append(fig)

        st.plotly_chart(fig, use_container_width=True)

    # ===============================
    # TOP CLIENTES
    # ===============================
    st.divider()
    st.subheader("👥 Top 10 Clientes com Mais Chamados")

    coluna_cliente = "Cliente (Organização)" if "Cliente (Organização)" in df.columns else "Nome do Cliente" if "Nome do Cliente" in df.columns else None

    if coluna_cliente and coluna_cliente in df.columns:
        top_clientes = (
            df.groupby(coluna_cliente)
            .size()
            .reset_index(name="Quantidade de Chamados")
            .sort_values("Quantidade de Chamados", ascending=False)
            .head(10)
        )

        fig_clientes = px.bar(
            top_clientes,
            y=coluna_cliente,
            x="Quantidade de Chamados",
            text="Quantidade de Chamados" 
        )

        fig_clientes.update_layout(yaxis={'categoryorder':'total ascending'})

        fig_clientes.update_traces(
            textposition="outside",
            marker_color=gerar_cores(
                px.colors.qualitative.Prism, 
                len(top_clientes)
            )
        )

        st.plotly_chart(fig_clientes, use_container_width=True)
        
    else:
        st.warning("⚠️ Coluna referente a 'Cliente' não encontrada na base de dados. Verifica o nome da coluna no teu Excel.")

    # ===============================
    # ERROS
    # ===============================
    st.divider()
    st.subheader("🚨 Chamados por Categoria de Erro")

    categorias_erros = [
        "Erro do Sistema",
        "Erro Atualização Relatórios",
        "Erro encaminhado para correção",
        "Erro Ponto de entrada",
        "Outros erros ou problemas",
        "Problemas - APi",
        "Problemas com atualização",
        "Erros e Problemas com software",
        "Base bloqueada"
    ]

    df_erros = df[df["Categoria"].isin(categorias_erros)]

    total_erros = len(df_erros)

    percentual_erros = (
        (total_erros / total_chamados_analisados) * 100
        if total_chamados_analisados > 0 else 0
    )

    col_e1, col_e2 = st.columns(2)

    col_e1.metric(
        "🟥 Total de Chamados de Erro",
        total_erros
    )

    col_e2.metric(
        "📊 % Erros sobre Chamados",
        f"{percentual_erros:.2f}%"
    )

    if not df_erros.empty:

        erros_qtd = (
            df_erros.groupby("Categoria")
            .size()
            .reset_index(name="Quantidade de Chamados")
            .sort_values("Quantidade de Chamados", ascending=False)
        )

        fig_erros = px.bar(
            erros_qtd,
            x="Categoria",
            y="Quantidade de Chamados",
            text="Quantidade de Chamados"
        )

        fig_erros.update_traces(
            textposition="outside",
            marker_color=gerar_cores(
                px.colors.qualitative.Bold,
                len(erros_qtd)
            )
        )

        st.plotly_chart(fig_erros, use_container_width=True)

    # ===============================
    # EXPORTAÇÃO HTML
    # ===============================
    st.divider()
    st.subheader("📥 Exportar Dashboard")

    if st.button("Gerar HTML do Dashboard"):

        with tempfile.NamedTemporaryFile(delete=False, suffix=".html") as tmp:
            html_path = tmp.name

        with open(html_path, "w", encoding="utf-8") as f:

            f.write("<html><head>")
            f.write("<script src='https://cdn.plot.ly/plotly-latest.min.js'></script>")
            f.write("</head><body>")
            f.write("<h1>Dashboard de Demandas</h1>")

            f.write(fig_top.to_html(full_html=False))

            for fig in figuras_categoria:
                f.write(fig.to_html(full_html=False))

            if 'fig_clientes' in locals():
                f.write(fig_clientes.to_html(full_html=False))

            if not df_erros.empty:
                f.write(fig_erros.to_html(full_html=False))

            f.write("</body></html>")

        with open(html_path, "rb") as file:
            st.download_button(
                "📥 Baixar HTML",
                file,
                "dashboard_suporte.html",
                "text/html"
            )

        os.remove(html_path)

else:
    # Mostra um erro amigável se o arquivo não for encontrado na pasta
    st.error(f"🚨 Arquivo '{nome_do_arquivo}' não encontrado! Certifique-se de que o arquivo Excel está na mesma pasta que este script.")