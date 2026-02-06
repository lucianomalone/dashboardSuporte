import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.io as pio
import tempfile
import os

# ===============================
# CONFIGURAÇÕES GLOBAIS
# ===============================
pio.templates.default = "plotly_white"

# ===============================
# FUNÇÃO DEFINITIVA DE CORES
# ===============================
def gerar_cores(paleta, tamanho):
    return [paleta[i % len(paleta)] for i in range(tamanho)]

# ===============================
# Configuração da página
# ===============================
st.set_page_config(
    page_title="Dashboard de Demandas do Suporte",
    layout="wide"
)

st.title("📊 Dashboard Mensal de Demandas do Suporte")

# ===============================
# Upload do arquivo
# ===============================
arquivo = st.file_uploader(
    "📂 Selecione o arquivo de chamados (Excel)",
    type=["xlsx", "xls"]
)

if arquivo:
    # ===============================
    # Base completa
    # ===============================
    df_total = pd.read_excel(arquivo)
    df_total.columns = df_total.columns.str.strip()
    total_chamados_base = len(df_total)

    # ===============================
    # Base analisada
    # ===============================
    df = df_total.copy()
    df = df[~df["Status"].isin(["Fora do horário", "DEV - Em desenvolvimento"])]

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
    # Top 3 Serviços
    # ===============================
    top_servicos = (
        df["Serviço (Completo)"]
        .value_counts()
        .head(3)
        .reset_index()
    )
    top_servicos.columns = ["Serviço (Completo)", "Quantidade de Chamados"]

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
    col1.metric("📞 Total de Chamados (Base)", total_chamados_base)
    col2.metric("📂 Chamados Analisados", total_chamados_analisados)
    col3.metric("🏆 % Top 3 Serviços", f"{percentual_top3:.2f}%")

    # ===============================
    # Gráfico Top 3 Serviços
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

    fig_top.update_layout(
        xaxis_title="Serviço",
        yaxis_title="Quantidade de Chamados"
    )

    st.plotly_chart(fig_top, use_container_width=True)

    # ===============================
    # Categorias por Serviço
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
        "Erros e Problemas com software"
    ]

    df_erros = df[df["Categoria"].isin(categorias_erros)]
    total_erros = len(df_erros)

    percentual_erros = (
        (total_erros / total_chamados_analisados) * 100
        if total_chamados_analisados > 0 else 0
    )

    col_e1, col_e2 = st.columns(2)
    col_e1.metric("🟥 Total de Chamados de Erro", total_erros)
    col_e2.metric("📊 % Erros sobre Chamados Analisados", f"{percentual_erros:.2f}%")

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

        fig_erros.update_layout(
            xaxis_title="Categoria de Erro",
            yaxis_title="Quantidade de Chamados"
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
            f.write(f"""
<!DOCTYPE html>
<html lang="pt-br">
<head>
<meta charset="utf-8">
<title>Dashboard de Demandas do Suporte</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
<script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
<style>
body {{ background:#f8f9fa; }}
.card {{
    border-radius:12px;
    box-shadow:0 4px 10px rgba(0,0,0,.08);
}}
.plotly-graph-div {{ margin-bottom:40px; }}
</style>
</head>
<body>
<div class="container my-5">

<h1 class="text-center text-primary mb-4">📊 Dashboard Mensal de Demandas</h1>

<div class="row mb-4 text-center">
<div class="col-md-4"><div class="card p-3">Total de Chamados<br><strong>{total_chamados_base}</strong></div></div>
<div class="col-md-4"><div class="card p-3">Chamados Analisados<br><strong>{total_chamados_analisados}</strong></div></div>
<div class="col-md-4"><div class="card p-3">% Top 3 Serviços<br><strong>{percentual_top3:.2f}%</strong></div></div>
</div>

<h4 class="mt-4">🏆 Top 3 Serviços</h4>
""")

            f.write(fig_top.to_html(full_html=False, include_plotlyjs=False))

            f.write("<h4 class='mt-5'>📌 Categorias por Serviço</h4>")
            for fig in figuras_categoria:
                f.write(fig.to_html(full_html=False, include_plotlyjs=False))

            f.write(f"""
<h4 class="mt-5 text-danger">🚨 Chamados de Erro</h4>
<div class="row text-center mb-3">
<div class="col-md-6"><div class="card p-3">Total de Erros<br><strong>{total_erros}</strong></div></div>
<div class="col-md-6"><div class="card p-3">% Erros sobre Analisados<br><strong>{percentual_erros:.2f}%</strong></div></div>
</div>
""")

            if not df_erros.empty:
                f.write(fig_erros.to_html(full_html=False, include_plotlyjs=False))

            f.write("</div></body></html>")

        with open(html_path, "rb") as file:
            st.download_button(
                "📥 Baixar HTML",
                file,
                "dashboard_suporte.html",
                "text/html"
            )

        os.remove(html_path)

else:
    st.info("👆 Faça upload de um arquivo Excel para iniciar o dashboard.")
