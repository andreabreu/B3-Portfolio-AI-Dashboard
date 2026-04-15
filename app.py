import json
import os
from dataclasses import dataclass
from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st
import yfinance as yf
from openai import OpenAI


@dataclass
class SheetConfig:
    name: str
    value_columns: list[str]
    code_column: str


SHEET_CONFIGS = [
    SheetConfig(
        name="Acoes",
        value_columns=["Valor Atualizado"],
        code_column="Código de Negociação",
    ),
    SheetConfig(
        name="Fundo de Investimento",
        value_columns=["Valor Atualizado"],
        code_column="Código de Negociação",
    ),
    SheetConfig(
        name="Renda Fixa",
        value_columns=["Valor Atualizado CURVA", "Valor Atualizado FECHAMENTO", "Valor Atualizado MTM"],
        code_column="Código",
    ),
]


def _to_number(value: object) -> float:
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text or text == "-":
        return 0.0
    text = text.replace(".", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return 0.0


def _choose_value_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for col in candidates:
        if col in df.columns:
            return col
    return None


def parse_portfolio(file_data: BytesIO) -> tuple[pd.DataFrame, dict[str, float]]:
    xl = pd.ExcelFile(file_data)
    rows = []
    totals = {}

    for config in SHEET_CONFIGS:
        if config.name not in xl.sheet_names:
            continue

        raw_df = xl.parse(config.name)
        value_col = _choose_value_column(raw_df, config.value_columns)
        if value_col is None:
            continue

        for _, row in raw_df.iterrows():
            asset_code = str(row.get(config.code_column, "")).strip()
            asset_name = str(row.get("Produto", "")).strip()
            value = _to_number(row.get(value_col, 0))

            if not asset_code or asset_code.lower() == "nan":
                continue
            if not asset_name or asset_name.lower() == "nan":
                continue
            if value <= 0:
                continue

            rows.append(
                {
                    "classe": config.name,
                    "ativo": asset_code,
                    "produto": asset_name,
                    "valor": value,
                }
            )

        totals[config.name] = sum(item["valor"] for item in rows if item["classe"] == config.name)

    portfolio = pd.DataFrame(rows)
    if portfolio.empty:
        return portfolio, totals

    portfolio = portfolio.sort_values(by="valor", ascending=False).reset_index(drop=True)
    grand_total = float(portfolio["valor"].sum())
    portfolio["peso_pct"] = (portfolio["valor"] / grand_total) * 100
    return portfolio, totals


def build_metrics(portfolio: pd.DataFrame) -> dict:
    total = float(portfolio["valor"].sum())
    by_class = (
        portfolio.groupby("classe", as_index=False)["valor"]
        .sum()
        .sort_values("valor", ascending=False)
    )
    by_asset = portfolio.sort_values("valor", ascending=False).head(10)
    top_asset_weight = float(by_asset.iloc[0]["peso_pct"]) if not by_asset.empty else 0
    top_3_weight = float(by_asset.head(3)["peso_pct"].sum()) if len(by_asset) >= 3 else float(by_asset["peso_pct"].sum())

    return {
        "total": total,
        "by_class": by_class.to_dict(orient="records"),
        "top_assets": by_asset[["ativo", "classe", "valor", "peso_pct"]].to_dict(orient="records"),
        "top_asset_weight": top_asset_weight,
        "top_3_weight": top_3_weight,
    }


def _to_brl(value: float) -> str:
    return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _to_yahoo_symbol(asset_code: str) -> str | None:
    code = str(asset_code).strip().upper()
    if not code:
        return None
    if ".SA" in code:
        return code
    if not code.isalnum():
        return None
    return f"{code}.SA"


@st.cache_data(ttl=900)
def fetch_yahoo_snapshot(asset_codes: tuple[str, ...]) -> dict[str, dict[str, float | None]]:
    snapshot: dict[str, dict[str, float | None]] = {}
    for asset in asset_codes:
        symbol = _to_yahoo_symbol(asset)
        if not symbol:
            snapshot[asset] = {"preco_atual": None, "variacao_pct_dia": None}
            continue

        try:
            hist = yf.Ticker(symbol).history(period="5d", interval="1d")
            if hist.empty:
                snapshot[asset] = {"preco_atual": None, "variacao_pct_dia": None}
                continue

            close = hist["Close"].dropna()
            if close.empty:
                snapshot[asset] = {"preco_atual": None, "variacao_pct_dia": None}
                continue

            last_price = float(close.iloc[-1])
            prev_price = float(close.iloc[-2]) if len(close) >= 2 else last_price
            change_pct = ((last_price - prev_price) / prev_price * 100) if prev_price else 0.0

            snapshot[asset] = {
                "preco_atual": round(last_price, 4),
                "variacao_pct_dia": round(change_pct, 4),
            }
        except Exception:
            snapshot[asset] = {"preco_atual": None, "variacao_pct_dia": None}

    return snapshot


def enrich_portfolio_with_yahoo(portfolio: pd.DataFrame) -> pd.DataFrame:
    enriched = portfolio.copy()
    market_rows = enriched["classe"].isin(["Acoes", "Fundo de Investimento"])
    assets = tuple(sorted(set(enriched.loc[market_rows, "ativo"].astype(str).tolist())))
    snapshot = fetch_yahoo_snapshot(assets)

    enriched["preco_atual_yf"] = enriched["ativo"].map(lambda code: snapshot.get(str(code), {}).get("preco_atual") if str(code) in snapshot else None)
    enriched["variacao_dia_yf_pct"] = enriched["ativo"].map(
        lambda code: snapshot.get(str(code), {}).get("variacao_pct_dia") if str(code) in snapshot else None
    )
    enriched["dados_yahoo"] = market_rows & enriched["preco_atual_yf"].notna()
    return enriched


def build_rule_based_analysis(metrics: dict) -> str:
    suggestions = []
    top_asset = metrics["top_asset_weight"]
    top3 = metrics["top_3_weight"]
    by_class = metrics["by_class"]

    if top_asset > 20:
        suggestions.append(
            f"- Concentração alta: seu maior ativo representa {top_asset:.1f}% da carteira. Considere reduzir para algo entre 10% e 15%."
        )
    else:
        suggestions.append("- Concentração do maior ativo está controlada, dentro de um intervalo razoável.")

    if top3 > 45:
        suggestions.append(
            f"- Seus 3 maiores ativos somam {top3:.1f}%. Avalie aumentar posições menores para melhorar diversificação."
        )
    else:
        suggestions.append("- Boa distribuição entre os 3 maiores ativos.")

    class_text = ", ".join([f"{item['classe']}: R$ {item['valor']:,.2f}" for item in by_class]).replace(",", "X").replace(".", ",").replace("X", ".")
    suggestions.append(f"- Distribuição por classe: {class_text}.")
    suggestions.append("- Próximo passo: definir metas de alocação por classe (ex.: 40% FIIs, 40% ações, 20% renda fixa) e rebalancear mensalmente.")

    return "### Análise inicial da carteira\n\n" + "\n".join(suggestions) + "\n\n*Esta análise é educacional e não é recomendação de investimento.*"


def build_ai_analysis(metrics: dict, objective: str, risk_profile: str, api_key: str) -> str:
    client = OpenAI(api_key=api_key)
    prompt = {
        "objetivo": objective,
        "perfil_risco": risk_profile,
        "resumo_carteira": metrics,
        "instrucao": (
            "Analise a carteira e responda em portugues do Brasil com: "
            "1) leitura geral da carteira, 2) principais riscos, 3) 5 sugestoes praticas de ajuste, "
            "4) plano de rebalanceamento em passos simples. "
            "Nao invente dados externos. Nao faça recomendacao personalizada definitiva."
        ),
    }

    response = client.responses.create(
        model="gpt-4.1-mini",
        input=json.dumps(prompt, ensure_ascii=False),
    )
    return response.output_text


def main() -> None:
    st.set_page_config(page_title="Portal de Investimentos com IA", page_icon="📈", layout="wide")
    st.title("📈 Portal de Investimentos com IA")
    st.caption("Faça upload do Excel da B3, veja sua carteira consolidada e receba sugestões com IA.")

    with st.sidebar:
        st.header("Configuração da IA")
        api_key = st.text_input(
            "OPENAI_API_KEY",
            type="password",
            value=os.getenv("OPENAI_API_KEY", ""),
            help="Opcional. Se não informar, o sistema gera análise automática por regras.",
        )
        objective = st.text_area(
            "Objetivo financeiro",
            value="Aumentar renda passiva e crescimento de patrimônio no longo prazo.",
        )
        risk_profile = st.selectbox(
            "Perfil de risco",
            ["Conservador", "Moderado", "Arrojado"],
            index=1,
        )
        use_yahoo_data = st.toggle(
            "Enriquecer com Yahoo Finance",
            value=True,
            help="Busca preço atual e variação diária dos ativos negociados em bolsa.",
        )

    uploaded_file = st.file_uploader("Envie seu arquivo de posição (.xlsx)", type=["xlsx"])
    if not uploaded_file:
        st.info("Faça o upload do arquivo para iniciar a análise.")
        st.stop()

    portfolio, _ = parse_portfolio(uploaded_file)
    if portfolio.empty:
        st.error("Não foi possível extrair posições válidas do arquivo. Verifique o formato do Excel.")
        st.stop()

    if use_yahoo_data:
        portfolio = enrich_portfolio_with_yahoo(portfolio)
    else:
        portfolio["preco_atual_yf"] = pd.NA
        portfolio["variacao_dia_yf_pct"] = pd.NA
        portfolio["dados_yahoo"] = False

    metrics = build_metrics(portfolio)
    market_assets = portfolio[portfolio["dados_yahoo"]].copy()
    yahoo_coverage_pct = (len(market_assets) / len(portfolio) * 100) if len(portfolio) else 0

    col1, col2, col3 = st.columns(3)
    col1.metric("Patrimônio total", _to_brl(metrics["total"]))
    col2.metric("Qtd. ativos", f"{len(portfolio)}")
    col3.metric("Cobertura Yahoo", f"{yahoo_coverage_pct:.1f}%")

    chart_col1, chart_col2 = st.columns(2)
    with chart_col1:
        pie_data = pd.DataFrame(metrics["by_class"])
        fig = px.pie(pie_data, names="classe", values="valor", title="Distribuição por classe")
        st.plotly_chart(fig, use_container_width=True)

    with chart_col2:
        bar_data = portfolio.sort_values("valor", ascending=False).head(10)
        fig2 = px.bar(bar_data, x="ativo", y="valor", color="classe", title="Top 10 ativos por valor")
        st.plotly_chart(fig2, use_container_width=True)

    st.subheader("Carteira consolidada")
    st.dataframe(
        portfolio[["classe", "ativo", "produto", "valor", "peso_pct", "preco_atual_yf", "variacao_dia_yf_pct"]].rename(
            columns={
                "classe": "Classe",
                "ativo": "Ativo",
                "produto": "Produto",
                "valor": "Valor (R$)",
                "peso_pct": "Peso (%)",
                "preco_atual_yf": "Preço atual (Yahoo)",
                "variacao_dia_yf_pct": "Variação dia % (Yahoo)",
            }
        ),
        use_container_width=True,
    )

    st.subheader("Mercado (Yahoo Finance)")
    if market_assets.empty:
        st.info("Nenhum ativo com cotação encontrada no Yahoo Finance.")
    else:
        gainers = (
            market_assets[["ativo", "variacao_dia_yf_pct"]]
            .dropna()
            .sort_values("variacao_dia_yf_pct", ascending=False)
            .head(5)
        )
        losers = (
            market_assets[["ativo", "variacao_dia_yf_pct"]]
            .dropna()
            .sort_values("variacao_dia_yf_pct", ascending=True)
            .head(5)
        )
        mk1, mk2 = st.columns(2)
        with mk1:
            st.caption("Top 5 altas do dia")
            st.dataframe(gainers.rename(columns={"ativo": "Ativo", "variacao_dia_yf_pct": "Variação %"}), use_container_width=True)
        with mk2:
            st.caption("Top 5 baixas do dia")
            st.dataframe(losers.rename(columns={"ativo": "Ativo", "variacao_dia_yf_pct": "Variação %"}), use_container_width=True)

    if st.button("Gerar análise e sugestões com IA", type="primary"):
        with st.spinner("Analisando carteira..."):
            try:
                metrics["mercado_yahoo"] = market_assets[
                    ["ativo", "preco_atual_yf", "variacao_dia_yf_pct", "valor", "peso_pct"]
                ].to_dict(orient="records")
                if api_key:
                    analysis = build_ai_analysis(metrics, objective, risk_profile, api_key)
                else:
                    analysis = build_rule_based_analysis(metrics)
                st.markdown(analysis)
            except Exception as exc:
                st.warning("Falha na chamada de IA. Exibindo análise automática por regras.")
                st.markdown(build_rule_based_analysis(metrics))
                st.caption(f"Detalhe técnico: {exc}")

    st.markdown("---")
    st.caption("Aviso: conteúdo educacional, não constitui recomendação de investimento.")


if __name__ == "__main__":
    main()
