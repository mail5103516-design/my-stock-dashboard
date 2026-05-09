import streamlit as st
import pandas as pd
import yfinance as yf
import plotly.graph_objects as go
import os
import re

# --- 設定と定数 ---
st.set_page_config(page_title="ポートフォリオ・ダッシュボード", layout="wide")
PORTFOLIO_FILE = "my_saved_portfolio.csv"

# --- 1. ティッカー抽出ロジック ---
def extract_ticker(item_name, item_type):
    match = re.search(r'^([A-Za-z0-9]+)', str(item_name))
    if not match:
        return None
    ticker_base = match.group(1)
    if item_type in ['日本株', '国内株式', '信用建玉']:
        return f"{ticker_base}.T"
    elif item_type in ['米国株', '米国株式']:
        return ticker_base
    return None

# --- 1.5 ノイズ排除ロジック ---
def is_noise(item_name, item_type):
    name_str = str(item_name).upper()
    type_str = str(item_type)
    valid_types = ['日本株', '国内株式', '米国株', '米国株式', '信用建玉']
    if type_str not in valid_types:
        return True
    exclude_keywords = [
        'ETF', 'NF', 'NEXT FUNDS', 'MAXIS', 'ISHARES', 'VANGUARD', 'SPDR',
        'レバ', 'ブル', 'ベア', 'インバース', 'インデックス', '連動', '投信'
    ]
    if any(kw in name_str for kw in exclude_keywords):
        return True
    return False

# --- 2. CSVパース ---
def parse_and_save_csv(uploaded_file):
    encodings = ['utf-8', 'utf-8-sig', 'cp932', 'shift_jis']
    df = None
    for enc in encodings:
        try:
            uploaded_file.seek(0)
            lines = uploaded_file.getvalue().decode(enc, errors='replace').split('\n')
            header_idx = 0
            for i, line in enumerate(lines):
                if "銘柄" in line or "ティッカー" in line or "ファンド名" in line:
                    header_idx = i
                    break
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, encoding=enc, skiprows=header_idx)
            break
        except Exception:
            continue
    if df is None or df.empty:
        raise ValueError("ファイルの読み込みに失敗しました。")
    df.to_csv(PORTFOLIO_FILE, index=False, encoding='utf-8-sig')
    return df

# --- 3. 株価・指標の取得 ---
def fetch_portfolio_metrics(df):
    type_col = '種別' if '種別' in df.columns else df.columns[0]
    if '銘柄コード・ティッカー' in df.columns: ticker_col = '銘柄コード・ティッカー'
    elif '銘柄名/ティッカー' in df.columns: ticker_col = '銘柄名/ティッカー'
    else: ticker_col = df.columns[1]
    if '銘柄' in df.columns: name_col = '銘柄'
    else: name_col = ticker_col
    qty_col   = '保有数量'    if '保有数量'    in df.columns else df.columns[2]
    if '平均取得単価' in df.columns:   price_col = '平均取得単価'
    elif '平均取得価額' in df.columns: price_col = '平均取得価額'
    else: price_col = df.columns[3]

    filtered_df = df[~df.apply(lambda row: is_noise(row[ticker_col], row[type_col]), axis=1)].copy()
    filtered_df['検索用ティッカー'] = filtered_df.apply(
        lambda row: extract_ticker(row[ticker_col], row[type_col]), axis=1)
    valid_tickers = filtered_df['検索用ティッカー'].dropna().unique().tolist()

    metrics_list = []
    if not valid_tickers:
        metrics_df = pd.DataFrame(columns=['検索用ティッカー','最新価格','EPS','PER','20MA','50MA','200MA','状態'])
    else:
        progress_bar = st.progress(0)
        for i, ticker in enumerate(valid_tickers):
            try:
                stock = yf.Ticker(ticker)
                info  = stock.info
                current_price = info.get('currentPrice', info.get('regularMarketPrice', None))
                eps = info.get('trailingEps', None)
                per = info.get('trailingPE',  None)
                hist = stock.history(period="1y")
                if not hist.empty:
                    ma20  = hist['Close'].rolling(window=20).mean().iloc[-1]
                    ma50  = hist['Close'].rolling(window=50).mean().iloc[-1]
                    ma200 = hist['Close'].rolling(window=200).mean().iloc[-1]
                    state = "横ばい"
                    if current_price and ma20 and ma50 and ma200:
                        if current_price > ma20 > ma50 > ma200: state = "🔥 上昇PO"
                        elif current_price < ma20 < ma50 < ma200: state = "🧊 下降PO"
                else:
                    ma20, ma50, ma200, state = None, None, None, "-"
                metrics_list.append({
                    '検索用ティッカー': ticker,
                    '最新価格': round(current_price, 2) if current_price else None,
                    'EPS': round(eps, 2) if eps else None,
                    'PER': round(per, 2) if per else None,
                    '20MA':  round(ma20,  2) if ma20  else None,
                    '50MA':  round(ma50,  2) if ma50  else None,
                    '200MA': round(ma200, 2) if ma200 else None,
                    '状態': state
                })
            except Exception:
                metrics_list.append({'検索用ティッカー': ticker, '最新価格': None,
                    'EPS': None, 'PER': None, '20MA': None, '50MA': None, '200MA': None, '状態': "エラー"})
            progress_bar.progress((i + 1) / len(valid_tickers))
        progress_bar.empty()
        metrics_df = pd.DataFrame(metrics_list)

    merged_df = pd.merge(filtered_df, metrics_df, on='検索用ティッカー', how='left')

    display_df = pd.DataFrame()
    display_df['コード']     = merged_df[ticker_col]
    display_df['銘柄名']     = merged_df[name_col]
    display_df['数量']       = merged_df[qty_col]
    display_df['取得単価']   = merged_df[price_col]
    display_df['現在値']     = merged_df['最新価格']
    display_df['EPS']        = merged_df['EPS']
    display_df['PER']        = merged_df['PER']
    display_df['20MA']       = merged_df['20MA']
    display_df['50MA']       = merged_df['50MA']
    display_df['200MA']      = merged_df['200MA']
    display_df['トレンド']   = merged_df['状態']
    display_df['種別']       = merged_df[type_col]
    display_df['ティッカー'] = merged_df['検索用ティッカー']
    return display_df

# --- 4. チャート生成（MA + フィボナッチ）---
def build_chart(ticker, name):
    try:
        stock = yf.Ticker(ticker)
        hist  = stock.history(period="1y")
        if hist.empty:
            return None

        hist['MA20']  = hist['Close'].rolling(20).mean()
        hist['MA50']  = hist['Close'].rolling(50).mean()
        hist['MA200'] = hist['Close'].rolling(200).mean()

        # フィボナッチ（直近60日）
        recent   = hist.tail(60)
        hi_fib   = float(recent['High'].max())
        lo_fib   = float(recent['Low'].min())
        diff_val = hi_fib - lo_fib

        fib_levels = {
            "161.8%": hi_fib + diff_val * 0.618,
            "100.0%": hi_fib,
            "76.4%":  hi_fib - diff_val * 0.236,
            "61.8%":  hi_fib - diff_val * 0.382,
            "50.0%":  hi_fib - diff_val * 0.5,
            "38.2%":  hi_fib - diff_val * 0.618,
            "23.6%":  hi_fib - diff_val * 0.764,
            "0.0%":   lo_fib,
            "-61.8%": lo_fib - diff_val * 0.618,
        }

        fig = go.Figure()

        # ローソク足
        fig.add_trace(go.Candlestick(
            x=recent.index,
            open=recent['Open'], high=recent['High'],
            low=recent['Low'],   close=recent['Close'],
            name="株価",
            increasing_line_color="#26a69a", increasing_fillcolor="#26a69a",
            decreasing_line_color="#ef5350", decreasing_fillcolor="#ef5350",
        ))

        # MA
        fig.add_trace(go.Scatter(x=recent.index, y=recent['MA20'],  name="20MA",  line=dict(color='#f5a623', width=1.5)))
        fig.add_trace(go.Scatter(x=recent.index, y=recent['MA50'],  name="50MA",  line=dict(color='#1a73e8', width=3)))
        fig.add_trace(go.Scatter(x=recent.index, y=recent['MA200'], name="200MA", line=dict(color='#e53935', width=1.5)))

        # フィボナッチ水平線
        for label, val in fib_levels.items():
            fig.add_hline(y=val, line_dash="dot", line_color="gray",
                          annotation_text=label, annotation_position="right")

        fig.update_layout(
            title=f"{name}（{ticker}）",
            xaxis_rangeslider_visible=False,
            height=500,
            margin=dict(l=20, r=80, t=40, b=20)
        )
        return fig
    except Exception:
        return None

# --- 5. UI構築 ---
st.title("📊 ポートフォリオ・モニター")

if 'display_data'   not in st.session_state: st.session_state['display_data']   = None
if 'chart_ticker'   not in st.session_state: st.session_state['chart_ticker']   = None
if 'chart_name'     not in st.session_state: st.session_state['chart_name']     = None

# 起動時の自動読み込み
base_df = None
if os.path.exists(PORTFOLIO_FILE):
    if os.path.getsize(PORTFOLIO_FILE) > 0:
        try:
            base_df = pd.read_csv(PORTFOLIO_FILE, encoding='utf-8-sig')
            st.success("✅ 保存されたポートフォリオを読み込みました。")
        except pd.errors.EmptyDataError:
            st.warning("⚠️ 保存されたファイルが破損しています。再度CSVをアップロードしてください。")
            os.remove(PORTFOLIO_FILE)
        except Exception as e:
            st.warning(f"⚠️ データの読み込みに失敗しました。（詳細: {e}）")
    else:
        st.warning("⚠️ 保存されたファイルが空です。再度CSVをアップロードしてください。")
        os.remove(PORTFOLIO_FILE)
else:
    st.warning("⚠️ ポートフォリオが登録されていません。下のメニューからCSVをアップロードしてください。")

with st.expander("📁 ポートフォリオの新規登録 / 上書き更新"):
    st.info("楽天証券のCSV、または自作のCSVをアップロードすると、内部データが上書き保存されます。")
    uploaded_file = st.file_uploader("CSVをアップロード", type=['csv'])
    if uploaded_file is not None:
        try:
            base_df = parse_and_save_csv(uploaded_file)
            st.success("ポートフォリオデータをシステムに保存しました！次回以降はアップロード不要です。")
            st.rerun()
        except Exception as e:
            st.error(f"エラー: {e}")

# 更新ボタン
if base_df is not None:
    st.markdown("---")
    col1, col2 = st.columns([1, 3])
    with col1:
        if st.button("🔄 最新の株価・指標を取得", type="primary", use_container_width=True):
            with st.spinner("データを取得中..."):
                st.session_state['display_data'] = fetch_portfolio_metrics(base_df)
                st.session_state['chart_ticker'] = None  # チャートリセット

# テーブル表示
if st.session_state['display_data'] is not None:
    st.markdown("### 最新ステータス")
    df = st.session_state['display_data']

    view_config = {
        "種別":       None,
        "ティッカー": None,
        "コード":     st.column_config.TextColumn("コード",   width="small"),
        "銘柄名":     st.column_config.TextColumn("銘柄名",   width="medium"),
        "現在値":     st.column_config.NumberColumn("現在値", format="%.2f"),
        "EPS":        st.column_config.NumberColumn("EPS",    format="%.2f"),
        "PER":        st.column_config.NumberColumn("PER",    format="%.1f"),
        "トレンド":   st.column_config.TextColumn("トレンド", width="small"),
    }

    tab1, tab2, tab3 = st.tabs(["🌐 すべて", "🇯🇵 日本株", "🇺🇸 米国株"])

    def render_table_with_chart_buttons(tab, filtered_df):
        with tab:
            st.dataframe(filtered_df, use_container_width=True,
                         hide_index=True, column_config=view_config)
            st.markdown("**📈 チャートを表示する銘柄を選択：**")
            cols = st.columns(min(len(filtered_df), 6))
            for i, (_, row) in enumerate(filtered_df.iterrows()):
                with cols[i % 6]:
                    label = str(row['コード']).split('.')[0]
                    if st.button(label, key=f"chart_{row['ティッカー']}_{tab}"):
                        st.session_state['chart_ticker'] = row['ティッカー']
                        st.session_state['chart_name']   = row['銘柄名']

    render_table_with_chart_buttons(tab1, df)
    render_table_with_chart_buttons(tab2, df[df['種別'].str.contains('日本|国内|信用', na=False, regex=True)])
    render_table_with_chart_buttons(tab3, df[df['種別'].str.contains('米国', na=False, regex=True)])

    # チャート表示エリア
    if st.session_state['chart_ticker']:
        st.markdown("---")
        st.markdown(f"### 📊 {st.session_state['chart_name']}（{st.session_state['chart_ticker']}）")
        with st.spinner("チャートを生成中..."):
            fig = build_chart(st.session_state['chart_ticker'], st.session_state['chart_name'])
        if fig:
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.error("チャートデータの取得に失敗しました。")
