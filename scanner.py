import datetime
import os
import pandas as pd
import requests
import yfinance as yf

# =====================================================================
# 📌 1. 修改持股請看這裡！
# 你可以在這個清單內「自由增減」你想追蹤的個股或 ETF 代號（記得用雙引號與逗號隔開）
# =====================================================================
watchlist = ["MU", "TSM", "SOXQ", "VOO", "BRK.B", "XLV"]

# 📌 固定輸出的總數據庫檔名
OUTPUT_FILENAME = "my_holdings_history.csv"

# FMP API Key (自動抓取 GitHub 保險箱)
FMP_API_KEY = os.environ.get("FMP_API_KEY", "YOUR_FMP_API_KEY")


def get_fmp_earnings(ticker):
    """獲取財報數據 (下次財報、共識EPS、上季Beat/Miss)"""
    try:
        # 1. 下次財報與共識 EPS
        res = requests.get(
            f"https://financialmodelingprep.com{ticker}&apikey={FMP_API_KEY}"
        ).json()
        next_date, consensus_eps = "N/A", "N/A"
        today = datetime.date.today().strftime("%Y-%m-%d")
        for event in res:
            if event.get("date") >= today:
                next_date = event.get("date")
                consensus_eps = event.get("epsEstimated")
                break

        # 2. 上季財報表現
        res_hist = requests.get(
            f"https://financialmodelingprep.com{ticker}?apikey={FMP_API_KEY}"
        ).json()
        beat_miss = "N/A"
        if res_hist:
            last_q = res_hist[0]
            actual = last_q.get("actualEps", 0)
            est = last_q.get("estimatedEps", 0)
            if actual is not None and est is not None:
                beat_miss = "Beat" if actual >= est else "Miss"

        return next_date, consensus_eps, beat_miss
    except:
        return "N/A", "N/A", "N/A"


def calculate_fibonacci(df):
    """計算斐波那契回撤線（以近252個交易日最高/最低點為波段錨點）"""
    high_idx = df["High"].idxmax()
    low_idx = df["Low"].idxmin()
    high_val = df["High"].max()
    low_val = df["Low"].min()

    anchor_info = (
        f"低點轉高點 (多頭波段)" if low_idx < high_idx else f"高點轉低點 (空頭波段)"
    )
    diff = high_val - low_val

    if low_idx < high_idx:
        levels = {
            "23.6%": high_val - diff * 0.236,
            "38.2%": high_val - diff * 0.382,
            "50.0%": high_val - diff * 0.5,
            "61.8%": high_val - diff * 0.618,
            "78.6%": high_val - diff * 0.786,
        }
    else:
        levels = {
            "23.6%": low_val + diff * 0.236,
            "38.2%": low_val + diff * 0.382,
            "50.0%": low_val + diff * 0.5,
            "61.8%": low_val + diff * 0.618,
            "78.6%": low_val + diff * 0.786,
        }

    return levels, f"最高:{high_idx.date()} 最低:{low_idx.date()} ({anchor_info})"


def fetch_stock_data(ticker):
    print(f"正在抓取 {ticker}...")
    stock = yf.Ticker(ticker)
    info = stock.info
    is_etf = info.get("quoteType") == "ETF"

    df = stock.history(period="1y")
    if df.empty:
        print(f"⚠️ 找不到 {ticker} 的歷史數據")
        return None

    # 純 Pandas 技術指標計算 (免第三方庫)
    df["20MA"] = df["Close"].rolling(window=20).mean()
    df["50MA"] = df["Close"].rolling(window=50).mean()
    df["30VolMA"] = df["Volume"].rolling(window=30).mean()

    delta = df["Close"].diff()
    gain = (delta.where(delta > 0, 0)).rolling(window=14).mean()
    loss = (-delta.where(delta < 0, 0)).rolling(window=14).mean()
    rs = gain / loss
    df["RSI14"] = 100 - (100 / (1 + rs))

    latest = df.iloc[-1]
    prev_close = df.iloc[-2]["Close"]
    fib_levels, fib_anchors = calculate_fibonacci(df)
    fast_info = stock.fast_info

    today_str = datetime.date.today().strftime("%Y-%m-%d")

    data = {
        "資料日期": today_str,  # ⭐ 新增日期欄位，防止合併後混淆
        "標的代號": ticker,
        "類型": "ETF" if is_etf else "個股",
        "開市價": round(latest["Open"], 2),
        "收市價": round(latest["Close"], 2),
        "日漲跌%": f"{((latest['Close'] - prev_close) / prev_close) * 100:.2f}%",
        "盤前/盤後價": info.get("preMarketPrice")
        or info.get("postMarketPrice")
        or "實時休市",
        "當日最低": round(fast_info["dayLow"], 2),
        "當日最高": round(fast_info["dayHigh"], 2),
        "52週最低": round(fast_info["yearLow"], 2),
        "52週最高": round(fast_info["yearHigh"], 2),
        "20D MA": round(latest["20MA"], 2) if not pd.isna(latest["20MA"]) else "N/A",
        "50D MA": round(latest["50MA"], 2) if not pd.isna(latest["50MA"]) else "N/A",
        "成交量": int(latest["Volume"]),
        "30日均量": int(latest["30VolMA"])
        if not pd.isna(latest["30VolMA"])
        else "N/A",
        "RelVol (相對成交量)": round(latest["Volume"] / latest["30VolMA"], 2)
        if not pd.isna(latest["30VolMA"])
        else "N/A",
        "RSI(14)": round(latest["RSI14"], 2)
        if not pd.isna(latest["RSI14"])
        else "N/A",
        "市值": info.get("marketCap", "N/A"),
        "Fibonacci波段錨點": fib_anchors,
        "Fib 23.6%": round(fib_levels["23.6%"], 2),
        "Fib 38.2%": round(fib_levels["38.2%"], 2),
        "Fib 50.0%": round(fib_levels["50.0%"], 2),
        "Fib 61.8%": round(fib_levels["61.8%"], 2),
        "Fib 78.6%": round(fib_levels["78.6%"], 2),
        "下次財報日": "N/A",
        "共識EPS": "N/A",
        "上季財報Beat/Miss": "N/A",
        "產業": "N/A",
        "Beta": "N/A",
        "營收年增": "N/A",
        "配息率": "N/A",
        "費用率": "N/A",
        "AUM (總資產)": "N/A",
        "前5大持股": "N/A",
    }

    if not is_etf:
        next_date, eps, beat_miss = get_fmp_earnings(ticker)
        data.update(
            {
                "下次財報日": next_date,
                "共識EPS": eps,
                "上季財報Beat/Miss": beat_miss,
                "產業": info.get("industry", "N/A"),
                "Beta": info.get("beta", "N/A"),
                "營收年增": f"{info.get('revenueGrowth', 0)*100:.2f}%"
                if info.get("revenueGrowth")
                else "N/A",
            }
        )
    else:
        try:
            holdings_df = pd.DataFrame(stock.holdings)
            top5 = (
                ", ".join(holdings_df["holding"].head(5).tolist())
                if not holdings_df.empty
                else "N/A"
            )
        except:
            top5 = "N/A"

        data.update(
            {
                "配息率": f"{info.get('trailingAnnualDividendYield', 0)*100:.2f}%"
                if info.get("trailingAnnualDividendYield")
                else "N/A",
                "費用率": f"{info.get('feesExpensesTotalPercentage', 0)*100:.2f}%"
                if info.get("feesExpensesTotalPercentage")
                else "N/A",
                "AUM (總資產)": info.get("totalAssets", "N/A"),
                "前5大持股": top5,
            }
        )

    return data


if __name__ == "__main__":
    new_results = []
    for ticker in watchlist:
        stock_data = fetch_stock_data(ticker)
        if stock_data:
            new_results.append(stock_data)

    current_df = pd.DataFrame(new_results)

    # 📌 2. 核心邏輯：自動讀取舊檔並 Append 新數據
    if os.path.exists(OUTPUT_FILENAME):
        try:
            # 讀取現有的歷史數據
            old_df = pd.read_csv(OUTPUT_FILENAME)
            # 將今天的新數據合併到舊數據的下方
            final_df = pd.concat([old_df, current_df], ignore_index=True)
            # 移除可能重複抓取的同一天同一檔標的（防呆機制）
            final_df.drop_duplicates(
                subset=["資料日期", "標的代號"], keep="last", inplace=True
            )
            print("==== 成功讀取舊有數據庫，已完成新增（Append） ====")
        except Exception as e:
            print(f"⚠️ 讀取舊檔案失敗 ({e})，將重新建立新檔案。")
            final_df = current_df
    else:
        print("==== 找不到舊檔案，正在建立全新的數據庫 ====")
        final_df = current_df

    # 📌 3. 儲存為固定的 CSV 檔案名稱 (此網址之後在外部調用將永遠不變)
    final_df.to_csv(OUTPUT_FILENAME, index=False, encoding="utf-8-sig")
    print(f"🎉 歷史數據合併完成！儲存路徑: {os.path.abspath(OUTPUT_FILENAME)}")
