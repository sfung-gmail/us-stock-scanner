import datetime
import os
import pandas as pd
import requests
import yfinance as yf

# 填入你的免費 FMP API Key (於 financialmodelingprep.com 申請)
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

    # 技術指標計算
    df["20MA"] = df["Close"].rolling(window=20).mean()
    df["50MA"] = df["Close"].rolling(window=50).mean()
    df["30VolMA"] = df["Volume"].rolling(window=30).mean()
    
    # 純 Pandas 計算標準 RSI(14)
    delta = df["Close"].diff()
    gain = (delta.where(delta > 0, 0)).rolling(window=14).mean()
    loss = (-delta.where(delta < 0, 0)).rolling(window=14).mean()
    rs = gain / loss
    df["RSI14"] = 100 - (100 / (1 + rs))


    latest = df.iloc[-1]
    prev_close = df.iloc[-2]["Close"]
    fib_levels, fib_anchors = calculate_fibonacci(df)
    fast_info = stock.fast_info

    # 建立統欄位字典
    data = {
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
        # 以下欄位將根據類型動態填入
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
    # 📌 1. 在這裡輸入你每天想要抓取的標的清單 (可自由增減個股或ETF)
    watchlist = ["AAPL", "NVDA", "TSLA", "QQQ", "SPY"]

    results = []
    for ticker in watchlist:
        stock_data = fetch_stock_data(ticker)
        if stock_data:
            results.append(stock_data)

    # 📌 2. 轉為 Pandas DataFrame
    final_df = pd.DataFrame(results)

    # 📌 3. 設定儲存檔名 (加上當天日期，防止舊檔案被覆蓋)
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    output_filename = f"market_data_{today_str}.csv"

    # 📌 4. 導出為 CSV (使用 utf-8-sig 確保中文在 Excel 中開啟不亂碼)
    final_df.to_csv(output_filename, index=False, encoding="utf-8-sig")

    print(f"\n======================================")
    print(f"🎉 數據抓取完成！已成功儲存至檔案: {os.path.abspath(output_filename)}")
    print(f"======================================")
