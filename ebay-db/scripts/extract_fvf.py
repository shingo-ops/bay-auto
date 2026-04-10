"""
extract_fvf.py
eBay 公式料率ページ HTML → Gemini 2.5 Flash-Lite で FVF レートを抽出
出力: fvf_rates.json  { "EBAY_US": { "category_id": rate, ... }, ... }

503 等でページ取得不可の場合は FVF_RATES_DEFAULT にフォールバック。
"""

import os
import json
import requests

GEMINI_API = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite:generateContent"

FVF_PAGES = {
    "EBAY_US": "https://www.ebay.com/help/selling/fees-credits-invoices/selling-fees?id=4822",
    "EBAY_GB": "https://www.ebay.co.uk/help/selling/fees-credits-invoices/selling-fees?id=4822",
}

# eBay公式FVFレート（2026年料率）ハードコードフォールバック
# 503等でページ取得不可時に使用
FVF_RATES_DEFAULT: dict = {
    "Most categories":             {"rate": 13.6,  "threshold": 7500,  "rate_above": 2.35},
    "Books & Magazines":           {"rate": 15.3,  "threshold": 7500,  "rate_above": 2.35},
    "Movies & TV":                 {"rate": 15.3,  "threshold": 7500,  "rate_above": 2.35},
    "Music":                       {"rate": 15.3,  "threshold": 7500,  "rate_above": 2.35},
    "Coins & Paper Money":         {"rate": 13.25, "threshold": 7500,  "rate_above": 2.35},
    "Bullion":                     {"rate": 13.6,  "threshold": 7500,  "rate_above": 7.0},
    "Women's Bags & Handbags":     {"rate": 15.0,  "threshold": 2000,  "rate_above": 9.0},
    "Trading Cards":               {"rate": 13.25, "threshold": 7500,  "rate_above": 2.35},
    "Collectible Card Games":      {"rate": 13.25, "threshold": 7500,  "rate_above": 2.35},
    "Comic Books & Memorabilia":   {"rate": 13.25, "threshold": 7500,  "rate_above": 2.35},
    "Jewelry & Watches":           {"rate": 15.0,  "threshold": 5000,  "rate_above": 9.0},
    "Watches, Parts & Accessories":{"rate": 15.0,  "threshold": 1000,  "rate_above": 6.5,
                                    "threshold2": 7500, "rate_above2": 3.0},
    "NFTs":                        {"rate": 5.0},
    "Heavy Equipment":             {"rate": 3.0,   "threshold": 15000, "rate_above": 0.5},
    "Guitars & Basses":            {"rate": 6.7,   "threshold": 7500,  "rate_above": 2.35},
    "Athletic Shoes":              {"rate": 8.0,   "note": "$150以上。$150未満は13.6%"},
}


def convert_default_rates(defaults: dict) -> dict:
    """FVF_RATES_DEFAULT を generate_csv.py が期待する形式に変換

    変換後形式: {category_name: {"fvf_rate": float, "fvf_note": str}}
    """
    result = {}
    for name, v in defaults.items():
        rate = v["rate"]
        note_parts = []
        if "threshold" in v and "rate_above" in v:
            note_parts.append(f"${v['threshold']:,}超は{v['rate_above']}%")
        if "threshold2" in v and "rate_above2" in v:
            note_parts.append(f"${v['threshold2']:,}超は{v['rate_above2']}%")
        if "note" in v:
            note_parts.append(v["note"])
        result[name] = {
            "fvf_rate": rate,
            "fvf_note": "、".join(note_parts),
        }
    return result


def notify_discord_fallback(marketplace_id: str, status_code: int) -> None:
    """フォールバック使用時に Discord へ通知"""
    webhook_url = os.environ.get("DISCORD_WEBHOOK") or os.environ.get("DISCORD_WEBHOOK_AUDIT")
    if not webhook_url:
        print(f"  [INFO] Discord Webhook 未設定のため通知スキップ")
        return
    msg = (
        f"⚠️ **ebay-db sync** [{marketplace_id}]\n"
        f"FVFページ取得失敗（HTTP {status_code}）\n"
        f"FVFレートはハードコード値を使用（eBay公式2026年料率）"
    )
    try:
        requests.post(webhook_url, json={"content": msg}, timeout=10)
    except Exception as e:
        print(f"  [WARN] Discord通知失敗: {e}")


def fetch_fvf_page(url: str) -> str:
    """料率ページの HTML を取得"""
    headers = {"User-Agent": "Mozilla/5.0 (compatible; ebay-db-sync/1.0)"}
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    return resp.text[:50000]  # 最大 50,000 文字（トークン節約）


def extract_fvf_with_gemini(html: str, marketplace_id: str) -> dict:
    """Gemini 2.5 Flash-Lite で HTML からカテゴリ別 FVF レートを抽出

    戻り値: { category_name: {"fvf_rate": float, "fvf_note": str} }
    fvf_note は段階的料率・例外条件など補足情報（例: "$150以上は8%"）。
    シンプルな固定料率のカテゴリは空文字。
    """
    api_key = os.environ["GEMINI_API_KEY"]

    prompt = (
        f"以下は eBay ({marketplace_id}) の販売手数料ページの HTML です。\n"
        "カテゴリ名・最終価値手数料(FVF)率(%)・補足情報のリストを JSON 配列で返してください。\n"
        "fvf_rate は代表的な料率（%）を数値で入力してください。\n"
        "fvf_note には段階的料率・例外条件・上限金額など補足情報を日本語で入力してください。"
        "補足なしの場合は空文字にしてください。\n"
        "例: [{\"category_name\": \"Electronics\", \"fvf_rate\": 13.25, \"fvf_note\": \"$7,500超は2.35%\"}]\n\n"
        f"HTML:\n{html}"
    )

    body = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "responseMimeType": "application/json",
            "responseSchema": {
                "type": "ARRAY",
                "items": {
                    "type": "OBJECT",
                    "properties": {
                        "category_name": {"type": "STRING"},
                        "fvf_rate": {"type": "NUMBER"},
                        "fvf_note": {"type": "STRING"},
                    },
                    "required": ["category_name", "fvf_rate", "fvf_note"],
                },
            },
        },
    }

    resp = requests.post(
        f"{GEMINI_API}?key={api_key}",
        json=body,
        headers={"Content-Type": "application/json"},
        timeout=60,
    )
    resp.raise_for_status()

    data = resp.json()
    text = data["candidates"][0]["content"]["parts"][0]["text"]
    rates_list = json.loads(text)

    # category_name → {"fvf_rate": ..., "fvf_note": ...} の辞書に変換
    return {
        item["category_name"]: {
            "fvf_rate": item["fvf_rate"],
            "fvf_note": item.get("fvf_note", ""),
        }
        for item in rates_list
    }


def main():
    print("=== extract_fvf.py 開始 ===")
    all_rates = {}

    for marketplace_id, url in FVF_PAGES.items():
        print(f"取得中: {marketplace_id} ({url})")
        try:
            html = fetch_fvf_page(url)
            rates = extract_fvf_with_gemini(html, marketplace_id)
            all_rates[marketplace_id] = rates
            print(f"  → {len(rates)} カテゴリのレートを抽出")
        except requests.exceptions.HTTPError as e:
            status = e.response.status_code if e.response is not None else 0
            print(f"  ⚠️ {marketplace_id} HTTP {status} エラー → ハードコード値にフォールバック")
            all_rates[marketplace_id] = convert_default_rates(FVF_RATES_DEFAULT)
            notify_discord_fallback(marketplace_id, status)
        except Exception as e:
            print(f"  ⚠️ {marketplace_id} 取得失敗: {e}")
            all_rates[marketplace_id] = {}

    output_path = os.environ.get("OUTPUT_DIR", ".") + "/fvf_rates.json"
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(all_rates, f, ensure_ascii=False, indent=2)

    print(f"=== 完了 → {output_path} ===")


if __name__ == "__main__":
    main()
