"""
check_api_limit.py
eBay API のレートリミット状況を軽量に確認するプリフライトスクリプト

fetchItemAspects（重量・上限消費）は使わず、
get_default_category_tree_id（超軽量）だけを叩いてトークン取得 + API疎通を確認する。

Usage:
  python ebay-db/scripts/check_api_limit.py
  EBAY_CLIENT_ID=xxx EBAY_CLIENT_SECRET=xxx python ...
"""

import os
import sys
import requests

TAXONOMY_API = "https://api.ebay.com/commerce/taxonomy/v1"
CHECK_MARKETPLACES = ["EBAY_US", "EBAY_GB", "EBAY_DE", "EBAY_AU"]


def get_access_token(client_id: str, client_secret: str) -> str:
    resp = requests.post(
        "https://api.ebay.com/identity/v1/oauth2/token",
        auth=(client_id, client_secret),
        data={
            "grant_type": "client_credentials",
            "scope": "https://api.ebay.com/oauth/api_scope",
        },
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        timeout=15,
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


def check_default_tree(token: str, marketplace_id: str) -> dict:
    """get_default_category_tree_id: 非常に軽量な単発GETリクエスト"""
    url = f"{TAXONOMY_API}/get_default_category_tree_id"
    resp = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
        params={"marketplace_id": marketplace_id},
        timeout=15,
    )
    return {
        "marketplace_id": marketplace_id,
        "status_code": resp.status_code,
        "ok": resp.status_code == 200,
        "body": resp.json() if resp.status_code == 200 else resp.text[:200],
    }


def main():
    client_id     = os.environ.get("EBAY_CLIENT_ID")
    client_secret = os.environ.get("EBAY_CLIENT_SECRET")

    if not client_id or not client_secret:
        print("❌ 環境変数 EBAY_CLIENT_ID / EBAY_CLIENT_SECRET が未設定")
        sys.exit(1)

    print("=== eBay API 疎通チェック ===\n")

    # Step1: OAuth トークン取得
    print("[1] OAuth トークン取得...")
    try:
        token = get_access_token(client_id, client_secret)
        print("    ✅ トークン取得成功\n")
    except Exception as e:
        print(f"    ❌ トークン取得失敗: {e}")
        sys.exit(1)

    # Step2: 各マーケットで get_default_category_tree_id を確認
    print("[2] マーケット別 API 疎通確認（fetchItemAspects は叩かない）")
    all_ok = True
    for mp in CHECK_MARKETPLACES:
        result = check_default_tree(token, mp)
        if result["ok"]:
            tree_id = result["body"].get("categoryTreeId", "?")
            version = result["body"].get("categoryTreeVersion", "?")
            print(f"    ✅ {mp}: tree_id={tree_id}, version={version}")
        else:
            sc = result["status_code"]
            body = result["body"]
            if sc == 429:
                print(f"    ❌ {mp}: 429 Too Many Requests → まだレートリミット中")
            else:
                print(f"    ❌ {mp}: HTTP {sc} → {body}")
            all_ok = False

    print()
    if all_ok:
        print("✅ 全マーケット疎通OK → sync-ebay-db.yml の手動実行が可能です")
        print("   gh workflow run sync-ebay-db.yml --repo shingo-ops/bay-auto")
    else:
        print("❌ 一部マーケットで問題あり → まだ待つ必要があります")
        sys.exit(1)


if __name__ == "__main__":
    main()
