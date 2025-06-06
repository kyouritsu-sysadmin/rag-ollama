# teams_webhook.py - Logic Apps Workflowに対応した通知機能（再修正版）
import requests
import logging
import json
from datetime import datetime
import traceback
import re
import os

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)  # 詳細なログを有効化

class TeamsWebhook:
    def __init__(self, webhook_url):
        """
        Teams Webhookクラスの初期化

        Args:
            webhook_url: Teams Workflow URL (Logic Apps URL)
        """
        self.webhook_url = webhook_url
        logger.info(f"Teams Workflowを初期化: {webhook_url[:30]}...")

    def send_ollama_response(self, query, response, conversation_data=None, search_path=None):
        """
        Ollamaの応答をTeams Workflowに送信する

        Args:
            query: ユーザーの質問
            response: Ollamaからの応答
            conversation_data: 会話データ（オプション）
            search_path: 検索に使用したディレクトリパス（オプション）

        Returns:
            dict: 送信結果
        """
        try:
            # 検索パスの短縮表示
            short_path = self._get_shortened_path(search_path) if search_path else "設定されたディレクトリ"
            
            # 現在の日時
            now = datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')
            
            # PDF検出用の正規表現
            pdf_detected = "PDFファイル" in response and ("見つかりました" in response or "存在します" in response)
            
            # PDFが見つかった場合はカラーをアクセントに
            card_color = "Accent" if pdf_detected else "Default"
            
            # --- ここから分割処理を追加 ---
            def split_text_blocks(text, max_len=900):
                blocks = []
                lines = text.split('\n')
                buf = ""
                for line in lines:
                    # Markdownリストや強調をTeams向けに変換
                    line = line.replace("**", "").replace("* ", "・")
                    if len(buf) + len(line) + 1 > max_len:
                        blocks.append(buf)
                        buf = line
                    else:
                        buf += ("\n" if buf else "") + line
                if buf:
                    blocks.append(buf)
                return blocks

            # --- ここから恒久対策ロジックを追加 ---
            def remove_control_chars(s):
                # 制御文字・サロゲートペア等を除去
                return re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\uD800-\uDFFF\uFFFE\uFFFF]', '', s)

            response = remove_control_chars(response)

            # TextBlock分割（最大10個、合計9000文字まで）
            MAX_BLOCKS = 10
            MAX_TOTAL_LEN = 9000
            response_blocks = split_text_blocks(response)
            total_len = 0
            limited_blocks = []
            for block in response_blocks:
                if len(limited_blocks) >= MAX_BLOCKS or total_len + len(block) > MAX_TOTAL_LEN:
                    break
                limited_blocks.append(block)
                total_len += len(block)
            if len(response_blocks) > MAX_BLOCKS or total_len < len(response):
                limited_blocks.append("（一部省略されています。全文は管理者にお問い合わせください）")
            response_textblocks = [
                {
                    "type": "TextBlock",
                    "text": block,
                    "wrap": True,
                    "spacing": "Medium"
                } for block in limited_blocks
            ]

            # --- AdaptiveCard本体を組み立て直し ---
            root_attachments_payload = {
                "attachments": [
                    {
                        "contentType": "application/vnd.microsoft.card.adaptive",
                        "content": {
                            "type": "AdaptiveCard",
                            "body": [
                                {
                                    "type": "TextBlock",
                                    "size": "Medium",
                                    "weight": "Bolder",
                                    "text": "Ollama回答",
                                    "wrap": True,
                                    "color": card_color
                                },
                                {
                                    "type": "TextBlock",
                                    "text": f"質問: {query}",
                                    "wrap": True,
                                    "weight": "Bolder",
                                    "color": "Accent"
                                },
                                {
                                    "type": "TextBlock",
                                    "text": f"検索対象: {short_path}",
                                    "wrap": True,
                                    "isSubtle": True,
                                    "size": "Small"
                                },
                            ] + response_textblocks + [
                                {
                                    "type": "TextBlock",
                                    "text": f"回答生成時刻: {now}",
                                    "wrap": True,
                                    "size": "Small",
                                    "isSubtle": True
                                }
                            ],
                            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                            "version": "1.0"
                        }
                    }
                ]
            }

            # AdaptiveCard全体のバイト長をチェック（25KB超なら省略）
            card_bytes = json.dumps(root_attachments_payload, ensure_ascii=False).encode('utf-8')
            if len(card_bytes) > 25000:
                # 省略メッセージのみのTextBlockに差し替え
                root_attachments_payload["attachments"][0]["content"]["body"] = [
                    {
                        "type": "TextBlock",
                        "text": "（回答が長すぎるため一部省略されています。全文は管理者にお問い合わせください）",
                        "wrap": True,
                        "spacing": "Medium"
                    }
                ]

            # バックアップ用のシンプルなペイロード
            simple_payload = {
                "text": f"### Ollama回答\n\n**質問**: {query}\n\n**検索対象**: {short_path}\n\n{response}\n\n*回答生成時刻: {now}*"
            }

            # 旧形式のペイロード（既存形式）
            legacy_payload = {
                "body": {
                    "attachments": [
                        {
                            "contentType": "application/vnd.microsoft.card.adaptive",
                            "content": {
                                "type": "AdaptiveCard",
                                "body": [
                                    {
                                        "type": "TextBlock",
                                        "size": "Medium",
                                        "weight": "Bolder",
                                        "text": "Ollama回答",
                                        "wrap": True,
                                        "color": card_color
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": f"質問: {query}",
                                        "wrap": True,
                                        "weight": "Bolder"
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": f"検索対象: {short_path}",
                                        "wrap": True,
                                        "size": "Small"
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": response,
                                        "wrap": True
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": f"回答生成時刻: {now}",
                                        "wrap": True,
                                        "size": "Small",
                                        "isSubtle": True
                                    }
                                ],
                                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                                "version": "1.0"
                            }
                        }
                    ]
                }
            }

            # リクエストヘッダー
            headers = {
                'Content-Type': 'application/json'
            }

            # 1. まず新しいルートレベルのattachments形式で試行
            logger.debug(f"Logic Apps送信ペイロード(ルートレベルattachments): {json.dumps(root_attachments_payload)[:300]}...")

            try:
                r = requests.post(
                    self.webhook_url, 
                    json=root_attachments_payload, 
                    headers=headers,
                    timeout=30
                )
                logger.debug(f"Logic Apps応答(ルートレベルattachments): {r.status_code}, {r.text[:100] if r.text else '空のレスポンス'}")

                if r.status_code >= 200 and r.status_code < 300:
                    logger.info(f"ルートレベルattachments形式でのLogic Apps通知送信成功: {r.status_code}")
                    return {"status": "success", "code": r.status_code, "format": "ルートレベルattachments"}
                else:
                    logger.warning(f"ルートレベルattachments形式での送信失敗: {r.status_code}。従来形式で再試行します。")

            except Exception as e:
                logger.warning(f"ルートレベルattachments形式送信エラー: {str(e)}。従来形式で再試行します。")

            # 2. 従来形式で試行
            logger.debug(f"Logic Apps送信ペイロード(従来形式): {json.dumps(legacy_payload)[:300]}...")

            try:
                r2 = requests.post(
                    self.webhook_url, 
                    json=legacy_payload, 
                    headers=headers,
                    timeout=30
                )
                logger.debug(f"Logic Apps応答(従来形式): {r2.status_code}, {r2.text[:100] if r2.text else '空のレスポンス'}")

                if r2.status_code >= 200 and r2.status_code < 300:
                    logger.info(f"従来形式でのLogic Apps通知送信成功: {r2.status_code}")
                    return {"status": "success", "code": r2.status_code, "format": "従来形式"}
                else:
                    logger.warning(f"従来形式での送信失敗: {r2.status_code}。シンプル形式で再試行します。")

            except Exception as e2:
                logger.warning(f"従来形式送信エラー: {str(e2)}。シンプル形式で再試行します。")

            # 3. シンプル形式で試行（最後の手段）
            logger.debug(f"Logic Apps送信ペイロード(シンプル): {json.dumps(simple_payload)[:300]}...")

            try:
                r3 = requests.post(
                    self.webhook_url, 
                    json=simple_payload, 
                    headers=headers,
                    timeout=30
                )
                logger.debug(f"Logic Apps応答(シンプル): {r3.status_code}, {r3.text[:100] if r3.text else '空のレスポンス'}")

                if r3.status_code >= 200 and r3.status_code < 300:
                    logger.info(f"シンプル形式でのLogic Apps通知送信成功: {r3.status_code}")
                    return {"status": "success", "code": r3.status_code, "format": "シンプル"}
                else:
                    logger.error(f"Logic Apps通知の送信に全て失敗しました: 最終ステータスコード={r3.status_code}")
                    return {"status": "error", "code": r3.status_code, "message": r3.text}

            except Exception as e3:
                logger.error(f"シンプル形式送信エラー: {str(e3)}")
                return {"status": "error", "message": str(e3)}

        except Exception as e:
            logger.error(f"Logic Apps通知の送信中にエラーが発生しました: {str(e)}")
            logger.error(traceback.format_exc())
            return {"status": "error", "message": str(e)}
    
    def _get_shortened_path(self, path):
        """
        長いパスを短縮して表示

        Args:
            path: 元のパス文字列

        Returns:
            短縮されたパス
        """
        if not path:
            return "デフォルト検索ディレクトリ"
            
        # ユーザー名を取得
        username = os.getenv("USERNAME", "owner")
        
        # パスが長い場合は短縮表示
        if len(path) > 50:
            # OneDriveパスの特定のパターンを検出
            if "OneDrive" in path:
                # 会社名を含むOneDriveパスのパターン
                company_match = re.search(r'OneDrive - ([^\\]+)', path)
                if company_match:
                    company = company_match.group(1)
                    # 短縮した会社名
                    short_company = company[:10] + "..." if len(company) > 10 else company
                    # パスの後半部分を取得
                    path_parts = path.split("\\")
                    if len(path_parts) > 3:
                        last_parts = path_parts[-3:]
                        return f"OneDrive - {short_company}\\...\\{last_parts[-3]}\\{last_parts[-2]}\\{last_parts[-1]}"
                    
            # 一般的な短縮
            path_parts = path.split("\\")
            if len(path_parts) > 3:
                first_part = path_parts[0]
                if ":" in first_part:  # ドライブレター
                    first_part = path_parts[0] + "\\" + path_parts[1]
                last_parts = path_parts[-2:]
                return f"{first_part}\\...\\{last_parts[0]}\\{last_parts[1]}"
        
        return path

    def send_direct_message(self, query, response, search_path=None):
        """
        send_ollama_responseのエイリアス（下位互換性のため）
        """
        return self.send_ollama_response(query, response, None, search_path)
