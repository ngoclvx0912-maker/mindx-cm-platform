#!/usr/bin/env python3
"""
MindX CM Report Platform — Telegram Bridge
Cầu nối giữa Platform và Telegram của SOD

Hai tính năng chính:
  1. send_to_sod: Gửi câu hỏi ẩn danh từ CM lên Telegram SOD
  2. telegram_webhook: Nhận reply từ SOD qua Telegram → lưu vào Sheets

Config:
  - TELEGRAM_BOT_TOKEN: Token của Telegram Bot
  - TELEGRAM_CHAT_ID: Chat ID của SOD (hoặc group SOD)
"""
import json
import sys
import os
import re
import urllib.request
import urllib.parse

# ===================== CONFIG =====================
# Cấu hình Telegram Bot
# Thay bằng token thực sau khi tạo bot qua @BotFather
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "8661328396:AAECdS1Lxac8uHtbbuez4k2dtWUptkURHKc")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "-5190375492")

APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyJpFo_X0M_uvSPCStTOPVIAganyFRaN2AaCGO-Ukv911nlVS3me-C0jfIc8AiTSF1V/exec"

TELEGRAM_API_BASE = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}"


def send_response(data, status=200):
    """In HTTP response headers + JSON body"""
    print(f"Status: {status}")
    print("Content-Type: application/json")
    print("Access-Control-Allow-Origin: *")
    print("Access-Control-Allow-Methods: POST, OPTIONS")
    print("Access-Control-Allow-Headers: Content-Type")
    print()
    print(json.dumps(data, ensure_ascii=False))


def send_telegram_message(chat_id: str, text: str, reply_to_message_id: int = None) -> dict:
    """Gửi tin nhắn qua Telegram Bot API"""
    payload = {
        "chat_id": chat_id,
        "text": text,
        "parse_mode": "HTML"
    }
    if reply_to_message_id:
        payload["reply_to_message_id"] = reply_to_message_id

    try:
        data = json.dumps(payload).encode('utf-8')
        req = urllib.request.Request(
            f"{TELEGRAM_API_BASE}/sendMessage",
            data=data,
            headers={"Content-Type": "application/json"},
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=15) as response:
            return json.loads(response.read().decode('utf-8'))
    except Exception as e:
        return {"ok": False, "error": str(e)}


def save_sod_answer(question_id: str, content: str) -> dict:
    """Lưu câu trả lời của SOD vào Google Sheets qua Apps Script"""
    try:
        post_body = json.dumps({
            "type": "sod_answer",
            "parent_id": question_id,
            "bu": "SOD",
            "content": content
        }).encode('utf-8')

        url = f"{APPS_SCRIPT_URL}?action=qa_post"
        req = urllib.request.Request(
            url,
            data=post_body,
            headers={"Content-Type": "text/plain"},
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=30) as response:
            return json.loads(response.read().decode('utf-8'))
    except Exception as e:
        return {"success": False, "error": str(e)}


def handle_send_to_sod(data: dict):
    """
    Gửi câu hỏi ẩn danh của CM lên Telegram SOD
    Body: { question_id, question_content }
    """
    question_id = data.get('question_id', '')
    question_content = data.get('question_content', '')

    if not question_id or not question_content:
        send_response({"success": False, "error": "Missing question_id or question_content"}, 400)
        return

    # Kiểm tra token đã cấu hình chưa
    if TELEGRAM_BOT_TOKEN == "PLACEHOLDER_TOKEN" or TELEGRAM_CHAT_ID == "PLACEHOLDER_CHAT_ID":
        # Mode giả lập — trả về success để frontend không bị lỗi
        send_response({
            "success": True,
            "simulated": True,
            "message": "Telegram chưa được cấu hình. Câu hỏi đã được đánh dấu chờ SOD.",
            "question_id": question_id
        })
        return

    # Format tin nhắn gửi cho SOD
    message_text = (
        f"📩 <b>Câu hỏi mới từ CM (Ẩn danh)</b>\n"
        f"━━━━━━━━━━━━━━━━━━━━\n\n"
        f"{question_content}\n\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"💡 <i>Reply tin nhắn này để trả lời CM.</i>\n"
        f"🔑 <code>QID:{question_id}</code>"
    )

    result = send_telegram_message(TELEGRAM_CHAT_ID, message_text)

    if result.get('ok'):
        message_id = result.get('result', {}).get('message_id')
        send_response({
            "success": True,
            "message_id": message_id,
            "question_id": question_id
        })
    else:
        send_response({
            "success": False,
            "error": result.get('description', result.get('error', 'Telegram API error'))
        }, 500)


def handle_telegram_webhook(data: dict):
    """
    Xử lý webhook từ Telegram khi SOD reply
    Telegram gửi update object → extract reply text → tìm question_id → lưu vào Sheets
    """
    # Xử lý Telegram Update object
    message = data.get('message', {})
    if not message:
        # Telegram có thể gửi edited_message, callback_query, etc.
        send_response({"ok": True, "handled": False, "reason": "not a message"})
        return

    # Lấy text của reply
    text = message.get('text', '').strip()
    if not text:
        send_response({"ok": True, "handled": False, "reason": "no text"})
        return

    # Kiểm tra có phải reply đến tin nhắn của bot không
    reply_to = message.get('reply_to_message', {})
    if not reply_to:
        # Không phải reply → bỏ qua
        send_response({"ok": True, "handled": False, "reason": "not a reply"})
        return

    # Tìm question_id trong tin nhắn gốc của bot
    original_text = reply_to.get('text', '')
    qid_match = re.search(r'QID:(qa_[^\s\n]+)', original_text)

    if not qid_match:
        send_response({"ok": True, "handled": False, "reason": "QID not found in original message"})
        return

    question_id = qid_match.group(1)

    # Lưu câu trả lời SOD vào Google Sheets
    save_result = save_sod_answer(question_id, text)

    if save_result.get('success'):
        # Gửi xác nhận lại cho SOD
        from_user = message.get('from', {})
        first_name = from_user.get('first_name', 'SOD')
        confirm_text = f"✅ Câu trả lời của bạn đã được ghi nhận và hiển thị trên platform."
        msg_id = message.get('message_id')
        send_telegram_message(str(message.get('chat', {}).get('id', TELEGRAM_CHAT_ID)), confirm_text, msg_id)

        send_response({
            "ok": True,
            "handled": True,
            "question_id": question_id,
            "answer_id": save_result.get('id', '')
        })
    else:
        send_response({
            "ok": False,
            "error": save_result.get('error', 'Failed to save answer')
        }, 500)


def main():
    method = os.environ.get('REQUEST_METHOD', 'GET').upper()
    query_string = os.environ.get('QUERY_STRING', '')

    # Parse query params
    params = urllib.parse.parse_qs(query_string)
    action = params.get('action', [''])[0]

    if method == 'OPTIONS':
        print("Status: 200")
        print("Content-Type: application/json")
        print("Access-Control-Allow-Origin: *")
        print("Access-Control-Allow-Methods: POST, OPTIONS")
        print("Access-Control-Allow-Headers: Content-Type")
        print()
        return

    if method != 'POST':
        send_response({"error": "Method not allowed — dùng POST"}, 405)
        return

    # Đọc body
    try:
        content_length = int(os.environ.get('CONTENT_LENGTH', 0))
        body = sys.stdin.read(content_length) if content_length > 0 else sys.stdin.read()
        data = json.loads(body) if body.strip() else {}
    except json.JSONDecodeError as e:
        send_response({"error": f"Invalid JSON body: {str(e)}"}, 400)
        return
    except Exception as e:
        send_response({"error": f"Error reading body: {str(e)}"}, 400)
        return

    # Route theo action
    if action == 'send_to_sod':
        handle_send_to_sod(data)
    elif action == 'telegram_webhook':
        handle_telegram_webhook(data)
    else:
        send_response({"error": f"Unknown action: '{action}'. Use: send_to_sod, telegram_webhook"}, 400)


if __name__ == '__main__':
    main()
