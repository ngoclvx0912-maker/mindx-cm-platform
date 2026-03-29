#!/usr/bin/env python3
"""
MindX CM Report Platform — MiMi Bot (Q&A AI Assistant)
Model: sonar-pro (Perplexity API) | Fallback: sonar

Tính năng "Hỏi MiMi":
  CM đặt câu hỏi về vận hành MindX → MiMi trả lời bằng tiếng Việt
  Nếu câu hỏi quá sâu về nội bộ → gợi ý nhờ SOD trả lời

Flow: CM nhấn "Hỏi MiMi" → gửi câu hỏi → MiMi trả lời → lưu vào Google Sheets
"""
import json
import sys
import os
import re
import urllib.request
import urllib.parse

# ===================== CONFIG =====================
PERPLEXITY_API_KEY = "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"
APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyJpFo_X0M_uvSPCStTOPVIAganyFRaN2AaCGO-Ukv911nlVS3me-C0jfIc8AiTSF1V/exec"

# ===================== CONTEXT MindX =====================
MINDX_CONTEXT = """
=== THÔNG TIN VỀ MindX Technology School ===

MindX Technology School là hệ thống trường công nghệ hàng đầu Việt Nam cho học sinh K12 (6-18 tuổi) và người lớn 18+.
Có 44-46 trung tâm (BU - Business Unit) trên toàn quốc.

CẤU TRÚC TỔ CHỨC:
- SOD (Sales Director): Giám đốc kinh doanh, quản lý toàn bộ hệ thống
- FM (Field Manager): Quản lý vùng, phụ trách một nhóm BU
- CM (Center Manager): Quản lý trung tâm, báo cáo trực tiếp lên FM/SOD
- Sale: Nhân viên kinh doanh tại BU
- CS (Customer Service): Chăm sóc học viên

MÔN HỌC: Coding (Python, Web, App), Robotics, AI, Game Design, Digital Art

OKRs 2026 (3 mục tiêu chính):
[N1 - OPTIMIZE] Target 154B - Tối ưu Leads MKT:
- CR16 (Conversion Rate Lead → Trial): 15% target
- CR46 (Trial → Deal): 50% target  
- AOV (Average Order Value): 18M target

[N2 - OPS] Target 174B - Re-enroll / Upsell / Referral:
- Retention: >60%
- Upsell: 25%
- Referral: 8%

[N3 - GROWTH] Target 100B - Sale tự kiếm:
- Direct Sales, Local Events, Partnership B2B

BÁO CÁO HÀNG TUẦN:
- Weekly Report (WR): CM báo cáo KPI hàng tuần (Target vs Actual)
- Action Plan (AP): CM lập kế hoạch hành động theo tuần để cải thiện KPIs

QUY TRÌNH VẬN HÀNH:
- Thứ 2-3: CM điền WR tuần trước + AP tuần mới
- Thứ 4: Deadline nộp báo cáo (lock sau 12h trưa Thứ 4)
- FM review và ghi chú feedback cho CM
- SOD tổng hợp, đánh giá toàn bộ hệ thống

THUẬT NGỮ THƯỜNG DÙNG:
- L1 Leads: Số lượng leads đến
- L4 Trials: Số học thử
- L6 Deals: Số hợp đồng ký
- Re-enroll: Gia hạn học
- Upsell: Nâng cấp gói học
- Referral: Giới thiệu học viên mới
"""

INTERNAL_KEYWORDS = [
    'lương', 'thưởng', 'kpi cá nhân', 'đánh giá nhân sự', 'sa thải', 'thôi việc',
    'chính sách nội bộ mới', 'chiến lược bí mật', 'ngân sách', 'chi phí',
    'hợp đồng cụ thể', 'số liệu tài chính', 'mức target cụ thể của tôi'
]


def is_internal_question(question: str) -> bool:
    """Kiểm tra xem câu hỏi có quá nội bộ không"""
    q_lower = question.lower()
    for kw in INTERNAL_KEYWORDS:
        if kw in q_lower:
            return True
    return False


def call_ai(question: str, question_id: str) -> str:
    """Gọi Perplexity API để trả lời câu hỏi"""
    # Kiểm tra câu hỏi quá nội bộ
    if is_internal_question(question):
        return "Câu hỏi này cần SOD trả lời trực tiếp. Bạn có thể nhấn 'Nhờ SOD trả lời' để được hỗ trợ."

    system_prompt = f"""Bạn là MiMi — trợ lý AI thân thiện của MindX Technology School.
Nhiệm vụ của bạn là hỗ trợ các Center Manager (CM) giải đáp thắc mắc về vận hành trung tâm.

{MINDX_CONTEXT}

HƯỚNG DẪN TRẢ LỜI:
1. Luôn trả lời bằng tiếng Việt, giọng thân thiện, dễ hiểu
2. Câu trả lời ngắn gọn, đi thẳng vào vấn đề (tối đa 3-4 đoạn)
3. Nếu có thể, đưa ra ví dụ thực tế hoặc bước hành động cụ thể
4. Nếu câu hỏi liên quan đến chính sách nội bộ sâu, quyết định của SOD, hoặc thông tin tài chính nhạy cảm: 
   Trả lời: "Câu hỏi này cần SOD trả lời trực tiếp. Bạn có thể nhấn 'Nhờ SOD trả lời' để được hỗ trợ."
5. Dùng emoji phù hợp để làm nội dung sinh động hơn
6. Luôn khuyến khích CM và tạo động lực"""

    payload = {
        "model": "sonar-reasoning-pro",
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": f"Câu hỏi từ CM (ID: {question_id}):\n\n{question}"}
        ],
        "max_tokens": 600,
        "temperature": 0.7
    }

    try:
        data = json.dumps(payload).encode('utf-8')
        req = urllib.request.Request(
            "https://api.perplexity.ai/chat/completions",
            data=data,
            headers={
                "Authorization": f"Bearer {PERPLEXITY_API_KEY}",
                "Content-Type": "application/json",
                "Accept": "application/json"
            },
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=30) as response:
            result = json.loads(response.read().decode('utf-8'))
            return result['choices'][0]['message']['content'].strip()
    except Exception as e:
        return f"MiMi tạm thời không khả dụng. Bạn có thể thử lại sau hoặc nhấn 'Nhờ SOD trả lời' để được hỗ trợ trực tiếp. (Lỗi: {str(e)[:100]})"


def save_ai_answer(question_id: str, content: str):
    """Lưu câu trả lời AI vào Google Sheets qua Apps Script"""
    try:
        post_body = json.dumps({
            "type": "ai_answer",
            "parent_id": question_id,
            "bu": "",
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
            result = json.loads(response.read().decode('utf-8'))
            return result
    except Exception as e:
        return {"success": False, "error": str(e)}


def main():
    print("Content-Type: application/json")
    print("Access-Control-Allow-Origin: *")
    print("Access-Control-Allow-Methods: POST, OPTIONS")
    print("Access-Control-Allow-Headers: Content-Type")
    print()

    method = os.environ.get('REQUEST_METHOD', 'GET')
    if method == 'OPTIONS':
        print(json.dumps({"ok": True}))
        return

    if method != 'POST':
        print(json.dumps({"error": "Method not allowed — dùng POST"}))
        return

    try:
        content_length = int(os.environ.get('CONTENT_LENGTH', 0))
        body = sys.stdin.read(content_length) if content_length > 0 else sys.stdin.read()
        data = json.loads(body)

        question_id = data.get('question_id', '')
        question_content = data.get('question_content', '')

        if not question_id or not question_content:
            print(json.dumps({"error": "Missing question_id or question_content"}))
            return

        # Gọi AI
        ai_answer = call_ai(question_content, question_id)

        # Lưu vào Sheets
        save_result = save_ai_answer(question_id, ai_answer)

        print(json.dumps({
            "success": True,
            "answer": ai_answer,
            "saved": save_result.get("success", False),
            "answer_id": save_result.get("id", "")
        }, ensure_ascii=False))

    except json.JSONDecodeError as e:
        print(json.dumps({"error": f"Invalid JSON: {str(e)}"}))
    except Exception as e:
        print(json.dumps({"error": f"Server error: {str(e)}"}))


if __name__ == '__main__':
    main()
