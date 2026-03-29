#!/usr/bin/env python3
"""
MindX CM Report Platform — AI Analysis Engine (v2)
Model: sonar-pro (Perplexity API) | Fallback: sonar
Features:
  - Truy xuất dữ liệu lịch sử WR tuần trước từ Google Sheets
  - Benchmark chuẩn từ OKRs 2026 (CR16, CR46, AOV, pace)
  - Học từ best-practice AP của CM khác (BU Xanh)
  - Action Plan chiến lược từ SOD/FM (OKRs 2026)
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

# ===================== BENCHMARKS (OKRs 2026) =====================
BENCHMARKS = {
    "CR16_TARGET": "15%",
    "CR46_TARGET": "50%",
    "AOV_TARGET": "18M",
    "RETENTION_TARGET": ">60%",
    "UPSELL_TARGET": "25%",
    "REFERRAL_TARGET": "8%",
    "BU_MONTHLY_TARGET": "1.2B",
    "N1_ANNUAL": "154B (32.1%)",
    "N2_ANNUAL": "174B (36.3%)",
    "N3_ANNUAL": "100B (20.8%)",
    "TOTAL_ANNUAL": "480B",
}

WEEK_PACE = {1: 0.15, 2: 0.40, 3: 0.70, 4: 1.00}

# ===================== OKRs STRATEGIC ACTIONS =====================
OKRS_STRATEGIES = """
=== CHIẾN LƯỢC OKRs 2026 TỪ SOD/FM ===

[OPTIMIZE - N1] Target 154B | Tối ưu Leads MKT:
1. Redesign Trial Experience 60 phút (15' warm-up + 30' trải nghiệm + 15' tư vấn). FM Optimize audit 20% trials/tuần.
2. Upsell & Product Bundling: Tư vấn gói dài hạn (12T), combo Robotics+Coding. Giảm gói 4T xuống 12%. AOV target 18M.
3. Lead Scoring & Allocation: Phân loại Hot/Warm/Cold. Sale CR cao nhận leads chất lượng. Hot Leads liên hệ <2h.
4. Dashboard N1 real-time + Content Nurturing: Nurturing 7 ngày (email+Zalo). Open rate >40%.

[OPS - N2] Target 174B | Re-enroll/Upsell/Referral:
1. CS Playbook theo Lifecycle: Week 1 welcome → Month 1 check-in → Month 3 milestone → Pre-expire re-enroll offer.
2. Early Bird Re-enroll: CS liên hệ 60 ngày trước hết hạn. Ưu đãi 5-10% + tư vấn path (Robotics→Coding, Basic→Advanced).
3. Referral có hệ thống: Tracking code, incentive rõ ràng (Parent refer → discount + free trial). Target 8%.
4. CS Performance Dashboard: Track Retention, Upsell, Referral, NPS per CS. FM Ops review hàng tuần.

[GROWTH - N3] Target 100B | Sale tự kiếm:
1. Direct Sales + Territory Mapping: Mỗi Sale có bản đồ 3-5km, list trường/chung cư. 3 field trips/tuần, 10 contacts/trip.
2. Local Event 2 events/BU/tháng: Workshop STEM, demo coding tại trường/hội chợ. Event Toolkit chuẩn.
3. Partnership B2B: MOU với trường (ngoại khóa STEM), trung tâm Anh ngữ, khu dân cư.
4. Sale Outbound KPI: 15 contacts/tuần, 5 trial books, 2 closes/tháng. Training 3 modules (Prospecting, Pitching, Networking).
5. Referral + CTV Network: 3 referrals/Sale/tháng. Pilot 50 CTV (giáo viên/phụ huynh). Commission 500K-1M/deal.
"""


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
        print(json.dumps({"error": "Method not allowed"}))
        return

    try:
        content_length = int(os.environ.get('CONTENT_LENGTH', 0))
        body = sys.stdin.read(content_length) if content_length > 0 else sys.stdin.read()
        data = json.loads(body)

        bu_name = data.get('bu', '')
        week = data.get('week', 1)
        month = data.get('month', '')
        ap_rows = data.get('ap_rows', [])
        wr_rows = data.get('wr_rows', [])
        health_status = data.get('health_status', '')
        # Dữ liệu bổ sung từ frontend
        prev_wr_rows = data.get('prev_wr_rows', [])
        best_aps = data.get('best_aps', [])

        # Nếu frontend không gửi dữ liệu lịch sử, tự truy xuất từ Google Sheets
        if not prev_wr_rows and month:
            prev_wr_rows = fetch_previous_wr(bu_name, month, week)

        if not best_aps and month:
            best_aps = fetch_best_aps(month, week, bu_name)

        # Build prompt
        prompt = build_analysis_prompt(
            bu_name, week, month, ap_rows, wr_rows, health_status,
            prev_wr_rows, best_aps
        )

        # Call Perplexity API với sonar-pro (retry + fallback)
        result = call_perplexity(prompt)

        print(json.dumps({"success": True, "analysis": result}, ensure_ascii=False))
    except Exception as e:
        print(json.dumps({"error": str(e), "success": False}, ensure_ascii=False))


# ===================== GOOGLE SHEETS DATA FETCH =====================
def fetch_from_apps_script(action, params):
    """Gọi Google Apps Script API để lấy dữ liệu"""
    try:
        url = APPS_SCRIPT_URL + "?" + urllib.parse.urlencode({"action": action, **params})
        req = urllib.request.Request(url, method='GET')
        req.add_header('User-Agent', 'MindX-AI-Analyzer/2.0')
        with urllib.request.urlopen(req, timeout=15) as resp:
            return json.loads(resp.read().decode('utf-8'))
    except Exception:
        return None


def fetch_previous_wr(bu_name, month, week):
    """Lấy WR tuần trước của cùng BU để so sánh xu hướng"""
    prev_month, prev_week = get_prev_period(month, week)
    data = fetch_from_apps_script('get', {
        'tab': 'WR', 'bu': bu_name,
        'month': prev_month, 'week': str(prev_week)
    })
    if data and data.get('rows'):
        return data['rows']
    return []


def fetch_best_aps(month, week, exclude_bu):
    """Lấy AP hay từ các BU khác cùng kỳ (BU có WR tốt = BU Xanh)"""
    # Lấy all_data cho kỳ hiện tại
    data = fetch_from_apps_script('all_data', {'month': month, 'week': str(week)})
    if not data:
        return []

    # Tìm BU có health tốt (Xanh) dựa trên WR doanh số
    good_aps = []
    for bu_name, bu_data in data.items():
        if bu_name == exclude_bu:
            continue
        wr_rows = bu_data.get('WR', [])
        ap_rows = bu_data.get('AP', [])
        if not wr_rows or not ap_rows:
            continue

        # Kiểm tra xem BU này có health tốt không
        is_healthy = check_bu_health(wr_rows, week)
        if is_healthy and len(ap_rows) >= 2:
            # Chỉ lấy AP có nội dung chất lượng
            quality_aps = [
                r for r in ap_rows
                if r.get('key_action') and len(str(r.get('key_action', ''))) > 15
                and r.get('root_cause') and len(str(r.get('root_cause', ''))) > 15
            ]
            if quality_aps:
                good_aps.append({
                    'bu': bu_name,
                    'aps': quality_aps[:3]  # Max 3 actions per BU
                })

    # Giới hạn 3 BU mẫu
    return good_aps[:3]


def check_bu_health(wr_rows, week):
    """Kiểm tra nhanh health status từ WR rows"""
    pace = WEEK_PACE.get(int(week), 0.15) if week else 0.15
    revenue_rows = [
        r for r in wr_rows
        if 'Doanh số' in str(r.get('kpi', '')) or 'TỔNG' in str(r.get('kpi', '')).upper()
    ]
    if not revenue_rows:
        return False

    all_meet = True
    for row in revenue_rows:
        try:
            target = float(row.get('target', 0))
            actual = float(row.get('actual', 0))
            if target > 0 and actual / target < pace:
                all_meet = False
                break
        except (ValueError, TypeError, ZeroDivisionError):
            continue

    return all_meet


def get_prev_period(month, week):
    """Tính tháng/tuần trước"""
    w = int(week) if week else 1
    if w > 1:
        return month, w - 1
    else:
        # W1 → W4 tháng trước
        try:
            y, m = month.split('-')
            y, m = int(y), int(m)
            if m == 1:
                return f"{y-1}-12", 4
            return f"{y}-{str(m-1).zfill(2)}", 4
        except Exception:
            return month, 4


# ===================== PROMPT BUILDER =====================
def build_analysis_prompt(bu_name, week, month, ap_rows, wr_rows, health_status,
                         prev_wr_rows, best_aps):
    pace = WEEK_PACE.get(int(week), 0.15) if week else 0.15
    pace_pct = int(pace * 100)

    # === WR hiện tại ===
    wr_summary = ""
    for row in wr_rows:
        kpi = row.get('kpi', '')
        target = row.get('target', '')
        actual = row.get('actual', '')
        notes = row.get('notes', '')
        if target and actual:
            try:
                t, a = float(target), float(actual)
                pct = f"{a/t*100:.0f}%" if t > 0 else "N/A"
                gap = f"{a-t:+.0f}"
            except (ValueError, TypeError):
                pct, gap = "N/A", ""
            wr_summary += f"- {kpi}: Target={target}, Actual={actual}, %Đạt={pct}, Gap={gap}"
            if notes:
                wr_summary += f" | Ghi chú: {notes}"
            wr_summary += "\n"

    # === WR tuần trước (so sánh xu hướng) ===
    prev_wr_summary = ""
    if prev_wr_rows:
        prev_month, prev_week = get_prev_period(month, week)
        prev_wr_summary = f"\n--- WR TUẦN TRƯỚC ({prev_month} W{prev_week}) ---\n"
        for row in prev_wr_rows:
            kpi = row.get('kpi', '')
            target = row.get('target', '')
            actual = row.get('actual', '')
            if target and actual:
                try:
                    t, a = float(target), float(actual)
                    pct = f"{a/t*100:.0f}%" if t > 0 else "N/A"
                except (ValueError, TypeError):
                    pct = "N/A"
                prev_wr_summary += f"- {kpi}: Target={target}, Actual={actual}, %Đạt={pct}\n"

    # === AP hiện tại ===
    ap_summary = ""
    for i, row in enumerate(ap_rows):
        func = row.get('func', '')
        chi_so = row.get('chi_so', '')
        van_de = row.get('van_de', '')
        root_cause = row.get('root_cause', '')
        key_action = row.get('key_action', '')
        mo_ta = row.get('mo_ta_trien_khai', '')
        target_do_luong = row.get('target_do_luong', '')
        deadline = row.get('deadline', '')
        owner = row.get('owner', '')
        status = row.get('status', '')
        if chi_so or van_de or key_action:
            ap_summary += f"\nAction {i+1} [{func}]:\n"
            ap_summary += f"  Chỉ số: {chi_so}\n"
            ap_summary += f"  Vấn đề: {van_de}\n"
            ap_summary += f"  Root Cause: {root_cause}\n"
            ap_summary += f"  Key Action: {key_action}\n"
            ap_summary += f"  Triển khai: {mo_ta}\n"
            ap_summary += f"  Target đo lường: {target_do_luong}\n"
            ap_summary += f"  Deadline: {deadline} | Owner: {owner} | Status: {status}\n"

    # === Best-practice AP từ BU Xanh ===
    best_ap_summary = ""
    if best_aps:
        best_ap_summary = "\n=== ACTION PLAN MẪU TỪ CÁC BU XANH (ĐÚNG TIẾN ĐỘ) ===\n"
        for bp in best_aps:
            best_ap_summary += f"\n[BU: {bp['bu']}]\n"
            for j, ap in enumerate(bp['aps']):
                best_ap_summary += f"  AP{j+1} [{ap.get('func','')}]: {ap.get('chi_so','')} — {ap.get('key_action','')}\n"
                if ap.get('root_cause'):
                    best_ap_summary += f"    Root Cause: {ap.get('root_cause','')}\n"
                if ap.get('target_do_luong'):
                    best_ap_summary += f"    Target: {ap.get('target_do_luong','')}\n"

    # === Build full prompt ===
    prompt = f"""Bạn là chuyên gia phân tích kinh doanh cấp cao (Senior Business Analyst) tại MindX Technology School — trường dạy lập trình & công nghệ cho trẻ em (K12) và người lớn (18+).

BU "{bu_name}" đang ở trạng thái [{health_status}] trong {month} Tuần {week}.
Pace kỳ vọng tuần {week}: {pace_pct}% target tháng.

=== BENCHMARK CHUẨN MINDX 2026 ===
- CR16 (Lead → Close): Mục tiêu {BENCHMARKS['CR16_TARGET']}
- CR46 (Trial → Close): Mục tiêu {BENCHMARKS['CR46_TARGET']}
- AOV (Giá trị đơn hàng TB): Mục tiêu {BENCHMARKS['AOV_TARGET']}
- Retention (tái đăng ký): {BENCHMARKS['RETENTION_TARGET']}
- Upsell Rate: {BENCHMARKS['UPSELL_TARGET']}
- Referral Rate: {BENCHMARKS['REFERRAL_TARGET']}
- Doanh số BU/tháng: {BENCHMARKS['BU_MONTHLY_TARGET']}
- Revenue hệ thống: N1={BENCHMARKS['N1_ANNUAL']}, N2={BENCHMARKS['N2_ANNUAL']}, N3={BENCHMARKS['N3_ANNUAL']}
- 150 Sales (phụ trách N1+N3), 90 CS (phụ trách N2)

=== DỮ LIỆU WEEKLY REPORT TUẦN NÀY ===
{wr_summary if wr_summary else "(Chưa có dữ liệu WR)"}
{prev_wr_summary if prev_wr_summary else ""}
=== ACTION PLAN CỦA CM ===
{ap_summary if ap_summary else "(Chưa có Action Plan)"}
{best_ap_summary if best_ap_summary else ""}
{OKRS_STRATEGIES}

=== YÊU CẦU PHÂN TÍCH ===

1. SO SÁNH XU HƯỚNG: Nếu có WR tuần trước, so sánh xu hướng (cải thiện/xấu đi) cho từng KPI doanh số chính.

2. ĐÁNH GIÁ TỪNG ACTION trong Action Plan:
   - So sánh với benchmark chuẩn (CR16={BENCHMARKS['CR16_TARGET']}, CR46={BENCHMARKS['CR46_TARGET']}, AOV={BENCHMARKS['AOV_TARGET']})
   - Verdict: "KHA_THI" hoặc "CAN_DIEU_CHINH" hoặc "KHONG_KHA_THI"
   - Lý do (2-3 câu, cụ thể với số liệu)
   - Nếu cần điều chỉnh: gợi ý cụ thể, tham khảo chiến lược OKRs 2026 và best-practice BU Xanh

3. KHUYẾN NGHỊ CHO FM/SOD:
   - 2-3 điểm hành động cụ thể
   - Nếu có AP mẫu từ BU Xanh, tham chiếu làm gợi ý
   - Ưu tiên action có thể thực hiện ngay trong tuần tới

QUAN TRỌNG:
- Trả lời BẰNG TIẾNG VIỆT
- Phân tích phải DỰA TRÊN DỮ LIỆU, không nói chung chung
- Nếu số liệu cho thấy vấn đề cụ thể (VD: CR thấp hơn benchmark, doanh số dưới pace), phải chỉ rõ
- KHÔNG search internet — chỉ dùng dữ liệu được cung cấp ở trên

Format JSON:
{{
  "trend": "IMPROVING | DECLINING | STABLE | NO_DATA",
  "trend_note": "Tóm tắt xu hướng so với tuần trước (1 câu)",
  "actions": [
    {{
      "index": 1,
      "verdict": "KHA_THI" | "CAN_DIEU_CHINH" | "KHONG_KHA_THI",
      "reason": "Phân tích cụ thể dựa trên data (2-3 câu)...",
      "suggestion": "Gợi ý cải thiện, tham chiếu OKRs/BU Xanh..." (chỉ khi verdict != KHA_THI)
    }}
  ],
  "summary": "Khuyến nghị cho FM/SOD (2-3 bullet points, mỗi bullet bắt đầu bằng •)"
}}

CHỈ trả về JSON, không thêm text nào khác."""

    return prompt


# ===================== PERPLEXITY API =====================
import time

def call_perplexity(prompt, max_retries=2):
    """Primary: sonar-pro with retry + exponential backoff. Fallback: sonar."""
    api_key = PERPLEXITY_API_KEY
    if not api_key:
        return generate_mock_response()

    url = "https://api.perplexity.ai/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    payload = json.dumps({
        "model": "sonar-reasoning-pro",
        "messages": [
            {
                "role": "system",
                "content": (
                    "Bạn là Senior Business Analyst tại MindX Technology School. "
                    "Phân tích CHÍNH XÁC dựa trên dữ liệu được cung cấp, không dùng thông tin ngoài. "
                    "Trả lời bằng tiếng Việt. CHỈ trả về JSON thuần túy, không markdown, không text thừa."
                )
            },
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.2,
        "max_tokens": 3000
    })

    last_error = None
    for attempt in range(max_retries + 1):
        if attempt > 0:
            time.sleep(2 ** attempt)  # Backoff: 2s, 4s
        try:
            req = urllib.request.Request(url, data=payload.encode('utf-8'), headers=headers, method='POST')
            with urllib.request.urlopen(req, timeout=50) as resp:
                result = json.loads(resp.read().decode('utf-8'))
                content = result['choices'][0]['message']['content']
                return extract_json(content)
        except urllib.error.HTTPError as e:
            last_error = e
            # 429 rate limit or 5xx → retry
            if e.code in (429, 500, 502, 503) and attempt < max_retries:
                continue
            break
        except Exception as e:
            last_error = e
            if attempt < max_retries:
                continue
            break

    # All retries failed → fallback to sonar
    return call_perplexity_fallback(prompt)


def call_perplexity_fallback(prompt, max_retries=1):
    """Fallback: sonar (lighter model, faster) with retry"""
    api_key = PERPLEXITY_API_KEY
    if not api_key:
        return generate_mock_response()

    url = "https://api.perplexity.ai/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    payload = json.dumps({
        "model": "sonar-reasoning-pro",
        "messages": [
            {
                "role": "system",
                "content": (
                    "Bạn là Senior Business Analyst tại MindX Technology School. "
                    "Phân tích CHÍNH XÁC dựa trên dữ liệu được cung cấp. "
                    "Trả lời bằng tiếng Việt. CHỈ trả về JSON thuần túy."
                )
            },
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.2,
        "max_tokens": 3000
    })

    for attempt in range(max_retries + 1):
        if attempt > 0:
            time.sleep(2)
        try:
            req = urllib.request.Request(url, data=payload.encode('utf-8'), headers=headers, method='POST')
            with urllib.request.urlopen(req, timeout=50) as resp:
                result = json.loads(resp.read().decode('utf-8'))
                content = result['choices'][0]['message']['content']
                return extract_json(content)
        except Exception:
            if attempt < max_retries:
                continue

    return generate_mock_response()


# ===================== JSON EXTRACTION =====================
def extract_json(text):
    """Robust JSON extraction — xử lý cả <think> tags từ reasoning model"""
    text = text.strip()

    # Loại bỏ <think>...</think> tags (reasoning tokens)
    text = re.sub(r'<think>[\s\S]*?</think>', '', text).strip()

    # Remove markdown code fences
    if text.startswith('```'):
        text = re.sub(r'^```(?:json)?\s*\n?', '', text)
        text = re.sub(r'\n?```\s*$', '', text)
        text = text.strip()

    # Try direct JSON parse
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Try to find JSON object in text
    match = re.search(r'\{[\s\S]*\}', text)
    if match:
        try:
            return json.loads(match.group())
        except json.JSONDecodeError:
            pass

    # Fallback
    return generate_mock_response()


def generate_mock_response():
    """Fallback response khi API không khả dụng"""
    return {
        "trend": "NO_DATA",
        "trend_note": "Không có đủ dữ liệu để so sánh xu hướng.",
        "actions": [
            {
                "index": 1,
                "verdict": "CAN_DIEU_CHINH",
                "reason": "Không thể phân tích chi tiết do hệ thống AI tạm thời không khả dụng. Action cần được review thủ công bởi FM.",
                "suggestion": "FM nên review trực tiếp Action Plan này với CM, đối chiếu với benchmark CR16=15%, CR46=50%, AOV=18M."
            }
        ],
        "summary": "• Hệ thống AI tạm thời không khả dụng — FM/SOD nên review thủ công.\n• Đối chiếu AP với benchmark: CR16 mục tiêu 15%, CR46 mục tiêu 50%, AOV mục tiêu 18M.\n• Ưu tiên hỗ trợ CM xác định root cause chính xác và key action cụ thể."
    }


if __name__ == '__main__':
    main()
