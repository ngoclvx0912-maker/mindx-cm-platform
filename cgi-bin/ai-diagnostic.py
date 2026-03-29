#!/usr/bin/env python3
"""
MindX CM Report Platform — AI Diagnostic Coach
Model: sonar-pro (Perplexity API) | Fallback: sonar

Tính năng "Chẩn đoán BU":
  Giai đoạn 1: Phân tích WR → Chỉ ra vấn đề gốc (Problem Identification)
  Giai đoạn 2: Phân tích nguyên nhân gốc rễ (Root Cause Analysis)
  Giai đoạn 3: Gợi ý Key Actions trọng điểm (Action Suggestions)

Flow: CM nhấn "Chẩn đoán BU" → AI đọc WR → trả về 3 giai đoạn
CM đọc kết quả → điền AP có định hướng
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
        wr_rows = data.get('wr_rows', [])
        prev_wr_rows = data.get('prev_wr_rows', [])
        best_aps = data.get('best_aps', [])

        # Nếu frontend không gửi dữ liệu lịch sử, tự truy xuất
        if not prev_wr_rows and month:
            prev_wr_rows = fetch_previous_wr(bu_name, month, week)

        if not best_aps and month:
            best_aps = fetch_best_aps(month, week, bu_name)

        # Build prompt cho Diagnostic Coach
        prompt = build_diagnostic_prompt(
            bu_name, week, month, wr_rows, prev_wr_rows, best_aps
        )

        # Call Perplexity API
        result = call_perplexity(prompt)

        print(json.dumps({"success": True, "diagnostic": result}, ensure_ascii=False))
    except Exception as e:
        print(json.dumps({"error": str(e), "success": False}, ensure_ascii=False))


# ===================== GOOGLE SHEETS DATA FETCH =====================
def fetch_from_apps_script(action, params):
    """Gọi Google Apps Script API để lấy dữ liệu"""
    try:
        url = APPS_SCRIPT_URL + "?" + urllib.parse.urlencode({"action": action, **params})
        req = urllib.request.Request(url, method='GET')
        req.add_header('User-Agent', 'MindX-AI-Diagnostic/1.0')
        with urllib.request.urlopen(req, timeout=15) as resp:
            return json.loads(resp.read().decode('utf-8'))
    except Exception:
        return None


def fetch_previous_wr(bu_name, month, week):
    """Lấy WR tuần trước của cùng BU"""
    prev_month, prev_week = get_prev_period(month, week)
    data = fetch_from_apps_script('get', {
        'tab': 'WR', 'bu': bu_name,
        'month': prev_month, 'week': str(prev_week)
    })
    if data and data.get('rows'):
        return data['rows']
    return []


def fetch_best_aps(month, week, exclude_bu):
    """Lấy AP hay từ các BU Xanh cùng kỳ"""
    data = fetch_from_apps_script('all_data', {'month': month, 'week': str(week)})
    if not data:
        return []

    good_aps = []
    for bu_name, bu_data in data.items():
        if bu_name == exclude_bu:
            continue
        wr_rows = bu_data.get('WR', [])
        ap_rows = bu_data.get('AP', [])
        if not wr_rows or not ap_rows:
            continue

        is_healthy = check_bu_health(wr_rows, week)
        if is_healthy and len(ap_rows) >= 2:
            quality_aps = [
                r for r in ap_rows
                if r.get('key_action') and len(str(r.get('key_action', ''))) > 15
                and r.get('root_cause') and len(str(r.get('root_cause', ''))) > 15
            ]
            if quality_aps:
                good_aps.append({
                    'bu': bu_name,
                    'aps': quality_aps[:3]
                })

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

    for row in revenue_rows:
        try:
            target = float(row.get('target', 0))
            actual = float(row.get('actual', 0))
            if target > 0 and actual / target < pace:
                return False
        except (ValueError, TypeError, ZeroDivisionError):
            continue

    return True


def get_prev_period(month, week):
    """Tính tháng/tuần trước"""
    w = int(week) if week else 1
    if w > 1:
        return month, w - 1
    else:
        try:
            y, m = month.split('-')
            y, m = int(y), int(m)
            if m == 1:
                return f"{y-1}-12", 4
            return f"{y}-{str(m-1).zfill(2)}", 4
        except Exception:
            return month, 4


# ===================== DIAGNOSTIC PROMPT BUILDER =====================
def build_diagnostic_prompt(bu_name, week, month, wr_rows, prev_wr_rows, best_aps):
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
                vs_pace = "ĐẠT PACE" if t > 0 and a / t >= pace else "DƯỚI PACE"
            except (ValueError, TypeError):
                pct, gap, vs_pace = "N/A", "", ""
            wr_summary += f"- {kpi}: Target={target}, Actual={actual}, %Đạt={pct}, Gap={gap}, {vs_pace}"
            if notes:
                wr_summary += f" | Ghi chú: {notes}"
            wr_summary += "\n"

    # === WR tuần trước ===
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

    # === Best-practice AP từ BU Xanh ===
    best_ap_summary = ""
    if best_aps:
        best_ap_summary = "\n=== ACTION PLAN MẪU TỪ CÁC BU XANH ===\n"
        for bp in best_aps:
            best_ap_summary += f"\n[BU: {bp['bu']}]\n"
            for j, ap in enumerate(bp['aps']):
                best_ap_summary += f"  AP{j+1} [{ap.get('func','')}]: Vấn đề: {ap.get('van_de','')} — Root Cause: {ap.get('root_cause','')} — Key Action: {ap.get('key_action','')}\n"

    # === Build full prompt ===
    prompt = f"""Bạn là AI Diagnostic Coach — chuyên gia huấn luyện Center Manager (CM) tại MindX Technology School.

NHIỆM VỤ: Phân tích dữ liệu WR của BU "{bu_name}" và hướng dẫn CM nhận diện đúng vấn đề, phân tích đúng nguyên nhân gốc rễ, và chọn đúng hành động trọng điểm.

BỐI CẢNH: {month} Tuần {week}. Pace kỳ vọng: {pace_pct}% target tháng.

=== BENCHMARK CHUẨN MINDX 2026 ===
- CR16 (Lead → Close): {BENCHMARKS['CR16_TARGET']}
- CR46 (Trial → Close): {BENCHMARKS['CR46_TARGET']}
- AOV: {BENCHMARKS['AOV_TARGET']}
- Retention: {BENCHMARKS['RETENTION_TARGET']}
- Upsell Rate: {BENCHMARKS['UPSELL_TARGET']}
- Referral Rate: {BENCHMARKS['REFERRAL_TARGET']}
- BU target/tháng: {BENCHMARKS['BU_MONTHLY_TARGET']}

=== DỮ LIỆU WEEKLY REPORT TUẦN NÀY ===
{wr_summary if wr_summary else "(Chưa có dữ liệu WR)"}
{prev_wr_summary}
{best_ap_summary}
{OKRS_STRATEGIES}

=== YÊU CẦU: PHÂN TÍCH CHẨN ĐOÁN 3 GIAI ĐOẠN ===

GIAI ĐOẠN 1 — NHẬN DIỆN VẤN ĐỀ (Problem Identification):
Dựa trên dữ liệu WR, xác định TỐI ĐA 3 VẤN ĐỀ CỐT LÕI NHẤT (không phải triệu chứng bề mặt).
Mỗi vấn đề cần:
- "indicator": Chỉ số nào cho thấy vấn đề (VD: "CR46 = 35%")
- "benchmark": So sánh với benchmark ("Benchmark: 50%")
- "gap_severity": "NGHIÊM_TRỌNG" hoặc "CẦN_THEO_DÕI"
- "problem_statement": Mô tả vấn đề gốc bằng 1-2 câu rõ ràng, cụ thể với số liệu
- "func": Function liên quan (GROWTH / OPTIMIZE / OPS)
- "trend": So sánh với tuần trước nếu có dữ liệu (CẢI_THIỆN / XẤU_ĐI / ỔN_ĐỊNH / KHÔNG_CÓ_DỮ_LIỆU)

Lưu ý: Ưu tiên vấn đề có tác động lớn nhất tới doanh số tổng. Bỏ qua các chỉ số đạt benchmark.

GIAI ĐOẠN 2 — PHÂN TÍCH NGUYÊN NHÂN GỐC RỄ (Root Cause Analysis):
Cho MỖI vấn đề ở giai đoạn 1, đưa ra 2-3 NGUYÊN NHÂN GỐC RỄ TIỀM NĂNG.
Mỗi nguyên nhân cần:
- "cause": Mô tả nguyên nhân (1-2 câu)
- "evidence": Bằng chứng từ dữ liệu WR (cụ thể, có số liệu)
- "likelihood": "CAO" hoặc "TRUNG_BÌNH"

Lưu ý: Phân biệt rõ "triệu chứng" vs "nguyên nhân gốc". VD: "CR thấp" là triệu chứng, "Chất lượng Trial chưa chuẩn 60 phút" là nguyên nhân gốc.

GIAI ĐOẠN 3 — GỢI Ý KEY ACTIONS TRỌNG ĐIỂM:
Cho MỖI vấn đề, gợi ý 1-2 KEY ACTIONS cụ thể nhất.
Mỗi action cần:
- "action": Mô tả hành động (cụ thể, actionable)
- "rationale": Tại sao action này giải quyết root cause (1-2 câu)
- "target_metric": Chỉ số đo lường + mục tiêu cụ thể (VD: "CR46 từ 35% → 45% trong 2 tuần")
- "source": Nguồn gốc gợi ý ("OKRs_2026" hoặc "BU_XANH" hoặc "BEST_PRACTICE")
- "priority": "P1" hoặc "P2"

QUAN TRỌNG:
- Trả lời BẰNG TIẾNG VIỆT
- Phân tích CHÍNH XÁC dựa trên dữ liệu, không nói chung chung
- Nếu BU đã đạt pace tốt ở tất cả chỉ số, vẫn chỉ ra 1-2 điểm có thể tối ưu thêm
- KHÔNG search internet — chỉ dùng dữ liệu được cung cấp

Format JSON:
{{
  "bu_summary": "Tóm tắt tình trạng BU trong 1-2 câu (bao gồm trạng thái sức khỏe doanh số)",
  "problems": [
    {{
      "id": 1,
      "indicator": "CR46 = 35%",
      "benchmark": "Benchmark: 50%",
      "gap_severity": "NGHIÊM_TRỌNG",
      "problem_statement": "Tỷ lệ chuyển đổi từ Trial sang Close rất thấp...",
      "func": "OPTIMIZE",
      "trend": "XẤU_ĐI",
      "root_causes": [
        {{
          "cause": "Chất lượng Trial chưa đạt chuẩn 60 phút...",
          "evidence": "Số Trial = 40 nhưng Close chỉ 14...",
          "likelihood": "CAO"
        }}
      ],
      "suggested_actions": [
        {{
          "action": "Redesign Trial Experience theo chuẩn 60 phút...",
          "rationale": "Trial chuẩn giúp PH trải nghiệm đầy đủ...",
          "target_metric": "CR46 từ 35% → 45% trong 2 tuần",
          "source": "OKRs_2026",
          "priority": "P1"
        }}
      ]
    }}
  ],
  "overall_priority": "Tóm tắt ưu tiên hành động cho CM (1-2 câu)"
}}

CHỈ trả về JSON, không thêm text nào khác."""

    return prompt


# ===================== PERPLEXITY API =====================
import time

def call_perplexity(prompt, max_retries=2):
    """Primary: sonar-pro with retry + exponential backoff. Fallback: sonar."""
    api_key = PERPLEXITY_API_KEY
    if not api_key:
        return generate_fallback()

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
                    "Bạn là AI Diagnostic Coach tại MindX Technology School. "
                    "Nhiệm vụ: hướng dẫn Center Manager nhận diện đúng vấn đề, phân tích nguyên nhân gốc rễ, và chọn hành động trọng điểm. "
                    "Phân tích DỰA TRÊN DỮ LIỆU, không nói chung chung. "
                    "Trả lời bằng tiếng Việt. CHỈ trả về JSON thuần túy, không markdown, không text thừa."
                )
            },
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.2,
        "max_tokens": 4000
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
        return generate_fallback()

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
                    "Bạn là AI Diagnostic Coach tại MindX Technology School. "
                    "Phân tích DỰA TRÊN DỮ LIỆU. "
                    "Trả lời bằng tiếng Việt. CHỈ trả về JSON thuần túy."
                )
            },
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.2,
        "max_tokens": 4000
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

    return generate_fallback()


# ===================== JSON EXTRACTION =====================
def extract_json(text):
    """Robust JSON extraction"""
    text = text.strip()
    text = re.sub(r'<think>[\s\S]*?</think>', '', text).strip()

    if text.startswith('```'):
        text = re.sub(r'^```(?:json)?\s*\n?', '', text)
        text = re.sub(r'\n?```\s*$', '', text)
        text = text.strip()

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    match = re.search(r'\{[\s\S]*\}', text)
    if match:
        try:
            return json.loads(match.group())
        except json.JSONDecodeError:
            pass

    return generate_fallback()


def generate_fallback():
    """Fallback response khi API không khả dụng"""
    return {
        "bu_summary": "Không thể phân tích tự động. FM nên review trực tiếp với CM.",
        "problems": [
            {
                "id": 1,
                "indicator": "Cần review thủ công",
                "benchmark": "—",
                "gap_severity": "CẦN_THEO_DÕI",
                "problem_statement": "Hệ thống AI tạm thời không khả dụng. FM nên review WR cùng CM, đối chiếu với benchmark CR16=15%, CR46=50%, AOV=18M.",
                "func": "OPTIMIZE",
                "trend": "KHÔNG_CÓ_DỮ_LIỆU",
                "root_causes": [
                    {
                        "cause": "Cần phân tích thủ công bởi FM",
                        "evidence": "Hệ thống AI đang bảo trì",
                        "likelihood": "TRUNG_BÌNH"
                    }
                ],
                "suggested_actions": [
                    {
                        "action": "FM review WR trực tiếp với CM, so sánh từng chỉ số với benchmark",
                        "rationale": "Đảm bảo CM nhận diện đúng vấn đề dù AI không khả dụng",
                        "target_metric": "Hoàn thành review trong buổi coaching tuần này",
                        "source": "BEST_PRACTICE",
                        "priority": "P1"
                    }
                ]
            }
        ],
        "overall_priority": "FM hỗ trợ CM review WR thủ công, đối chiếu benchmark CR16=15%, CR46=50%, AOV=18M."
    }


if __name__ == '__main__':
    main()
