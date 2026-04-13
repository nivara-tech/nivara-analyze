"""
LLM-powered insights layer using Claude.
Takes analysis results and generates narrative commentary for PPT slides.
"""
import anthropic
import json


def generate_insights(analysis, api_key):
    """
    Call Claude to generate executive narrative insights from the P&L analysis.
    Returns a dict with slide-specific commentary.
    """
    client = anthropic.Anthropic(api_key=api_key)

    # Build a compact data summary for the prompt
    a = analysis
    ns = a.current_month.get("Total Net Sales", 0)
    ns_aop = a.aop_month.get("Total Net Sales", 0)
    ns_ly = a.ly_same_month.get("Total Net Sales", 0)
    ns_pm = a.prev_month.get("Total Net Sales", 0)

    ebitda = a.current_month.get("EBITDA (post allocation)", 0)
    ebitda_aop = a.aop_month.get("EBITDA (post allocation)", 0)
    ebitda_pct = a.current_month.get("EBITDA %", 0)

    gm_pct = a.current_month.get("Gross Margin %", 0)
    pat = a.current_month.get("Profit After Tax", 0)
    pat_pct = a.current_month.get("PAT %", 0)

    pbt = a.current_month.get("Profit Before Tax", 0)

    # Budget achievement summary
    ba_summary = {}
    for label, data in a.budget_achievement.items():
        ba_summary[label] = {
            "actual": round(data["actual"], 1),
            "budget": round(data["budget"], 1),
            "achievement_pct": round(data["achievement_pct"], 1),
            "variance_pct": round(data["variance_pct"], 1),
        }

    # Top outliers
    top_outliers = []
    for o in a.outliers[:25]:
        top_outliers.append({
            "item": o.line_item,
            "type": o.comparison_type,
            "deviation_pct": round(o.deviation_pct, 1),
            "direction": o.direction,
            "severity": o.severity,
            "actual": round(o.current_value, 1),
            "reference": round(o.reference_value, 1),
        })

    # Key cost items
    cost_summary = {}
    cost_labels = [
        "Total COGS", "Total Employee Related Costs", "Total Adv & Sales Promo",
        "Total Service Charges", "Freight", "Total Other Direct Costs",
        "Total Indirect Costs (A+B)", "Common + Mfg Overheads"
    ]
    for label in cost_labels:
        actual = a.current_month.get(label, 0)
        aop = a.aop_month.get(label, 0)
        ly = a.ly_same_month.get(label, 0)
        cost_summary[label] = {
            "actual": round(actual, 1),
            "aop": round(aop, 1),
            "last_year": round(ly, 1),
            "pct_ns": round((actual / ns * 100) if ns else 0, 1),
        }

    # YTD summary
    ytd_ns = a.ytd_actual.get("Total Net Sales", 0)
    ytd_ebitda = a.ytd_actual.get("EBITDA (post allocation)", 0)
    ytd_pat = a.ytd_actual.get("Profit After Tax", 0)

    data_payload = json.dumps({
        "company": a.company,
        "period": f"{a.review_month} {a.review_fy}",
        "is_full_year": a.review_month == "Mar",
        "current_month": {
            "net_sales": round(ns, 1),
            "net_sales_aop": round(ns_aop, 1),
            "net_sales_ly": round(ns_ly, 1),
            "net_sales_prev_month": round(ns_pm, 1),
            "gross_margin_pct": round(gm_pct, 1),
            "ebitda": round(ebitda, 1),
            "ebitda_aop": round(ebitda_aop, 1),
            "ebitda_pct": round(ebitda_pct, 1),
            "pbt": round(pbt, 1),
            "pat": round(pat, 1),
            "pat_pct": round(pat_pct, 1),
        },
        "budget_achievement": ba_summary,
        "top_outliers": top_outliers,
        "cost_summary": cost_summary,
        "ytd": {
            "net_sales": round(ytd_ns, 1),
            "ebitda": round(ytd_ebitda, 1),
            "pat": round(ytd_pat, 1),
        },
        "highlights": a.highlights,
    }, indent=2)

    prompt = f"""You are a CFO-level financial analyst at Eureka Forbes Limited (water purifiers, vacuum cleaners, air purifiers).
You are writing commentary for an internal monthly P&L review presentation.

Here is the analysis data for {a.review_month} {a.review_fy}:

{data_payload}

Generate the following sections as JSON with these exact keys:

1. "executive_summary" — 4-5 bullet points for the CEO/board. Cover: revenue performance, margin trend, key wins, key concerns, outlook. Be specific with numbers.

2. "budget_commentary" — 3-4 sentences on AOP achievement. What beat, what missed, and why it matters.

3. "cost_commentary" — 3-4 sentences analyzing cost trends. Flag any cost line that grew disproportionately vs revenue. Reference freight, employee costs, ad spend specifically.

4. "outlier_commentary" — 4-5 bullet points explaining the most significant outliers. Group by theme (revenue, costs, margins). Suggest possible business reasons (seasonality, campaigns, inflation, etc.).

5. "risks_and_actions" — 3-4 bullet points on risks and recommended actions based on the data.

6. "quarter_commentary" — 2-3 sentences comparing quarterly performance.

7. "full_year_commentary" — (Only if is_full_year=true) 3-4 sentences on full year performance vs AOP and previous year.

Rules:
- All values are in Rs Crores unless stated otherwise
- Be concise, specific, and use actual numbers from the data
- Write in professional business English, not academic
- Do NOT use generic filler — every sentence must reference a specific number or trend
- Return ONLY valid JSON, no markdown wrapping"""

    message = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=2000,
        messages=[{"role": "user", "content": prompt}],
    )

    response_text = message.content[0].text.strip()

    # Parse JSON from response
    if response_text.startswith("```"):
        response_text = response_text.split("```")[1]
        if response_text.startswith("json"):
            response_text = response_text[4:]

    try:
        insights = json.loads(response_text)
    except json.JSONDecodeError:
        # Try to extract JSON from the response
        start = response_text.find("{")
        end = response_text.rfind("}") + 1
        if start >= 0 and end > start:
            insights = json.loads(response_text[start:end])
        else:
            insights = {
                "executive_summary": ["Analysis data available — LLM parsing failed. Review raw highlights."],
                "budget_commentary": "See budget achievement table for details.",
                "cost_commentary": "See cost deep dive slide for details.",
                "outlier_commentary": ["Review outlier slides for full details."],
                "risks_and_actions": ["Review highlighted outliers and take action on high-severity items."],
                "quarter_commentary": "See quarterly comparison slide.",
                "full_year_commentary": "See full year results slide.",
            }

    return insights
