# ============================================================
#   JOB FINDER — Easy Apply Job Scraper
#   Built by Nivas Verelli
#   Powered by: JSearch API (free) + Groq AI (free)
# ============================================================
#
#   SETUP (run these in your terminal first):
#   pip install groq openpyxl python-dotenv
#
#   ADD your keys directly in the CONFIG section below
#
#   THEN RUN:
#   python job_finder.py
# ============================================================

import os
import json
import re
import webbrowser
import urllib.request
import urllib.parse
from datetime import datetime
from groq import Groq


# ============================================================
#   CONFIG — Put your API keys and profile here
# ============================================================



MY_SKILLS = """
Name: Nivas Verelli
Degree: MS Business Analytics, UT Dallas (graduating December 2026)
Undergrad: BS Computer Science

Skills:
- SQL (window functions, CTEs, pivot analysis)
- Python (pandas, ETL pipelines, automation)
- Power BI, Excel dashboards
- Azure, Databricks, PySpark
- Data modeling, medallion architecture
- Basic machine learning (classification, clustering)

Work Experience:
- C-DAC Bangalore: Analyzed 50,000 service records, found root cause of SLA breaches
- Ronnia Langston Foundation: Built automated ETL pipelines, KPI dashboards

Projects:
- Healthcare capstone: Azure medallion pipeline + claim denial prediction + RAG layer
- Levi Strauss: Geospatial store rationalization, identified $42M in savings
- Customer segmentation: K-Means clustering + XGBoost on 3.35M records

Target Roles: Data Analyst, Product Operations Analyst, Business Analyst
Location: Dallas/Irving TX — open to hybrid or remote
"""

JOB_KEYWORDS = "data analyst"
JOB_LOCATION  = "Dallas, TX"
MAX_JOBS      = 30
MIN_SCORE     = 1


# ============================================================
#   STEP 1 — Scrape Jobs via JSearch API (free)
# ============================================================

def scrape_jobs():
    print("\n🔍 Searching for jobs...")
    print(f"   Keywords : {JOB_KEYWORDS}")
    print(f"   Location : {JOB_LOCATION}\n")

    query = f"{JOB_KEYWORDS} in {JOB_LOCATION}"

    url = "https://jsearch.p.rapidapi.com/search?" + urllib.parse.urlencode({
        "query"      : query,
        "page"       : "1",
        "num_pages"  : "3",
        "date_posted": "today",
    })

    req = urllib.request.Request(url, headers={
        "X-RapidAPI-Key" : RAPIDAPI_KEY,
        "X-RapidAPI-Host": "jsearch.p.rapidapi.com"
    })

    try:
        response = urllib.request.urlopen(req, timeout=20)
        data     = json.loads(response.read())
    except Exception as e:
        print(f"❌ JSearch error: {e}")
        print("   Check your RAPIDAPI_KEY is correct")
        return []

    jobs = []
    for item in data.get("data", []):
        jobs.append({
            "title"      : item.get("job_title", ""),
            "company"    : item.get("employer_name", ""),
            "location"   : f"{item.get('job_city','')}, {item.get('job_state','')}".strip(", "),
            "posted"     : item.get("job_posted_at_datetime_utc", "")[:10],
            "applicants" : "Unknown",
            "description": item.get("job_description", "")[:1500],
            "url"        : item.get("job_apply_link", ""),
            "salary"     : (
                f"${item.get('job_min_salary','?')}–${item.get('job_max_salary','?')}"
                if item.get("job_min_salary") else "Not listed"
            ),
            "job_type"   : item.get("job_employment_type", ""),
        })

    print(f"✅ Found {len(jobs)} jobs!\n")
    return jobs[:MAX_JOBS]


# ============================================================
#   STEP 2 — Score Each Job Using Groq AI
# ============================================================

def score_jobs(jobs):
    print("🤖 Groq AI is scoring each job...")
    print("   (Takes about 30-60 seconds)\n")

    client = Groq(api_key=GROQ_API_KEY)
    scored = []

    for i, job in enumerate(jobs, 1):
        print(f"   [{i}/{len(jobs)}] {job['title']} @ {job['company']}")

        prompt = f"""You are a career advisor. Score this job against the candidate.
Return ONLY a raw JSON object. No explanation. No markdown. No backticks.

CANDIDATE:
{MY_SKILLS}

JOB:
Title: {job['title']}
Company: {job['company']}
Location: {job['location']}
Description: {job['description']}

Return exactly this JSON:
{{"score": 7, "priority": "High", "match_reason": "one sentence", "top_skills_matched": ["skill1"], "gaps": ["gap1"], "apply": true}}

Rules:
- score is integer 1-10
- priority is exactly High, Medium, or Low
- apply is true if score >= 6, else false
"""

        try:
            response = client.chat.completions.create(
                model="llama3-8b-8192",
                messages=[{"role": "user", "content": prompt}],
                max_tokens=250,
                temperature=0.1
            )

            raw = response.choices[0].message.content.strip()

            # Strip markdown if Groq adds it
            raw = re.sub(r"```json|```", "", raw).strip()

            result = json.loads(raw)
            job.update(result)

        except Exception:
            job.update({
                "score": 5,
                "priority": "Medium",
                "match_reason": "Could not score — review manually",
                "top_skills_matched": [],
                "gaps": [],
                "apply": True
            })

        scored.append(job)

    scored.sort(key=lambda x: x.get("score", 0), reverse=True)

    apply_count = len([j for j in scored if j.get("apply")])
    print(f"\n✅ Done! {apply_count} jobs worth applying to.\n")
    return scored


# ============================================================
#   STEP 3 — Save to Excel
# ============================================================

def save_excel(jobs):
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
    except ImportError:
        print("⚠️  Skipping Excel — run: pip install openpyxl")
        return None

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Jobs"

    headers = ["#", "Score", "Priority", "Apply?", "Job Title",
               "Company", "Location", "Salary", "Why It Fits", "Link"]
    ws.append(headers)

    for cell in ws[1]:
        cell.font      = Font(bold=True, color="FFFFFF")
        cell.fill      = PatternFill("solid", fgColor="1B4F8A")
        cell.alignment = Alignment(horizontal="center")

    colors = {"High": "D4EDDA", "Medium": "FFF3CD", "Low": "F8F9FA"}

    for i, job in enumerate(jobs, 1):
        ws.append([
            i,
            job.get("score", "?"),
            job.get("priority", "?"),
            "✅ YES" if job.get("apply") else "⏭ SKIP",
            job.get("title", ""),
            job.get("company", ""),
            job.get("location", ""),
            job.get("salary", ""),
            job.get("match_reason", ""),
            job.get("url", ""),
        ])
        fill = PatternFill("solid", fgColor=colors.get(job.get("priority", "Low"), "F8F9FA"))
        for cell in ws[i + 1]:
            cell.fill = fill

    widths = [5, 8, 10, 10, 35, 25, 20, 15, 50, 60]
    for idx, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(idx)].width = w

    filename = f"jobs_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    wb.save(filename)
    print(f"✅ Excel saved: {filename}")
    return filename


# ============================================================
#   STEP 4 — Build Webpage
# ============================================================

def create_webpage(jobs):
    today      = datetime.now().strftime("%B %d, %Y")
    apply_jobs = [j for j in jobs if j.get("apply") and j.get("score", 0) >= MIN_SCORE]
    high       = len([j for j in apply_jobs if j.get("priority") == "High"])
    medium     = len([j for j in apply_jobs if j.get("priority") == "Medium"])

    def card(job, rank):
        priority = job.get("priority", "Low")
        score    = job.get("score", "?")
        matched  = job.get("top_skills_matched", [])
        gaps     = job.get("gaps", [])

        badge = {"High": ("#1e7e34", "#d4edda"),
                 "Medium": ("#856404", "#fff3cd"),
                 "Low": ("#495057", "#e2e3e5")}.get(priority, ("#495057", "#e2e3e5"))

        score_color = "#1e7e34" if score >= 8 else "#856404" if score >= 6 else "#dc3545"

        skills_html = "".join(
            f'<span style="background:#e3f2fd;color:#1565c0;padding:2px 8px;'
            f'border-radius:12px;font-size:12px;margin:2px;display:inline-block">{s}</span>'
            for s in matched)

        gaps_html = "".join(
            f'<span style="background:#fce4ec;color:#b71c1c;padding:2px 8px;'
            f'border-radius:12px;font-size:12px;margin:2px;display:inline-block">{g}</span>'
            for g in gaps) or '<span style="color:#888;font-size:12px">None</span>'

        return f"""
        <div class="card" data-priority="{priority}" id="c{rank}">
          <div style="display:flex;align-items:flex-start;gap:14px;margin-bottom:12px">
            <div style="font-size:22px;font-weight:800;color:#1B4F8A;min-width:36px">#{rank}</div>
            <div style="flex:1">
              <div style="font-size:17px;font-weight:600">{job.get('title','')}</div>
              <div style="font-size:13px;color:#666;margin-top:3px">
                🏢 {job.get('company','')} &nbsp;·&nbsp; 📍 {job.get('location','')} &nbsp;·&nbsp; 📅 {job.get('posted','')}
              </div>
            </div>
            <div style="text-align:center;min-width:70px">
              <div style="font-size:28px;font-weight:800;color:{score_color};line-height:1">{score}</div>
              <div style="font-size:11px;color:#999">/10</div>
              <span style="background:{badge[1]};color:{badge[0]};padding:3px 10px;
                border-radius:20px;font-size:11px;font-weight:600;margin-top:4px;display:inline-block">{priority}</span>
            </div>
          </div>
          <div style="background:#f8f9fa;border-radius:8px;padding:10px 14px;font-size:14px;
               color:#444;margin-bottom:10px;border-left:3px solid #0d6efd">
            💡 {job.get('match_reason','')}
          </div>
          <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;flex-wrap:wrap">
            <span style="font-size:12px;font-weight:600;color:#555;min-width:80px">✅ Matched:</span>
            {skills_html}
          </div>
          <div style="display:flex;align-items:center;gap:8px;margin-bottom:12px;flex-wrap:wrap">
            <span style="font-size:12px;font-weight:600;color:#555;min-width:80px">⚠️ Gaps:</span>
            {gaps_html}
          </div>
          <div style="display:flex;align-items:center;gap:14px;flex-wrap:wrap">
            <span style="font-size:13px;color:#666">💰 {job.get('salary','Not listed')}</span>
            <span style="font-size:13px;color:#666">💼 {job.get('job_type','')}</span>
            <a href="{job.get('url','#')}" target="_blank"
               style="margin-left:auto;background:#0077B5;color:white;padding:9px 20px;
               border-radius:8px;text-decoration:none;font-weight:600;font-size:14px">
              Apply Now →
            </a>
          </div>
        </div>"""

    cards_html = "\n".join(card(j, i+1) for i, j in enumerate(apply_jobs))

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Job Finder — {today}</title>
<style>
  * {{ box-sizing:border-box; margin:0; padding:0 }}
  body {{ font-family:-apple-system,sans-serif; background:#f0f4f8; padding:20px; color:#1a1a1a }}
  .header {{ background:linear-gradient(135deg,#1B4F8A,#0d6efd); color:white;
             padding:28px 32px; border-radius:14px; margin-bottom:24px }}
  .stats {{ display:grid; grid-template-columns:repeat(4,1fr); gap:12px; margin-bottom:24px }}
  .stat {{ background:white; border-radius:10px; padding:16px; text-align:center;
           box-shadow:0 2px 6px rgba(0,0,0,.07) }}
  .stat-n {{ font-size:32px; font-weight:700; color:#1B4F8A }}
  .stat-l {{ font-size:12px; color:#888; margin-top:4px }}
  .filters {{ display:flex; gap:10px; margin-bottom:20px; flex-wrap:wrap }}
  .btn {{ padding:7px 16px; border:1.5px solid #1B4F8A; border-radius:20px;
          background:white; color:#1B4F8A; cursor:pointer; font-size:13px; font-weight:500 }}
  .btn.active, .btn:hover {{ background:#1B4F8A; color:white }}
  .card {{ background:white; border-radius:12px; padding:20px; margin-bottom:14px;
           box-shadow:0 2px 8px rgba(0,0,0,.07); border-left:5px solid #1B4F8A }}
  .card:hover {{ transform:translateY(-2px); box-shadow:0 6px 16px rgba(0,0,0,.1);
                 transition:all .15s }}
</style>
</head>
<body>
<div class="header">
  <h1 style="font-size:26px;margin-bottom:6px">🎯 Your Job List</h1>
  <p style="opacity:.85;font-size:14px">{today} &nbsp;·&nbsp; Sorted by AI match score &nbsp;·&nbsp; Click Apply Now to apply</p>
</div>
<div class="stats">
  <div class="stat"><div class="stat-n">{len(jobs)}</div><div class="stat-l">Total found</div></div>
  <div class="stat"><div class="stat-n">{len(apply_jobs)}</div><div class="stat-l">Worth applying</div></div>
  <div class="stat"><div class="stat-n">{high}</div><div class="stat-l">High priority</div></div>
  <div class="stat"><div class="stat-n">{medium}</div><div class="stat-l">Medium priority</div></div>
</div>
<div class="filters">
  <button class="btn active" onclick="filter('all',this)">All ({len(apply_jobs)})</button>
  <button class="btn" onclick="filter('High',this)">🟢 High ({high})</button>
  <button class="btn" onclick="filter('Medium',this)">🟡 Medium ({medium})</button>
</div>
<div id="list">
  {cards_html or '<p style="text-align:center;padding:60px;color:#888">No jobs found today. Try again tomorrow!</p>'}
</div>
<script>
function filter(p, btn) {{
  document.querySelectorAll('.btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  document.querySelectorAll('.card').forEach(c => {{
    c.style.display = (p === 'all' || c.dataset.priority === p) ? '' : 'none';
  }});
}}
</script>
</body></html>"""

    filename = f"jobs_{datetime.now().strftime('%Y-%m-%d')}.html"
    with open(filename, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✅ Webpage created: {filename}")
    return filename


# ============================================================
#   MAIN
# ============================================================

def main():
    print("\n" + "=" * 55)
    print("   🚀 JOB FINDER — Starting up")
    print("=" * 55)

    if RAPIDAPI_KEY == "paste_your_rapidapi_key_here":
        print("❌ Please add your RAPIDAPI_KEY in the CONFIG section at the top of this file")
        return

    if GROQ_API_KEY == "paste_your_groq_key_here":
        print("❌ Please add your GROQ_API_KEY in the CONFIG section at the top of this file")
        return

    jobs = scrape_jobs()

    if not jobs:
        print("❌ No jobs found. Check your RAPIDAPI_KEY is correct.")
        return

    jobs    = score_jobs(jobs)
    excel   = save_excel(jobs)
    webpage = create_webpage(jobs)

    print("\n" + "=" * 55)
    print("   ✅ ALL DONE!")
    print("=" * 55)
    print(f"   📊 Excel   : {excel}")
    print(f"   🌐 Webpage : {webpage}")
    print("\n   Opening browser now...\n")

    webbrowser.open(f"file://{os.path.abspath(webpage)}")


if __name__ == "__main__":
    main()