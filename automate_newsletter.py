import os
import re
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from typing import List, Dict, Tuple, Optional

import pandas as pd
from mailchimp_marketing import Client

from get_latest_campaign import get_latest_campaign
from extract_excel import get_first_30_rows_from_excel

from mailchimp_marketing.api_client import ApiClientError


# Configure Mailchimp client (reuse config style from existing scripts)
MAILCHIMP_API_KEY = os.getenv("MAILCHIMP_API_KEY", "aa90f697f32e710c03320a2758f209f5-us6")
MAILCHIMP_SERVER_PREFIX = os.getenv("MAILCHIMP_SERVER_PREFIX", "us6")

mailchimp = Client()
mailchimp.set_config({
    "api_key": MAILCHIMP_API_KEY,
    "server": MAILCHIMP_SERVER_PREFIX,
})

# Static divider HTML used between event blocks (copied from template)
DIVIDER_HTML = (
    '<table border="0" cellpadding="0" cellspacing="0" width="100%" class="mcnDividerBlock" '
    'style="min-width: 100%;border-collapse: collapse;mso-table-lspace: 0pt;mso-table-rspace: 0pt;'
    '-ms-text-size-adjust: 100%;-webkit-text-size-adjust: 100%;table-layout: fixed !important;">'
    '<tbody class="mcnDividerBlockOuter"><tr>'
    '<td class="mcnDividerBlockInner" style="min-width: 100%;padding: 0px 18px;'
    'mso-line-height-rule: exactly;-ms-text-size-adjust: 100%;-webkit-text-size-adjust: 100%;">'
    '<table class="mcnDividerContent" border="0" cellpadding="0" cellspacing="0" width="100%" '
    'style="min-width: 100%;border-top: 2px dotted #990000;border-collapse: collapse;'
    'mso-table-lspace: 0pt;mso-table-rspace: 0pt;-ms-text-size-adjust: 100%;-webkit-text-size-adjust: 100%;">'
    '<tbody><tr><td style="mso-line-height-rule: exactly;-ms-text-size-adjust: 100%;-webkit-text-size-adjust: 100%;">'
    '<span></span></td></tr></tbody></table>'
    '</td></tr></tbody></table>'
)


DATE_RE = re.compile(r"(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2}(st|nd|rd|th),\s+\d{4}")

def _pick_header_section_key(sections: Dict[str, str]) -> Optional[str]:
    for k, v in sections.items():
        if v and DATE_RE.search(v):
            return k
    for k, v in sections.items():
        if v and ("font-size:24px" in v or "#011F5B" in v):
            return k
    return None


def _pick_event_section_keys(sections: Dict[str, str]) -> list[str]:
    keys = []
    for k, v in sections.items():
        if not v:
            continue
        if "Mantra Health" in v:
            continue
        # Heuristic: any section that contains legacy block classes from the events area
        if ("mcnCaption" in v) or ("mcnDividerBlock" in v):
            keys.append(k)
    return keys


def _set_content_sections(mc: Client, campaign_id: str, header_html: str, events_html: str) -> bool:
    c = mc.campaigns.get_content(campaign_id)
    tmpl = c.get("template") or {}
    sections = dict(tmpl.get("sections") or {})
    if not sections:
        mc.campaigns.set_content(campaign_id, {"html": c.get("html", "")})
        return False

    hk = _pick_header_section_key(sections)
    bks = _pick_event_section_keys(sections)

    touched = False

    if hk:
        cur = sections.get(hk, "")
        if DATE_RE.search(cur):
            sections[hk] = DATE_RE.sub(header_html, cur)
        else:
            sections[hk] = header_html
        touched = True

    # Replace ALL event-ish sections with our compiled events_html
    # (Your template usually has one main body section; if there are multiple,
    # this guarantees we don't leave an old block hanging around.)
    for bk in bks:
        sections[bk] = events_html
        touched = True

    if not touched:
        raise RuntimeError("Template has sections, but none matched header or events. Aborting to avoid revert.")

    mc.campaigns.set_content(campaign_id, {"template": {"id": tmpl.get("id"), "sections": sections}})

    # Verify on the server
    v = mc.campaigns.get_content(campaign_id)
    vtmpl = v.get("template") or {}
    vsections = vtmpl.get("sections") or {}
    text = "".join(vsections.values())
    with open(os.path.join("artifacts", f"sections_after_{campaign_id}.html"), "w", encoding="utf-8") as f:
        f.write(text or "")

    ok_header = (header_html in text)
    ok_oldblock = not ("Stay Healthy & Connected This Summer" in text and "Mantra Health" in text and text.find("Stay Healthy & Connected This Summer") < text.find("Mantra Health"))

    if not ok_header:
        raise RuntimeError("Header date not updated in sections (verification failed).")
    if not ok_oldblock:
        raise RuntimeError("Old block still present before Mantra after section update.")

    print(f"[DEBUG] Updated sections. header_key={hk} event_keys={bks}")
    return True


def ordinal(n: int) -> str:
    if 10 <= n % 100 <= 20:
        suf = "th"
    else:
        suf = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suf}"


def format_header_date(dt: datetime) -> str:
    return dt.strftime("%B ") + ordinal(dt.day) + dt.strftime(", %Y")


def tomorrow_eastern(now: Optional[datetime] = None) -> datetime:
    tz = ZoneInfo("America/New_York")
    now_et = (now or datetime.now(tz)).astimezone(tz)
    # Tomorrow date at same local time
    tomorrow = (now_et + timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    return tomorrow


def schedule_time_iso_9am_eastern(tomorrow_dt: datetime) -> str:
    tz = ZoneInfo("America/New_York")
    dt_9am_et = tomorrow_dt.replace(hour=9, minute=0, second=0, microsecond=0, tzinfo=tz)
    # Mailchimp accepts RFC3339/ISO 8601 with timezone offset
    return dt_9am_et.isoformat()


# Column normalization
COL_MAP_KEYS = {
    "title": ["event title"],
    "description": ["event description"],
    "date": ["date"],
    "time": ["time", "time:"],
    "location": ["location", "location:"],
    "link": ["event link", "event link:"],
    "image_url": [
        "kindly provide the link to your event flyer",
        "link to your event flyer",
        "image",
    ],
}


def _norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(s).strip().lower()).strip()


def map_columns(df: pd.DataFrame) -> Dict[str, str]:
    norms = {col: _norm(col) for col in df.columns}
    mapping: Dict[str, str] = {}
    for key, hints in COL_MAP_KEYS.items():
        for col, n in norms.items():
            if any(h in n for h in hints):
                mapping[key] = col
                break
    missing = [k for k in ["title", "description", "date", "time", "location", "link", "image_url"] if k not in mapping]
    if missing:
        print(f"[WARN] Missing expected columns (will treat as blank where needed): {missing}")
    return mapping


def parse_upcoming_events(df: pd.DataFrame) -> List[Dict[str, str]]:
    mapping = map_columns(df)
    # Parse dates and filter strictly future (upcoming)
    # Using Eastern today
    today_et = datetime.now(ZoneInfo("America/New_York")).date()
    # Convert date column
    date_col = mapping.get("date")
    if date_col is None:
        return []
    df = df.copy()
    df[date_col] = pd.to_datetime(df[date_col], errors="coerce").dt.date
    upcoming = df[df[date_col] > today_et]

    events: List[Dict[str, str]] = []
    for _, row in upcoming.iterrows():
        def val(key: str) -> str:
            col = mapping.get(key)
            if col is None:
                return ""
            v = row.get(col)
            return "" if pd.isna(v) else str(v)
        d = row[date_col]
        events.append({
            "title": val("title").strip(),
            "description": val("description").strip(),
            "date_disp": "" if pd.isna(d) else datetime(d.year, d.month, d.day).strftime("%m/%d/%Y"),
            "time": val("time").strip(),
            "location": val("location").strip(),
            "link": val("link").strip(),
            "image_url": val("image_url").strip(),
        })
    return events


def build_event_block(event: Dict[str, str]) -> str:
    title = (event.get("title") or "").replace("&", "&amp;")
    desc = (event.get("description") or "").replace("&", "&amp;")
    date_disp = event.get("date_disp") or ""
    time = event.get("time") or ""
    location = event.get("location") or ""
    link = event.get("link") or ""
    image_url = event.get("image_url") or ""

    # Basic sanitization; real HTML escaping kept minimal to preserve any desired markup in desc
    def esc(s: str) -> str:
        return s.replace("<", "&lt;").replace(">", "&gt;")

    # Compose content exactly like sample's left-variant (image right, text left)
    # Location (if present) is bold on its own line above the bold date/time.
    location_html = f"<strong>{esc(location)}</strong><br>" if location else ""
    link_html = (
        f'<a href="{esc(link)}" style="mso-line-height-rule: exactly;-ms-text-size-adjust: 100%;-webkit-text-size-adjust: 100%;color: #0c89e9;font-weight: normal;text-decoration: underline;">{esc(link)}</a>'
        if link else ""
    )

    return f'''
<table border="0" cellpadding="0" cellspacing="0" width="100%" class="mcnCaptionBlock" style="border-collapse: collapse;mso-table-lspace: 0pt;mso-table-rspace: 0pt;-ms-text-size-adjust: 100%;-webkit-text-size-adjust: 100%;"><tbody class="mcnCaptionBlockOuter"><tr><td class="mcnCaptionBlockInner" valign="top" style="padding: 9px;mso-line-height-rule: exactly;-ms-text-size-adjust: 100%;-webkit-text-size-adjust: 100%;">
<table border="0" cellpadding="0" cellspacing="0" class="mcnCaptionLeftContentOuter" width="100%" style="border-collapse: collapse;mso-table-lspace: 0pt;mso-table-rspace: 0pt;-ms-text-size-adjust: 100%;-webkit-text-size-adjust: 100%;"><tbody><tr>
<td valign="top" class="mcnCaptionLeftContentInner" style="padding: 0 9px;mso-line-height-rule: exactly;-ms-text-size-adjust: 100%;-webkit-text-size-adjust: 100%;">
<table align="right" border="0" cellpadding="0" cellspacing="0" class="mcnCaptionLeftImageContentContainer" width="264" style="border-collapse: collapse;mso-table-lspace: 0pt;mso-table-rspace: 0pt;-ms-text-size-adjust: 100%;-webkit-text-size-adjust: 100%;float: right;"><tbody><tr>
<td class="mcnCaptionLeftImageContent" align="center" valign="top" style="mso-line-height-rule: exactly;-ms-text-size-adjust: 100%;-webkit-text-size-adjust: 100%;">
{f'<img alt="" src="{esc(image_url)}" width="264" style="max-width: 1080px;border-radius: 2%;border: 0;height: auto;outline: none;text-decoration: none;-ms-interpolation-mode: bicubic;vertical-align: bottom;" class="mcnImage">' if image_url else ''}
</td></tr></tbody></table>
<table class="mcnCaptionLeftTextContentContainer" align="left" border="0" cellpadding="0" cellspacing="0" width="264" style="border-collapse: collapse;mso-table-lspace: 0pt;mso-table-rspace: 0pt;-ms-text-size-adjust: 100%;-webkit-text-size-adjust: 100%;float: left;"><tbody><tr>
<td valign="top" class="mcnTextContent" style="font-family: &quot;Helvetica Neue&quot;, Helvetica, Arial, Verdana, sans-serif;font-size: 14px;line-height: 150%;text-align: left;mso-line-height-rule: exactly;-ms-text-size-adjust: 100%;-webkit-text-size-adjust: 100%;word-break: break-word;color: #000000;">
<h1 class="null" style="text-align: center;display: block;margin: 0;padding: 0;color: #000000;font-family: 'Helvetica Neue', Helvetica, Arial, Verdana, sans-serif;font-size: 26px;font-style: normal;font-weight: bold;line-height: 125%;letter-spacing: normal;">{esc(title)}</h1>
<p style="text-align: left;font-family: &quot;Helvetica Neue&quot;, Helvetica, Arial, Verdana, sans-serif;font-size: 14px;line-height: 150%;margin: 10px 0;padding: 0;mso-line-height-rule: exactly;-ms-text-size-adjust: 100%;-webkit-text-size-adjust: 100%;color: #000000;">{desc}<br><br>{location_html}<strong>{esc(date_disp)}<br>{esc(time)}</strong><br><br>{link_html}</p>
</td></tr></tbody></table>
</td></tr></tbody></table>
</td></tr></tbody></table>
'''


def find_nth(hay: str, needle: str, n: int, start: int = 0) -> int:
    idx = start
    for _ in range(n):
        idx = hay.find(needle, idx)
        if idx == -1:
            return -1
        idx += len(needle)
    return idx - len(needle)


def find_table_block_bounds(html: str, table_start_idx: int) -> Tuple[int, int]:
    """Return (start, end) indices for the outer <table ...>...</table> block starting at given '<table' index."""
    start = html.rfind("<table", 0, table_start_idx + 1)
    if start == -1:
        start = table_start_idx
    depth = 0
    i = start
    token_re = re.compile(r"<table|</table>", re.IGNORECASE)
    for m in token_re.finditer(html, i):
        token = m.group(0).lower()
        if token == "<table":
            depth += 1
            if depth == 1:
                start = m.start()
        else:
            depth -= 1
            if depth == 0:
                end = m.end()
                return (start, end)
    return (start, start)


def find_divider_table_open_start(html: str, class_idx: int) -> int:
    """Given an index of 'class="mcnDividerBlock"', find the opening '<table' start of that divider block."""
    i = class_idx
    while True:
        lt = html.rfind('<table', 0, i + 1)
        if lt == -1:
            return -1
        gt = html.find('>', lt)
        if gt == -1:
            return -1
        opening = html[lt:gt].lower()
        if 'class="mcndividerblock"' in opening:
            return lt
        # keep searching earlier tables
        i = lt - 1



def find_enclosing_table_open(html: str, pos: int) -> int:
    """Find the opening index of the outermost <table ...> whose bounds enclose pos.
    Returns -1 if not found.
    """
    # Walk backwards through table openings until we find one whose closing </table> is after pos
    search_end = pos
    while True:
        t_open = html.rfind('<table', 0, search_end)
        if t_open == -1:
            return -1
        t_start, t_end = find_table_block_bounds(html, t_open)
        if t_end > pos >= t_start:
            # Found a table that encloses pos; try to see if there is a larger enclosing one
            # Continue searching before this start to see if there's an outer table also enclosing pos
            outer_open = html.rfind('<table', 0, t_start)
            if outer_open == -1:
                return t_start
            outer_bounds = find_table_block_bounds(html, outer_open)
            if outer_bounds[1] > pos >= outer_bounds[0]:
                # Move outward
                search_end = outer_open
                continue
            else:
                return t_start
        else:
            # Move further back
            search_end = t_open


def update_html(current_html: str, header_date_str: str, events: List[Dict[str, str]]) -> str:
    html = current_html

    # 1) Update header date inside #templateHeader (the span with font-size:24px)
    header_idx = html.find('id="templateHeader"')
    if header_idx != -1:
        # limit search to a window after header_idx
        window_end = html.find('id="templateBody"', header_idx)
        window_end = window_end if window_end != -1 else header_idx + 8000
        window = html[header_idx:window_end]
        span_re = re.compile(r"(<span[^>]*font-size:\s*24px[^>]*>)(.*?)(</span>)", re.IGNORECASE | re.DOTALL)
        def _repl(m):
            return m.group(1) + header_date_str + m.group(3)
        new_window, n = span_re.subn(_repl, window, count=1)
        if n:
            html = html[:header_idx] + new_window + html[header_idx + len(window):]
        else:
            print("[WARN] Could not locate header date span; leaving as-is.")
    else:
        print("[WARN] #templateHeader not found; leaving header date unchanged.")

    # 2) Replace events section between the two top dividers and the divider before Mantra block
    body_idx = html.find('id="templateBody"')
    if body_idx == -1:
        print("[WARN] #templateBody not found; skipping events replacement.")
        return html

    # Find first two mcnDividerBlock occurrences after body
    first_div_class_idx = html.find('class="mcnDividerBlock"', body_idx)
    second_div_class_idx = html.find('class="mcnDividerBlock"', first_div_class_idx + 1) if first_div_class_idx != -1 else -1
    if first_div_class_idx == -1 or second_div_class_idx == -1:
        print("[WARN] Could not find two top divider blocks; skipping events replacement.")
        return html

    # Compute the exact end of the second divider table
    second_div_open = find_divider_table_open_start(html, second_div_class_idx)
    if second_div_open == -1:
        print("[WARN] Could not find opening <table for the second divider; skipping events replacement.")
        return html
    start_delete_bounds = find_table_block_bounds(html, second_div_open)
    start_delete = start_delete_bounds[1]  # after the second divider table

    # Find Mantra block heading
    mantra_idx = html.lower().find("access support with mantra health".lower(), start_delete)
    if mantra_idx == -1:
        print("[WARN] Mantra Health block not found; skipping events replacement.")
        return html

    # Iterate to find the divider that is immediately above the Mantra block
    scan_pos = start_delete
    last_div_open = -1
    while True:
        idx = html.find('class="mcnDividerBlock"', scan_pos)
        if idx == -1 or idx >= mantra_idx:
            break
        div_open = find_divider_table_open_start(html, idx)
        if div_open != -1:
            last_div_open = div_open
        scan_pos = idx + 1

    if last_div_open == -1:
        print("[WARN] Divider above Mantra not found during scan; skipping events replacement.")
        return html

    end_delete_bounds = find_table_block_bounds(html, last_div_open)
    # end_delete should be the start of the divider table RIGHT BEFORE the Mantra section,
    # so deletion will remove everything up to (but not including) that divider.
    end_delete = end_delete_bounds[0]

    # If there is any stray content like a 'right-variant' block still between end_delete and the Mantra heading,
    # broaden to the enclosing table that contains the Mantra heading and step back to the previous divider.
    # (Safety: only do this if we still detect the "Stay Healthy" headline in between.)
    probe_text = html[end_delete:mantra_idx]
    if "Stay Healthy" in probe_text or "Connected This Summer" in probe_text:
        # Move end_delete earlier to the last divider before mantra (already is), but ensure we did not start too late
        # by recapturing the enclosing table of the Mantra heading and not overlapping.
        enclosing_open = find_enclosing_table_open(html, mantra_idx)
        if enclosing_open != -1 and enclosing_open < mantra_idx and enclosing_open > end_delete:
            end_delete = enclosing_open

    # Build replacement for events area: event block + divider for each event
    events_html = ""
    for ev in events:
        block = build_event_block(ev)
        events_html += block + DIVIDER_HTML

    new_html = html[:start_delete] + events_html + html[end_delete:]
    return new_html


def replicate_update_and_optionally_schedule(excel_url: str, dry_run: bool = True) -> Optional[str]:
    # Compute target date and schedule time
    excel_url = "https://penno365-my.sharepoint.com/:x:/g/personal/gapsa_pr_gapsa_upenn_edu/EWx0O2kdYFxOtPh92obhyNwBL73UMrhbNMyzRKcYLO87wA?download=1"
    tmr = tomorrow_eastern()
    header_date = format_header_date(tmr)
    title = f"GAPSA Newsletter - {header_date}"
    subject = f"✉️GAPSA Newsletter - {header_date}"
    schedule_iso = schedule_time_iso_9am_eastern(tmr)

    # Fetch latest sent campaign
    latest = get_latest_campaign()
    if not latest:
        print("No campaigns found to replicate.")
        return None
    source_id = latest["id"]

    # Create a brand-new campaign (no template), cloning key settings from latest
    src = mailchimp.campaigns.get(source_id)
    list_id = (src.get("recipients") or {}).get("list_id")
    if not list_id:
        raise RuntimeError("Could not read list_id from latest campaign.")

    src_settings = src.get("settings") or {}
    from_name   = src_settings.get("from_name")   or "GAPSA"
    reply_to    = src_settings.get("reply_to")    or "no-reply@example.com"
    to_name     = src_settings.get("to_name")     or ""
    folder_id   = src_settings.get("folder_id")   # may be None, that’s fine

    payload = {
        "type": "regular",
        "recipients": {"list_id": list_id},
        "settings": {
            "title": title,                 # our computed title
            "subject_line": subject,        # our computed subject
            "from_name": from_name,
            "reply_to": reply_to,
            "to_name": to_name,
            "folder_id": folder_id,         # optional
            # IMPORTANT: do NOT include template_id here
        },
        # OPTIONAL: copy tracking options if you care about them
        # "tracking": src.get("tracking") or {},
    }

    new_campaign = mailchimp.campaigns.create(payload)
    new_id = new_campaign["id"]
    print(f"Created new campaign (no template): {new_id}")


    # Update settings (title, subject)
    mailchimp.campaigns.update(new_id, {"settings": {"title": title, "subject_line": subject}})
    print(f"Updated settings: title='{title}', subject='{subject}'")

    # Build from the SOURCE campaign's HTML (the template you like)
    src_content = mailchimp.campaigns.get_content(source_id)
    source_html = src_content.get("html", "") or ""
    if not source_html:
        raise RuntimeError("Latest campaign has empty HTML; nothing to base the new email on.")

    # Read Excel and prepare events
    df = get_first_30_rows_from_excel(excel_url)
    events = parse_upcoming_events(df)

    # Optional safety: don't schedule an empty newsletter
    if not events:
        print("No upcoming events found; not scheduling.")
        return None

    # Update the SOURCE HTML to tomorrow's header + new events
    updated_html = update_html(source_html, header_date, events)

    # Always write a local preview artifact for review
    os.makedirs("artifacts", exist_ok=True)
    with open(os.path.join("artifacts", f"proposed_{new_id}.html"), "w", encoding="utf-8") as f:
        f.write(updated_html)

    # Respect dry_run: do not touch Mailchimp content or schedule
    if dry_run:
        print("Dry run enabled: not updating Mailchimp content or scheduling.")
        return new_id

    # --- Real update path (no template sections) ---
    mailchimp.campaigns.set_content(new_id, {"html": updated_html})

    # Verify on server
    verify = mailchimp.campaigns.get_content(new_id)
    final_blob = verify.get("html", "") or ""
    header_html = format_header_date(tmr)

    if header_html not in final_blob:
        raise RuntimeError("Header date not present in final HTML after set_content")
    if ("Stay Healthy & Connected This Summer" in final_blob) and ("Mantra Health" in final_blob) \
    and (final_blob.find("Stay Healthy & Connected This Summer") < final_blob.find("Mantra Health")):
        raise RuntimeError("Old block still present before Mantra after set_content")

    # Dump the exact HTML that will be sent
    with open(os.path.join("artifacts", f"final_before_schedule_{new_id}.html"), "w", encoding="utf-8") as f:
        f.write(final_blob)

    # Schedule for tomorrow 9 AM Eastern
    mailchimp.campaigns.schedule(new_id, {"schedule_time": schedule_iso})
    print(f"Scheduled campaign at {schedule_iso} (America/New_York)")


    return new_id


if __name__ == "__main__":
    # Configure your Excel link here (must be a direct download link)
    EXCEL_URL = (
        "https://penno365-my.sharepoint.com/:x:/g/personal/gapsa_pr_gapsa_upenn_edu/"
        "EWx0O2kdYFxOtPh92obhyNwBL73UMrhbNMyzRKcYLO87wA?download=1"
    )
    # First run as a dry run to generate proposed HTML locally without changing/scheduling the campaign
    replicate_update_and_optionally_schedule(EXCEL_URL, dry_run=True)

