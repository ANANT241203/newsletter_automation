import os
from get_latest_campaign import get_latest_campaign_full


def dump_latest_campaign_html(out_dir: str = "artifacts", filename: str = "latest_campaign.html") -> int:
    """
    Fetch the latest Mailchimp campaign's HTML and write it to a local file for inspection.
    Returns 0 on success, non-zero on failure.
    """
    data = get_latest_campaign_full()
    if not data:
        print("No campaigns found.")
        return 1

    html = data.get("content", {}).get("html", "")
    if not html:
        print("Latest campaign has no HTML content available.")
        return 2

    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, filename)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)

    basic = data.get("basic", {})
    title = basic.get("settings", {}).get("title")
    cid = basic.get("id")
    print(f"Wrote HTML to: {os.path.abspath(out_path)}")
    print(f"Campaign ID: {cid}\nTitle: {title}")
    return 0


if __name__ == "__main__":
    raise SystemExit(dump_latest_campaign_html())

