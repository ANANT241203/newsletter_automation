import os
import sys
from datetime import datetime

from automate_newsletter import replicate_update_and_optionally_schedule

# Excel link (public, direct download)
EXCEL_URL = (
    "https://penno365-my.sharepoint.com/:x:/g/personal/gapsa_pr_gapsa_upenn_edu/"
    "EWx0O2kdYFxOtPh92obhyNwBL73UMrhbNMyzRKcYLO87wA?download=1"
)


def main() -> int:
    os.makedirs("artifacts", exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = os.path.join("artifacts", f"run_{ts}.log")

    def log(msg: str):
        print(msg)
        try:
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(msg + "\n")
        except Exception:
            # Best-effort logging; ignore file errors
            pass

    log("Starting GAPSA newsletter automation (replicate + update + schedule)...")
    try:
        new_id = replicate_update_and_optionally_schedule(EXCEL_URL, dry_run=False)
        if not new_id:
            log("Failed: replicate_update_and_optionally_schedule returned no campaign id.")
            return 2
        log(f"Success. New campaign id: {new_id}")
        log("A copy of the final HTML was pushed to Mailchimp and the campaign was scheduled for 9:00 AM ET tomorrow.")
        log("See artifacts/ for any saved HTML or logs.")
        return 0
    except Exception as e:
        log(f"Error: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())

