import mailchimp_marketing
from mailchimp_marketing import Client
from get_latest_campaign import get_latest_campaign_full
from extract_excel import get_first_30_rows_from_excel
import pandas as pd
import requests
from io import BytesIO

import os
from mailchimp_marketing import Client
from dotenv import load_dotenv

# load values from .env into the environment
load_dotenv()

MAILCHIMP_API_KEY = os.getenv("MAILCHIMP_API_KEY")
MAILCHIMP_SERVER_PREFIX = os.getenv("MAILCHIMP_SERVER_PREFIX", "us6")

mailchimp = Client()
mailchimp.set_config({
    "api_key": MAILCHIMP_API_KEY,
    "server": MAILCHIMP_SERVER_PREFIX,
})


def update_newsletter_design_from_excel(excel_url):
    """
    Fetch the latest campaign, read the Excel file, and update the campaign design.
    """
    # Get latest campaign info
    campaign_data = get_latest_campaign_full()
    if not campaign_data:
        print("No campaign found.")
        return
    campaign_id = campaign_data['basic']['id']
    print(f"Editing campaign: {campaign_id}")

    # Get first 30 rows from Excel file
    df = get_first_30_rows_from_excel(excel_url)
    print("Excel preview (first 30 rows):")
    print(df)

    # Example: Use the current HTML as a base
    current_html = campaign_data['content'].get('html', '')
    # TODO: Update current_html with new content from Excel

    # Update campaign design (this will not change anything until you edit current_html)
    #resp = mailchimp.campaigns.set_content(campaign_id, {"html": current_html})
    #print("Mailchimp API response:")
    #print(resp)

if __name__ == "__main__":
    # Use the Excel link from your newsletter.py
    excel_url = "https://penno365-my.sharepoint.com/:x:/g/personal/gapsa_pr_gapsa_upenn_edu/EWx0O2kdYFxOtPh92obhyNwBL73UMrhbNMyzRKcYLO87wA?download=1"
    update_newsletter_design_from_excel(excel_url)
