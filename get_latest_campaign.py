import mailchimp_marketing
from mailchimp_marketing import Client

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


def get_latest_campaign():
    """
    Fetch the most recent Mailchimp campaign (by send_time).
    :return: Campaign data dict or None if not found
    """
    mailchimp = Client()
    mailchimp.set_config({
        "api_key": MAILCHIMP_API_KEY,
        "server": "us6"
    })
    campaigns = mailchimp.campaigns.list(sort_field="send_time", sort_dir="DESC", count=1)
    if campaigns.get('campaigns'):
        return campaigns['campaigns'][0]
    return None

def get_latest_campaign_full():
    """
    Fetch all available details for the most recent Mailchimp campaign.
    :return: Dict with all campaign details, or None if not found
    """
    mailchimp = Client()
    mailchimp.set_config({
        "api_key": MAILCHIMP_API_KEY,
        "server": "us6"
    })
    campaigns = mailchimp.campaigns.list(sort_field="send_time", sort_dir="DESC", count=1)
    if not campaigns.get('campaigns'):
        return None
    latest = campaigns['campaigns'][0]
    campaign_id = latest['id']
    # Fetch all available details
    details = mailchimp.campaigns.get(campaign_id)
    content = mailchimp.campaigns.get_content(campaign_id)
    feedback = mailchimp.campaigns.get_feedback(campaign_id)
    checklist = mailchimp.campaigns.get_send_checklist(campaign_id)
    report = None
    try:
        report = mailchimp.reports.get_campaign_report(campaign_id)
    except Exception:
        pass  # Report only exists for sent campaigns
    return {
        'basic': latest,
        'details': details,
        'content': content,
        'feedback': feedback,
        'send_checklist': checklist,
        'report': report
    }

if __name__ == "__main__":
    latest = get_latest_campaign()
    if latest:
        print("Latest campaign:")
        print(f"ID: {latest['id']}")
        print(f"Title: {latest['settings']['title']}")
        print(f"Send Time: {latest.get('send_time')}")
    else:
        print("No campaigns found.")
    
    data = get_latest_campaign_full()
    if data:
        print("\n=== Basic Info ===\n", data['basic'])
        print("\n=== Details ===\n", data['details'])
        print("\n=== Content ===\n", data['content'])
        print("\n=== Feedback ===\n", data['feedback'])
        print("\n=== Send Checklist ===\n", data['send_checklist'])
        print("\n=== Report ===\n", data['report'])
    else:
        print("No campaigns found.")
