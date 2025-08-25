import mailchimp_marketing

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


response = mailchimp.ping.get()
print(response)

data_link = "https://penno365-my.sharepoint.com/:x:/g/personal/gapsa_pr_gapsa_upenn_edu/EWx0O2kdYFxOtPh92obhyNwBL73UMrhbNMyzRKcYLO87wA"