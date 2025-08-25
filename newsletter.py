import mailchimp_marketing

from mailchimp_marketing import Client

mailchimp = Client()
mailchimp.set_config({
  "api_key": "aa90f697f32e710c03320a2758f209f5-us6",
  "server": "us6"
})

response = mailchimp.ping.get()
print(response)

data_link = "https://penno365-my.sharepoint.com/:x:/g/personal/gapsa_pr_gapsa_upenn_edu/EWx0O2kdYFxOtPh92obhyNwBL73UMrhbNMyzRKcYLO87wA"