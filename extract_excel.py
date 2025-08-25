import pandas as pd
import requests
from io import BytesIO

def get_first_30_rows_from_excel(excel_url):
    """
    Download the Excel file and return the first 30 rows as a DataFrame.
    :param excel_url: Direct download link to the Excel file
    :return: pandas.DataFrame with the first 30 rows
    """
    response = requests.get(excel_url)
    response.raise_for_status()
    df = pd.read_excel(BytesIO(response.content), engine='openpyxl')
    return df.head(30)

if __name__ == "__main__":
    url = "https://penno365-my.sharepoint.com/:x:/g/personal/gapsa_pr_gapsa_upenn_edu/EWx0O2kdYFxOtPh92obhyNwBL73UMrhbNMyzRKcYLO87wA?download=1"
    df = get_first_30_rows_from_excel(url)
    print(df)
