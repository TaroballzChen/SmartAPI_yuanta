import requests
import sys

TOKEN=" " #填入於LINE notify申請的Token

def lineNotify(msg,token=TOKEN):
    url = "https://notify-api.line.me/api/notify"
    headers = {
        "Authorization": "Bearer " + token,
        "Content-Type": "application/x-www-form-urlencoded"
    }

    payload = {'message': msg}
    r = requests.post(url, headers=headers, params=payload)
    return r.status_code

lineNotify(msg=sys.argv[1])
