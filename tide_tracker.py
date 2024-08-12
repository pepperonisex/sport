import datetime
import requests

def generate_url(port_id, formatted_date):
    ut = 2
    scale = 1
    return f"https://maree.info/maree-graph.php?p={port_id}&d={formatted_date}&ut={ut}&scale={scale}"

def fetch_maree_data(port_id, formatted_date, j):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36'
    }
    endpoint = f"https://maree.info/do/load-maree-jours.php?p={port_id}&d={formatted_date}&j={j}"
    
    session = requests.Session()
    session.get("https://maree.info/", headers=headers)
    response = session.get(endpoint, headers=headers)
    
    if response.status_code == 200:
        print("Maree data:", response.text)
    else:
        print("Failed to retrieve maree data. Status code:", response.status_code)

def main():
    today = datetime.datetime.now()
    formatted_date = today.strftime('%Y%m%d')
    port_id = 58
    j_value = 1
    
    url_today = generate_url(port_id, formatted_date)
    print("URL for today's tide image:", url_today)
    
    fetch_maree_data(port_id, formatted_date, j_value)

if __name__ == "__main__":
    main()
