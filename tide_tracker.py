import datetime
import requests
import os
from bs4 import BeautifulSoup


def fetch_maree_data(session, port_id):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36',
    }
    endpoint = f"https://maree.info/{port_id}"
    response = session.get(endpoint, headers=headers)
    
    if response.status_code == 200:
        soup = BeautifulSoup(response.content, 'html.parser')
        table = soup.find('table', id='MareeJours')
        maree_data = []

        if table:
            rows = table.find_all('tr')[1:]
            for row in rows:
                date = row.find('th').text.strip()
                times = [td.strip() for td in row.find_all('td')[0].stripped_strings]
                heights = [td.strip() for td in row.find_all('td')[1].stripped_strings]
                coefficients = [td.strip() for td in row.find_all('td')[2].stripped_strings]
                maree_data.append({
                    'date': date,
                    'times': times,
                    'heights': heights,
                    'coefficients': coefficients
                })

            for entry in maree_data:
                print(f"Date: {entry['date']}")
                print(f"Heures de marée: {', '.join(entry['times'])}")
                print(f"Hauteurs: {', '.join(entry['heights'])}")
                print(f"Coefficients: {', '.join(entry['coefficients'])}")
                print("-" * 40)
        else:
            print("Table des marées non trouvée.")
    else:
        print(f"Failed to retrieve maree data. Status code: {response.status_code}")


def main():
    base_url = "https://maree.info/"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36',
        'Referer': base_url,
        'Accept': 'image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8',
    }

    today = datetime.datetime.now()
    formatted_date = today.strftime('%Y%m%d')
    port_id = 58
    
    session = requests.Session()
    session.get(base_url, headers=headers)
    
    script_dir = os.path.dirname(os.path.realpath(__file__))
    image_directory = os.path.join(script_dir, 'images')
    image_path = os.path.join(image_directory, f"maree_{formatted_date}.png")
    os.makedirs(image_directory, exist_ok=True)
    
    response = session.get(f"https://maree.info/maree-graph.php?p={port_id}&d={formatted_date}&ut=2&scale=1", headers=headers)
    
    with open(image_path, 'wb') as file:
        file.write(response.content)
    print(f"Image saved to {image_path}")
        
    fetch_maree_data(session, port_id)

if __name__ == "__main__":
    main()
