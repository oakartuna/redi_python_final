import requests
import openpyxl
from bs4 import BeautifulSoup
import pandas as pd

def retrieve_product_info(url):
    try:
        response = requests.get(url)
        response.raise_for_status # Raise an exception for non-successful status codes
    except requests.exceptions.RequestException as exc:
        print(f"Error occurred during HTTP request: {exc}")
        return []  # Return an empty list in case of an error

    soup = BeautifulSoup(response.content, 'html.parser')

    product_containers = soup.find_all('a', class_='boxed')
    
    product_data = []
    
    for container in product_containers:
        name = container.find('h3', class_='h6').text.strip()
        price = container.find('div', class_='price').text.strip().split('\xa0')[-2].split('€ ')[1]
        price = float(price.replace(",", "."))
        
        camera, receiver = parse_keywords(name)
        
        link = container['href']
        product_url = f"https://www.fpv24.com{link}"
        
        product_data.append({
            'Product Name': name,
            'Price': f"€ {price}",
            'Camera': camera,
            'Receiver': receiver,
            'Link': f'=HYPERLINK("{product_url}","Product Link")'
        })
        
    sorted_product_data = sorted(product_data, key=lambda x: x['Price'])
    return sorted_product_data

def parse_keywords(name):
    camera_keywords = ['Analog', 'DJI', 'Nebula', 'Vista', 'Runcam', 'Polar', 'Caddx']
    receiver_keywords = ['PNP', 'TBS Crossfire', 'Crossfire', 'TBS Nano', 'TBS', 'FrSky']
    
    camera = 'N/S'
    receiver = 'N/S'
    
    for keyword in camera_keywords:
        if keyword.lower() in name.lower():
            camera = 'Analog' if keyword == 'Analog' else 'DJI'
            break
    
    for keyword in receiver_keywords:
        if keyword.lower() in name.lower():
            receiver = 'TBS Crossfire' if keyword == 'Crossfire' else keyword
            break
        
    return camera, receiver

def write_to_file(data, filename):
    df = pd.DataFrame(data)
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='FPV Drone Prices & Features', index=False)
        workbook = writer.book
        worksheet = workbook['FPV Drone Prices & Features']
    
    
    for column in worksheet.columns:
            if column[0].column_letter != 'E':  # Exclude 'Link' column (column E) as it cannot be abjusted properly.
                max_length = 0
                column = [cell for cell in column]
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

            else:
                worksheet.column_dimensions[column[0].column_letter].width = 15

    workbook.save(filename)

url = 'https://www.fpv24.com/en/race-copter-rtf'
data = retrieve_product_info(url)
write_to_file(data, 'FPV_drone_products.xlsx')
