import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl.styles import Alignment
import os
import urllib.parse
from colorama import init, Fore, Back, Style

init(autoreset=True)

print("***********************************************")
print(Style.BRIGHT + Fore.RED + "WARNING: This is intended for educational purposes only.")
print(Style.BRIGHT + Fore.BLUE + "This script serves as a learning tool for web scraping and should be used with care.")
print(Style.BRIGHT + Fore.WHITE + "Be sure to comply with the website's terms and conditions.")
print("***********************************************\n")

def is_file_open(file_path):
    try:
        with open(file_path, 'a'):
            return False
    except IOError:
        return True

location = input("Voer de locatie in voor de vacaturezoekopdracht (bijv. 'Veenendaal'): ").strip()
encoded_location = urllib.parse.quote(location)

url = f'https://www.werkzoeken.nl/vacatures/?what=&where={encoded_location}&filtered=1&r=30&submit_homesearch='
print(f"Vacatures ophalen in {location}...")

try:
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'}
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    
    soup = BeautifulSoup(response.text, 'html.parser')
    print("Informatie aan het verwerken...")

    jobs = soup.find_all('a', class_='vacancy')
    if not jobs:
        print("Geen vacatures gevonden.")

    job_data = []

    for job in jobs:
        title = job.find('h3').text.strip() if job.find('h3') else 'Geen titel gevonden'
        
        location_and_company = job.find('div', class_='location-and-company-name')
        if location_and_company:
            location = location_and_company.find('strong').text.strip()
            company_name = location_and_company.text.split('•')[-1].strip()
        else:
            location = 'Geen locatie gevonden'
            company_name = 'Geen bedrijfsnaam gevonden'

        requested_info = job.find('div', class_='requested-wrapper')
        if requested_info:
            hours = requested_info.find('div').text.strip() if requested_info.find('div') else 'Geen uren gevonden'
            offers = requested_info.find_all('div', class_='offer')
            salary = offers[0].text.strip() if len(offers) > 0 else 'Geen salaris gevonden'
            additional_info = ', '.join(offer.text.strip() for offer in offers)
        else:
            hours = 'Geen uren gevonden'
            salary = 'Geen salaris gevonden'
            additional_info = 'Geen aanvullende informatie gevonden.'

        job_data.append({
            'Functie': title,
            'Bedrijf': company_name,
            'Locatie': location,
            'Uren': hours,
            'Salaris': salary,
            'Aanvullende informatie': additional_info,
        })

    df = pd.DataFrame(job_data)
    excel_filename = 'Vacaturen.xlsx'

    if is_file_open(excel_filename):
        print(Fore.RED + f"Foutmelding: Het bestand... '{excel_filename}' is momenteel geopend. Sluit eerst het volledige bestand voordat je het script uitvoert.")
        input("Druk op Enter om het script te beëindigen...")
        exit()

    df.to_excel(excel_filename, index=False)

    with pd.ExcelWriter(excel_filename, engine='openpyxl', mode='a') as writer:
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # Kolombreedte
        column_widths = {
            'A': 55,  # Functie
            'B': 35,  # Bedrijf
            'C': 35,  # Locatie
            'D': 20,  # Uren
            'E': 30,  # Salaris
            'F': 50   # Aanvullende informatie
        }

        for col, width in column_widths.items():
            worksheet.column_dimensions[col].width = width

        for cell in worksheet['F']:
            cell.alignment = Alignment(wrap_text=True)

    print(Fore.GREEN + f"De gegevens zijn succesvol verzameld. Het Excel-bestand '{excel_filename}' is aangemaakt.")

except requests.exceptions.RequestException as e:
    print(Fore.RED + f"Er is een fout opgetreden: {e}")

input("")
