import json
from openpyxl import Workbook, styles, open as read
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from htmldocx import HtmlToDocx
from openai import OpenAI
import time
import requests
import argparse



def flat(slownik, tab=0, parent = "", is_parent = True):
  
    output = []
    for i in slownik.keys():
        output += [(i, tab, parent)] if is_parent else [(i, tab)]
        if type(slownik[i]) == dict:
            output += flat(slownik[i], tab + 1, i, is_parent) if is_parent else flat(slownik[i], tab + 1, is_parent=False)
        else:
            output += [(slownik[i], tab +1, i)] if is_parent else [(slownik[i], tab +1)]

    return output
    

def tlo():
   response = client.images.generate(
      model="dall-e-3",
      prompt=f"create background without any text for inviation to: {rodzaj}. {szczegoly}",
      size="1024x1024",
      quality="standard",
      n=1,
   )
   url = response.data[0].url
   filename = 'background.png'
   r = requests.get(url, allow_redirects=True)
   open(filename, 'wb').write(r.content)

def zakupy():
  completion = client.chat.completions.create(model='gpt-3.5-turbo', 
                                              messages=[
                                              {"role": "system", "content": "you're a party designer"},
                                              {"role": "user", "content": f"plan what to buy for a party with  {goscie} guests, with {pieniadze} budget. {rodzaj}. The party will take place in {miejsce} Show it in JSON {{product: {{'price': estimated price in PLN's for one item, 'number':number of items}}}} without any comments. DONT use square braces. You mustn spend more than {pieniadze} "}
                                              ])
  return completion.choices[0].message.content

def zaproszenie():
  completion = client.chat.completions.create(model='gpt-3.5-turbo', 
                                              messages=[
                                              {"role": "system", "content": "you're a party designer"},
                                              {"role": "user", "content": f"write an invitation for a party using HTML without any comments. {rodzaj}. Party will take place in {miejsce}. {szczegoly}. Use background.png file."}
                                              ])
  return completion.choices[0].message.content

def excel(tekst):
  try:
    workbook = read(filename='Arkusz.xlsx')
  except:
    workbook = Workbook()
  sheet = workbook.active

  slownik = json.loads(tekst)
  wykres = {}
  m, n = "", 0
  wiersz = 1
  max_tab = 0

  for value, tab, parent in flat(slownik):  
    max_tab = max(max_tab, tab)
    if parent == 'price':
        n = chr(ord("A") + tab) + str(wiersz)
        
    sheet[chr(ord("A") + tab) + str(wiersz)].value = value
    wiersz += 1
    if tab == 0:
        m = value
        wykres[m] = "="
    if parent=="number":
        wykres[m] += f"({chr(ord("A") + tab+1) + str(wiersz - 1)})+"
        sheet.cell(row = wiersz-1, column = tab + 2).value = f'={n}*{chr(ord("A") + tab) + str(wiersz-1)}'

  litera = chr(ord("A")+max_tab+1)
  sheet[litera + str(wiersz + 1)] = f"=SUM({litera}1:{litera}{str(wiersz)})"
  sheet[litera + str(wiersz + 1)].font = styles.Font(color="00FF0000", size=20)
  wiersz = 2

  for value, tab in flat(wykres, is_parent=False):   
    
    sheet[chr(ord("L") + (wiersz//2)) + str(tab + 1)].value = value + ("0" if tab == 1 else "")
    wiersz += 1


    
  litera = chr(ord("A")+max_tab+1)
  
  workbook.save("Arkusz.xlsx")


def maps(miejsce):
  driver = webdriver.Edge()
  driver.get('https://www.google.com/maps/dir/Zespół Szkół Zakonu Pijarów Matki Bożej Królowej Szkół Pobożnych/' + miejsce)
  for i in range(0, 1100, 100):
     for j in range(0, 1100, 100):        
      driver.execute_script(f'el = document.elementFromPoint({i}, {j}); el.click();')
  time.sleep(3)
  screenshot = driver.save_screenshot('plik.png')
  
  driver.quit()

parser = argparse.ArgumentParser()
parser.add_argument('--key', type = str, help = 'klucz do API OpenAI')

args = parser.parse_args()
client = OpenAI(
  api_key=args.key
)

rodzaj = input("informantion about party: ")
goscie = input("Number of guests: ")
pieniadze = input("Amount of money: ")
miejsce = input("Place: ")
szczegoly = input("Additional information about invitaion: ")

tlo() 

tekst = zakupy()
excel(tekst)

maps(miejsce)
kod = zaproszenie()

with open("zaproszenie.html", "w") as plik:
   plik.write(kod)

new_parser = HtmlToDocx()
new_parser.parse_html_file("zaproszenie.html", "zaproszenie")
