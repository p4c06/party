# party-gen

End-of-year project in computer science

## Purpose

This is a program written for the end of the school year in elementary school computer science class. Its main purpose is to demonstrate the impact of computers and AI on our lives. An example of such an impact is the organization of a party.

The program performs most of the tasks needed to organize a party. It generates:
1. A shopping list in Excel
2. A map for directions
3. A background for the invitation
4. An invitation in HTML
5. Converts it to Word

## How to use the code
to run the code install all needed python modules.
```pip3 install -r requirements.txt```
then you should run program with those flags:
```python3 projekt-informatyka.py --key [YOUR OPENAI KEY]```
program will ask you for a details:
1. kind of party (eg. birthday, anniversary)
2. number of guests
3. how much money can you spend
4. where the party will be organized
5. the details about the invitaion (eg. colour, background, theme)

And then generated files will be:
- invitaion: zaproszenie.html / zaproszenie.docx
- directions: plik.png
- shopping list: Arkusz.xlsx
- background for invivtaion: background.png

## Contact

e-mail: franek@gawron.pro