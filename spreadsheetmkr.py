from asyncio.windows_events import NULL
from csv import list_dialects
from xmlrpc.server import list_public_methods
import requests
import time
from datetime import date
import sys
import xlsxwriter
from tqdm import tqdm

#Ask for AN API key from input instead.

api_key = input("Please enter your API key: ")
time.sleep(1)


g = requests.get("https://api.hypixel.net/guild?key=" + api_key + "&name=Hypixel+Knights")
g = g.json()


today = date.today()
today = today.isoformat()
workbook = xlsxwriter.Workbook('spreadsheets/'+ today +'.xlsx')
worksheet = workbook.add_worksheet()
worksheet.set_column(0, 0, 30)
worksheet.set_column(1, 0, 30)
worksheet.set_column(2, 0, 30)
worksheet.set_column(3, 0, 30)
worksheet.set_column(4, 0, 30)


bold = workbook.add_format({'bold': True, 'bg_color': 'gray', 'align': 'center'})

default = workbook.add_format({'align': 'center'})


worksheet.write('A1', 'Player Name', bold)
worksheet.write('B1', 'Guild Rank', bold)
worksheet.write('C1', 'Total Week Guild XP', bold)
worksheet.write('D1', 'List of people under 25k', bold)


name_slot = 1
rank_slot = 1
gxp_slot = 1
list_slot = 1


members = len(g['guild']['members'])

def addToNaughtyList():
  global list_slot
  list_slot = 1 + list_slot
  total_list_spot = "D"+str(list_slot)
  total_list_color = workbook.add_format({'bg_color': 'ff6666', 'align': 'center'}) #Format to red bg for list cell
  worksheet.write(total_list_spot, name, total_list_color)



def ColorSelector(list_slot, p_uuid):
  expHistory = expHistory = g['guild']['members'][i]['expHistory']
  expHistory = sum(expHistory.values())
  ExemptList = ["Officer", "Manager", "Guild Master"] # gonna add another list of players w horus insurance/hono later using UUID scan.
  if (int(expHistory) >= 0 and int(expHistory) < 25000 and player_rank not in ExemptList):
    total_gxp_color = '#ff6666' #red
    addToNaughtyList()
  else:
    total_gxp_color = '#00cc00' #green
  #Turn back to blue if exempted
  tempColor = checkExemptions(ExemptList, p_uuid)
  #check if flag triggered or not
  if (tempColor != 1):
    total_gxp_color = tempColor
  return total_gxp_color

def checkExemptions(Exemptions, PLAYER_UUID):
  #Flag if exemption not detected
  endColor = 1
  #SpecialExempts = Hono/Horus insurance txt file translated to list
  #Simple read & input to list
  SpecialExempts = []
  for line in open ('hono.txt', 'r').readlines():
      SpecialExempts.append(line.strip())
  #Add to existing list
  for line in open ('insurance.txt', 'r').readlines():
      SpecialExempts.append(line.strip())
  if (player_rank in Exemptions or PLAYER_UUID in SpecialExempts):
    endColor = '#1240EC' #blue

  return endColor




for i in tqdm(range(len(g['guild']['members'])), desc="Progress"):
  #Get member name, Hypixel API only gives UUIDs
  uuid = g['guild']['members'][i]['uuid']

  x = requests.get("https://playerdb.co/api/player/minecraft/" + uuid)
  x = x.json()
  name = x['data']['player']['username']

  # Add cells
  player_rank = g['guild']['members'][i]['rank']
  name_slot = 1 + name_slot
  total_name_slot = "A"+str(name_slot)

  rank_slot = 1 + rank_slot
  total_rank_slot = "B"+str(rank_slot)

  gxp_slot = 1 + gxp_slot
  total_gxp_slot = "C"+str(gxp_slot)
  # Select which color to apply to the gxp amount cells
  textColor = ColorSelector(list_slot, uuid)
  # Sets the color of gxp amount cells & the values
  expHistory = "{:,}".format(sum(g['guild']['members'][i]['expHistory'].values()))
  total_gxp_color = workbook.add_format({'bg_color': textColor, 'align': 'center'})
  worksheet.write(total_name_slot, name, default)
  worksheet.write(total_rank_slot, player_rank, default)
  worksheet.write(total_gxp_slot, expHistory, total_gxp_color,)




#Make sure to close the spreadsheet
workbook.close()
input('Done! Press any key to exit')
