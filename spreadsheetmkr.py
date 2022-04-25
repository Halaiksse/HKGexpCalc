import requests
import time
import xlsxwriter
from tqdm import tqdm

#Ask for AN API key from input instead.
try:
    with open('API_KEY') as inf:
        api_key = inf.read()
except FileNotFoundError:
    api_key = input("Please enter your API key: ")
    with open('API_KEY', 'w') as outf:
        outf.write(api_key)

time.sleep(1)
## TODO: Dont ask, manually set HK instead.
guild = input("ENTER A GUILD NAME: ")

g = requests.get("https://api.hypixel.net/guild?key=" + api_key + "&name=" + guild)
g = g.json()

#One spreadsheet for everyone and one spreadsheet for < 25k gexp.

## TODO: Add date in spreadsheet name / always set for HK
workbook = xlsxwriter.Workbook('spreadsheets/'+guild+'.xlsx')
worksheet = workbook.add_worksheet()
worksheet.set_column(0, 0, 30)
worksheet.set_column(1, 0, 30)
worksheet.set_column(2, 0, 30)
worksheet.set_column(3, 0, 30)
worksheet.set_column(4, 0, 30)

bold = workbook.add_format({'bold': True, 'bg_color': 'gray', 'align': 'center'})

default = workbook.add_format({'align': 'center'})

#Remove D1
worksheet.write('A1', 'Player Name', bold)
worksheet.write('B1', 'Guild Rank', bold)
worksheet.write('C1', 'Total Week Guild XP', bold)
worksheet.write('D1', 'Join Date', bold)

name_slot = 1
rank_slot = 1
gxp_slot = 1
join_slot = 1

members = len(g['guild']['members'])



#Function + Skip anyone above ET.
for i in tqdm(range(len(g['guild']['members'])), desc="Progress"):
  uuid = g['guild']['members'][i]['uuid']

  x = requests.get("https://playerdb.co/api/player/minecraft/" + uuid)
  x = x.json()
  name = x['data']['player']['username']


  player_rank = g['guild']['members'][i]['rank']
  name_slot = 1 + name_slot
  total_name_slot = "A"+str(name_slot)

  rank_slot = 1 + rank_slot
  total_rank_slot = "B"+str(rank_slot)

  gxp_slot = 1 + gxp_slot
  total_gxp_slot = "C"+str(gxp_slot)
  #Remove
  join_slot = 1 + join_slot
  total_join_slot = "D"+str(join_slot)



  expHistory = expHistory = g['guild']['members'][i]['expHistory']
  expHistory = sum(expHistory.values())
  #CHNG For 0 to < 25k alert, above is green.
  if (int(expHistory) >= 0):
    total_gxp_color = '#ff6666'
  if (int(expHistory) >= 10000):
    total_gxp_color = '#ffff66'
  if (int(expHistory) >= 100000):
    total_gxp_color = '#33ff33'
  if (int(expHistory) >= 200000):
    total_gxp_color = '#00cc00'

  join_date = g['guild']['members'][i]['joined']
  join_date = join_date/1000
  join_date = time.strftime("%a, %d %b %Y %H:%M:%S", time.localtime(join_date))


  expHistory = "{:,}".format(sum(g['guild']['members'][i]['expHistory'].values()))
  total_gxp_color = workbook.add_format({'bg_color': total_gxp_color, 'align': 'center'})

  worksheet.write(total_name_slot, name, default)
  worksheet.write(total_rank_slot, player_rank, default)
  worksheet.write(total_gxp_slot, expHistory, total_gxp_color,)
  worksheet.write(total_join_slot, join_date, default,)




workbook.close()
print('Done! Press any key to exit')
input()
