import pandas
import pyperclip

team = "Buffalo"
conference = "MAC"
b = '[b]'
u = '[u]'
bc = '[/b]'
uc = '[/u]'

def generate_weekly_report():
    output = ''
    data_sheet = pandas.read_excel('recruit_update.xlsx', sheet_name='recruit')
    output += '[font=Arial]\n\n'
    output += f"[b][u][size=120]{team} Recruiting Update (Sorted by National"
    output += ' Ranking)[/size][/u][/b]\n'
    output += create_recruits(data_sheet)
    output += create_class_ranking(data_sheet)
    pyperclip.copy(output)
    pyperclip.paste()   

def create_recruits(sheet):
    output = ''
    commit_change = False
    for i in range(0, sheet.values.shape[0]):
        row = sheet.values[i]
        player = f"{row[0]} #{row[1]} {row[2]}: {row[3]}, {row[4]} lbs |"
        player += f" {row[5]} | {row[6]}-star | "
        if row[7] in [3, 5, 8]:
            player += f"Top {row[7]} ({row[8]})\n"
        elif row[7] == 'O':
            player += f"Open ({row[8]})\n"
        elif row[7] == 'C' and row[8] == team:
            player += "Committed\n"
        elif row[7] == 'C':
            player += f"Commited to {row[8]}\n"
        if row[9] == True:
            player = b + player + bc
            commit_change = True
        output += player

    if commit_change:
        output += b + "Bold indicates a change in commitment status" + bc + '\n'

    output += '\n'
    return output

def create_class_ranking(sheet):
    output = ''
    output += b + u + "Class Ranking" + uc + bc + '\n'
    row = sheet.columns
    output += f"#{row[13]} Overall, #{row[11]} {conference}"

    return output

generate_weekly_report()