import pandas
import pyperclip

team = "Buffalo"

def generate_weekly_report():
    output = ''
    data_sheet = pandas.read_excel('weekly_update.xlsx', sheet_name=None)
    output += '[font=Arial]\n\n'
    output += generate_conference_scores(data_sheet['scores'])
    output += generate_potw(data_sheet['potw'])
    output += generate_injuries(data_sheet['injuries'])
    output += generate_conference_standings(data_sheet['standings'])
    output += generate_top_25(data_sheet['top25'])
    output += generate_next_week(data_sheet['next_week'])
    pyperclip.copy(output)
    pyperclip.paste()   

def generate_conference_scores(sheet):
    output = ''

    week = sheet.columns[3]
    output += f"[b][u][size=120]Week {week} Results[/size][/u][/b]\n"
    for i in range(1, sheet.values.shape[0]):
        row = sheet.values[i]
        output += f"{row[0]} {row[1]} - {row[2]} {row[3]}\n"

    output += '\n'
    return output

def generate_potw(sheet):
    output = ''
    nat_potw = False
    output += '[b][u]Players of the Week[/u][/b]\n'
    for i in range(0, sheet.values.shape[0]):
        row = sheet.values[i]
        if (row[4] == True):
            output += f"* {row[0]} {row[1]}, {row[2]}: {row[3]}\n"
            nat_potw = True
        else:
            output += f"{row[0]} {row[1]}, {row[2]}: {row[3]}\n"
    if nat_potw:
        output += '\n* indicates national player of the week\n\n'
    else:
        output += '\n'

    return output

def generate_injuries(sheet):
    output = ''
    output += f"[b][u]{team} Injuries[/u][/b]\n"
    if sheet.values.shape[0] == 0:
        output += 'None\n\n'
        return output

    for i in range(0, sheet.values.shape[0]):
        row = sheet.values[i]
        output += f"{row[0]} {row[1]} - {row[2]} ({row[3]})\n"

    output += '\n'

    return output

def generate_conference_standings(sheet):
    output = ''
    output += '[b][u]MAC Standings[/u][/b]\n'

    for i in range(0, sheet.values.shape[0]):
        row = sheet.values[i]
        output += f"{i+1}. {row[1]}, {row[2]}-{row[3]} ({row[4]}-{row[5]})\n"

    output += '\n'

    return output

def generate_top_25(sheet):
    output = ''
    output += '[b][u]Top 25[/u][/b]\n'

    for i in range(0, sheet.values.shape[0]):
        row = sheet.values[i]
        output += f"{i+1}. {row[1]} ({row[2]}-{row[3]})\n"

    output += '\n'

    return output

def generate_next_week(sheet):
    output = ''
    output += f"[b][u]{team} Upcoming Week[/u][/b]\n"
    row = sheet.values[0]
    output += row[0]

    return output

generate_weekly_report()