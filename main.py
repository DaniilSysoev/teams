import os
import httplib2
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials
import time


#Выбор режима работы
while True:
    creds_json = os.path.dirname(__file__) + '/credentials.json'
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds_service = ServiceAccountCredentials.from_json_keyfile_name(creds_json, scopes).authorize(httplib2.Http())
    service = build('sheets', 'v4', http=creds_service)
    sheet = service.spreadsheets()
    sheet_id = "тут id листа откуда читают"
    read = sheet.values().get(spreadsheetId=sheet_id, range="...!A1:V").execute()
    array = {}
    # Считывание таблицы и записывание в словарь, где ключ - оп, значение - список людей
    for i in read["values"]:
        if i[3] == "1":
            if i[21] in array.keys():
                array[i[21]].append(" ".join([i[0], i[1], i[2]]))
            else:
                array[i[21]] = [" ".join([i[0], i[1], i[2]])]
    keys = list(array.keys())
    print("Распределить по ОП(1)\nРаспределить по командам(2)")
    mode = input()
    if mode == "1":
        #Очиcтка листа ОП
        sheet.values().clear(
            spreadsheetId="id листа куда писать",
            range="ОП"
        ).execute()

        #Заполнение листа по оп
        count = 1
        for i in keys:
            time.sleep(1)
            resp = sheet.values().update(
                spreadsheetId="",
                range='ОП!A' + str(count),
                valueInputOption='RAW',
                body={'values': [[i]]}).execute()
            resp = sheet.values().update(
                spreadsheetId="",
                range='ОП!B'+str(count),
                valueInputOption='RAW',
                body={'values': [array[i]]}).execute()
            count += 1
    elif mode == "2":
        op = ["ИБ", "КБ", "ИТСС", "ИВТ", "ПМ"]
        miem = {}
        not_miem = {}
        for i in array.keys():
            if i in op:
                miem[i] = array[i]
            else:
                not_miem[i] = array[i]
        #Распределение по командам
        teams = {1: [], 2: [], 3: [], 4: [], 5: [], 6: [], 7: [], 8: [], 9: [], 10: [], 11: [], 12: [], 13: [], 14: [], 15: []}
        i = 1
        was = []
        people_count = 0
        length_miem = sum([len(miem[i]) for i in miem])
        number_of_teams = 15
        miem_keys = list(miem.keys())
        while i <= number_of_teams:
            team_count = 0
            while team_count < (length_miem//number_of_teams + 1):
                for j in range(len(miem_keys)):
                    op_count = 0
                    for t in range(len(miem[miem_keys[j]])):
                        if miem[miem_keys[j]][t] in was and op_count >= ((length_miem//number_of_teams + 1)//2 + 1):
                            break
                        elif op_count < ((length_miem//number_of_teams + 1)//2 + 1):
                            if not(miem[miem_keys[j]][t] in was):
                                teams[i].append((miem_keys[j], miem[miem_keys[j]][t]))
                                team_count += 1
                                was.append(miem[miem_keys[j]][t])
                                op_count += 1
                                people_count += 1
                        else:
                            break
                        if people_count == length_miem:
                            break
                        if team_count >= (length_miem//number_of_teams + 1):
                            break
                    if people_count == length_miem:
                        break
                    if team_count >= (length_miem//number_of_teams + 1):
                        break
                if people_count == length_miem:
                    break
            i += 1
            if people_count == length_miem:
                break

        i = 1
        was = []
        people_count = 0
        length_not_miem = sum([len(not_miem[i]) for i in not_miem])
        number_of_teams = 15
        not_miem_keys = list(not_miem.keys())
        while i <= number_of_teams:
            team_count = 0
            while team_count < (length_not_miem // number_of_teams +1):
                for j in range(len(not_miem_keys)):
                    op_count = 0
                    for t in range(len(not_miem[not_miem_keys[j]])):
                        if not_miem[not_miem_keys[j]][t] in was and op_count >= ((length_not_miem // number_of_teams + 1) // 2 + 1):
                            break
                        elif op_count < ((length_not_miem // number_of_teams + 1) // 2 + 1):
                            if not (not_miem[not_miem_keys[j]][t] in was):
                                teams[i].append((not_miem_keys[j], not_miem[not_miem_keys[j]][t]))
                                team_count += 1
                                was.append(not_miem[not_miem_keys[j]][t])
                                op_count += 1
                                people_count += 1
                        else:
                            break
                        if people_count == length_not_miem:
                            break
                        if team_count >= (length_not_miem // number_of_teams + 1):
                            break
                    if people_count == length_not_miem:
                        break
                    if team_count >= (length_not_miem // number_of_teams + 1):
                        break
                if people_count == length_not_miem:
                    break
            i += 1
            if people_count == length_not_miem:
                break
        #Здесь я чисто для себя выводил в консоль все команды и размеры команд, чтобы найти ошибки, но если вдруг понадобится что то изменить в логике работы, то можно это раскоментить
        #print(" ".join([str(len(i)) for i in teams.values()]))
        #print("\n".join([str(teams[i]) for i in range(1, 16)]))
        #print(sum([len(teams[i]) for i in range(1, 16)]))

        #Очиста листа Команды
        keys = list(array.keys())
        sheet.values().clear(
            spreadsheetId="",
            range="Команды(старое)"
        ).execute()
        #Запись на лист по командам в вертикальном порядке, но не больше 26 команд (нужно доработать, чтобы было больше команд)
        alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        count = 0
        for i in teams:
            send = [["{0} {1}".format(j[0], j[1])] for j in teams[i]]
            resp = sheet.values().update(
                spreadsheetId="",
                range='Команды(старое)!' + alphabet[count] + "1",
                valueInputOption='RAW',
                body={'values': [[i]]}).execute()
            resp = sheet.values().update(
                spreadsheetId="",
                range='Команды(старое)!' + alphabet[count] + "2",
                valueInputOption='RAW',
                body={'values': send}).execute()
            count += 1
