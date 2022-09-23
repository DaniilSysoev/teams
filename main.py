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
    sheet_id = "1i0Zy82V9kmvR0YuFJeYRrCG_QeAzC5a-tTo0ZG5KzWE"
    read = sheet.values().get(spreadsheetId=sheet_id, range="Регистрация!A1:V").execute()
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
            spreadsheetId="1qCqS-gDxqhrrtnQrPzRWd7_HiR8xCBI88nCHGZ2J_Uc",
            range="ОП"
        ).execute()

        #Заполнение листа по оп
        count = 1
        for i in keys:
            time.sleep(1)
            resp = sheet.values().update(
                spreadsheetId="1qCqS-gDxqhrrtnQrPzRWd7_HiR8xCBI88nCHGZ2J_Uc",
                range='ОП!A' + str(count),
                valueInputOption='RAW',
                body={'values': [[i]]}).execute()
            resp = sheet.values().update(
                spreadsheetId="1qCqS-gDxqhrrtnQrPzRWd7_HiR8xCBI88nCHGZ2J_Uc",
                range='ОП!B'+str(count),
                valueInputOption='RAW',
                body={'values': [array[i]]}).execute()
            count += 1
    elif mode == "2":
        #Распределение по командам
        teams = {1: [], 2: [], 3: [], 4: [], 5: [], 6: [], 7: [], 8: [], 9: [], 10: [], 11: [], 12: [], 13: [], 14: [], 15: []}
        i = 1
        was = []
        people_count = 0
        length = sum([len(array[i]) for i in array])
        number_of_teams = 15
        while i <= number_of_teams:
            team_count = 0
            while team_count < (length//number_of_teams + 1):
                for j in range(len(keys)):
                    op_count = 0
                    for t in range(len(array[keys[j]])):
                        if array[keys[j]][t] in was and op_count >= ((length//number_of_teams + 1)//2):
                            break
                        elif op_count < 11:
                            if not(array[keys[j]][t] in was):
                                teams[i].append((keys[j], array[keys[j]][t]))
                                team_count += 1
                                was.append(array[keys[j]][t])
                                op_count += 1
                                people_count += 1
                        else:
                            break
                        if people_count == length:
                            break
                        if team_count >= (length//number_of_teams + 1):
                            break
                    if people_count == length:
                        break
                    if team_count >= (length//number_of_teams + 1):
                        break
                if people_count == length:
                    break
            i += 1
            if people_count == length:
                break

        #Здесь я чисто для себя выводил в консоль все команды и размеры команд, чтобы найти ошибки, но если вдруг понадобится что то изменить в логике работы, то можно это раскоментить
        #print(" ".join([str(len(i)) for i in teams.values()]))
        #print("\n".join([str(teams[i]) for i in range(1, 16)]))
        #print(sum([len(teams[i]) for i in range(1, 16)]))

        #Очиста листа Команды
        keys = list(array.keys())
        sheet.values().clear(
            spreadsheetId="1qCqS-gDxqhrrtnQrPzRWd7_HiR8xCBI88nCHGZ2J_Uc",
            range="Команды"
        ).execute()
        #Запись на лист по командам в вертикальном порядке, но не больше 26 команд (нужно доработать, чтобы было больше команд)
        alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        count = 0
        for i in teams:
            send = [["{0} {1}".format(j[0], j[1])] for j in teams[i]]
            resp = sheet.values().update(
                spreadsheetId="1qCqS-gDxqhrrtnQrPzRWd7_HiR8xCBI88nCHGZ2J_Uc",
                range='Команды!' + alphabet[count] + "1",
                valueInputOption='RAW',
                body={'values': [[i]]}).execute()
            resp = sheet.values().update(
                spreadsheetId="1qCqS-gDxqhrrtnQrPzRWd7_HiR8xCBI88nCHGZ2J_Uc",
                range='Команды!' + alphabet[count] + "2",
                valueInputOption='RAW',
                body={'values': send}).execute()
            count += 1