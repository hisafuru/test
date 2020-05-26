# -*- coding: utf-8 -*-
import gspread,json,yaml,datetime,time,sys,importlib,csv
sys.path.append("../GoogleSheets/GoogleDrive")
import editDrive
import requests

#ServiceAccountCredentials：Googleの各サービスへアクセスできるservice変数を生成します。
from oauth2client.service_account import ServiceAccountCredentials

class GoogleSheets():
    record = []
    pastRecord = []
    row = 0

    def __init__(self):
        print('Initializing sheets api...')

        sys.setrecursionlimit(300)

        #paperless_setting.yamlの読み込み
        try:
            with open('paperless_setting.yaml') as file:
                self.config = yaml.safe_load(file.read())
                file.close()
        except Exception as e:
            print(e)
            print('Please check if paperless_setting.yml exists in the setting folder.')
            sys.exit(1)

        #2つのAPIを記述しないとリフレッシュトークンを3600秒毎に発行し続けなければならない
        scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']


        try:

            #認証情報設定
            #ダウンロードしたjsonファイル名をクレデンシャル変数に設定（秘密鍵、Pythonファイルから読み込みしやすい位置に置く）
            credentials = ServiceAccountCredentials.from_json_keyfile_name('client-secret-sheet.json', scope)

            #OAuth2の資格情報を使用してGoogle APIにログインします。
            self.gc = gspread.authorize(credentials)

            #共有設定したスプレッドシートキーを変数[SPREADSHEET_KEY]に格納する。
            self.SPREADSHEET_KEY = self.config["SPREADSHEET_KEY"]
            self.FILE_PATH = self.config['MACHINE_NAME'] + ".csv"
            self.worksheet = self.gc.open_by_key(self.SPREADSHEET_KEY).sheet1

            with open(self.FILE_PATH, "w") as f:
                writer = csv.writer(f, lineterminator="\n")
                self.row = len(self.worksheet.get_all_values())
                writer.writerows(self.worksheet.get_all_values())
                f.close()

            self.gdrive = editDrive.GoogleDriveUpload()
            self.error = False
            self.pastError = False
            self.diffErrorAndNowRow = 0
            self.quantityPerHour = 0

            if self.row != 1:
                self.record = [self.worksheet.cell(self.row, 1).value,self.worksheet.cell(self.row, 2).value,self.worksheet.cell(self.row, 3).value,self.worksheet.cell(self.row, 4).value,self.worksheet.cell(self.row, 5).value,self.worksheet.cell(self.row,  6).value]
        except requests.exceptions.ReadTimeout as e:
            print(e)
            time.sleep(2)
            __init__()
        except Exception as e:
            print('An unexpected error occurred')
            print(e)
            sys.exit(1)

        print('Finish Initializing.')

    #1時間経った時に呼び出し、新しいファイルを作る
    def create_new_record(self):
        try:
            self.worksheet = self.gc.open_by_key(self.SPREADSHEET_KEY).sheet1

            with open(self.FILE_PATH, "w") as f:
                writer = csv.writer(f, lineterminator="\n")
                self.row = len(self.worksheet.get_all_values())
                writer.writerows(self.worksheet.get_all_values())
                f.close()

            #もしエラー中なら
            if self.error == True and self.pastError == False:
                self.diffErrorAndNowRow+=1
                self.pastError = True
                self.pastRecord = self.record
            elif self.pastError == True and self.error == True:
                self.diffErrorAndNowRow+=1
                self.pastError = True

            #既にその時間のレコードがあればそのまま
            if not (datetime.datetime.now().strftime('%Y%m%d') == self.worksheet.cell(self.row, 2).value and str(datetime.datetime.now().hour) == self.worksheet.cell(self.row, 3).value):
                self.quantityPerHour = 0

                csv_data = self.read_csv(self.FILE_PATH)
                csv_data[self.row-1] = self.record
                self.record = [self.config['ITEM_NAME'],int(datetime.datetime.now().strftime('%Y%m%d')),datetime.datetime.now().hour,None,None,None]
                csv_data.append(self.record)
                self.write_csv(self.FILE_PATH, csv_data)
                #ファイル容量の関係でアップロードをするかどうか考え中
                self.gdrive.upload_file()

            else:
                self.record = [self.worksheet.cell(self.row, 1).value,self.worksheet.cell(self.row, 2).value,self.worksheet.cell(self.row, 3).value,self.worksheet.cell(self.row, 4).value,self.worksheet.cell(self.row, 5).value,self.worksheet.cell(self.row,  6).value]

        except requests.exceptions.ReadTimeout as e:
            print(e)
            time.sleep(2)
            create_new_record()

        except Exception as e:
            print('An unexpected error occurred')
            print(e)
            sys.exit(1)

    def update_createtime(self):
        try:
            time = str(datetime.datetime.now().strftime('%M:%S'))
            if self.record[3] == None or self.record[3] == "":
                self.record[3] = time + ','
            else:
                self.record[3] += time + ','

            self.worksheet = self.gc.open_by_key(self.SPREADSHEET_KEY).sheet1
            self.row = len(self.worksheet.get_all_values())
            self.worksheet.update_cell(self.row,4,self.record[3])
            self.quantityPerHour += 1

        except requests.exceptions.ReadTimeout as e:
            print(e)
            time.sleep(2)
            update_createtime()

        except Exception as e:
            print('An unexpected error occurred')
            print(e)
            sys.exit(1)

    def update_errortime(self):
        try:
            time = str(datetime.datetime.now().strftime('%M:%S'))

            self.worksheet = self.gc.open_by_key(self.SPREADSHEET_KEY).sheet1
            self.row = len(self.worksheet.get_all_values())

            if self.pastError == True and self.error == True:
                self.pastRecord[4] += time + ','

                self.error = False
                self.pastError = False

                self.worksheet.update_cell(self.row-self.diffErrorAndNowRow,5,self.pastRecord[4])
                csv_data = self.read_csv(self.FILE_PATH)
                csv_data[self.row-self.diffErrorAndNowRow-1] = self.pastRecord

                self.write_csv(self.FILE_PATH, csv_data)
                self.diffErrorAndNowRow = 0

            else:
                if self.record[3] == None or self.record[3] == "" and self.error == False:
                    self.record[4] = time + ','
                    self.error = True
                else:
                    if self.error == False:
                        self.record[4] += time + ','
                    else:
                        self.record[4] += time + ', '
                self.error = not self.error
                self.worksheet.update_cell(self.row,5,self.record[4])

        except requests.exceptions.ReadTimeout as e:
            print(e)
            time.sleep(2)
            update_errortime()

        except Exception as e:
            print('An unexpected error occurred')
            print(e)
            sys.exit(1)

    def upload_count(self):
        try:
            self.worksheet = self.gc.open_by_key(self.SPREADSHEET_KEY).sheet1
            self.row = len(self.worksheet.get_all_values())
            self.record[5] = self.quantityPerHour
            self.worksheet.update_cell(self.row,6,self.quantityPerHour)

        except requests.exceptions.ReadTimeout as e:
            print(e)
            time.sleep(2)
            upload_count()

        except Exception as e:
            print('An unexpected error occurred')
            print(e)
            sys.exit(1)

    def update_cell(self,sheetName,rowNum,columnNum,value):
        #共有設定したスプレッドシートのシート1を開く
        self.worksheet = gc.open_by_key(self.SPREADSHEET_KEY).worksheet(sheetName)
        self.worksheet.update_cell(rowNum,columnNum,value)

    def read_csv(self,filename):
        f = open(filename, "r")
        csv_data = csv.reader(f)
        list = [ e for e in csv_data]
        f.close()
        return list

    def write_csv(self,filename, list):
        with open(filename, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(list)
        f.close()
