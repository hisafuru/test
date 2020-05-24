# -*- coding: utf-8 -*-
import pydrive
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import os,sys,yaml,time

class GoogleDriveUpload():
    def __init__(self):

        print('Initializing google api...')

        #再帰関数の上限を変更
        sys.setrecursionlimit(300)

        #oauth認証
        try:
            gauth = GoogleAuth()
            gauth.CommandLineAuth()
            self.drive = GoogleDrive(gauth)
        except Exception as e:
            print('Oauth authentication failed')
            print(e)
            sys.exit(1)

        #settings.yamlの読み込み
        try:
            with open('paperless_setting.yaml') as file:
                self.config = yaml.safe_load(file.read())
                file.close()
        except Exception as e:
            print('Error reading paperless _ setting.yaml')
            print(e)
            print('Please check if paperless _ setting.yml exists in the setting folder.')
            sys.exit(1)

        print('Finish Initializing.')

    #機械ごとのcsvファイルのアップロード
    def upload_file(self):
        stand_by_sec = 5

        print('Uploading file [' + self.config['MACHINE_NAME'] + '.csv]')

        try:
            f = self.drive.CreateFile({'id':self.config["SPREADSHEET_KEY"]})
            f.SetContentFile(self.config['MACHINE_NAME'] + '.csv')
            f.Upload()
        except pydrive.files.FileNotUploadedError as e:
            print('Connection timed out')
            print(e)

            #stand_by_sec秒待ってやり直し
            time.sleep(stand_by_sec)
            upload_file()
        except Exception as e:
            print('Other errors')
            print(e)
            sys.exit(1)

        print('Finish uploading [' + self.config['MACHINE_NAME'] + '.csv].')

    #任意のファイルのアップロード
    def upload_anyfile(self,file_path,id):
        print('Uploading file [' + file_path + ']')

        try:
            f = self.drive.CreateFile({'id':id})
            f.SetContentFile(file_name)
            f.Upload()
        except pydrive.files.FileNotUploadedError as e:
            print('Connection timed out')
            print(e)

            #stand_by_sec秒待ってやり直し
            time.sleep(stand_by_sec)
            upload_anyfile(file_path=file_path,id=id)
        except Exception as e:
            print('Other errors')
            print(e)
            sys.exit(1)

        print('Finish uploading.')
