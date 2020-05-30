# -*- coding: utf-8 -*-
import mcp3002
import datetime
import editSheet
import time

class LightSensor():

    def __init__(self,config):
        self.inProduction = False
        self.error = False

        self.config = config
        self.gs = editSheet.GoogleSheets()

    def start_measure(self):
        try:
            val_red_time = 0
            interval = 8
            inProduction = False
            error = False

            while True:
                previousHour = datetime.datetime.now().hour
                self.gs.create_new_record()

                #1時間製造を監視,　製造時間の記録, エラーの記録
                while True:
                    #電圧取得
                    val_green,voltage_green,val_red,voltage_red = mcp3002.measure()

                    #緑ランプの監視, 1つ出来たらその時間を記録
                    if val_green >= self.config['LIGHT_THRESHOLD'] and inProduction == False:
                        inProduction = True
                    elif inProduction == True and val_green < self.config['LIGHT_THRESHOLD']:
                        self.gs.update_createtime()
                        self.gs.upload_count()
                        inProduction = False

                    #赤ランプの監視, エラーが出たら記録

                    if val_red >= self.config['LIGHT_THRESHOLD'] and error == False:
                        error = True
                        self.gs.update_errortime()

                    elif val_red >= self.config['LIGHT_THRESHOLD'] and error == True:
                        val_red_time = 0

                    elif error == True and val_red < self.config['LIGHT_THRESHOLD']:
                        if val_red_time == 0:
                            val_red_time = time.time()
                        elif time.time() - val_red_time > interval:
                            val_red_time = 0
                            error = False
                            self.gs.update_errortime()


                    if previousHour < datetime.datetime.now().hour or (previousHour == 23 and datetime.datetime.now().hour == 0):
                        self.gs.upload_count()
                        print('break')
                        break

                    time.sleep(0.5)

        except KeyboardInterrupt:
            print("\nCtl+C")
        except Exception as e:
            print('error')
            print(str(e))
        finally:
            mcp3002.cleanup()
            print("\nexit program")
