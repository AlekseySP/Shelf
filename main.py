from kivy.core.window import Window
from kivy.uix.boxlayout import BoxLayout
from kivy.clock import Clock
from kivy.core.audio import SoundLoader
from kivy.properties import ObjectProperty
from kivymd.uix.filemanager import MDFileManager
from kivymd.app import MDApp
from kivymd.toast import toast
from openpyxl import load_workbook
import pymysql

Host = "remotemysql.com"
Port = 3306
User = "3ALraEDF4c"
Password = "L3pBenmNFz"

class MainScreen(BoxLayout):
    con = pymysql.connect(host=Host,port=Port,user=User, passwd=Password,db=User)
    cur = con.cursor()
    cur.execute("SELECT * FROM Articles")
    rows = cur.fetchall()
    adres_dict = {}
    for row in rows:
        ar = str(row[0])
        ad = str(row[1])
        adres_dict[ar] = ad
    loadfile = ObjectProperty(None)
    text_input = ObjectProperty(None)
    place_label = ObjectProperty(None)

    def show_keyboard(self, event):
        self.text_input.focus = True

    def show_place(self):
        self.beep2 = SoundLoader.load('b1.wav')
        self.beep1 = SoundLoader.load('b2.wav')
        self.actAdr = self.place_label.text
        self.adress = self.adres_dict
        self.inp = self.text_input.text.upper()
        if self.inp:
            if ";" in self.inp:
                self.out = list(self.inp.split(";"))[-1]
            else:
                self.out = list(self.inp.split())[-1]   
            if self.out in self.adress:
                self.place = self.adress[self.out]
                self.text_input.helper_text = self.out
                self.place_label.text=self.place
                if self.actAdr == self.place:
                	self.beep1.play()
                else:
                	self.beep2.play()
                Clock.schedule_once(self.show_keyboard)
            else:
                self.place_label.text= "Неизвестно"
                self.text_input.helper_text = ""
                self.beep2.play()
                Clock.schedule_once(self.show_keyboard)
        else:
            self.place_label.text="Нет ввода"
            self.text_input.helper_text = ""
            self.beep2.play()
        self.text_input.text = ""
        Clock.schedule_once(self.show_keyboard)



class ShelfApp(MDApp):
    def __init__(self, **kwargs):
        self.title = "Shelf 3"
        self.theme_cls.theme_style = "Dark"
        self.theme_cls.primary_palette = "Teal"
        super().__init__(**kwargs)
        Window.bind(on_keyboard=self.events)
        self.manager_open = False
        self.file_manager = MDFileManager(
            exit_manager=self.exit_manager,
            select_path=self.select_path,
            previous=False,
            )
        self.file_manager.ext = ['.xlsx']

    def build(self):
        return MainScreen()

    def file_manager_open(self):
        self.file_manager.show("/")#('/data/media/0')
        self.manager_open = True
        # output manager to the screen

    def server_update(self, path):
        self.path = path
        self.delete_old_data()
        self.upload_data(self.path)

    def dict_from_xl(self, path):
        self.path = path
        self.wb = load_workbook(filename = self.path)
        self.ws = self.wb.worksheets[0]
        self.adres_dict = {}
        for i in range (1, len(self.ws["A"])):
            self.k = str(self.ws["A" + str(i)].value)
            self.v = str(self.ws["B" + str(i)].value)
            self.adres_dict.update([(self.k, self.v)])
        print("dictionary passed")
        return self.adres_dict

    def upload_data(self, path):
        self.path = path
        self.d = self.dict_from_xl(self.path)
        self.arts  = list(self.d.keys())
        self.con = pymysql.connect(host=Host,
        port=Port, user=User, passwd=Password,
        db=User)
        self.cur = self.con.cursor()
        # sql = "UPDATE Articles set adress = %s where article = %s"
        self.sql = "INSERT INTO Articles (Article, Adress) VALUES (%s, %s)"
       # for i in arts:
        #   art = i
         #  adr = d[i]
         #  val = (art, adr)
        self.val = [(i, self.d[i]) for i in self.arts]
        self.cur.executemany(self.sql, self.val)
        self.con.commit()
        self.cur.close()
        self.con.close()
        print("Data uploading complete.")
        

    def delete_old_data(self):
        self.con = pymysql.connect(host=Host,
        port=Port, user=User, passwd=Password,
        db=User)
        self.cur = self.con.cursor()
        self.Delete_all_rows = """truncate table Articles """
        self.cur.execute(self.Delete_all_rows)
        self.con.commit()
        print("All Record Deleted successfully")
        self.cur.close()
        self.con.close()


    def select_path(self, path):
        '''It will be called when you click on the file name
        or the catalog selection button.
        :type path: str;
        :param path: path to the selected directory or file;
        '''
        self.exit_manager()
        self.server_update(path)
        toast('Сервер обновлён')
    def exit_manager(self, *args):
        '''Called when the user reaches the root of the directory tree.'''
        self.manager_open = False
        self.file_manager.close()
    def events(self, instance, keyboard, keycode, text, modifiers):
        '''Called when buttons are pressed on the mobile device.'''
        if keyboard in (1001, 27):
            if self.manager_open:
                self.file_manager.back()
        return True


if __name__ == "__main__":
    ShelfApp().run()