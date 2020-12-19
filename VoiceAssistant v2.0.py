import os
import time
import sys
import threading as thread
from functools import wraps
try:
    import webbrowser as wb
    import speech_recognition as sr
    from fuzzywuzzy import fuzz
    import pyautogui as pg
    from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
except ImportError:
    print("Установите необходимые модули!")
    exit()


def clamp(n, minv=0, maxv=1):
    if n > maxv:
        n = maxv
    elif n < minv:
        n = minv
    return n


class System:

    @staticmethod
    def SetVolume(value):
        print("Меняю параметры звука..")
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            volume = session._ctl.QueryInterface(ISimpleAudioVolume)
            volume.SetMasterVolume(clamp(value), None)

    @staticmethod
    def GetVolume(value=None):
        print("Получайте параметры звука..")
        sessions = AudioUtilities.GetAllSessions()
        v = 0
        volume = sessions[0]._ctl.QueryInterface(ISimpleAudioVolume)
        v = volume.GetMasterVolume()
        return v

    @staticmethod
    def GetBrightness():
        pass

    @staticmethod
    def SetBrightness():
        pass

    @staticmethod
    def Sleep():
        pass


class Commands:

    chrome = ["google chrome", "chrome", "хром"]
    notepad = ["notepad plus plus"]
    discord = ["discord", "дискорд"]
    cmd = ["cmd", "командная строка"]
    word = ["word", "текстовый редактор"]
    excel = ["excel", "x", "таблицы"]
    power_point = ["power point", "презентации"]
    photoshop = ["adobe photoshop 2021", "photoshop", "фотошоп", "фото"]
    pycharm = ["pycharm", "чары", "пайчарм"]
    unity = ["unity", "юнити"]

    time = ["время", "текущее время", "времени", "сейчас час", "сейчас", "который час"]
    close = ["стоп", "прекрати", "отключись", "замолчи", "молчи", "выключись"]

    youtube = ["youtube", "ютьюб", "ютуб"]
    ivi = ["ivi", "иви", "киви", "фильмы", "qiwi"]
    translate = ["переведи", "переводчик", "как будет "]
    fireplace = ["камин", "огонь", "релакс", "костер", "звуки", "успокой     меня"]
    sound_lower = ["потише", "тише", "убавь громкость"]
    sound_higher = ["погромче", "громче", "увеличь громкость", "не слышу"]

    c = [chrome, youtube, notepad, discord, time, close, cmd, word,
         excel, power_point, photoshop, ivi, translate, pycharm, unity, fireplace,
         sound_higher, sound_lower]


class Reaction:

    command = None
    @staticmethod
    def PrepareRawData(text):
        """text - input parameter as string is raw data"""
        to_remove = ["open", "start", "launch", "get",
                     "открой", "запусти", "какой", "какое", "сколько", "включи", "сделай",
                     "скажи", "пожалуйста"]; """list of words that must be removed for 
                                                        precision of reaction"""
        t = text.split(' ')
        for i in to_remove:
            if i in t:
                t.remove(i)

        l = Commands.c

        costs = list([max(list([fuzz.ratio(" ".join(t).lower().strip(), inner) for inner in i])) for i in l])
        print(costs, l)
        try:
            if max(costs) > 75:
                objs = l[costs.index(max(costs))]
                print("[log]:", objs[0], "распознано")
                return costs.index(max(costs))  # costs.index(max(costs))
        except ValueError:
            return " ".join(t).lower().strip()

    @staticmethod
    def Respond(text):

        def open(name=""):
            def inner(function):
                @wraps(function)
                def i():
                    print(f"Открываю {name}..")
                    function()

                return i

            return inner

        @open("youtube")
        def youtube():
            path = r"https://www.youtube.com/"
            wb.open(path)
            return True

        @open("chrome")
        def chrome():
            path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
            os.startfile(path)
            return True

        @open("notepad++")
        def notepad():
            path = r"C:\Program Files\Notepad++\notepad++.exe"
            os.startfile(path)
            return True

        @open("discord")
        def discord():
            path = r"C:\Users\olegb\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Discord Inc\Discord.lnk"
            os.startfile(path)
            return True

        @open("microsoft word")
        def word():
            path = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk"
            os.startfile(path)
            return True

        @open("microsoft excel")
        def excel():
            path = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk"
            os.startfile(path)
            return True

        @open("microsoft power point")
        def powerpoint():
            path = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PowerPoint.lnk"
            os.startfile(path)
            return True

        @open("cmd")
        def cmd():
            path = r"C:\Users\olegb\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\System Tools\Command Prompt.lnk"
            os.startfile(path)
            return True

        @open("photoshop")
        def photoshop():
            path = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Adobe Photoshop 2021.lnk"
            os.startfile(path)
            return True

        @open("ivi")
        def ivi():
            wb.open(r"https://www.ivi.ru/profile")
            return True

        @open("translator")
        def translator():
            wb.open(r"https://translate.google.com/")
            return True

        @open("pycharm")
        def pycharm():
            path = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\JetBrains\PyCharm Community Edition 2020.2.3.lnk"
            os.startfile(path)
            return True

        @open("unity")
        def unity():
            path = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Unity 2019.3.13f1 (64-bit)\Unity.lnk"
            os.startfile(path)
            return True

        @open("fireplace")
        def fireplace():
            print("open")
            thread.main_thread()
            wb.open(r"https://www.youtube.com/watch?v=Ux8xAuQBdkk&ab_channel=TheSilentWatcher")
            time.sleep(3.4)
            System.SetVolume(System.GetVolume() + 0.5)
            pg.press('f')

        def sound_higher():
            print("делаю..")
            System.SetVolume(1)  # It hasn't been finished because of deadline

        def now():
            print("Сейчас", time.strftime("%H:%M:%S"))
            return True

        def close():
            print("Выключаюсь..")
            exit()
            sys.exit()

        print(text)
        open(text)
        reactions = {0 : chrome,
                     1 : youtube,
                     2 : notepad,
                     3 : discord,
                     4 : now,
                     5 : close,
                     6 : cmd,
                     7 : word,
                     8 : excel,
                     9 : powerpoint,
                     10 : photoshop,
                     11 : ivi,
                     12 : translator,
                     13 : pycharm,
                     14 : unity,
                     15 : fireplace,
                     16 : sound_higher}

        try:
            return reactions[text]
        except KeyError:
            return None


class Recorder:

    def __init__(self):

        self.microphone_id = 1  # external one is 0
        self.r = sr.Recognizer()
        self.r.energy_threshold = 180  # 380
        self.listen = False

    def record_audio(self):
        self.listen = True
        with sr.Microphone(device_index=self.microphone_id) as source:

            self.r.phrase_threshold = 0.45  # 0.45
            self.r.adjust_for_ambient_noise(source)
            self.r.pause_threshold = 0.2  # 0.4f
            self.r.non_speaking_duration = 0.18  # 0.18
            print("Слушаю..")
            audio = self.r.listen(source)

            print("Услышал.")

            print("Распознаю..")
            self.listen = False
            voice_data = ''
            try:
                voice_data = self.r.recognize_google(audio, language="ru-RU").lower()
                print("Улышал " + voice_data)
            except sr.UnknownValueError:
                print('Не понял')
            except sr.RequestError:
                print("Невозможно распознать голос. Проверьте подключение к Интернету!")
            Reaction.command = Reaction.Respond(Reaction.PrepareRawData(voice_data))
            return


def start():
    pg.FAILSAFE = False
    r = Recorder()

    while True:
        if not r.listen:  # and thread.active_count() < 4
            # print("New thread")
            task1 = thread.Thread(target=r.record_audio, kwargs={})
            task1.start()
            # print(thread.active_count())
        if Reaction.command is not None:
            # print(Reaction.command)
            print("выполняю..")
            Reaction.command()
            Reaction.command = None


start()