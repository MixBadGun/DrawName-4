import random
import xlrd
import os
from kivy.clock import Clock
#音效
from kivy.core.audio import SoundLoader
soundfile="sound/chose.wav"
soundfile2="sound/chose2.wav"
soundfile3="sound/passing.wav"
try:
    track = SoundLoader.load(filename=soundfile)
    track2 = SoundLoader.load(filename=soundfile2)
    track3 = SoundLoader.load(filename=soundfile3)
    track3.loop = True
except:
    print('错误：音效导入错误。')
#声明变量
times = 0
data = xlrd.open_workbook("list/name.xls")
table = data.sheets()[0].col_values(0)
using_times = 0
splitnum = 0
chosen_name = ''
chosen_list = []
namelist = []
timed = 0
the_bonus = None
#创建窗口
from kivy.app import App
from kivy.lang import Builder
from kivy.uix.screenmanager import Screen
from kivy.uix.label import Label
from kivy.uix.image import Image
from kivy.uix.anchorlayout import AnchorLayout
from kivy.config import Config
from kivy.core.window import Window
from kivy.uix.popup import Popup
#文本写入
import datetime
timenow = datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S')
if not os.path.exists("config"):
      os.makedirs("config")
#字体注册
from kivy.core.text import LabelBase
try:
    LabelBase.register(name='normal',fn_regular='./font/SourceHanSansCN-Normal.otf')
    LabelBase.register(name='bold',fn_regular='./font/SourceHanSansCN-Bold.otf')
    LabelBase.register(name='heavy',fn_regular='./font/SourceHanSansCN-Heavy.otf')
except:
    LabelBase.register(name='normal')
    LabelBase.register(name='bold')
    LabelBase.register(name='heavy')
#弹窗
#界面构建
Builder.load_string('''
<TheScreen>
    BoxLayout:
        Image:
            source: 'image/background.png'
            allow_stretch: True
            keep_ratio: False
    GridLayout:
        rows: 4
        Image:
            source: 'image/logo.png'
            size_hint_y: None
            height: 100
        BoxLayout:
            id: box
            anchor: 'top'
            padding: 10
            size_hint_y: None
            height: 200
        ScrollView:
            do_scroll_y: True
            StackLayout:
                id: scroll_list
                size_hint_y: None
                orientation: 'lr-tb'
                height: dp(scroll_list.minimum_height)
        BoxLayout:
            size_hint_y: None
            height: 100
            AnchorLayout:
                id: anc
                size_hint_y: None
                height: 100
                Button:
                    text: '抽取'
                    size_hint_x: None
                    size_hint_y: None
                    font_name: 'bold'
                    font_size: 25
                    background_color: 0.4,0.8,1,1
                    background_normal: ''
                    width: self.height*2
                    height: anc.height/1.5
                    on_release: root.start_random()
            AnchorLayout:
                size_hint_y: None
                height: 100
                Button:
                    text: '连抽'
                    size_hint_x: None
                    size_hint_y: None
                    font_name: 'bold'
                    font_size: 25
                    background_color: 0.4,0.8,1,1
                    background_normal: ''
                    width: self.height*2
                    height: anc.height/1.5
                    on_release: root.bonus_random()
''')

class TheScreen(Screen):
    #文本写入
    def text_create(self,msg):
        global timenow
        desktop_path = "config/"
        file = open(desktop_path+ timenow +'.txt', 'a')
        file.writelines(msg+'\n')
    #弹窗
    def popups(self,x,y):
        popup = Popup(title=x,content=Label(text=y),size_hint=(None, None),size=(400, 400),on_dismiss=os._exit(0))
        popup.open()
    #拆分名字类
    def split_name(self, name):
        try:
            namelist = []
            num = len(name)
            temp = ""
            for x in range(0,num):
                temp += name[x]
                namelist.append(temp)
            return namelist
        except:
            self.popups('错误','拆分错误，请检查表格。')
    #划分区域
    def section(self):
        try:
            global table
            global splitnum
            global namelist
            try:
                data = xlrd.open_workbook("list/name.xls")
            except:
                self.popups('错误','表格读取失败，请检查 list/name.xls 文件是否存在。')
            table = data.sheets()[0].col_values(0)
            numlist=[]
            #遍历名称以查询
            for i in table:
                num = len(i)
                numlist.append(num)
                temp = ""
                for x in range(0,num):
                    temp += i[x]
                    namelist.append(temp)
            splitnum = max(numlist)
            for i in range(0,splitnum):
                box = self.ids['box']
                layout = AnchorLayout()
                box.add_widget(layout)
                image = Image(source='image/block.png')
                layout.add_widget(image)
            Config.set('graphics', 'resizable', 0)
            Window.size = (200*splitnum, 650)
        except:
            self.popups('错误','表格读取失败。')
    #随机抽取
    def random_name(self):
        global table
        for i in range(0,10):
            random.shuffle(table)
        item = random.choice(table)
        return item
    #显示幸运儿
    def add_name(self, name):
        global times
        global txtfile
        times += 1
        scroll_list = self.ids['scroll_list']
        lab = Label(size_hint_y=None,height=50,font_name='bold',font_size=30,text="第 "+str(times)+" 个幸运儿："+name)
        self.text_create("第 "+str(times)+" 个幸运儿：\t"+name)
        scroll_list.add_widget(lab)
    #动画
    def add_scrolling(self,index):
        box = self.ids['box']
        layout = box.children[index]
        image = Image(source='image/chose1.gif',anim_delay=1.0 / random.randint(30,60))
        layout.add_widget(image)
    def add_title(self,name,index):
        box = self.ids['box']
        layout = box.children[index]
        if layout.width < layout.height:
            title_size = layout.width*0.7
        else:
            title_size = layout.height*0.7
        label = Label(text=name,font_size=int(title_size),font_name='heavy',color="#000000",pos=(-5,15))
        layout.add_widget(label)
    def remove_scrolling(self,index):
        box = self.ids['box']
        layout = box.children[index]
        layout.clear_widgets()
        image = Image(source='image/block.png')
        layout.add_widget(image)
    #主判断
    def start_random(self):
        global splitnum
        global using_times
        global chosen_name
        global chosen_list
        global namelist
        passing = 0
        if using_times != 0:
            track2.play()
            if using_times > len(chosen_name) or using_times == splitnum:
                bullon = True
            else:
                bullon = namelist.count(chosen_list[using_times-1]) == 1
            if bullon:
                for i in range(using_times,splitnum+1):
                    self.remove_scrolling(splitnum-i)
                    try:
                        self.add_title(chosen_name[i-1],splitnum-i)
                    except:
                        self.add_title("",splitnum-i)
                using_times = 0
                self.add_name(chosen_name)
                track.play()
                track3.stop()
                passing = 1
            else:
                self.remove_scrolling(splitnum-using_times)
                self.add_title(chosen_name[using_times-1],splitnum-using_times)
                using_times += 1
        if using_times == 0 and passing == 0:
            for i in range(0,splitnum):
                self.add_scrolling(i)
            track3.play()
            chosen_name = self.random_name()
            chosen_list = self.split_name(chosen_name)
            using_times += 1
        passing = 0
    #连抽
    def bonus(self,dt):
        global chosen_name
        chosen_name = self.random_name()
        self.add_name(chosen_name)
        track.play()
    def bonus_random(self):
        global timed
        global the_bonus
        for i in range(0,splitnum):
            self.add_scrolling(i)
        if timed == 0:
            track3.play()
            the_bonus = Clock.schedule_interval(self.bonus, 0.1)
            timed = 1
        else:
            the_bonus.cancel()
            timed = 0
            track3.stop()
            track.play()
            for i in range(0,splitnum):
                self.remove_scrolling(i)
                self.add_title('止',i)
    def __init__(self, **kwargs):
        super(TheScreen, self).__init__(**kwargs)
        self.section()

class TheApp(App):
    def build(self):
        self.title = '坏枪随机器 4.0'
        return TheScreen()
#运行窗体
TheApp().run()