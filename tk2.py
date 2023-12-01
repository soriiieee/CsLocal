import os,sys
import tkinter
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog

import openpyxl
import numpy as np
from scipy import optimize
import copy

from datetime import datetime
import pandas as pd
import matplotlib.pyplot as plt
cmap = plt.get_cmap("tab10")

from calculator import Calculator

def filedialog1(excel_name):
    fTyp = [("エクセルブック", "*.xlsx")]
    iFile = os.path.abspath(os.path.dirname(__file__))
    iFilePath = filedialog.askopenfilename(filetypes = fTyp, initialdir = iFile)
    excel_name.set(iFilePath)

def filedialog2(need_name):
    fTyp = [("エクセルブック", "*.xlsx")]
    iFile = os.path.abspath(os.path.dirname(__file__))
    iFilePath = filedialog.askopenfilename(filetypes = fTyp, initialdir = iFile)    
    need_name.set(iFilePath)

        
class Application(tkinter.Frame):
    def __init__(self,root = None):
        super().__init__(root)
        self.pack()
        self.root = root
    
        style = ttk.Style(self.root)        
        style.configure("My.TEntry",
            foreground="black",
            background="white")
        style.map("My.TEntry",
            foreground=[("readonly", "blue")],
            fieldbackground=[("readonly", "red")])
        
        
        self.root.geometry("800x800")
        self.root.title("Compile Strategy !!")
        
        
        self.calc = Calculator()
        self.df_album_music = self.calc.get_music_df()
        self.df_needs_title = self.calc.get_needs()
        self.remains = self.calc.Remains_dicts
        
        
        kwargs={}        
        self.width = 800
        
        #描画の仕組みを作る
        self.create_music_table()
        self.create_titles()
        self.create_title_results()
        self.create_music_results()
        
    
    def create_music_table(self):
        
        f3 = ttk.Frame(self.root, padding=10,width=self.width, height=500); f3.pack()
        l = ttk.Label(f3, text="利用曲の比率", padding=(5, 2),font=("MSゴシック", "15", "bold")); l.pack(side=TOP) # 「ファイル参照」ラベルの作成
        # f3.grid(row=2, column=1, sticky=W)
        my_tree = ttk.Treeview(f3)
        list_music = [ m.replace("_","\n") for m in self.df_album_music.columns]
        list_album = self.df_album_music.index.tolist()
        
        for i,music in enumerate(["Title"]+ list_music):
            if i==0:
                l = ttk.Label(f3, text=music, padding=(5, 2),font=("MSゴシック", "15", "bold"),width=8) # 「ファイル参照」ラベルの作成
            else:
                l = ttk.Label(f3, text=music, padding=(5, 2),width=8) 
            # side = TOP if i==0 else LEFT ; l.pack(side=side)
            l.pack(side=LEFT)
        l = ttk.Label(f3, text="   合計  ", padding=(5, 2),width=8,font=("MSゴシック", "10", "bold")); l.pack(side=LEFT) #ファイル参照」ラベルの作成
        
        ss = [ [] for _ in range(len(list_album))]
        es = [ [] for _ in range(len(list_album))]
        fs = []
        for i , album in enumerate(list_album):
            f = ttk.Frame(self.root, padding=0,width=self.width, height=500); f.pack()
            fs.append(f)
            for j , music in enumerate(list_music):
                s = StringVar(value="1") # 「ファイル参照」エントリーの作成
                ss[i].append(s)
        
        sum_ss,sum_ee = [],[]
        for i , album in enumerate(self.calc.album_titles): 
            l = ttk.Label(fs[i], text=album, padding=(5, 2),width=8) # 「ファイル参照」ラベルの作成
            l.pack(side=LEFT)
            sum0=0
            for j , music in enumerate(list_music):            
                num = self.df_album_music.iloc[i,j]
                e = ttk.Entry(fs[i], textvariable=ss[i][j], width=8)
                e.pack(side=LEFT)
                ss[i][j].set(str(num))
                es[i].append(e)
                sum0 += num
            
            ##合計の追加
            sums = StringVar() # 「ファイル参照」エントリーの作成
            sums.set(str(sum0))
            sum_ss.append(sums)
            l = ttk.Label(fs[i],text="=", padding=(5, 2),width=2); l.pack(side=LEFT)
            e = ttk.Entry(fs[i], textvariable=sums, width=6,style="My.TEntry");e.pack(side=LEFT)
            sum_ee.append(e)
            
            def insert():
                for i , album in enumerate(self.calc.album_titles): 
                    for j , music in enumerate(list_music):    
                        ss[i][j].set(str(self.df_album_music.iloc[i,j]))
        
        self.ss = ss
        self.es = es
        self.sum_ss = sum_ss
        self.sum_ee = sum_ee
        # button = Button(f2, text="Insert", width=10, command=insert)
        
        ### 必要アルバムの数を描画する
    def create_titles(self):
        
        f4 = ttk.Frame(self.root, padding=0,width= self.width, height=100); f4.pack()
        f42 = ttk.Frame(self.root, padding=0,width= self.width, height=500); f42.pack()
        l = ttk.Label(f4, text="必要タイトル数", padding=(5, 2),font=("MSゴシック", "15", "bold")); l.pack(side=TOP) # 「ファイル参照」ラベルの作成
        
        needs = self.df_needs_title.to_dict()["Num2"]
        sum_title = 0
        
        require_title_s , require_title_e = [] , []
        
        
        for i , album in enumerate(self.calc.album_titles): 
            l = ttk.Label(f4, text=album, padding=(5, 2),width=8); l.pack(side=LEFT) # 「ファイル参照」ラベルの作成
            s = StringVar() # 「ファイル参照」エントリーの作成
            require_title_s.append(s)
            num = needs[album]
            e = ttk.Entry(f42, textvariable=require_title_s[i], width=8)
            require_title_s[i].set(str(num))
            require_title_e.append(e)
            sum_title += num
            e.pack(side=LEFT)
        
        self.require_title_s,self.require_title_e = require_title_s , require_title_e #必要なタイトル数
        l = ttk.Label(f4, text="   合計  ", padding=(5, 2),width=8); l.pack(side=LEFT) # 「ファイル参照」ラベルの作成
        req_title = StringVar() # 「ファイル参照」エントリーの作成
        req_title.set(str(sum_title))
        l = ttk.Label(f42,text="=", padding=(5, 2),width=2); l.pack(side=LEFT)
        e = ttk.Entry(f42, textvariable=req_title, width=6);e.pack(side=LEFT)
        
        self.req_title_sum_s = req_title
        self.req_title_sum_e = e
        
        # def insert4():
        #     for i , album in enumerate(self.calc.album_titles):   
        #         s4s[i].set(str(needs[i]))
        # button = Button(f2, text="Insert", width=10, command=insert4) 

    ###結果出力画面
    def create_title_results(self):
        
        f6 = ttk.Frame(self.root, padding=0,width=self.width, height=100); f6.pack()
        f62 = ttk.Frame(self.root, padding=0,width=self.width, height=500); f62.pack()
        f63 = ttk.Frame(self.root, padding=0,width=self.width, height=500); f63.pack()
        l = ttk.Label(f6,text="解析結果(作成数/不足数)", padding=(5, 2),width=20); l.pack(side=TOP) # 「ファイル参照」ラベルの作成
        # s6s = []
        
        result_title_s , result_title_e = [] , []
        diff_title_s , diff_title_e = [] , []
        sum_make = 0
        for i , album in enumerate(self.calc.album_titles): 
            # l = ttk.Label(f6, text=album, padding=(5, 2)); l.pack(side=LEFT) # 「ファイル参照」ラベルの作成
            s,s2 = StringVar(),StringVar() # 「ファイル参照」エントリーの作成
            result_title_s.append(s)
            diff_title_s.append(s2)
            needs = self.df_needs_title.to_dict()["Num2"]
            num = needs[album]
            
            e = ttk.Entry(f62, textvariable=result_title_s[i], width=8)
            e2 = ttk.Entry(f63, textvariable=diff_title_s[i], width=8,style="My.TEntry")
            sum_make += 0
            result_title_e.append(e)
            diff_title_e.append(e2)
            e.pack(side=LEFT); e2.pack(side=LEFT)

        ## 変数で固定
        self.result_title_s , self.result_title_e   = result_title_s,result_title_e #必要なアルバムタイトル数(s)
        self.diff_title_s  , self.diff_title_e   = diff_title_s , diff_title_e
        # l = ttk.Label(f6, text=" = 合計", padding=(5, 2)); l.pack(side=LEFT) # 「ファイル参照」ラベルの作成
        make_title,diff_title = StringVar(),StringVar() # 「ファイル参照」エントリーの作成
        make_title.set(str(sum_make))
        diff_title.set(str(0))
        # s6s.append(s)
        
        l = ttk.Label(f62,text="=", padding=(5, 2),width=2); l.pack(side=LEFT)
        e2 = ttk.Entry(f62, textvariable=make_title, width=6);e2.pack(side=LEFT)
        l = ttk.Label(f63,text="=", padding=(5, 2),width=2); l.pack(side=LEFT)
        e3 = ttk.Entry(f63, textvariable=diff_title, width=6);e3.pack(side=LEFT)

        self.result_title_sum_s,self.result_title_sum_e = make_title,e2
        self.diff_title_sum_s,self.diff_title_sum_e = diff_title,e3
        
        ###利用可能曲数
        f5 = ttk.Frame(self.root, padding=0,width=self.width, height=100); f5.pack()
        f52 = ttk.Frame(self.root, padding=0,width=self.width, height=500); f52.pack()
        l = ttk.Label(f5, text="利用可能midi曲数", padding=(5, 2),font=("MSゴシック", "15", "bold"),width=8); l.pack(side=TOP) # 「ファイル参照」ラベルの作成
        s5s = []
        
        require_midi_s , require_midi_e = [] , []
        list_music = [ m.replace("_","\n") for m in self.df_album_music.columns]
        # needs = self.remains
        for i , mc in enumerate(list_music): 
            l = ttk.Label(f5, text=mc, padding=(5, 2)); l.pack(side=LEFT) # 「ファイル参照」ラベルの作成
            s = StringVar() # 「ファイル参照」エントリーの作成
            require_midi_s.append(s)
            num = self.remains[mc.replace("\n","_")]
            e = ttk.Entry(f52, textvariable=require_midi_s[i], width=8)
            require_midi_s[i].set(str(num))
            e.pack(side=LEFT)
            require_midi_e.append(e)
            
        def insert5():
            for i , mc in enumerate(self.calc.album_titles):   
                s5s[i].set(str(self.remains[mc.replace("\n","_")]))
        # button = Button(f2, text="Insert", width=10, command=insert5)
        self.require_midi_s , self.require_midi_e = require_midi_s , require_midi_e 
        
    def create_music_results(self):
        
        f7 = ttk.Frame(self.root, padding=0,width=self.width, height=100); f7.pack()
        f72 = ttk.Frame(self.root, padding=0,width=self.width, height=500); f72.pack()
        f73 = ttk.Frame(self.root, padding=0,width=self.width, height=500); f73.pack()
        l = ttk.Label(f7,text="解析結果(利用曲/残曲)", padding=(5, 2)); l.pack(side=TOP) # 「ファイル参照」ラベルの作成
        s7s = []
        
        result_midi_s , result_midi_e = [],[]
        diff_midi_s , diff_midi_e = [],[]
        list_music = [ m.replace("_","\n") for m in self.df_album_music.columns]
        for i,mc in enumerate(list_music):
            # l = ttk.Label(f7, text=album, padding=(5, 2)); l.pack(side=LEFT) # 「ファイル参照」ラベルの作成
            s,s2 = StringVar(),StringVar() # 「ファイル参照」エントリーの作成
            result_midi_s.append(s)
            diff_midi_s.append(s2)
            e = ttk.Entry(f72, textvariable=result_midi_s[i], width=8)
            e2 = ttk.Entry(f73, textvariable=diff_midi_s[i], width=8)
            e.pack(side=LEFT);e2.pack(side=LEFT)
            result_midi_e.append(e)
            diff_midi_e.append(e2)

        self.result_midi_s , self.result_midi_e = result_midi_s , result_midi_e
        self.diff_midi_s , self.diff_midi_e = diff_midi_s , diff_midi_e
        
        f8 = ttk.Frame(self.root, padding=10, width=self.width, height=500); f8.pack()
        l = ttk.Label(f8, text="実行", padding=(5, 2),font=("MSゴシック", "15", "bold")); l.pack(side=TOP)
        Button8 = ttk.Button(f8, text="実行",
                            command=lambda: self.conductMain(1)); Button8.pack(side=LEFT)

    
    def color(self,entry):
        # NOTE: ウィジェットは複数の状態を同時に持つことができる為、state メソッドで readonly 状態を on/off 
        if float(entry.get()) > 0:
            entry.state(["readonly"])
        else:
            entry.state(["!readonly"])
            
    def color_num30(self,entry):
        # NOTE: ウィジェットは複数の状態を同時に持つことができる為、state メソッドで readonly 状態を on/off 
        if float(entry.get()) != 30:
            entry.state(["readonly"])
        else:
            entry.state(["!readonly"])    
    
    # 実行ボタン押下時の実行関数
    def conductMain(self,num):
        
        values = np.zeros((len(self.ss),len(self.ss[0])))
        for i,(album,sums,sume) in enumerate(zip(self.ss,self.sum_ss,self.sum_ee)):
            
            for j,mc in enumerate(album):
                values[i,j] = int(mc.get())
            self.sum_ss[i].set(str(int(np.sum(values[i,:]))))
            self.color_num30(sume)

        titles_number2 = np.zeros(len(self.require_title_s))
        for i,num in enumerate(self.require_title_s):
            titles_number2[i] = int(num.get())
                    
        available_songs2 = np.zeros(len(self.require_midi_s))
        for i,song in enumerate(self.require_midi_s):
            available_songs2[i] = int(song.get())

            
        # if not excel_name:
        #     messagebox.showerror("error", "パスの指定がありません")

        # print("calc start !")
        # print(values,available_songs2,titles_number2)
        albums ,use_musics, remains_musics = self.calc.calc_compile_number(values,available_songs2,titles_number2)  
        # shotage = titles_number2 - albums
        
        ## 描画
        for i,(rs,re,ds,de,qs,qe) in enumerate(zip(
            self.result_title_s,self.result_title_e,
            self.diff_title_s,self.diff_title_e,
            self.require_title_s,self.require_title_e)):
            
            rs.set(str(albums[i]))
            ds.set(str( int(qs.get())- albums[i]))
            self.color(de)
        
        self.req_title_sum_s.set(str(int(np.sum(titles_number2))))
        self.result_title_sum_s.set(str(int(np.sum(albums))))
        self.diff_title_sum_s.set(str(int(np.sum(titles_number2) - int(np.sum(albums)))))
        # req_title.set(str(int(np.sum(titles_number2))))
        # make_title.set(str(np.sum(albums)))
            
        for i,(ms,me,ds,de,qs,qe) in enumerate(zip(
            self.result_midi_s , self.result_midi_e,
            self.diff_midi_s , self.diff_midi_e,
            self.require_midi_s , self.require_midi_e,
            )):
            ms.set(str(int( int(qs.get()) - remains_musics[i])))
            ds.set(str(int(remains_musics[i])))
            



def main():    
    ###calculater
    # rootの作成
    root = Tk()
    app = Application(root)
    app.mainloop()


if __name__ == "__main__":
    
    main()

    
    
    
