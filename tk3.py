import os,sys
import tkinter as tk
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

from calculator import Calculator

def get_bg_colors():
    cmap = plt.get_cmap("tab10")
    colors = []
    for rgb in np.array(cmap.colors)*255:
        rgb = rgb.astype(int)
        
        c = '#{:02x}{:02x}{:02x}'.format(*tuple(rgb))
        colors.append(c)
    return colors
        
class Application(tk.Frame):
    def __init__(self,root = None):
        super().__init__(root)
        self.pack()
        self.root = root
    
        style = ttk.Style(self.root)        
                
        self.root.geometry("1200x2200")
        self.root.title("Order Strategy !!")
        
        
        self.calc = Calculator()
        self.album_titles = self.calc.album_titles
        self.music_dicts = self.calc.music_dicts
        
        self.df,self.save_csv = self.calc.make_csv()
        
        kwargs={}        
        self.width = 1200
        self.cmaps = get_bg_colors()
        
        #描画の仕組みを作
        self.create_orders()
        
    def create_orders(self):
        
        f0 = ttk.Frame(self.root, padding=0,width=self.width, height=100); f0.pack()
        l = tk.Label(f0,text= self.save_csv, width=50,font=("MSゴシック", "20", "bold"));l.pack(side=LEFT)
        
        frames = []
        frames2 = []

        orders_s,orders_e,orders_n = {},{},{}
        orders_l = {}
        orders_style = {}
        orders_diff = {}
        
        f = ttk.Frame(self.root, padding=0,width=self.width, height=100); f.pack()
        
        tmp = self.music_dicts
        tmp["None"] = len(list(self.music_dicts))
        tmp["None"] = -1
        for m in tmp:
            l = tk.Label(f,text= m.replace("_","\n"), width=8,bg = str(self.cmaps[int(self.music_dicts[m])]))
            l.pack(side=LEFT)
        
        
        for title in self.calc.album_titles:
            f = ttk.Frame(self.root, padding=0,width=self.width, height=100); f.pack()
            frames.append(f)
            l = tk.Label(f,text=title,font=("MSゴシック", "20", "bold")); l.pack(side=TOP)
            # l = ttk.Label(f,text=title, padding=(5, 2)); l.pack(side=TOP)
            
            mucs = self.df[self.df["CompileTheme"]==title]["Style"].values
            nums = self.df[self.df["CompileTheme"]==title]["Quantity"].values
            orders = self.df[self.df["CompileTheme"]==title]["Order"].values
            
            orders_style[title] = []
            orders_s[title],orders_n[title] = [],[]
            orders_e[title] = []
            orders_l[title] = []
            orders_diff[title] = []
            
            cs = np.ones(31) * -1

            f = ttk.Frame(self.root, padding=0,width=self.width, height=100); f.pack()
            for m,n,o in zip(mucs,nums,orders):
                l = tk.Label(f,text= str(n), width=3,
                             bg = str(self.cmaps[int(self.music_dicts[m])])); l.pack(side=LEFT)
                l = tk.Label(f,text= "±0", width=3,bg="white"); l.pack(side=LEFT)
                # l = ttk.Label(f,text=n, padding=(5, 2),width=3); l.pack(side=LEFT)
                orders_diff[title].append(l)
                for on in o.split(","):
                    cs[int(on)] = int(self.music_dicts[m])
                
                orders_style[title].append(m)
                orders_n[title].append(n)
                
                os = StringVar() # 「ファイル参照」エントリーの作成
                os.set(str(o)) ; orders_s[title].append(os)
                e = ttk.Entry(f, textvariable=os, width=int(int(n)*2),style="My.TEntry"); e.pack(side=LEFT)
                orders_e[title].append(e)
                
            
            f2 = ttk.Frame(self.root, padding=0,width=self.width, height=100); f2.pack()
            for ii in range(1,30+1):
                # side = TOP if ii==1 else LEFT
                # l = ttk.Label(f2,text=str(ii),width=3,
                #               background = str(self.cmaps[int(ii%5)])); l.pack(side=LEFT)
                l = tk.Label(f2,text=str(ii),width=3,
                              bg = str(self.cmaps[int(cs[ii])])); l.pack(side=LEFT)

                orders_l[title].append(l)
                frames2.append(f2)
                
        
        self.orders_style = orders_style
        self.orders_s = orders_s
        self.orders_e = orders_e
        self.orders_l = orders_l
        self.orders_n = orders_n
        self.orders_diff = orders_diff
        self.frames2 = frames2
        
        def insert():
            for title in list(orders_s.keys()):   
                for os in orders_s[title]:
                    os.set(str(os.get()))
        
        f8 = ttk.Frame(self.root, padding=10, width=self.width, height=500); f8.pack()
        # l = ttk.Label(f8, text="実行", padding=(5, 2),font=("MSゴシック", "15", "bold")); l.pack(side=TOP)
        Button8 = ttk.Button(f8, text="実行",command=lambda: insert)
        Button10 = ttk.Button(f8, text="再計算",command=lambda: recalc(self));Button10.pack(side=LEFT)
        Button9 = ttk.Button(f8, text="修正",command=lambda: self.order_fix());Button9.pack(side=LEFT)
        Button11 = ttk.Button(f8, text="閉じる",command=lambda: self.quit());Button11.pack(side=LEFT)
    
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
    
    def quit(self):
        self.root.destroy()
        
    # 実行ボタン押下時の実行関数
    def order_fix(self):
        
        
        orders = []
        for idx,title in enumerate(list(self.orders_s.keys())):
            
            cs = np.ones(31) * -1
            for s,style in zip(self.orders_s[title],self.orders_style[title]):
                ids = list(map(int, set(s.get().split(","))))
                cs[ids] = int(self.music_dicts[style])
            
            """ update """
            for ii in range(1,30+1):
                self.orders_l[title][ii-1]['bg'] = str(self.cmaps[int(cs[ii])])

            """ check """
            for jj,(s,style,n) in enumerate(zip(self.orders_s[title],self.orders_style[title],self.orders_n[title])):
                ids = sorted(list(map(int, set(s.get().split(",")))))
                diff = len(ids) - n
                
                if diff == 0:
                    t,c = "±" + str(diff),"white"
                elif diff>0:
                    t,c = "+" + str(diff) , "red"
                else:
                    t,c = str(diff) , "blue"
                
                self.orders_diff[title][jj]['bg'] = c
                self.orders_diff[title][jj]['text'] = t
                
                orders_sorted = ",".join(list(map(str,ids)))
                s.set(orders_sorted)
                
                orders.append(orders_sorted)
        
        #update-orders
        self.df["Order"] = orders
        self.df.to_csv(self.save_csv,index=False)

            
def main():    
    root = Tk()
    app = Application(root)
    app.mainloop()

def recalc(self):
    self.root.destroy()
    root = Tk()
    app = Application(root)
    app.mainloop()
    

if __name__ == "__main__":
    
    main()

    
    
    
