import tkinter as tk
from tkinter import colorchooser,filedialog,font
import makeMusicFormat_app

class songFormat:
    def __init__(self,master):
        self.master=master
        self.master.title("Make Music Format")
        self.master.minsize(400,300)

        #excel生成用に渡す引数
        self.musicName=None
        self.barLow=None
        self.dir=None
        self.section_entry_list=[]
        self.bar_entry_list=[]
        self.col16_list=[]
        #表示する色の初期化
        self.col16=["tmp","#80ff80"]

        #リストの初期化
        def clear_list():
            self.section_entry_list=[]
            self.bar_entry_list=[]
            self.col16_list=[]

        ###event処理達
        def musicName_event():
            #題名と一行の小節数を取得
            self.musicName=self.musicName_entry.get()
            self.barLow=self.barLow_entry.get()

        def color_event():
            #カラーピッカーの起動
            self.col16=colorchooser.askcolor()
            #ボタンの色、テキストの変更
            self.color_button.config(bg=self.col16[1]) 
            self.color_button.config(text=self.col16[1])
        
        def plus_event():
            #セクション名、小節数、色の取得
            a=self.section_entry.get()
            b=self.bar_entry.get()
            c=self.col16[1]
            #openpyxlだと#がいらない
            #"#000000"だと"0"になる
            if(c!="0"):
                c=c.split("#")
                c=c[1]
            
            #入力された文字の削除
            self.section_entry.delete(0,tk.END)
            self.bar_entry.delete(0,tk.END)
            txt=a+","+b+","+c
            #リストボックスに追加
            self.listbox.insert(tk.END,txt)

        def minus_event():
            #選択物を取得して削除
            selectedIndex = tk.ACTIVE
            self.listbox.delete(selectedIndex)

        def generate_event():
            clear_list()
            index=self.listbox.size()
            for i in range(index):
                sep=self.listbox.get(i)
                tmp=sep.split(",")
                self.section_entry_list.append(tmp[0])
                self.bar_entry_list.append(tmp[1])
                self.col16_list.append(tmp[2])
            
            musicName_event()
            
            #excelシートの生成
            m=makeMusicFormat_app.msf(self.musicName,self.barLow,self.section_entry_list,self.bar_entry_list,self.col16_list,self.dir)
            m.wright()

        def select_file():
            file_type = [("Excel", "*.xlsx")]
            selected_file = filedialog.asksaveasfilename(filetypes=file_type,defaultextension="xlsx")
            if selected_file:
                # for file in selected_files:
                #     self.listbox_files.insert(tk.END, file)
                self.dir=selected_file
                self.dir_label.config(text=selected_file)

        #gridの行列はこれで取得できる(デバッグ用)
        # def callback(event):
        #     info = event.widget.grid_info()


        #widget
        #label
        self.name_label=tk.Label(text="題名")
        self.barLow_label=tk.Label(text="一行何小節か")
        self.section_label=tk.Label(text="セクション名")
        self.bar_label=tk.Label(text="小節数")
        self.color_label=tk.Label(text="色")
        self.dir_label=tk.Label(text="保存ファイルを選択(.xlsx)")
        #entry
        self.ent_font=font.Font(size=15)
        self.musicName_entry=tk.Entry(width=20,font=self.ent_font)
        self.barLow_entry=tk.Entry(width=20,font=self.ent_font)
        self.section_entry=tk.Entry(width=20,font=self.ent_font)
        self.bar_entry=tk.Entry(width=20,font=self.ent_font)
        self.color_entry=tk.Entry(width=20,font=self.ent_font)
        #button
        self.color_button=tk.Button(text="クリックして変更",command=color_event,bg=self.col16[1],relief=tk.RAISED)
        self.plus_button=tk.Button(text="+",width=2,height=1,command=plus_event)
        self.minas_button=tk.Button(text="-",width=2,height=1,command=minus_event)
        self.generate_button=tk.Button(text="generate",command=generate_event)
        self.selectFile_button=tk.Button(text="参照",command=select_file)
        #scroll,listbox
        self.scr=tk.Scrollbar(orient="vertical")
        self.list_font=font.Font(size=20)
        self.listbox=tk.Listbox(font=self.list_font)
        self.listbox["yscrollcommand"] = self.scr.set
        self.scr["command"]=self.listbox.yview
        #gridで配置
        self.name_label.grid(row=0,column=0)
        self.musicName_entry.grid(row=1,column=0,sticky="nsew",padx=2)
        self.barLow_label.grid(row=0,column=1)
        self.barLow_entry.grid(row=1,column=1,sticky="nsew",padx=2)
        self.section_label.grid(row=3,column=0)
        self.section_entry.grid(row=4,column=0,sticky="nsew",padx=2)
        self.bar_label.grid(row=3,column=1)
        self.bar_entry.grid(row=4,column=1,sticky="nsew",padx=2)
        self.color_label.grid(row=3,column=2)
        self.color_button.grid(row=4,column=2,sticky="nsew",padx=2)
        self.listbox.grid(row=5,columnspan=3,sticky="nsew",padx=3,pady=3)
        self.scr.grid(row=5,column=3,sticky="wns")
        self.generate_button.grid(row=6,column=0,sticky="w")
        self.selectFile_button.grid(row=6,column=0,sticky="e")
        self.dir_label.grid(row=6,column=1,sticky="w")
        self.plus_button.grid(row=6,column=2,sticky="e")
        self.minas_button.grid(row=6,column=3,sticky="w")

        self.master.grid_columnconfigure(0,weight=1,minsize=100)    
        self.master.grid_columnconfigure(1,weight=1,minsize=100)    
        self.master.grid_columnconfigure(2,weight=1,minsize=100)    
        self.master.grid_columnconfigure(3,weight=1,minsize=100)    
        self.master.grid_rowconfigure(5,weight=1,minsize=100)

        # self.master.bind_all("<1>",callback) #gridの行列取得


if __name__ == '__main__':
    root = tk.Tk()
    make = songFormat(root)
    root.mainloop()       
        