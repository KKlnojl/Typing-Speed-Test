import pandas as pd
import openpyxl  # 需有此才能順利開啟xlsx
import tkinter as tk
import random


# ---------------------------------- 常數設定 ---------------------------------- #
WIN_WIDTH = 700
WIN_HEIGHT = 400
MAIN_FONT = ("arial", 13)
SCORE_V_FONT = ("arial", 18, "bold")
N = 50  # 起始顯示單字數
MIN = 2


# ---------------------------------- Class ---------------------------------- #
class TypingTest:
    def __init__(self, root, words, time=1, ):
        self.words = words
        self.time = time
        self.time_left = self.time * 60
        self.wpm = 0
        self.cpm = 0
        self.start_index = 0
        self.timer = None
        self.field_score = tk.Frame(root)
        self.label_cpm = tk.Label(self.field_score, text="Corrected CPM:", font=MAIN_FONT)
        self.label_cpm_v = tk.Label(self.field_score, text=self.cpm, font=SCORE_V_FONT)
        self.label_wpm = tk.Label(self.field_score, text="WPM:", font=MAIN_FONT)
        self.label_wpm_v = tk.Label(self.field_score, text=self.wpm, font=SCORE_V_FONT)
        self.label_time = tk.Label(self.field_score, text="Time left:", font=MAIN_FONT)
        self.label_time_left = tk.Label(self.field_score, text=self.time_left, font=SCORE_V_FONT)
        self.field_words = tk.Frame(root)
        self.label_instruction = tk.Label(self.field_words,
                                          text="When you are ready, please start entering any word,"
                                               " then the timer will start.",
                                          font=MAIN_FONT)
        self.btn_start = tk.Button(self.field_score, text="start", command=lambda:[self.count_down(),
                                                                                   self.callback()])  # 用labmda 繫結多個函式舉例
        self.btn_again = tk.Button(self.field_score, text="Try again", command=self.try_again)
        self.field_type = tk.Frame(root)
        self.show_words = tk.Text(self.field_type, wrap="word", height=10)  # wrap設定以單字換行
        self.typing_area = tk.Entry(self.field_type, width=30, font=MAIN_FONT)
        self.setting()

    def setting(self):
        self.field_score.pack(side="top", pady=20)
        self.label_cpm.pack(side="left", padx=5)
        self.label_cpm_v.pack(side="left", padx=5)
        self.label_wpm.pack(side="left", padx=5)
        self.label_wpm_v.pack(side="left", padx=5)
        self.label_time.pack(side="left", padx=5)
        self.label_time_left.pack(side="left", padx=5)
        self.btn_start.pack()
        self.btn_again.pack()
        self.field_words.pack(side="top")
        self.label_instruction.pack(side="top", padx=30)
        self.field_type.pack(side="bottom", pady=20)
        self.show_words.pack()
        self.show_init()
        self.typing_area.pack(side="bottom", pady=20)

    def count_down(self):
        self.btn_start.configure(state="disabled")  # 設定遊戲開始後，該按鈕就不能在被按一次
        if self.time_left >= 0:
            self.label_time_left.config(text=self.time_left)
            self.time_left -= 1
            self.timer = window.after(1000, self.count_down)
        if self.time_left == 0:
            wpm = round(self.wpm / (self.time * 60 - self.time_left) * 60 , 1)
            self.label_wpm_v.config(text=wpm)
            cpm = round(self.cpm / (self.time * 60 - self.time_left) * 60, 1)
            self.label_cpm_v.config(text=cpm)
            window.unbind("<space>")

    def try_again(self):
        # self.btn_start.configure(state="normal")
        window.after_cancel(self.timer)
        random.shuffle(self.words)
        self.time_left = self.time * 60
        self.label_time_left.config(text=self.time_left)
        self.btn_start.configure(state="active")
        self.cpm = 0
        self.wpm = 0
        self.label_wpm_v.config(text=self.wpm)
        self.label_cpm_v.config(text=self.cpm)
        self.start_index = 0
        self.show_words.delete("1.0", "end")  # 清除所有show_words內容，從第一個字元到最後一個字元
        self.clear_tag()
        self.show_init()
        window.unbind("<space>")

    def clear_tag(self):
        for tag in self.show_words.tag_names():
            self.show_words.tag_delete(tag)

    def callback(self):
        print("a")
        window.bind("<space>", self.check_typing)  # enter: <Return>

    def check_typing(self, event):  # event當鍵盤被觸發時，才執行此程式。
        enter = self.typing_area.get().rstrip()  # 去除空白
        self.check_word(enter, self.words[self.start_index], self.start_index)
        self.typing_area.delete(0, "end")  # 按下enter 後，清除已輸入的文字
        self.start_index += 1
        self.show_words.tag_config(f"tag{self.start_index}", background="yellow")

    def check_word(self, enter, word, index):
        # 計算wpm
        if enter == word:
            self.wpm += 1
            print(self.wpm)
            self.show_words.tag_config(f"tag{index}", background="green", foreground="white")
            wpm = round(self.wpm / (self.time * 60 - self.time_left) * 60, 1)
            self.label_wpm_v.config(text=wpm)
        else:
            self.show_words.tag_config(f"tag{index}", background="gray", foreground="white")
        # 計算cpm
        try:
            for i in range(len(word)):
                if enter[i] == word[i]:
                    self.cpm += 1
                    cpm = round(self.cpm / (self.time * 60 - self.time_left) * 60, 1)
                    self.label_cpm_v.config(text=cpm)
        except IndexError:
            pass
        # 每確認一次(每按一次space)，多新增一個單字。
        self.show_words.insert(tk.END, self.words[N+index])
        self.show_words.insert(tk.END, " ")
        self.show_words.see(tk.END)

    def show_init(self):
        for i in range(len(self.words[:N])):
            self.show_words.insert(tk.END, self.words[i], f"tag{i}")
            self.show_words.insert(tk.END," ")
            print(f"tag{self.words[i]}")
        self.show_words.tag_config(f"tag{self.start_index}", background="yellow")


# ---------------------------------- UI setting -Main Window ---------------------------------- #
# Preliminary: Create a list of words

vocabulary = pd.read_excel("words.xlsx")
list_words = vocabulary["Vocabulary"].to_list()
random.shuffle(list_words)

window = tk.Tk()
window.title("Typing Speed Test")
window.geometry(f"{WIN_WIDTH}x{WIN_HEIGHT}+0+0")
TypingTest(window, list_words, MIN)
window.mainloop()
