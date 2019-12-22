# Author: fzzjj2008
# Date: 2019.12.22

import os
import time
import tkinter as tk
import tkinter.messagebox
import tkinter.colorchooser
import win32con
import win32gui
import win32clipboard
import keyboard
import configparser
from io import BytesIO
from PIL import Image, ImageGrab


class Config:

    screen_w = 0
    screen_h = 0
    iniconf = configparser.ConfigParser()
    ini_path = os.path.abspath('./config.ini')
    scnpng_path = os.path.abspath('./temp.png')
    scnpng_dark_path = os.path.abspath('./temp_dark.png')

    @staticmethod
    def recover_default_setting():
        os.remove(Config.ini_path)

    @staticmethod
    def get_save_path():
        print('save crop file')
        tm_str = time.strftime("%Y%m%d%H%M%S", time.localtime())
        if os.path.exists('./capture') == False:
            os.mkdir('./capture')
        return os.path.abspath('./capture/%s.png' % tm_str)

    @staticmethod
    def get_hotkey():
        try:
            Config.iniconf.read(Config.ini_path)
            return Config.iniconf.get("CONF", "hotkey")
        except:
            return 'Ctrl + Alt + A'

    @staticmethod
    def set_hotkey(hotkey):
        if Config.iniconf.has_section("CONF") == False:
            Config.iniconf.add_section("CONF")
        Config.iniconf.set("CONF", "hotkey", hotkey)
        with open(Config.ini_path, "w+") as f:
            Config.iniconf.write(f)

    @staticmethod
    def get_hide_win():
        try:
            Config.iniconf.read(Config.ini_path)
            return Config.iniconf.get("CONF", "hide_win")
        except:
            return 'False'

    @staticmethod
    def set_hide_win(hide_win):
        if Config.iniconf.has_section("CONF") == False:
            Config.iniconf.add_section("CONF")
        Config.iniconf.set("CONF", "hide_win", hide_win)
        with open(Config.ini_path, "w+") as f:
            Config.iniconf.write(f)

    @staticmethod
    def get_outline_color():
        try:
            Config.iniconf.read(Config.ini_path)
            return Config.iniconf.get("CONF", "outline_color")
        except:
            return '#FFFFFF'

    @staticmethod
    def set_outline_color(color):
        if Config.iniconf.has_section("CONF") == False:
            Config.iniconf.add_section("CONF")
        Config.iniconf.set("CONF", "outline_color", color)
        with open(Config.ini_path, "w+") as f:
            Config.iniconf.write(f)


class Capture:

    @staticmethod
    def capture_all_screen():
        print('capture all screen')
        im = ImageGrab.grab()
        im.save(Config.scnpng_path)
        im_dark = im.point(lambda p: p * 0.5)
        im_dark.save(Config.scnpng_dark_path)
        im.close()
        im_dark.close()

    @staticmethod
    def send_to_clipboard(clip_type, data):
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(clip_type, data)
        win32clipboard.CloseClipboard()

    @staticmethod
    def save_image(left, top, right, bottom):
        print('save image')
        image_save = Image.open(Config.scnpng_path).crop(
            (left, top, right, bottom))
        # save crop image to save_path
        image_save.save(Config.get_save_path())

    @staticmethod
    def crop_image(left, top, right, bottom):
        print('crop image')
        img_crop = Image.open(Config.scnpng_path).crop(
            (left, top, right, bottom))
        # save crop image to clipboard
        output = BytesIO()
        img_crop.convert("RGB")
        img_crop.save(output, "BMP")
        data = output.getvalue()[14:]   # discard 14 bytes of bmp head
        output.close()
        Capture.send_to_clipboard(win32clipboard.CF_DIB, data)

    @staticmethod
    def clear_scnpng():
        print('clear temp png')
        os.remove(Config.scnpng_path)
        os.remove(Config.scnpng_dark_path)


class CapturePicTool:

    tool_w = 160
    tool_h = 30
    left = 0
    top = 0
    right = 0
    bottom = 0

    def __init__(self, parent):
        self.parent = parent
        # toplevel init
        self.toplevel = tk.Toplevel(parent, bg='white')
        self.toplevel.overrideredirect(True)
        self.toplevel.wm_attributes("-topmost", 1)
        # size label init
        self.size_label_str = tk.StringVar()
        self.size_label = tk.Label(self.toplevel,
                                   textvariable=self.size_label_str,
                                   fg='black')
        self.size_label.configure(background='white')
        self.size_label.place(x=4, y=4)
        # tool button init
        self.save_img = tk.PhotoImage(file="res/save_24x24.png")
        self.btn_save = tk.Button(self.toplevel,
                                  image=self.save_img,
                                  relief="flat",
                                  bd=1,
                                  cursor='hand2',
                                  command=self.on_save_handle)
        self.btn_save.configure(background='white')
        self.btn_save.place(x=70, y=2)
        self.cancel_img = tk.PhotoImage(file="res/cancel_24x24.png")
        self.btn_cancel = tk.Button(self.toplevel,
                                    image=self.cancel_img,
                                    relief="flat",
                                    bd=1,
                                    cursor='hand2',
                                    command=self.on_cancel_handle)
        self.btn_cancel.configure(background='white')
        self.btn_cancel.place(x=100, y=2)
        self.ok_img = tk.PhotoImage(file="res/ok_24x24.png")
        self.btn_ok = tk.Button(self.toplevel,
                                image=self.ok_img,
                                relief="flat",
                                bd=1,
                                cursor='hand2',
                                command=self.on_ok_handle)
        self.btn_ok.configure(background='white')
        self.btn_ok.place(x=130, y=2)

    def update(self, left, top, right, bottom):
        # record position
        self.left = left
        self.top = top
        self.right = right
        self.bottom = bottom
        # adjust position, default at bottom-left corner
        if (self.left + self.tool_w <= Config.screen_w) and (self.bottom + self.tool_h <= Config.screen_h):
            # bottom-left
            tool_x = self.left
            tool_y = self.bottom
        elif (self.right - self.tool_w >= 0) and (self.bottom + self.tool_h <= Config.screen_h):
            # bottom-right
            tool_x = self.right - self.tool_w
            tool_y = self.bottom
        elif (self.left + self.tool_w <= Config.screen_w) and (self.top - self.tool_h >= 0):
            # top-left
            tool_x = self.left
            tool_y = self.top - self.tool_h
        elif (self.right - self.tool_w >= 0) and (self.top - self.tool_h >= 0):
            # top-right
            tool_x = self.right - self.tool_w
            tool_y = self.top - self.tool_h
        elif (self.right + self.tool_w <= Config.screen_w):
            # right
            tool_x = self.right
            tool_y = self.top
        elif (self.left - self.tool_w >= 0):
            # left
            tool_x = self.left - self.tool_w
            tool_y = self.top
        else:
            tool_x = self.left
            tool_y = self.top
        self.toplevel.geometry('%dx%d+%d+%d' %
                               (self.tool_w, self.tool_h, tool_x, tool_y))
        # show label
        size_str = '%dx%d' % (self.right-self.left, self.bottom-self.top)
        self.size_label_str.set(size_str)

    def on_ok_handle(self):
        print('click ok tool btn')
        Capture.crop_image(self.left, self.top, self.right, self.bottom)
        # close window
        self.parent.destroy()

    def on_cancel_handle(self):
        print('click cancel tool btn')
        # close window
        self.parent.destroy()

    def on_save_handle(self):
        print('click save tool btn')
        Capture.save_image(self.left, self.top, self.right, self.bottom)
        # close window
        self.parent.destroy()

    def destroy(self):
        if self.toplevel:
            self.toplevel.destroy()
            self.toplevel = None


class CapturePicWindow:

    left = 0
    top = 0
    right = 0
    bottom = 0
    toplevel = None

    def __init__(self, parent):
        # toplevel init
        self.toplevel = tk.Toplevel(parent, bg='white', cursor='fleur')
        self.toplevel.wm_attributes("-topmost", 1)
        self.toplevel.overrideredirect(True)

    def show(self, left, top, right, bottom):
        # record position
        self.left = left
        self.top = top
        self.right = right
        self.bottom = bottom
        width = abs(right - left)
        height = abs(bottom - top)

        # change toplevel position
        self.toplevel.geometry('%dx%d+%d+%d' % (width, height, left, top))

        # load scnpng
        self.scnpng = tk.PhotoImage(file=Config.scnpng_path)
        # show canvas
        self.canvas = tk.Canvas(self.toplevel,
                                bg='white',
                                width=width,
                                height=height,
                                highlightthickness=0)
        self.canvas.create_image(-left, -top, anchor='nw', image=self.scnpng)
        self.canvas.pack(fill=tk.BOTH, expand=tk.YES)   # fill parent component

    def destroy(self):
        if self.toplevel:
            self.toplevel.destroy()
            self.toplevel = None


class CaptureCanvasWindow:

    select = False
    toplevel = None
    rect = None
    pic = None
    tool = None

    def __init__(self, parent):
        self.window = parent

    def show_scnpng(self):
        self.canvas.delete(self.scnpng_dark)
        self.canvas.create_image(
            Config.screen_w//2, Config.screen_h//2, image=self.scnpng)
        self.canvas.update()

    def show_dark_scnpng(self):
        self.canvas.delete(self.scnpng)
        self.canvas.create_image(
            Config.screen_w//2, Config.screen_h//2, image=self.scnpng_dark)
        self.canvas.update()

    def on_mouse_left_down(self, event):
        if self.select:
            return
        self.canvas.delete(self.rect)
        self.left = event.x
        self.top = event.y
        self.select = True
        # show scneen png
        self.show_scnpng()
        # destroy CapturePicWindow
        if self.pic:
            self.pic.destroy()
            self.pic = None
        if self.tool:
            self.tool.destroy()
            self.tool = None

    def on_mouse_move(self, event):
        if not self.select:
            return
        self.canvas.delete(self.rect)
        self.rect = self.canvas.create_rectangle(
            self.left, self.top, event.x, event.y, outline=Config.get_outline_color())

    def on_mouse_left_up(self, event):
        if not self.select:
            return
        self.select = False
        self.right = event.x
        self.bottom = event.y
        # do nothing if start position == end position
        if (self.left == self.right or self.top == self.bottom):
            return
        # ensure start position < end position
        if self.left > self.right:
            (self.left, self.right) = (self.right, self.left)
        if self.top > self.bottom:
            (self.top, self.bottom) = (self.bottom, self.top)
        print('left:%d,right:%d,top:%d,bottom:%d' %
              (self.left, self.right, self.top, self.bottom))
        # show dark pic
        self.show_dark_scnpng()
        # show CapturePicWindow
        self.pic = CapturePicWindow(self.toplevel)
        self.pic.show(self.left, self.top, self.right, self.bottom)
        # show CapturePicTool
        self.tool = CapturePicTool(self.toplevel)
        self.tool.update(self.left, self.top, self.right, self.bottom)

    def on_cancel_capture(self, event):
        if self.pic:
            self.pic.destroy()
            self.pic = None
        if self.tool:
            self.tool.destroy()
            self.tool = None
        if self.toplevel:
            self.toplevel.destroy()
            self.toplevel = None

    def show(self):
        # load screen png
        self.scnpng = tk.PhotoImage(file=Config.scnpng_path)
        self.scnpng_dark = tk.PhotoImage(file=Config.scnpng_dark_path)

        # create canvas
        self.toplevel = tk.Toplevel(self.window,
                                    width=Config.screen_w,
                                    height=Config.screen_h,
                                    cursor='crosshair')
        self.toplevel.resizable(False, False)
        self.toplevel.overrideredirect(True)
        self.toplevel.wm_attributes("-topmost", 1)
        self.canvas = tk.Canvas(self.toplevel,
                                bg='white',
                                width=Config.screen_w,
                                height=Config.screen_h,
                                highlightthickness=0)
        self.show_scnpng()
        self.canvas.bind('<Button-1>', self.on_mouse_left_down)
        self.canvas.bind('<B1-Motion>', self.on_mouse_move)
        self.canvas.bind('<ButtonRelease-1>', self.on_mouse_left_up)
        self.canvas.bind('<ButtonRelease-3>', self.on_cancel_capture)
        self.canvas.pack(fill=tk.BOTH, expand=tk.YES)   # fill parent component


class CaptureSetDialog(tk.Toplevel):
    def __init__(self, window):
        super().__init__()
        # window init
        self.window = window
        self.title('Setting')
        self.grab_set()
        self.geometry('360x170')
        self.resizable(False, False)
        self.configure(background='white')
        # hotkey
        self.label_hotkey = tk.Label(self, text='Hotkey setting: ')
        self.label_hotkey.configure(background='white')
        self.label_hotkey.place(x=20, y=10)
        hotkey_text = tk.StringVar()
        hotkey_text.set(Config.get_hotkey())
        self.entry_hotkey = tk.Entry(self,
                                     textvariable=hotkey_text,
                                     state='disable')
        self.entry_hotkey.place(x=200, y=10)
        # capture hide window
        self.label_hd_win = tk.Label(self, text='Hide window when capture: ')
        self.label_hd_win.configure(background='white')
        self.label_hd_win.place(x=20, y=50)
        self.btn_hd_win = tk.Checkbutton(self, bd=0, cursor='hand2',
                                         relief='raised',
                                         command=self.hide_win_setting)
        self.btn_hd_win.configure(background='white')
        self.btn_hd_win.place(x=200, y=50)
        # outline color choose
        self.label_recover = tk.Label(self, text='Select outline color: ')
        self.label_recover.configure(background='white')
        self.label_recover.place(x=20, y=90)
        self.btn_recover = tk.Button(self, text='select',
                                     bd=1,
                                     cursor='hand2',
                                     command=self.select_color_setting)
        self.btn_recover.place(x=200, y=85)
        # recover all settings
        self.label_recover = tk.Label(self, text='Recover all settings: ')
        self.label_recover.configure(background='white')
        self.label_recover.place(x=20, y=130)
        self.btn_recover = tk.Button(self, text='recover',
                                     bd=1,
                                     cursor='hand2',
                                     command=self.recover_setting)
        self.btn_recover.place(x=200, y=125)
        # load all data
        if Config.get_hide_win() == 'True':
            self.btn_hd_win.select()
        else:
            self.btn_hd_win.deselect()

    def select_color_setting(self):
        print('press color setting')
        (rgb, hx) = tk.colorchooser.askcolor(
            parent=self, title='Color', color=Config.get_outline_color())
        print(hx)
        Config.set_outline_color(hx)

    def hide_win_setting(self):
        print('press recover setting')
        if Config.get_hide_win() == 'True':
            Config.set_hide_win('False')
        else:
            Config.set_hide_win('True')

    def recover_setting(self):
        print('press recover setting')
        ret = tk.messagebox.askquestion('Tips', 'Recovery all settings?')
        if ret == 'yes':
            print('recover all settings')
            Config.recover_default_setting()


class CaptureMainWindow:

    def __init__(self, window):
        # window init
        self.window = window
        self.window.title('Capture')
        self.window.geometry('300x100')
        self.window.resizable(False, False)
        self.window.configure(background='white')

        # canvas window init
        self.canvas = CaptureCanvasWindow(self.window)

        # label init
        self.label_str = tk.StringVar()
        self.label_str.set('Press <%s> to capture' % Config.get_hotkey())
        self.label = tk.Label(self.window,
                              textvariable=self.label_str,
                              height=4)
        self.label.configure(background='white')
        self.label.pack()

        # btn_capture init
        self.capture_img = tk.PhotoImage(file="res/capture_24x24.png")
        self.btn_capture = tk.Button(self.window,
                                     image=self.capture_img,
                                     relief="flat",
                                     bd=1,
                                     cursor='hand2',
                                     command=self.on_capture)
        self.btn_capture.configure(background='white')
        self.btn_capture.place(x=120, y=60)

        # btn_setting init
        self.setting_img = tk.PhotoImage(file="res/setting_24x24.png")
        self.btn_setting = tk.Button(self.window,
                                     image=self.setting_img,
                                     relief="flat",
                                     bd=1,
                                     cursor='hand2',
                                     command=self.on_setting)
        self.btn_setting.configure(background='white')
        self.btn_setting.place(x=150, y=60)

    def on_capture(self):
        # hide window
        if Config.get_hide_win() == 'True':
            self.window.withdraw()
            time.sleep(0.2)

        Capture.capture_all_screen()
        self.canvas.show()
        self.window.wait_window(self.canvas.toplevel)
        Capture.clear_scnpng()     # clear temp scn png

        # recovery window
        if Config.get_hide_win() == 'True':
            self.window.deiconify()

    def on_setting(self):
        self.set_dialog = CaptureSetDialog(self.window)
        self.window.wait_window(self.set_dialog)

    def mainloop(self):
        self.window.mainloop()


if __name__ == "__main__":
    window = tk.Tk()
    Config.screen_w = window.winfo_screenwidth()
    Config.screen_h = window.winfo_screenheight()
    capture_win = CaptureMainWindow(window)
    keyboard.add_hotkey(Config.get_hotkey(), capture_win.on_capture)
    capture_win.mainloop()
