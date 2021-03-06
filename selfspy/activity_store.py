# Copyright 2012 David Fendrich

# This file is part of Selfspy

# Selfspy is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.

# Selfspy is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.

# You should have received a copy of the GNU General Public License
# along with Selfspy.  If not, see <http://www.gnu.org/licenses/>.
import threading
import time

import PyQt5
import pythoncom
import win32clipboard

import sys
import ctypes
from ctypes import wintypes

# Win32 imports
import win32api
import win32gui
import win32con
import win32file
import win32pipe
import win32process
import win32security
from win32event import CreateEvent, SetEvent, WaitForSingleObject
from win32event import MsgWaitForMultipleObjects, WAIT_OBJECT_0
from win32event import WAIT_TIMEOUT, INFINITE, QS_ALLINPUT, QS_POSTMESSAGE
from win32event import QS_ALLEVENTS

import pywintypes

byref = ctypes.byref
user32 = ctypes.windll.user32

from pynput import keyboard

from PySide import QtGui, QtCore

from datetime import datetime

NOW = datetime.now

import sqlalchemy

import platform

if platform.system() == 'Darwin':
    from selfspy import sniff_cocoa as sniffer
elif platform.system() == 'Windows':
    from selfspy import sniff_win as sniffer
else:
    from selfspy import sniff_x as sniffer

from selfspy import models
from selfspy.models import Process, Window, Geometry, Click, Keys, Clipboard

from PyQt5 import QtWidgets

from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *

from Tkinter import *

from pyHook import cpyHook

SKIP_MODIFIERS = {"", "Shift_L", "Super_L", "Alt_L", "Super_R", "Shift_R",
                  "[65027]"}  # [65027] is AltGr in X for some ungodly reason.

SCROLL_BUTTONS = {4, 5, 6, 7}
SCROLL_COOLOFF = 10  # seconds

app = QtWidgets.QApplication(["", ""])
clipboard = app.clipboard()

mimeData = clipboard.mimeData()


class Test(Text):
    def __init__(self, master, **kw):
        Text.__init__(self, master, **kw)
        self.bind('<Control-c>', self.copy)
        self.bind('<Control-x>', self.cut)
        self.bind('<Control-v>', self.paste)

    def copy(self, event=None):
        self.clipboard_clear()
        text = self.get("sel.first", "sel.last")
        self.clipboard_append(text)

    def cut(self, event):
        self.copy()
        self.delete("sel.first", "sel.last")

    def paste(self, event):
        print("PASTE")
        text = self.selection_get(selection='CLIPBOARD')
        self.insert('insert', text)


def test():
    r = Tk()
    t = Test(r)
    t.pack(fill='both', expand=1)
    r.update()


class ClipboardWatcher(threading.Thread):
    def __init__(self, callback, pause=1.):
        super(ClipboardWatcher, self).__init__()
        self._callback = callback
        self._pause = pause
        self._stopping = False

    def run(self):
        recent_value = ""
        while not self._stopping:
            tmp_value = clipboard.text()
            if tmp_value != recent_value:
                recent_value = tmp_value
                self._callback()
            time.sleep(self._pause)

    def stop(self):
        self._stopping = True


class Display:
    def __init__(self):
        self.proc_id = None
        self.win_id = None
        self.geo_id = None


class KeyPress:
    def __init__(self, key, time, is_repeat):
        self.key = key
        self.time = time
        self.is_repeat = is_repeat


class ActivityStore:
    def __init__(self, db_name, encrypter=None, store_text=True, repeat_char=True):
        self.session_maker = models.initialize(db_name)

        models.ENCRYPTER = encrypter

        self.store_text = store_text
        self.repeat_char = repeat_char
        self.curtext = u""

        self.key_presses = []
        self.mouse_path = []

        self.current_window = Display()

        self.last_scroll = {button: 0 for button in SCROLL_BUTTONS}

        self.last_key_time = time.time()
        self.last_commit = time.time()

        self.started = NOW()
        self.last_screen_change = None

        self.last_clipboard_change = None

        self.lastKeyWasCtrl = False

        self.was_ctrl_c = False
        self.was_ctrl_v = False
        self.was_ctrl_x = False


        # Start QT App in new thread, for receiving clipboard changes
        t = threading.Thread(target=self.qtApp)
        t.daemon = True
        t.start()

        # message_map = {
        #     win32con.WM_RENDERFORMAT: self.testHook
        # }
        # hwnd = win32gui.GetActiveWindow()
        #
        # msgShellHook = ctypes.windll.user32.RegisterWindowMessageA("SHELLHOOK")
        #
        # hHook = ctypes.windll.user32.SetWindowsHookExA(win32con.WM_RENDERFORMAT)

        # wc = win32gui.WNDCLASS()
        # wc.lpfnWndProc = message_map
        # wc.lpszClassName = 'MyWindowClass'
        # hinst = wc.hInstance = win32api.GetModuleHandle(None)
        # classAtom = win32gui.RegisterClass(wc)
        # self.hwnd = win32gui.CreateWindow(
        #     classAtom,
        #     "win32gui test",
        #     0,
        #     0,
        #     0,
        #     win32con.CW_USEDEFAULT,
        #     win32con.CW_USEDEFAULT,
        #     0,
        #     0,
        #     hinst,
        #     None
        # )
        # print self.hwnd

        # StopEvent = win32event.CreateEvent(None, 0, 0, None)
        #
        # while 1:
        #     rc = win32event.MsgWaitForMultipleObjects(
        #         StopEvent,
        #         0,  # Wait for all = false, so it waits for anyone
        #         200,  # (or win32event.INFINITE)
        #         win32event.QS_ALLEVENTS)  # Accepts all input

        # root = Tk()
        # root.bind('<Control-v>', self.testPaste())

    def testHook(self):
        print "hooked"

    def qtApp(self):
        app = QtGui.QApplication(sys.argv)
        clipboard = app.clipboard()

        # Connect clipboard changes to got_changed_clipboard
        clipboard.dataChanged.connect(self.got_changed_clipboard)

        app.exec_()

    def on_press(key):
        try:
            print('alphanumeric key {0} pressed'.format(
                key.char))
        except AttributeError:
            print('special key {0} pressed'.format(
                key))

    def on_release(key):
        print('{0} released'.format(
            key))
        if key == keyboard.Key.esc:
            # Stop listener
            return False

    def testPaste(self):
        print ("paste")

    def trycommit(self):

        self.last_commit = time.time()
        for _ in range(1000):
            try:
                self.session.commit()
                break
            except sqlalchemy.exc.OperationalError:
                print("exception")
                time.sleep(1)
            except:
                self.session.rollback()

    def run(self):
        self.session = self.session_maker()
        self.sniffer = sniffer.Sniffer()
        self.sniffer.screen_hook = self.got_screen_change
        self.sniffer.key_hook = self.got_key
        self.sniffer.mouse_button_hook = self.got_mouse_click
        self.sniffer.mouse_move_hook = self.got_mouse_move

        # self.sniffer.clipboard_hook = self.got_changed_clipboard

        # watcher = ClipboardWatcher(self.got_changed_clipboard, 1.)
        # watcher.start()


        # root = Tk()
        # root.bind('<Control-v>', self.testPaste())
        # root.update()

        # self.getWindowsMessages()
        pythoncom.PumpMessages()

        self.sniffer.run()

    def getWindowsMessages(self):
        # # Create a hidden window that will receive messages for things
        # # like adding new handles to wait on or quitting the thread.
        # # I use the Button class because I'm too lazy to register my own.
        # theWindow = win32gui.CreateWindow("Button",  # lpClassName
        #                                   "",  # lpWindowName
        #                                   0,  # dwStyle
        #                                   0,  # x
        #                                   0,  # y
        #                                   0,  # width
        #                                   0,  # height
        #                                   0,  # parent
        #                                   0,  # menu
        #                                   0,  # hInstance
        #                                   None  # lParam
        #                                   )
        # # list of process handles to wait for
        # handles = []
        #
        # while True:
        #     #val = MsgWaitForMultipleObjects(handles, 0, INFINITE, QS_POSTMESSAGE | QS_ALLEVENTS)
        #
        #     status, msg = win32gui.PeekMessage(None,
        #                                            0,
        #                                            0,
        #                                            win32con.PM_REMOVE)
        #     while status != 0:
        #             if msg[1] == win32con.WM_RENDERFORMAT:
        #                 print "IS WM_RENDERFORMAT"
        #             else:
        #                 # Drop all other messages, since we receive all messages, not
        #                 # just WM_NEW_PHANDLE and WM_CLOSE_THREAD.
        #                 pass
        #
        #             status, msg = win32gui.PeekMessage(
        #                 None,
        #                 0,
        #                 0,
        #                 win32con.PM_REMOVE)






        try:
            msg = wintypes.MSG()
            while win32gui.PeekMessage(None, 0, 0, win32con.PM_REMOVE) != 0:

                #win32gui.TranslateMessage(byref(msg))
                #win32gui.DispatchMessage(byref(msg))
                #print msg.message

                if msg.message == win32con.WM_RENDERFORMAT:
                    print "was WM_RENDERFORMAT"

                # user32.TranslateMessage(byref(msg))
                # user32.DispatchMessageA(byref(msg))

        finally:
            print "finally"

    def testChanged(self):
        print("TEST")

    def register_clipboard_formats(self):
        """
        Here we extract all of the formats from the qmimedata object and
        register them.  This will give us an ID for each type.
        This will return a dictionary with the ID's and descriptions for
        each mime type.
        """
        mime_formats = mimeData.formats()
        format_dictionary = dict()
        for format in mime_formats:
            id = win32clipboard.RegisterClipboardFormat(str(format))
            format_dictionary[id] = format

        return format_dictionary

    def get_clipboard_formats(self):
        format_dictionary = dict()
        if mimeData.hasText():
            id = win32con.CF_TEXT
            data = str(mimeData.text())
            format_dictionary[id] = data

            id = win32con.CF_UNICODETEXT
            data = unicode(mimeData.text())
            format_dictionary[id] = data

        if mimeData.hasUrls():
            files = []
            for url in mimeData.urls():
                file_name = url.toLocalFile()
                if file_name:
                    files.append(str(file_name))

            if files:
                id = win32con.CF_HDROP
                format_dictionary[id] = files

        return format_dictionary

    def got_changed_clipboard(self):
        'Receives Clipboard Data'

        print ("GOT CLIPBOARD")

        # Trying to implement Delayed Rendering
        # win32clipboard.OpenClipboard()
        # win32clipboard.EmptyClipboard()
        # try:
        #     win32clipboard.SetClipboardData(win32clipboard.CF_TEXT, None)
        # finally:
        #     win32clipboard.CloseClipboard()

        # user32.OpenClipboard(0)
        # user32.EmptyClipboard()
        # user32.SetClipboardData(1, None)
        # user32.CloseClipboard()

        #self.getWindowsMessages()

        # msg = wintypes.MSG()
        # print msg.message

        # win32clipboard.SetClipboardData(0, "test")


        # Store clipboard data
        self.store_clipboard()





    def got_screen_change(self, process_name, window_name, win_x, win_y, win_width, win_height):
        """Receives a screen change and stores any changes.
        If the process or window has changed it will also store any queued pressed keys.
        Keyword arguments:
        process_name -- the name of the process running the current window
        window_name -- the name of the window
        win_x -- the x position of the window
        win_y -- the y position of the window
        win_width -- the width of the window
        win_height -- the height of the window"""

        # skip the event if same arguments as last time are passed

        args = [process_name, window_name, win_x, win_y, win_width, win_height]
        if self.last_screen_change == args:
            return

        self.last_screen_change = args

        cur_process = self.session.query(
            Process
        ).filter_by(
            name=process_name
        ).scalar()
        if not cur_process:
            cur_process = Process(process_name)
            self.session.add(cur_process)

        cur_geometry = self.session.query(
            Geometry
        ).filter_by(
            xpos=win_x,
            ypos=win_y,
            width=win_width,
            height=win_height
        ).scalar()
        if not cur_geometry:
            cur_geometry = Geometry(win_x, win_y, win_width, win_height)
            self.session.add(cur_geometry)

        cur_window = self.session.query(Window).filter_by(title=window_name,
                                                          process_id=cur_process.id).scalar()
        if not cur_window:
            cur_window = Window(window_name, cur_process.id)
            self.session.add(cur_window)

        if not (self.current_window.proc_id == cur_process.id
                and self.current_window.win_id == cur_window.id):
            self.trycommit()
            self.store_keys()  # happens before as these keypresses belong to the previous window
            self.current_window.proc_id = cur_process.id
            self.current_window.win_id = cur_window.id
            self.current_window.geo_id = cur_geometry.id

    def filter_many(self):
        specials_in_row = 0
        lastpress = None
        newpresses = []
        for press in self.key_presses:
            key = press.key
            if specials_in_row and key != lastpress.key:
                if specials_in_row > 1:
                    lastpress.key = '%s]x%d>' % (lastpress.key[:-2], specials_in_row)

                newpresses.append(lastpress)
                specials_in_row = 0

            if len(key) > 1:
                specials_in_row += 1
                lastpress = press
            else:
                newpresses.append(press)

        if specials_in_row:
            if specials_in_row > 1:
                lastpress.key = '%s]x%d>' % (lastpress.key[:-2], specials_in_row)
            newpresses.append(lastpress)

        self.key_presses = newpresses

    def store_clipboard(self):

        clipboard_content = clipboard.text()
        types = str(mimeData.formats())

        # print(self.register_clipboard_formats())

        hasUrls = mimeData.hasUrls()
        hasText = mimeData.hasText()
        hasImage = mimeData.hasImage()
        hasHtml = mimeData.hasHtml()
        image = clipboard.image()

        image_height = image.height()
        image_width = image.width()

        # print("WAS CTRL+C", self.was_ctrl_c)
        # print("WAS CTRL+V", self.was_ctrl_v)
        # print("WAS CTRL+X", self.was_ctrl_x)


        # print (clipboard_content)

        keys = [press.key for press in self.key_presses]
        timings = [press.time for press in self.key_presses]
        add = lambda count, press: count + (0 if press.is_repeat else 1)
        nrkeys = reduce(add, self.key_presses, 0)

        lastTwoKeys = []

        hot_key_used = False

        for key in keys[-2:]:
            print ("KEY", key)
            print(str(key))

            lastTwoKeys.append(key)

        print ("LAST TWO KEYS", lastTwoKeys)

        hot_key_used = False

        if ("Ctrl" in str(lastTwoKeys[0:])):
            hot_key_used = True

        # Prepare for storing clipboard data in SQLite database
        self.session.add(
            Clipboard(clipboard_content.encode('utf8'), types, hasHtml, hasImage, hasText, hasUrls, image_height,
                      image_width, self.was_ctrl_c, self.was_ctrl_v, self.was_ctrl_x,
                      self.current_window.proc_id,
                      self.current_window.win_id,
                      self.current_window.geo_id))

        self.trycommit()

        self.started = NOW()

    def store_keys(self):
        """ Stores the current queued key-presses """
        if self.repeat_char:
            self.filter_many()

        if self.key_presses:
            keys = [press.key for press in self.key_presses]
            timings = [press.time for press in self.key_presses]
            add = lambda count, press: count + (0 if press.is_repeat else 1)
            nrkeys = reduce(add, self.key_presses, 0)

            curtext = u""
            if not self.store_text:
                keys = []
            else:
                curtext = ''.join(keys)

            self.session.add(Keys(curtext.encode('utf8'),
                                  keys,
                                  timings,
                                  nrkeys,
                                  self.started,
                                  self.current_window.proc_id,
                                  self.current_window.win_id,
                                  self.current_window.geo_id))

            self.trycommit()
            self.started = NOW()
            self.key_presses = []
            self.last_key_time = time.time()

    def got_key(self, keycode, state, string, is_repeat):
        """ Receives key-presses and queues them for storage.
            keycode is the code sent by the keyboard to represent the pressed key
            state is the list of modifier keys pressed, each modifier key should be represented
                  with capital letters and optionally followed by an underscore and location
                  specifier, i.e: SHIFT or SHIFT_L/SHIFT_R, ALT, CTRL
            string is the string representation of the key press
            repeat is True if the current key is a repeat sent by the keyboard """

        now = time.time()

        # print keycode

        self.was_ctrl_c = False
        self.was_ctrl_v = False
        self.was_ctrl_x = False

        if (keycode == '22' or string == 'v'):
            print "was v"
            if (self.lastKeyWasCtrl):
                # print ("STRG-V")
                self.was_ctrl_v = True
                self.store_clipboard()

        if (keycode == '3' or string == 'c'):
            print "was c"
            if (self.lastKeyWasCtrl):
                # print ("STRG-C")
                self.was_ctrl_c = True

        if (keycode == '24' or string == 'x'):
            print "was x"
            if (self.lastKeyWasCtrl):
                # print ("STRG-X")
                self.was_ctrl_x = True

        if (len(state)):
            if (state[0] == 'Ctrl'):
                # print "was control"
                self.lastKeyWasCtrl = True
            else:
                self.lastKeyWasCtrl = False
        else:
            self.lastKeyWasCtrl = False

        if string in SKIP_MODIFIERS:
            return
        if len(state) > 1 or (len(state) == 1 and state[0] != "Shift"):
            string = '<[%s: %s]>' % (' '.join(state), string)
        elif len(string) > 1:
            string = '<[%s]>' % string

        # print ("STRING AFTER", string)

        self.key_presses.append(KeyPress(string, now - self.last_key_time, is_repeat))
        self.last_key_time = now

    def store_click(self, button, x, y):
        """ Stores incoming mouse-clicks """
        self.session.add(Click(button,
                               True,
                               x, y,
                               len(self.mouse_path),
                               self.current_window.proc_id,
                               self.current_window.win_id,
                               self.current_window.geo_id))
        self.mouse_path = []
        self.trycommit()

    def got_mouse_click(self, button, x, y):
        """ Receives mouse clicks and sends them for storage.
            Mouse buttons: left: 1, middle: 2, right: 3, scroll up: 4, down:5, left:6, right:7
            x,y are the coordinates of the keypress
            press is True if it pressed down, False if released"""
        if button in [4, 5, 6, 7]:
            if time.time() - self.last_scroll[button] < SCROLL_COOLOFF:
                return
            self.last_scroll[button] = time.time()

        self.store_click(button, x, y)

    def got_mouse_move(self, x, y):
        """ Queues mouse movements.
            x,y are the new coorinates on moving the mouse"""
        self.mouse_path.append([x, y])

    def close(self):
        """ stops the sniffer and stores the latest keys. To be used on shutdown of program"""
        self.sniffer.cancel()
        self.store_keys()

    def change_password(self, new_encrypter):
        self.session = self.session_maker()
        keys = self.session.query(Keys).all()
        for k in keys:
            dtext = k.decrypt_text()
            dkeys = k.decrypt_keys()
            k.encrypt_text(dtext, new_encrypter)
            k.encrypt_keys(dkeys, new_encrypter)
        self.session.commit()
