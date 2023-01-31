import platform
if platform.system() != "Windows":
    print("This python script only works for Windows")
    exit(1)

import random
import subprocess
import win32con
import ctypes
from pathlib import Path
import tkinter as tk
from tkinter import ttk
import tkinter.messagebox
import time
import win32gui
import win32process
import psutil
import queue
import keyboard
import threading
from dataclasses import dataclass
import shutil
from configparser import RawConfigParser
import io

@dataclass
class DolphinInstance:
    proc: subprocess.Popen
    instance_index: int
    user_dir: Path
    main_window_handle: int
    game_window_handle: int

EnumWindows = ctypes.windll.user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))
GetWindowText = ctypes.windll.user32.GetWindowTextW
GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW
IsWindowVisible = ctypes.windll.user32.IsWindowVisible
SetWindowText = ctypes.windll.user32.SetWindowTextW
SW_MAXIMIZE = 3
SW_MINIMIZE = 6
SW_SHOW = 5
SW_HIDE = 0

dolphin_file_path = Path("C:/Dolphin-x64/dolphin.exe")

games_path = Path(f"D:/Roms/WiiGC")
insane_kart_wii = "InsaneKartWii_v1.0.3_Default(IKWPD1).wbfs"
wiimms_mario_kart_fun = "RMCP49.wbfs"
mario_kart_double_dash = "Mario Kart - Double Dash!! (Europe) (En,Fr,De,Es,It).iso"
available_games = [insane_kart_wii, wiimms_mario_kart_fun, mario_kart_double_dash]

user_dir_root_path = Path("C:/Dolphin-x64/Splitscreen")
dolphin_main_profile_path = Path.home()/Path("Documents/Dolphin Emulator")
q = queue.Queue()
dolphin_instances: list[DolphinInstance] = []

def action_replay_mkdd_pal(instance_index: int, config: RawConfigParser, total_instance_count: int) -> str:
    codes = {
        "Buffer Patch": "043DB6D4 43FA0000",
        "No Background Music [Ralf]": "043BA590 00000000",
        "8:9 Screen": 
"""043D5C48 3F638E39
043D5C4C 3FE8A71E""",
        "16:9 Screen": 
"""043D5C48 3FE38E39
043D5C4C 40649373"""
    }
    if total_instance_count != 2:
        codes.pop("8:9 Screen")
    else:
        codes.pop("16:9 Screen")
    if instance_index == 0:
        codes.pop("No Background Music [Ralf]")
    codes_str = ""
    for code_name in codes.keys():
        codes_str += f"${code_name}\n{codes[code_name]}\n"
        if not config.has_section("ActionReplay_Enabled"):
            config.add_section("ActionReplay_Enabled")
        config.set('ActionReplay_Enabled',f'${code_name}', None)
    return codes_str


def gecko_mkwii_pal(instance_index, config: RawConfigParser, total_instance_count: int) -> str:
    codes = {
        "no Music": "048A1D48 30000000",
    #    "no Menu Sounds": "048A1D4C 30000000",
    #    "no engine/bullet bill sound": "048A1D54 30000000",
    #    "no external sounds": "048A1D50 30000000",
    #    "no road sounds": "048A1D58 30000000",
    #    "no tire spark sounds": "048A1D5C 30000000",
    #    "no laku, boost, or trick sounds": "048A1D60 30000000",
    #    "no count down or lap sounds": "048A1D64 30000000",
    #    "no voice sound": "048A1D68 30000000",
    #    "no voice when star/mega mushroom used": "048A1D70 30000000",
    #    "no engine sound": "0424AAD0 30000000"
    }
    if instance_index == 0:
        codes.pop("no Music")
    codes_str = ""
    for code_name in codes.keys():
        codes_str += f"${code_name}\n{codes[code_name]}\n"
        if not config.has_section("Gecko_Enabled"):
            config.add_section("Gecko_Enabled")
        config.set('Gecko_Enabled',f'${code_name}', None)
    return codes_str

def gecko_mkwii_custom_port_pal(port) -> dict[str, str]:
    return {
        f"Custom UDP Port ({port})": (
            "C210E074 00000002\n"
            "2C031964 41820008\n"
            f"3860{port:04x} 00000000"
        )
    }

def get_hwnds_for_pid(pid):
    def callback(hwnd, hwnds):
        #if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
        _, found_pid = win32process.GetWindowThreadProcessId(hwnd)

        if found_pid == pid:
            hwnds.append(hwnd)
        return True
    hwnds = []
    win32gui.EnumWindows(callback, hwnds)
    return hwnds

def getWindowTitleByHandle(hwnd):
    length = GetWindowTextLength(hwnd)
    buff = ctypes.create_unicode_buffer(length + 1)
    GetWindowText(hwnd, buff, length + 1)
    return buff.value

def formatConfig() -> str:
    config = []
    config.append("--config=Dolphin.Display.RenderToMain=False")
    config.append("--config=Dolphin.Display.Fullscreen=True")
    config.append("--config=Dolphin.Analytics.PermissionAsked=True")
    config.append("--config=Dolphin.Interface.ShowActiveTitle=True")
    config.append("--config=GFX.Settings.BorderlessFullscreen=True")
    config.append("--config=Dolphin.Core.WiiSDCardAllowWrites=False")
    config.append("--config=Dolphin.Input.BackgroundInput=False")
    config.append(f"--config=Dolphin.Core.BBA_MAC=00:09:bf:{random.randint(0, 255):02x}:{random.randint(0, 255):02x}:{random.randint(0, 255):02x}")
    escapedConfig = []
    for c in config:
        escapedConfig.append(f'"{c}"')
    return " ".join(escapedConfig)

def optionxform(option: str) -> str:
    print(option)
    return option

def setupDolphinControls(game_settings_path: Path, game_id6: str, controllers_per_instance: int, instance_index: int, total_instance_count: int):
    gameini_config = RawConfigParser(allow_no_value=True, strict=True, delimiters="=")
    gameini_config.optionxform = optionxform
    gameini_actionReplay_section = ""
    gameini_gecko_section = ""
    # Read the gameini's, parse the Codes section separately
    gameini_path = game_settings_path/Path(f"{game_id6}.ini")
    if gameini_path.exists() and gameini_path.is_file():
        gameini_without_codes_str = ""
        with open(gameini_path, "r") as f:
            mode = 0
            for line in f:
                if line.startswith("["):
                    # Determine mode
                    if "[ActionReplay]" in line:
                        mode = 1
                    elif "[Gecko]" in line:
                        mode = 2
                    else:
                        mode = 0
                        gameini_without_codes_str += line
                else:
                    # Operation depending on mode
                    if mode == 0:
                        gameini_without_codes_str += line
                    elif mode == 1:
                        gameini_actionReplay_section += line
                    elif mode == 2:
                        gameini_gecko_section += line
        gameini_config.read_file(io.StringIO(gameini_without_codes_str))

    controller_index = instance_index
    controller_index = 0
    # Read the controls which are for this instance_index
    padType = {}
    padProfile = {}
    wiimoteSource = {}
    wiimoteProfile = {}
    for i in range(0, controllers_per_instance):
        padType[i] = gameini_config.get('Controls',f'PadType{controller_index*controllers_per_instance+i}',fallback=None)
        padProfile[i] = gameini_config.get('Controls',f'PadProfile{controller_index*controllers_per_instance+i+1}',fallback=None)
        wiimoteSource[i] = gameini_config.get('Controls',f'WiimoteSource{controller_index*controllers_per_instance+i}',fallback=None)
        wiimoteProfile[i] = gameini_config.get('Controls',f'WiimoteProfile{controller_index*controllers_per_instance+i+1}',fallback=None)
    # Remove the controls for all players
    for i in range(0, 12):
        if i < 4:
            gameini_config.set('Controls',f'PadType{i}', 0)
        else:
            gameini_config.remove_option('Controls',f'PadType{i}')
        gameini_config.remove_option('Controls',f'PadProfile{i+1}')
        if i < 4:
            gameini_config.set('Controls',f'WiimoteSource{i}', 0)
        else:
            gameini_config.remove_option('Controls',f'WiimoteSource{i}')
        gameini_config.remove_option('Controls',f'WiimoteProfile{i+1}')
    # Set the controls of the instance_index to the first slot
    for i in range(0, controllers_per_instance):
        if padType[i]:
            gameini_config.set('Controls',f'PadType{i}', padType[i])
        if padProfile[i]:
            gameini_config.set('Controls',f'PadProfile{i+1}', padProfile[i])
        if wiimoteSource[i]:
            gameini_config.set('Controls',f'WiimoteSource{i}', wiimoteSource[i])
        if wiimoteProfile[i]:
            gameini_config.set('Controls',f'WiimoteProfile{i+1}', wiimoteProfile[i])
    # Remove Gecko Codes and Action Replay Codes
    gameini_config.remove_section('Gecko')
    gameini_config.remove_section('ActionReplay')
    # Write to the gameini
    with open(game_settings_path/Path(f"{game_id6}.ini"), 'w') as configfile:
        # Prepare
        additional_action_replay = ""
        if "GM4P01" in game_id6:
            additional_action_replay = action_replay_mkdd_pal(instance_index, gameini_config, total_instance_count)
        additional_gecko = ""
        if "RMCP" in game_id6:
            additional_gecko = gecko_mkwii_pal(instance_index, gameini_config, total_instance_count)
        # Write Main GameIni
        gameini_config.write(configfile)
        # Action Replay
        configfile.write("[ActionReplay]\n")
        if gameini_actionReplay_section:
            configfile.write(gameini_actionReplay_section)
        configfile.write(additional_action_replay)
        # Gecko
        configfile.write("[Gecko]\n")
        if gameini_gecko_section:
            configfile.write(gameini_gecko_section)
        configfile.write(additional_gecko)
        

def setupDolphinConfiguration(user_dir: Path, game_id6: str, controllers_per_instance: int, instance_index: int, total_instance_count: int):
    config_path         = Path('Config')
    shutil.rmtree(user_dir/config_path, ignore_errors=True)
    shutil.copytree(dolphin_main_profile_path/config_path, user_dir/config_path, dirs_exist_ok=True)

    game_settings_path  = Path('GameSettings')
    shutil.rmtree(user_dir/game_settings_path, ignore_errors=True)
    shutil.copytree(dolphin_main_profile_path/game_settings_path, user_dir/game_settings_path, dirs_exist_ok=True)
    setupDolphinControls(user_dir/game_settings_path, game_id6, controllers_per_instance, instance_index, total_instance_count)

    gc_path             = Path('GC')
    if not (user_dir/gc_path).exists():
        shutil.copytree(dolphin_main_profile_path/gc_path, user_dir/gc_path, dirs_exist_ok=True)
    
    load_path           = Path('Load')
    if not (user_dir/load_path).exists():
        subprocess.check_call(f'mklink /J "{user_dir/load_path}" "{dolphin_main_profile_path/load_path}"', shell=True)

    resource_packs_path = Path('ResourcePacks')
    if not (user_dir/resource_packs_path).exists():
        subprocess.check_call(f'mklink /J "{user_dir/resource_packs_path}" "{dolphin_main_profile_path/resource_packs_path}"', shell=True)
    
    wii_path            = Path('Wii')
    if not (user_dir/wii_path).exists():
        shutil.copytree(dolphin_main_profile_path/wii_path, user_dir/wii_path, dirs_exist_ok=True)

    cache_path          = Path('Cache')
    if not (user_dir/cache_path).exists():
        shutil.copytree(dolphin_main_profile_path/cache_path, user_dir/cache_path, dirs_exist_ok=True)

def startDolphin(dolphin_file: Path, game_file: Path, game_id6: str, user_root_dir: Path, controllers_per_instance: int, instance_index: int, total_instance_count: int) -> DolphinInstance:
    user_dir = user_root_dir / Path(str(instance_index))
    setupDolphinConfiguration(user_dir, game_id6, controllers_per_instance, instance_index, total_instance_count)
    DETACHED_PROCESS = 8 # https://learn.microsoft.com/en-us/windows/win32/procthread/process-creation-flags
    proc = subprocess.Popen(f'"{dolphin_file}" --exec "{game_file}" --confirm --user "{user_dir}" {formatConfig()}', creationflags=DETACHED_PROCESS)
    game_window_handle = None
    main_window_handle = None
    for i in range(0, 10):
        time.sleep(2)
        hwnds = get_hwnds_for_pid(proc.pid)
        for hwnd in hwnds:
            window_title = win32gui.GetWindowText(hwnd).strip()
            #print(f"GetWindowText: {window_title}")
            if 'FPS' in window_title or 'Loading game specific input' in window_title or window_title == 'Dolphin':
                game_window_handle = hwnd
            elif 'Dolphin' in win32gui.GetWindowText(hwnd):
                main_window_handle = hwnd
        if game_window_handle and main_window_handle:
            break
    if not game_window_handle or not main_window_handle:
        proc.terminate()
        tkinter.messagebox.showerror("Error starting dolphin instance", "Could not determine the dolphin window handle.")
    return DolphinInstance(proc, instance_index, user_dir, main_window_handle, game_window_handle)

class App(tk.Tk):
    show_main_windows: tk.IntVar
    dolphin_instances_entry: tk.Entry
    controllers_per_instance_entry: tk.Entry
    selected_game: tk.StringVar

    def __init__(self):
        super().__init__()
        # Selected Game
        self.selected_game = tk.StringVar()
        combobox = ttk.Combobox(self, textvariable=self.selected_game)
        combobox['values'] = available_games
        combobox['state'] = 'readonly'
        combobox.set(available_games[0])
        combobox.pack(fill=tk.X)
        # Dolphin instance count
        label = tk.Label(self, text="Dolphin Instances:")
        label.pack()
        self.dolphin_instances_entry = tk.Entry(self, width=50)
        self.dolphin_instances_entry.pack()
        # Controllers per instance
        label = tk.Label(self, text="Controllers per Instance:")
        label.pack()
        self.controllers_per_instance_entry = tk.Entry(self, width=50)
        self.controllers_per_instance_entry.pack()
        # Go Button
        buttonGo = tk.Button(self,
            text="Go",
            width=10,
            height=2,
            command=self.handleButtonGo
        )
        buttonGo.pack()
        # Reposition Button
        buttonRepos = tk.Button(self,
            text="Reposition",
            width=10,
            height=2,
            command=self.handleButtonRepos
        )
        buttonRepos.pack()
        # Show Main Windows Checkbox
        self.show_main_windows = tk.IntVar(self)
        buttonShowMainWindows = tk.Checkbutton(self,
            text="Show Main Windows",
            command=self.showMainWindows,
            variable=self.show_main_windows)
        buttonShowMainWindows.pack()
        # Configure Button
        buttonSettings = tk.Button(self,
            text="Settings",
            width=10,
            height=2,
            command=self.handleButtonSettings
        )
        buttonSettings.pack()
        # - profile combobox -
        #self.current_profile = tk.StringVar(self)
        #combobox = ttk.Combobox(self, textvariable=self.current_profile, state='readonly')
        #profiles_list = []
        # add default value first
        #default_profile_path = Path("profiles/default")
        #if default_profile_path.exists() and default_profile_path.is_dir():
        #    profiles_list.append(default_profile_path.as_posix())
        # add all other profiles
        #for profile_path in Path("profiles").iterdir():
        #    if profile_path.is_dir() and profile_path != default_profile_path:
        #        profiles_list.append(profile_path.as_posix())
        #combobox['values'] = tuple(profiles_list)
        #combobox.current(0)
        #combobox.pack()

    def showMainWindows(self):
        for dolphin_instance in dolphin_instances:
            if win32gui.IsWindow(dolphin_instance.main_window_handle):
                if self.show_main_windows.get():
                    win32gui.ShowWindow(dolphin_instance.main_window_handle, SW_SHOW)
                else:
                    win32gui.ShowWindow(dolphin_instance.main_window_handle, SW_HIDE)

    def reposition_(self, dolphin_instance: DolphinInstance, count: int) -> bool:
        failure = True
        if count == 1:
            return True
        elif count == 2:
            cols = 2
            rows = 1
        elif count >= 3 and count <= 4:
            cols = 2
            rows = 2
        elif count >= 5 and count <= 6:
            cols = 2
            rows = 3
        elif count >= 7 and count <= 8:
            cols = 2
            rows = 4
        elif count == 9:
            cols = 3
            rows = 3
        elif count >= 10 and count <= 12:
            cols = 4
            rows = 3
        else:
            return False
        x = int(dolphin_instance.instance_index % cols)
        y = int(dolphin_instance.instance_index / cols)
        x = int(x * self.winfo_screenwidth() / cols)
        y = int(y * self.winfo_screenheight() / rows)
        width = int(self.winfo_screenwidth() / cols)
        height = int(self.winfo_screenheight() / rows)
        if win32gui.IsWindow(dolphin_instance.game_window_handle):
            win32gui.SetWindowPos(dolphin_instance.game_window_handle, win32con.HWND_TOPMOST, x, y, width, height, win32con.SWP_SHOWWINDOW)
            failure = False
        if win32gui.IsWindow(dolphin_instance.main_window_handle):
            win32gui.SetWindowPos(dolphin_instance.main_window_handle, win32con.HWND_BOTTOM, x, y, width, height, win32con.SWP_SHOWWINDOW)
            if not self.show_main_windows.get():
                win32gui.ShowWindow(dolphin_instance.main_window_handle, SW_HIDE)
            failure = False
        return not failure

    def reposition(self, count = 0) -> int:
        successCount = 0
        if count == 0:
            count = len(dolphin_instances)
        for dolphin_instance in dolphin_instances:
            if self.reposition_(dolphin_instance, count):
                successCount += 1
        return successCount

    def handleButtonRepos(self):
        self.attributes('-topmost', False)
        self.update()
        self.reposition()

    def handleButtonSettings(self):
        # todo
        pass

    def handleButtonGo(self):
        global dolphin_instances
        requested_dolphin_instances_str = self.dolphin_instances_entry.get()
        if (not requested_dolphin_instances_str.isdigit()):
            return
        requested_dolphin_instances = int(requested_dolphin_instances_str)
        controllers_per_instance_str = self.controllers_per_instance_entry.get()
        controllers_per_instance = 1
        if controllers_per_instance_str.isdigit():
            controllers_per_instance = int(requested_dolphin_instances_str)

        existing_dolphin_instances = 0
        for i in range(0, len(dolphin_instances)):
            dolphin_instance = dolphin_instances[i]
            if win32gui.IsWindow(dolphin_instance.game_window_handle):
                existing_dolphin_instances+=1

        if requested_dolphin_instances < 0 or requested_dolphin_instances > 12:
            return
        if requested_dolphin_instances == existing_dolphin_instances and requested_dolphin_instances == len(dolphin_instances):
            return
        
        # remove the excess dolphin instances
        for i in range(requested_dolphin_instances, len(dolphin_instances)):
            dolphin_instances.pop().proc.terminate()
        
        self.reposition(requested_dolphin_instances)

        game_file_path = games_path / Path(self.selected_game.get())

        # determine gameid
        game_id6 = subprocess.check_output(["wit", "ID6", game_file_path], encoding="utf8").strip()

        # check if old dolphin instances still exist and if not start them afresh
        for i in range(0, len(dolphin_instances)):
            dolphin_instance = dolphin_instances[i]
            if not win32gui.IsWindow(dolphin_instance.game_window_handle):
                dolphin_instances[i] = startDolphin(dolphin_file_path, game_file_path, game_id6, user_dir_root_path, controllers_per_instance, i, requested_dolphin_instances)
                self.reposition(requested_dolphin_instances)
        
        # start new dolphin instances if neccessary
        for i in range(len(dolphin_instances), requested_dolphin_instances):
            dolphin_instance = startDolphin(dolphin_file_path, game_file_path, game_id6, user_dir_root_path, controllers_per_instance, i, requested_dolphin_instances)
            dolphin_instances.insert(i+1, dolphin_instance)
            self.reposition(requested_dolphin_instances)

        while not q.empty():
            q.get()
            q.task_done()
        for dolphin_instance in dolphin_instances:
            q.put(dolphin_instance)
        
        for i in range(0,5):
            time.sleep(1)
            if self.reposition(requested_dolphin_instances) == len(dolphin_instances):
                break

def check_esc_pressed(esc_stop: queue.Queue, app: App):
    while esc_stop.empty():
        if keyboard.is_pressed('esc') or keyboard.is_pressed('alt'):
            for dolphin_instance in list(q.queue):
                if win32gui.IsWindow(dolphin_instance.game_window_handle):
                    win32gui.SetWindowPos(dolphin_instance.game_window_handle, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            app.lift()
            app.attributes('-topmost', True)
            app.update()
            app.attributes('-topmost', False)
        time.sleep(0.1)

def main():
    app = App()

    esc_stop = queue.Queue()
    threading.Thread(target=check_esc_pressed, daemon=True, args=(esc_stop, app,)).start()

    app.mainloop()

    esc_stop.put(True)


if __name__ == "__main__":
    main()