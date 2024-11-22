import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys
import ctypes
import shelve
import time
import pygetwindow as gw
import pyautogui
import pyperclip
from traceback import format_exc
from tkinter import TclError
import requests
import shutil
import hashlib



script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
config_dir = os.path.join(script_dir, "AutoCoDSlave_Config")
if not os.path.exists(config_dir):
    os.makedirs(config_dir)
    FILE_ATTRIBUTE_HIDDEN = 0x02
    ctypes.windll.kernel32.SetFileAttributesW(config_dir, FILE_ATTRIBUTE_HIDDEN)
db_path = os.path.join(config_dir, 'config_db')
selection_path = os.path.join(config_dir, 'selection_db')
button_click_status = None


def is_ctrl_pressed():
    return ctypes.windll.user32.GetAsyncKeyState(0xA2) & 0x8000 != 0


def force_lazy_mode_off():
    with shelve.open(db_path) as config:
        options = config['OPTIONS']
        options['LazyMode'] = 0
        config['OPTIONS'] = options

def download_update(url, save_path):
    response = requests.get(url, stream=True)
    with open(save_path, 'wb') as out_file:
        shutil.copyfileobj(response.raw, out_file)
    del response

def calculate_hash(file_path):
    hash_sha256 = hashlib.sha256()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_sha256.update(chunk)
    return hash_sha256.hexdigest()

def restart_program():
    python = sys.executable
    os.execl(python, python, *sys.argv)

if is_ctrl_pressed():
    force_lazy_mode_off()


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def run_as_admin():
    if not is_admin():
        try:
            params = ' '.join([sys.executable] + sys.argv)
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
            sys.exit(0)
        except Exception as e:
            raise RuntimeError(f"无法以管理员权限重新启动程序: {e}")


run_as_admin()


def init_config():
    default_path = {'RUT_Path': '', 'XGP_CoD_Path': '', 'Steam_CoD_Path': ''}
    default_options = {'LazyMode': 0, 'LaunchCoD': 0, 'LaunchRUT': 0}
    default_keys = {'CoD21AllKey': '', 'CoD21UAVKey': '', 'CoD20AllKey': '', 'CoD20UAVKey': ''}

    with shelve.open(db_path) as config:
        # 初始化 PATH 条目及其值
        if 'PATH' not in config:
            config['PATH'] = default_path
        else:
            for key in default_path:
                if key not in config['PATH']:
                    config['PATH'][key] = default_path[key]

        # 初始化 OPTIONS 条目及其值
        if 'OPTIONS' not in config:
            config['OPTIONS'] = default_options
        else:
            for key in default_options:
                if key not in config['OPTIONS']:
                    config['OPTIONS'][key] = default_options[key]

        # 初始化 KEYS 条目及其值
        if 'KEYS' not in config:
            config['KEYS'] = default_keys
        else:
            for key in default_keys:
                if key not in config['KEYS']:
                    config['KEYS'][key] = default_keys[key]


def get_config():
    with shelve.open(db_path) as config:
        path = config.get('PATH', {'RUT_Path': '', 'XGP_CoD_Path': '', 'Steam_CoD_Path': ''})
        options = config.get('OPTIONS', {'LazyMode': 0, 'LaunchCoD': 0, 'LaunchRUT': 0})
        keys = config.get('KEYS', {'CoD21AllKey': '', 'CoD21UAVKey': '', 'CoD20AllKey': '', 'CoD20UAVKey': ''})
        return path, options, keys


def set_config(path, options, keys):
    with shelve.open(db_path) as config:
        config['PATH'] = path
        config['OPTIONS'] = options
        config['KEYS'] = keys


def get_selection():
    try:
        with shelve.open(selection_path) as selection:
            return selection['data']
    except KeyError:
        return {}


def set_selection(selection_data):
    with shelve.open(selection_path, 'c') as selection:
        selection['data'] = selection_data


init_config()
config_paths, config_options, config_keys = get_config()

selection_data = get_selection()


def create_gui():
    root = tk.Tk()
    root.title("AutoCoDSlave")
    root.geometry('800x600')
    root.resizable(False, False)

    def save_config():
        config_paths['RUT_Path'] = e1.get()
        config_paths['XGP_CoD_Path'] = e2.get()
        config_paths['Steam_CoD_Path'] = e7.get()
        config_keys['CoD21AllKey'] = e3.get()
        config_keys['CoD21UAVKey'] = e4.get()
        config_keys['CoD20AllKey'] = e5.get()
        config_keys['CoD20UAVKey'] = e6.get()
        set_config(config_paths, config_options, config_keys)
        launch()
        set_config(config_paths, config_options, config_keys)
        popup.destroy()
        update_checkbuttons()
        update_platform_checkbuttons()
        launch() if var_lazy_mode.get() else root.deiconify()

    def configure_keys_and_paths():
        global popup, e1, e2, e3, e4, e5, e6, e7
        popup = tk.Toplevel(root)
        popup.title("配置密钥与地址")
        popup.geometry('800x600')
        popup.resizable(False, False)

        def browse_file(entry):
            file_path = filedialog.askopenfilename()
            entry.delete(0, tk.END)
            entry.insert(0, file_path)

        tk.Label(popup, text="RUTV3.exe所在目录").grid(row=0, pady=10)
        tk.Label(popup, text="CoD启动器位置(XGP平台)").grid(row=1, pady=10)
        tk.Label(popup, text="CoD启动器位置(Steam平台)").grid(row=2, pady=10)
        tk.Label(popup, text="CoD21全解Key").grid(row=3, pady=10)
        tk.Label(popup, text="CoD21无限UAV Key").grid(row=4, pady=10)
        tk.Label(popup, text="CoD20全解Key").grid(row=5, pady=10)
        tk.Label(popup, text="CoD20无限UAV Key").grid(row=6, pady=10)
        e1 = tk.Entry(popup, width=70)
        e2 = tk.Entry(popup, width=70)
        e7 = tk.Entry(popup, width=70)
        e3 = tk.Entry(popup, width=70, show='*')
        e4 = tk.Entry(popup, width=70, show='*')
        e5 = tk.Entry(popup, width=70, show='*')
        e6 = tk.Entry(popup, width=70, show='*')
        e1.grid(row=0, column=1, padx=20)
        e2.grid(row=1, column=1, padx=20)
        e7.grid(row=2, column=1, padx=20)
        e3.grid(row=3, column=1, padx=20)
        e4.grid(row=4, column=1, padx=20)
        e5.grid(row=5, column=1, padx=20)
        e6.grid(row=6, column=1, padx=20)
        tk.Button(popup, text="浏览", command=lambda: browse_file(e1)).grid(row=0, column=2, padx=20)
        tk.Button(popup, text="浏览", command=lambda: browse_file(e2)).grid(row=1, column=2, padx=20)
        tk.Button(popup, text="浏览", command=lambda: browse_file(e7)).grid(row=2, column=2, padx=20)
        red_label = tk.Label(popup, text="XGP平台不可以直接选择桌面快捷方式，如何找到正确位置：\n"
                                         "打开Xbox应用-右键使命召唤-管理-文件-浏览...\n"
                                         "在打开的文件管理器中双击Call of Duty文件夹-\n"
                                         "双击Content文件夹-找到gamelaunchhelper.exe即为CoD启动器\n"
                                         "例如：F:/XboxGames/Call of Duty/Content/gamelaunchhelper.exe",
                             fg="red")
        red_label.grid(row=8, columnspan=3, pady=20)
        button_frame = tk.Frame(popup)
        button_frame.grid(row=7, columnspan=3, pady=20)
        tk.Button(button_frame, text="取消配置", command=lambda: (popup.destroy(), root.deiconify())).pack(side='left',
                                                                                                           padx=10)
        tk.Button(button_frame, text="保存并退出", command=save_config).pack(side='right', padx=10)
        e1.insert(0, config_paths['RUT_Path'])
        e2.insert(0, config_paths['XGP_CoD_Path'])
        e7.insert(0, config_paths['Steam_CoD_Path'])
        e3.insert(0, config_keys['CoD21AllKey'])
        e4.insert(0, config_keys['CoD21UAVKey'])
        e5.insert(0, config_keys['CoD20AllKey'])
        e6.insert(0, config_keys['CoD20UAVKey'])
        root.withdraw()
        popup.protocol("WM_DELETE_WINDOW", lambda: (popup.destroy(), root.deiconify()))
        popup.mainloop()

    def start_launch_process(action):
        try:
            root.iconify()  # 将界面最小化
        except Exception as e:
            messagebox.showerror("Error", f"最小化窗口时出错: {e}\n{format_exc()}")
        if action == 'launch_rut_and_cod':
            config_options['LaunchRUT'] = 1
            config_options['LaunchCoD'] = 1
        elif action == 'launch_rut_only':
            config_options['LaunchRUT'] = 1
            config_options['LaunchCoD'] = 0
        elif action == 'start_cod_only':
            config_options['LaunchCoD'] = 1
            config_options['LaunchRUT'] = 0
        set_config(config_paths, config_options, config_keys)
        save_selections()
        launch()
        close_autocodslave()

    def launch():
        try:
            config_paths, config_options, config_keys = get_config()
            rut_path = config_paths['RUT_Path']
            cod_path = None
            if config_options.get('LaunchCoD', 0):
                if var_xgp.get():
                    cod_path = config_paths['XGP_CoD_Path']
                elif var_steam.get():
                    cod_path = config_paths['Steam_CoD_Path']
            if config_options.get('LaunchRUT', 0) and rut_path:
                options = get_selected_options()
                if options:
                    execute_rut_and_cod(options)
            if cod_path and config_options.get('LaunchCoD', 0):
                start_cod(cod_path)
        except Exception as e:
            messagebox.showerror("Error", f"启动出错: {e}\n{format_exc()}")

    def execute_rut_and_cod(options):
        rut_path = config_paths['RUT_Path']
        for label, key in options:
            print(f"启动 RUT: {label}")
            start_rut(rut_path)
            if not wait_for_rut_window():
                print("未检测到 RUT 窗口")
                break
            print(f"输入密钥: {label}")
            input_key('RUTV3', config_keys[key])
            if not watch_rut_closing():
                print("检测到 RUT 窗口关闭")
                break

    def start_cod(cod_path):
        os.startfile(cod_path)

    def start_rut(rut_path):
        os.startfile(rut_path)

    def get_selected_options():
        options = []
        if var4.get():
            options.append(("CoD20无限UAV", 'CoD20UAVKey'))
        if var3.get():
            options.append(("CoD20全解", 'CoD20AllKey'))
        if var1.get():
            options.append(("CoD21全解", 'CoD21AllKey'))
        if var2.get():
            options.append(("CoD21无限UAV", 'CoD21UAVKey'))
        return options

    def wait_for_rut_window():
        for _ in range(10):
            print("检查RUT窗口是否存在...")
            if any('RUTV3' in win.title for win in gw.getWindowsWithTitle('RUTV3')):
                print("检测到RUT窗口")
                return True
            time.sleep(3)
        return False

    def input_key(app_title, key):
        windows = gw.getWindowsWithTitle(app_title)
        if windows:
            win = windows[0]
            clipboard_content = pyperclip.paste()
            try:
                pyperclip.copy(key)
                if not win.isActive:
                    win.activate()
                    time.sleep(1)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.5)
                pyautogui.press('enter')
            finally:
                pyperclip.copy(clipboard_content)

    def watch_rut_closing():
        while True:
            print("监测RUT窗口是否关闭...")
            if not any('RUTV3' in win.title for win in gw.getWindowsWithTitle('RUTV3')):
                return True
            time.sleep(1)
        return False

    def close_autocodslave():
        print("关闭AutoCoDSlave")
        try:
            update_platform_checkbuttons()
            set_config(config_paths, config_options, config_keys)
        except TclError as err:
            print(f"Ignored TclError: {err}")
        root.destroy()

    def on_btn_start_rut_and_cod():
        start_launch_process('launch_rut_and_cod')

    def on_btn_start_rut():
        start_launch_process('launch_rut_only')

    def on_btn_start_cod_only():
        start_launch_process('start_cod_only')

    def save_selections():
        selections = {
            'CoD21AllKey': var1.get(),
            'CoD21UAVKey': var2.get(),
            'CoD20AllKey': var3.get(),
            'CoD20UAVKey': var4.get(),
            'XGP': var_xgp.get(),
            'Steam': var_steam.get()
        }
        set_selection(selections)
        set_config(config_paths, config_options, config_keys)

    def save_lazy_mode():
        config_options['LazyMode'] = var_lazy_mode.get()
        set_config(config_paths, config_options, config_keys)

    def check_selection_logic():
        if var1.get():
            cb2.config(state='disabled')
            cb3.config(state='disabled')
            cb4.config(state='disabled')
        elif var2.get():
            cb1.config(state='disabled')
            cb3.config(state='disabled')
            cb4.config(state='disabled')
        elif var3.get():
            cb1.config(state='disabled')
            cb2.config(state='disabled')
            cb4.config(state='normal')
        elif var4.get():
            cb1.config(state='disabled')
            cb2.config(state='disabled')
            cb3.config(state='normal')
        else:
            for cb, key in [(cb1, 'CoD21AllKey'), (cb2, 'CoD21UAVKey'), (cb3, 'CoD20AllKey'), (cb4, 'CoD20UAVKey')]:
                if config_keys[key]:
                    cb.config(state='normal')
                else:
                    cb.config(state='disabled')

    def check_platform_logic(*args):
        if var_xgp.get():
            cb_steam.config(state='disabled')
        elif var_steam.get():
            cb_xgp.config(state='disabled')
        else:
            cb_xgp.config(state='normal')
            cb_steam.config(state='normal')
        save_selections()
        save_lazy_mode()

    def on_lazy_mode_check():
        if var_lazy_mode.get():
            msg = "勾选该选项后，下次启动程序时将直接按本次运行时的勾选设置和按钮选项执行\n" \
                  "注意:勾选后下次启动将不会显示UI界面，若需显示主界面以更改配置，请按住键盘左Ctrl键后双击运行本程序。"
            messagebox.showinfo("懒人模式", msg)
        save_lazy_mode()

    def update_program():
        try:
            global script_dir  # 确保脚本目录被正确引用

            github_url = "https://github.com/2eit/AutoCoDSlave/releases/download/AutoCoDSlave-latest/AutoCoDSlave.exe"
            download_path = os.path.join(script_dir, "new_program.exe")
            script_path = sys.argv[0]

            # 下载文件
            download_update(github_url, download_path)

            # 计算当前程序文件和下载文件的哈希值
            current_hash = calculate_hash(script_path)
            new_hash = calculate_hash(download_path)

            # 比较哈希值，不一致则更新
            if current_hash != new_hash:
                # 创建批处理文件内容
                bat_path = os.path.join(script_dir, "update_program.bat")
                bat_content = f"""
                @echo off
                taskkill /F /IM "{os.path.basename(sys.executable)}"
                del "{script_path}"
                move /Y "{download_path}" "{script_path}"
                start "" "{script_path}"
                del "%~f0"
                """

                # 写入批处理文件
                with open(bat_path, "w") as bat_file:
                    bat_file.write(bat_content.strip())

                # 弹窗提示更新完成
                if messagebox.showinfo("更新完成", "程序更新已完成。点击确定以重启应用程序。"):
                    # 运行批处理文件并关闭自身
                    os.startfile(bat_path)
                    sys.exit()

            else:
                os.remove(download_path)
                messagebox.showinfo("已是最新版本", "当前程序已经是最新版本，不需要更新。")

        except Exception as e:
            messagebox.showerror("更新失败", f"更新程序时出现错误: {e}\n{format_exc()}")

    tk.Button(root, text="配置密钥与地址", command=configure_keys_and_paths).pack(pady=20)
    tk.Button(root, text="启动RUT并启动游戏", command=on_btn_start_rut_and_cod).pack(pady=20)
    tk.Button(root, text="仅启动RUT", command=on_btn_start_rut).pack(pady=20)
    tk.Button(root, text="我是绿玩😡(仅启动游戏)", command=on_btn_start_cod_only).pack(pady=20)

    var1 = tk.IntVar()
    var2 = tk.IntVar()
    var3 = tk.IntVar()
    var4 = tk.IntVar()
    var_xgp = tk.IntVar()
    var_steam = tk.IntVar()
    var_lazy_mode = tk.IntVar()

    cb1 = tk.Checkbutton(root, text="CoD21全解", variable=var1,
                         command=lambda: (save_selections(), check_selection_logic()))
    cb2 = tk.Checkbutton(root, text="CoD21无限UAV", variable=var2,
                         command=lambda: (save_selections(), check_selection_logic()))
    cb3 = tk.Checkbutton(root, text="CoD20全解", variable=var3,
                         command=lambda: (save_selections(), check_selection_logic()))
    cb4 = tk.Checkbutton(root, text="CoD20无限UAV", variable=var4,
                         command=lambda: (save_selections(), check_selection_logic()))
    cb_xgp = tk.Checkbutton(root, text="XGP平台", variable=var_xgp,
                            command=lambda: (save_selections(), check_platform_logic()))
    cb_steam = tk.Checkbutton(root, text="Steam平台", variable=var_steam,
                              command=lambda: (save_selections(), check_platform_logic()))
    cb_lazy_mode = tk.Checkbutton(root, text="懒人模式", variable=var_lazy_mode, command=on_lazy_mode_check)
    cb1.pack()
    cb2.pack()
    cb3.pack()
    cb4.pack()
    cb_xgp.pack()
    cb_steam.pack()
    cb_lazy_mode.pack()
    tk.Label(root, text="--by 怂呐(https://github.com/2eit/)", font=('TkDefaultFont', 8)).pack(side="right", anchor="se", padx=10, pady=10)
    update_label = tk.Label(root, text="点我更新", fg="blue", cursor="hand2", font=('TkDefaultFont', 10))
    update_label.pack(side=tk.BOTTOM, pady=20)
    update_label.place(relx=0.5, rely=1.0, anchor='s')
    update_label.bind("<Button-1>", lambda e: update_program())


    var_lazy_mode.set(config_options.get('LazyMode', 0))

    def update_checkbuttons():
        selection_data = get_selection()
        default_keys = {'CoD21AllKey': '', 'CoD21UAVKey': '', 'CoD20AllKey': '', 'CoD20UAVKey': ''}
        for key in default_keys.keys():
            if key not in config_keys:
                config_keys[key] = default_keys[key]
        for cb, key, var in [(cb1, 'CoD21AllKey', var1), (cb2, 'CoD21UAVKey', var2), (cb3,'CoD20AllKey', var3), (cb4, 'CoD20UAVKey', var4)]:
            if config_keys[key]:
                cb.config(state='normal')
            else:
                cb.config(state='disabled')
            if key in selection_data:
                var.set(selection_data[key])
            else:
                var.set(0)
            check_selection_logic()


    def update_platform_checkbuttons():
        selection_data = get_selection()
        var_xgp.set(selection_data.get('XGP', 0))
        var_steam.set(selection_data.get('Steam', 0))
        check_platform_logic()


    update_checkbuttons()
    update_platform_checkbuttons()
    if config_options.get('LazyMode', 0):
        launch()
        close_autocodslave()

    root.mainloop()

create_gui()
