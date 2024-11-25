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
import webbrowser
import win32com.client


script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
config_dir = os.path.join(script_dir, "AutoCoDSlave_Config")
if not os.path.exists(config_dir):
    os.makedirs(config_dir)
    FILE_ATTRIBUTE_HIDDEN = 0x02
    ctypes.windll.kernel32.SetFileAttributesW(config_dir, FILE_ATTRIBUTE_HIDDEN)
db_path = os.path.join(config_dir, 'config_db')
selection_path = os.path.join(config_dir, 'selection_db')
button_click_status = None


def open_project_url(event):
    webbrowser.open("https://github.com/2eit/AutoCoDSlave")

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
    response.close()
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

def init_config():
    default_path = {'RUT_Path': '', 'CoD_Path': ''}
    default_options = {'LazyMode': 0, 'LaunchCoD': 0, 'LaunchRUT': 0, 'XGP': 0, 'Steam': 0, 'BattleNet': 0}  # 添加平台状态
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
        path = config.get('PATH', {'RUT_Path': '', 'CoD_Path': ''})
        options = config.get('OPTIONS',
                             {'LazyMode': 0, 'LaunchCoD': 0, 'LaunchRUT': 0, 'XGP': 0, 'Steam': 0, 'BattleNet': 0})
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
            data = selection.get('data', {})
            print(f"读取选择状态: {data}")  # 调试信息
            return data
    except KeyError:
        return {}

def set_selection(selection_data):
    try:
        with shelve.open(selection_path, 'c') as selection:
            selection['data'] = selection_data
            print(f"保存选择状态: {selection_data}")  # 调试信息
    except Exception as e:
        print(f"保存选择状态出错: {e}")


def create_gui():
    root = tk.Tk()
    root.title("AutoCoDSlave")

    def save_selections_and_config():
        selections = {
            'CoD21AllKey': var1.get(),
            'CoD21UAVKey': var2.get(),
            'CoD20AllKey': var3.get(),
            'CoD20UAVKey': var4.get(),
            'XGP': var_xgp.get(),
            'Steam': var_steam.get(),
            'BattleNet': var_battle_net.get(),
            'LazyMode': var_lazy_mode.get(),
        }

        # 更新platform status
        config_options['XGP'] = var_xgp.get()
        config_options['Steam'] = var_steam.get()
        config_options['BattleNet'] = var_battle_net.get()
        config_options['LazyMode'] = var_lazy_mode.get()

        set_selection(selections)
        set_config(config_paths, config_options, config_keys)
        print(f"保存选择和配置：{selections}")

    def on_cb_toggled():
        save_selections_and_config()
        check_selection_logic()

    def on_btn_clicked(action):
        if action == 'launch_rut_and_cod':
            config_options['LaunchRUT'] = 1
            config_options['LaunchCoD'] = 1
        elif action == 'launch_rut_only':
            config_options['LaunchRUT'] = 1
            config_options['LaunchCoD'] = 0
        elif action == 'start_cod_only':
            config_options['LaunchCoD'] = 1
            config_options['LaunchRUT'] = 0
        save_selections_and_config()
        start_launch_process(action)

    def find_battlenet_path():
        try:
            # 创建 Shell.Application COM 对象
            shell = win32com.client.Dispatch("Shell.Application")
            apps_folder = shell.Namespace("shell:AppsFolder")

            for item in apps_folder.Items():
                if item.Name == "Call of Duty":
                    # 获取应用程序路径
                    app_path = item.Path
                    if app_path:
                        return app_path

            # 未找到应用程序
            messagebox.showerror("错误", "未找到战网平台上的 'Call of Duty' 应用")
            return None

        except Exception as e:
            messagebox.showerror("错误", f"查找 'Call of Duty' 应用路径时出错: {e}\n{format_exc()}")
            return None

    def save_selections():
        selections = {
            'CoD21AllKey': var1.get(),
            'CoD21UAVKey': var2.get(),
            'CoD20AllKey': var3.get(),
            'CoD20UAVKey': var4.get(),
        }

        # 更新 config_options 中的平台复选框状态
        config_options['XGP'] = var_xgp.get()
        config_options['Steam'] = var_steam.get()
        config_options['BattleNet'] = var_battle_net.get()

        set_selection(selections)
        set_config(config_paths, config_options, config_keys)
        print(f"保存时选择数据: {selections}")  # 调试信息

    def save_config():
        config_paths['RUT_Path'] = e1.get()  # 保存路径
        config_keys['CoD21AllKey'] = e3.get()  # 保存 CoD21 全解 Key
        config_keys['CoD21UAVKey'] = e4.get()  # 保存 CoD21 无限 UAV Key
        config_keys['CoD20AllKey'] = e5.get()  # 保存 CoD20 全解 Key
        config_keys['CoD20UAVKey'] = e6.get()  # 保存 CoD20 无限 UAV Key

        # 保存平台选择状态
        config_options['XGP'] = var_xgp.get()
        config_options['Steam'] = var_steam.get()
        config_options['BattleNet'] = var_battle_net.get()

        set_config(config_paths, config_options, config_keys)
        popup.destroy()  # 销毁弹出的配置窗口
        root.deiconify()  # 恢复主窗口的可见性
        update_checkbuttons()
        update_platform_checkbuttons()

    def configure_keys_and_paths():
        global popup, e1, entry_cod, e3, e4, e5, e6
        popup = tk.Toplevel(root)
        popup.title("配置密钥和地址")
        popup.geometry('500x300')
        popup.resizable(False, False)

        def browse_file(entry):
            file_path = filedialog.askopenfilename()
            entry.delete(0, tk.END)
            entry.insert(0, file_path)

        # 标签和输入框
        tk.Label(popup, text="RUTV3.exe目录").grid(row=0, pady=10)
        tk.Label(popup, text="CoD21全解Key").grid(row=1, pady=10)
        tk.Label(popup, text="CoD21无限UAV Key").grid(row=2, pady=10)
        tk.Label(popup, text="CoD20全解Key").grid(row=3, pady=10)
        tk.Label(popup, text="CoD20无限UAV Key").grid(row=4, pady=10)

        e1 = tk.Entry(popup, width=35)
        e3 = tk.Entry(popup, width=35, show='*')
        e4 = tk.Entry(popup, width=35, show='*')
        e5 = tk.Entry(popup, width=35, show='*')
        e6 = tk.Entry(popup, width=35, show='*')

        e1.grid(row=0, column=1, padx=20)
        e3.grid(row=1, column=1, padx=20)
        e4.grid(row=2, column=1, padx=20)
        e5.grid(row=3, column=1, padx=20)
        e6.grid(row=4, column=1, padx=20)

        # 浏览文件按钮
        tk.Button(popup, text="浏览", command=lambda: browse_file(e1)).grid(row=0, column=2, padx=20)

        # 确认和取消按钮
        button_frame = tk.Frame(popup)
        button_frame.grid(row=5, columnspan=3, pady=20)
        tk.Button(button_frame, text="取消配置", command=lambda: (popup.destroy(), root.deiconify())).pack(side='left',
                                                                                                           padx=10)
        tk.Button(button_frame, text="保存并退出", command=lambda: (save_config(), root.deiconify())).pack(side='right',
                                                                                                           padx=10)

        # 填充输入框的初始值
        e1.insert(0, config_paths['RUT_Path'])
        e3.insert(0, config_keys['CoD21AllKey'])
        e4.insert(0, config_keys['CoD21UAVKey'])
        e5.insert(0, config_keys['CoD20AllKey'])
        e6.insert(0, config_keys['CoD20UAVKey'])

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
            cod_path = ''

            launch_xgp = var_xgp.get() == 1 or config_options.get('XGP', 0) == 1
            launch_steam = var_steam.get() == 1 or config_options.get('Steam', 0) == 1
            launch_battle_net = var_battle_net.get() == 1 or config_options.get('BattleNet', 0) == 1

            if config_options.get('LaunchCoD', 0):
                if launch_xgp:
                    cod_path = "shell:AppsFolder\\38985CA0.COREBase_5bkah9njm3e9g!codShip"
                elif launch_steam:
                    cod_path = "steam://rungameid/2933620"
                elif launch_battle_net:
                    cod_path = find_battlenet_path()

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
            if any('RUT' in win.title and '.exe' in win.title for win in gw.getWindowsWithTitle('RUT')):
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

    def save_selections():
        selections = {
            'CoD21AllKey': var1.get(),
            'CoD21UAVKey': var2.get(),
            'CoD20AllKey': var3.get(),
            'CoD20UAVKey': var4.get(),
            'XGP': var_xgp.get(),
            'Steam': var_steam.get(),
            'BattleNet': var_battle_net.get(),
        }

        set_selection(selections)
        set_config(config_paths, config_options, config_keys)
        print(f"保存时选择数据: {selections}")  # 调试信息

    def save_lazy_mode():
        config_options['LazyMode'] = var_lazy_mode.get()
        set_config(config_paths, config_options, config_keys)

    def check_selection_logic():
        # 检查密钥的启用状态
        cod21_all_configured = bool(config_keys['CoD21AllKey'])
        cod21_uav_configured = bool(config_keys['CoD21UAVKey'])
        cod20_all_configured = bool(config_keys['CoD20AllKey'])
        cod20_uav_configured = bool(config_keys['CoD20UAVKey'])

        # 根据密钥状态启用/禁用复选框
        cb1.config(state='normal' if cod21_all_configured else 'disabled')
        cb2.config(state='normal' if cod21_uav_configured else 'disabled')
        cb3.config(state='normal' if cod20_all_configured else 'disabled')
        cb4.config(state='normal' if cod20_uav_configured else 'disabled')

        # 复选框互斥逻辑
        if var1.get():
            cb2.config(state='disabled')
            cb3.config(state='disabled')
            cb4.config(state='disabled')

        if var2.get():
            cb1.config(state='disabled')
            cb3.config(state='disabled')
            cb4.config(state='disabled')

        if var3.get():
            cb1.config(state='disabled')
            cb2.config(state='disabled')

        if var4.get():
            cb1.config(state='disabled')
            cb2.config(state='disabled')

        # 平台复选框相互禁用
        if var_xgp.get():
            cb_steam.config(state='disabled')
            cb_battle_net.config(state='disabled')
        elif var_steam.get():
            cb_xgp.config(state='disabled')
            cb_battle_net.config(state='disabled')
        elif var_battle_net.get():
            cb_xgp.config(state='disabled')
            cb_steam.config(state='disabled')
        else:
            cb_xgp.config(state='normal')
            cb_steam.config(state='normal')
            cb_battle_net.config(state='normal')

        # 按钮的启用/禁用逻辑
        keys_selected = var1.get() or var2.get() or var3.get() or var4.get()
        platform_selected = var_xgp.get() or var_steam.get() or var_battle_net.get()
        rut_configured = bool(config_paths['RUT_Path'])

        btn_start_rut_and_cod.config(
            state='normal' if keys_selected and platform_selected and rut_configured else 'disabled'
        )
        btn_start_rut.config(state='normal' if keys_selected and rut_configured else 'disabled')
        btn_start_cod_only.config(state='normal' if platform_selected else 'disabled')

    def on_lazy_mode_check():
        if var_lazy_mode.get():
            msg = "勾选该选项后，下次启动程序时将直接按本次运行时的勾选设置和按钮选项执行\n" \
                  "注意:勾选后下次启动将不会显示UI界面，若需显示主界面以更改配置，请按住键盘左Ctrl键后双击运行本程序。"
            messagebox.showinfo("懒人模式", msg)
        save_lazy_mode()

    def update_platform_checkbuttons():
        selection_data = get_selection()
        var_xgp.set(selection_data.get('XGP', 0))
        var_steam.set(selection_data.get('Steam', 0))
        var_battle_net.set(selection_data.get('战网平台', 0))
        check_selection_logic()

    def update_checkbuttons():
        selection_data = get_selection()
        print(f"更新时选择数据: {selection_data}")  # 调试信息

        for key in ['CoD21AllKey', 'CoD21UAVKey', 'CoD20AllKey', 'CoD20UAVKey']:
            if key not in config_keys:
                config_keys[key] = ''

        for cb, key, var in [(cb1, 'CoD21AllKey', var1), (cb2, 'CoD21UAVKey', var2), (cb3, 'CoD20AllKey', var3),
                             (cb4, 'CoD20UAVKey', var4)]:
            state = 'normal' if config_keys[key] else 'disabled'
            cb.config(state=state)
            var.set(selection_data.get(key, 0))

        # 更新平台复选框状态
        var_xgp.set(config_options.get('XGP', 0))
        var_steam.set(config_options.get('Steam', 0))
        var_battle_net.set(config_options.get('BattleNet', 0))

        check_selection_logic()

    def update_program():
        try:
            global script_dir, config_dir
            progress_window = tk.Tk()
            progress_window.title("更新中")
            tk.Label(progress_window, text="程序正在更新中，完成后将自动重启，请稍候...").pack(padx=10, pady=10)
            progress_window.geometry("300x100")
            progress_window.resizable(False, False)
            progress_window.update()

            github_url = "https://github.com/2eit/AutoCoDSlave/releases/download/AutoCoDSlave-latest/AutoCoDSlave.exe"
            download_exe_path = os.path.join(config_dir, "AutoCoDSlave.exe")
            current_exe_path = sys.argv[0]

            # 如果存在旧的AutoCoDSlave.exe文件，则删除它
            if os.path.exists(download_exe_path):
                os.remove(download_exe_path)
            download_update(github_url, download_exe_path)

            # 计算现有的和新的AutoCoDSlave.exe的哈希值
            current_exe_hash = calculate_hash(current_exe_path)
            new_exe_hash = calculate_hash(download_exe_path)

            # 如果哈希值不同，启动Update.exe
            if current_exe_hash != new_exe_hash:
                update_exe_path = os.path.join(config_dir, "Update.exe")
                if os.path.exists(update_exe_path):
                    os.remove(update_exe_path)
                github_update_url = "https://github.com/2eit/AutoCoDSlave/releases/download/Update-latest/Update.exe"
                download_update(github_update_url, update_exe_path)
                os.execl(update_exe_path, update_exe_path)
            else:
                messagebox.showinfo("更新检查", "已是最新版本，无需更新。")
                os.remove(download_exe_path)

        except Exception as e:
            messagebox.showerror("更新失败", f"更新程序时出现错误: {e}\n{format_exc()}")
        finally:
            progress_window.destroy()


    root.geometry('500x400')
    #root.resizable(False, False)
    # 设置列的重量，让它们按比例分布
    for i in range(11):
        root.grid_columnconfigure(i, weight=1)
    for i in range(15):
        root.grid_rowconfigure(i, weight=1)

    btn_configure = tk.Button(root, text="配置密钥与地址", command=configure_keys_and_paths)
    btn_start_rut_and_cod = tk.Button(root, text="启动RUT并启动游戏", command=lambda: on_btn_clicked('launch_rut_and_cod'))
    btn_start_rut = tk.Button(root, text="仅启动RUT", command=lambda: on_btn_clicked('launch_rut_only'))
    btn_start_cod_only = tk.Button(root, text="我是绿玩😡(仅启动游戏)", command=lambda: on_btn_clicked('start_cod_only'))

    # 按钮的展示
    btn_configure.grid(row=0, column=5)
    btn_start_rut_and_cod.grid(row=1, column=4, sticky='ew')
    btn_start_rut.grid(row=1, column=6, sticky='ew')
    btn_start_cod_only.grid(row=2, column=5)

    var1 = tk.IntVar()
    var2 = tk.IntVar()
    var3 = tk.IntVar()
    var4 = tk.IntVar()
    var_xgp = tk.IntVar()
    var_steam = tk.IntVar()
    var_battle_net = tk.IntVar()
    var_lazy_mode = tk.IntVar()


    cb1 = tk.Checkbutton(root, text="CoD21全解", variable=var1, command=on_cb_toggled)
    cb2 = tk.Checkbutton(root, text="CoD21无限UAV", variable=var2, command=on_cb_toggled)
    cb3 = tk.Checkbutton(root, text="CoD20全解", variable=var3, command=on_cb_toggled)
    cb4 = tk.Checkbutton(root, text="CoD20无限UAV", variable=var4, command=on_cb_toggled)
    cb_xgp = tk.Checkbutton(root, text="XGP平台    ", variable=var_xgp, command=on_cb_toggled)
    cb_steam = tk.Checkbutton(root, text="Steam平台", variable=var_steam, command=on_cb_toggled)
    cb_battle_net = tk.Checkbutton(root, text="战网平台", variable=var_battle_net, command=on_cb_toggled)
    cb_lazy_mode = tk.Checkbutton(root, text="懒人模式", variable=var_lazy_mode, command=on_lazy_mode_check)

    cb1.grid(row=5, column=4, sticky='e')
    cb2.grid(row=5, column=6, sticky='w')
    cb3.grid(row=6, column=4, sticky='e')
    cb4.grid(row=6, column=6, sticky='w')
    cb_xgp.grid(row=8, column=4, sticky='e')
    cb_steam.grid(row=8, column=5)
    cb_battle_net.grid(row=8, column=6, sticky='w')
    cb_lazy_mode.grid(row=10, column=5, sticky='nsew')

    name_label = tk.Label(root, text="--by 2eit", font=('TkDefaultFont', 8), fg="blue", cursor="hand2")
    name_label.grid(row=15, column=6, sticky='e')
    name_label.bind("<Button-1>", open_project_url)
    update_label = tk.Label(root, text="点我更新", fg="blue", cursor="hand2", font=('TkDefaultFont', 10))
    update_label.grid(row=15, column=5, sticky='nsew')
    update_label.bind("<Button-1>", lambda e: update_program())

    var_lazy_mode.set(config_options.get('LazyMode', 0))

    update_checkbuttons()
    update_platform_checkbuttons()
    if config_options.get('LazyMode', 0):
        launch()
        close_autocodslave()

    root.mainloop()

run_as_admin()
init_config()
config_paths, config_options, config_keys = get_config()
selection_data = get_selection()
create_gui()
