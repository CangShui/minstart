import time
import glob
import os
import psutil
import win32gui
import win32con
import win32process
import win32com.client
import json

# ---------------- 配置部分 ----------------
CONFIG_PATH = r"C:\Users\Public\minstart.json"
CONFIG_VERSION = "1.0.0.3"

DEFAULT_CONFIG = {
    "config_version": CONFIG_VERSION,
    "lnk_dir": r"C:\Users\Public\lnk",
    "monitor_seconds": 10,   # 总共监控多久
    "scan_interval": 0.5,      # 每隔多少秒扫描一次
}

def compare_versions(v1, v2):
    v1p = [int(x) for x in v1.split('.')]
    v2p = [int(x) for x in v2.split('.')]
    m = max(len(v1p), len(v2p))
    v1p += [0] * (m - len(v1p))
    v2p += [0] * (m - len(v2p))
    for a, b in zip(v1p, v2p):
        if a > b: return 1
        if a < b: return -1
    return 0

def save_config(cfg):
    os.makedirs(os.path.dirname(CONFIG_PATH), exist_ok=True)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=4, ensure_ascii=False)

def load_config():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        # 如果脚本里的版本更高，就给老配置补齐新字段
        if compare_versions(CONFIG_VERSION, cfg.get("config_version", "0")) > 0:
            for k, v in DEFAULT_CONFIG.items():
                if k not in cfg:
                    cfg[k] = v
            cfg["config_version"] = CONFIG_VERSION
            save_config(cfg)
        return cfg
    else:
        save_config(DEFAULT_CONFIG.copy())
        return DEFAULT_CONFIG.copy()

# 载入配置
CONFIG = load_config()
LNK_DIR = CONFIG.get("lnk_dir", r"C:\Users\Public\lnk")
MONITOR_SECONDS = CONFIG.get("monitor_seconds", 60)
SCAN_INTERVAL = CONFIG.get("scan_interval", 1)
# ---------------- 配置部分结束 ----------------


def resolve_lnk_target(lnk_path):
    """用 WScript.Shell 解析 .lnk 目标路径"""
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(lnk_path)
        target = shortcut.TargetPath
        if target:
            return target
    except Exception:
        pass
    return None

def build_target_processes_from_lnks():
    pattern = os.path.join(LNK_DIR, "*.lnk")
    lnk_files = glob.glob(pattern)
    if not lnk_files:
        print(f"{LNK_DIR} 下没有找到任何 .lnk")
        return set(), []

    target_processes = set()
    print("解析快捷方式目标：")
    for lnk in lnk_files:
        target = resolve_lnk_target(os.path.abspath(lnk))
        if target:
            exe_name = os.path.basename(target)
            target_processes.add(exe_name)
            print(f" - {lnk} -> {target} -> 进程名：{exe_name}")
        else:
            print(f" - {lnk} -> 无法解析目标（可能是特殊类型快捷方式）")

    if not target_processes:
        print("没有解析出任何 exe，后续将不会最小化任何窗口。")

    return target_processes, lnk_files

def get_hwnds_by_pid(pid):
    hwnds = []
    def callback(hwnd, acc):
        try:
            _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
            if found_pid == pid and win32gui.IsWindowVisible(hwnd):
                acc.append(hwnd)
        except Exception:
            pass
        return True
    win32gui.EnumWindows(callback, hwnds)
    return hwnds

def minimize_windows_of_pid(pid):
    hwnds = get_hwnds_by_pid(pid)
    hit = False
    for hwnd in hwnds:
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
            hit = True
        except Exception:
            pass
    return hit

def launch_all_shortcuts(lnk_files):
    print("并发启动以下快捷方式：")
    for lnk in lnk_files:
        print(" -", lnk)
        try:
            os.startfile(lnk)
        except Exception as e:
            print(f"启动失败 {lnk}: {e}")

def monitor_and_minimize_by_targets(target_processes):
    if not target_processes:
        print("TARGET_PROCESSES 为空，跳过监控。")
        return

    print(f"开始监控以下进程的窗口并最小化（持续 {MONITOR_SECONDS} 秒，间隔 {SCAN_INTERVAL} 秒）：")
    for name in target_processes:
        print(" -", name)

    end_time = time.time() + MONITOR_SECONDS

    while time.time() < end_time:
        for p in psutil.process_iter(["pid", "name"]):
            name = p.info.get("name")
            pid = p.info.get("pid")
            if not name or not pid:
                continue
            if name not in target_processes:
                continue

            if minimize_windows_of_pid(pid):
                print(f"已尝试最小化 {name} (PID={pid}) 的窗口")
        time.sleep(SCAN_INTERVAL)

    print("监控结束。")

def main():
    target_processes, lnk_files = build_target_processes_from_lnks()
    if not lnk_files:
        return

    launch_all_shortcuts(lnk_files)
    monitor_and_minimize_by_targets(target_processes)

if __name__ == "__main__":
    main()
