import pandas as pd
import pyperclip
import time
import sys
import os
from pynput import keyboard

# --- 全局配置常量 ---
TRIGGER_KEY = keyboard.Key.f9
KEY_PRESS_DELAY = 0.03
TAB_PRESS_DELAY = 0.1
# 手动模式下，终端显示的列表最大长度
MANUAL_DISPLAY_LIMIT = 20
# --- 新增：手动模式下字符串截断的长度 ---
TRUNCATE_FRONT = 20
TRUNCATE_BACK = 10

# --- 全局状态变量 ---
manual_paste_index = 0
manual_page_start_index = 0
data_to_paste = []
listener = None


def clear_screen():
    """清空终端屏幕"""
    os.system('cls' if os.name == 'nt' else 'clear')


def get_data_from_clipboard() -> pd.DataFrame | None:
    """
    从剪贴板读取表格数据并返回一个 DataFrame。
    如果失败则返回 None。
    """
    try:
        df = pd.read_clipboard(header=None, sep='\t')
        if df.empty:
            print("错误：剪贴板为空或格式不正确。")
            return None
        return df
    except Exception as e:
        print(f"读取剪贴板时发生错误: {e}")
        print("请确保您已从 Excel 中复制了一个区域。")
        return None


def display_data(df: pd.DataFrame):
    """
    在终端中显示 DataFrame，如果数据量大则以摘要形式显示。
    """
    print("--- 剪贴板数据预览 ---")
    with pd.option_context('display.max_rows', 10, 'display.max_columns', 10):
        print(df)
    print("------------------------\n")


def get_paste_shortcut():
    """根据操作系统返回正确的粘贴快捷键组合"""
    key_map = {'win32': keyboard.Key.ctrl, 'darwin': keyboard.Key.cmd}
    return key_map.get(sys.platform, keyboard.Key.ctrl)


def get_select_all_shortcut():
    """根据操作系统返回正确的全选快捷键组合"""
    key_map = {'win32': keyboard.Key.ctrl, 'darwin': keyboard.Key.cmd}
    return key_map.get(sys.platform, keyboard.Key.ctrl)


# --- **新增辅助函数** ---
def truncate_middle(text: str, front_len: int, back_len: int) -> str:
    """
    如果文本长度超过前后长度之和，则在中间截断并用...表示。
    """
    text = str(text).replace('\n', ' ').replace('\r', '')
    if len(text) > front_len + back_len:
        return f"{text[:front_len]}...{text[-back_len:]}"
    return text


# --- 模式实现 ---

# --- 手动模式 (V4 - 优化字符串显示) ---

def redraw_manual_ui(page_start_index: int, current_paste_index: int, all_data: list):
    """
    重绘手动模式的用户界面。
    """
    clear_screen()
    trigger_key_name = str(TRIGGER_KEY).split('.')[-1].upper()
    print(f"--- 按 {trigger_key_name} 键粘贴 | 在此窗口按 Ctrl+C 退出 ---")
    print("-" * 50)

    end_index = page_start_index + MANUAL_DISPLAY_LIMIT
    display_list = all_data[page_start_index:end_index]
    marker_pos = current_paste_index - page_start_index

    for i, item in enumerate(display_list):
        # --- **核心修改点** ---
        # 使用新的截断函数来格式化显示内容
        display_item = truncate_middle(item, TRUNCATE_FRONT, TRUNCATE_BACK)

        if i == marker_pos:
            print(f">>> {display_item}")
        else:
            print(f"    {display_item}")


def on_press_manual(key):
    """手动模式下的按键监听回调函数"""
    global manual_paste_index, listener, manual_page_start_index

    if key == TRIGGER_KEY:
        if manual_paste_index < len(data_to_paste):
            current_item = str(data_to_paste[manual_paste_index])

            # 1. 粘贴操作
            pyperclip.copy(current_item)
            time.sleep(KEY_PRESS_DELAY)
            kb_controller = keyboard.Controller()
            modifier_key = get_paste_shortcut()
            with kb_controller.pressed(modifier_key):
                kb_controller.press('v')
                kb_controller.release('v')

            manual_paste_index += 1

            # 2. 检查是否完成
            if manual_paste_index < len(data_to_paste):
                progress_on_page = manual_paste_index - manual_page_start_index
                refresh_threshold = int(MANUAL_DISPLAY_LIMIT * 0.75)

                if progress_on_page >= refresh_threshold:
                    manual_page_start_index = manual_paste_index

                redraw_manual_ui(manual_page_start_index, manual_paste_index, data_to_paste)
            else:
                clear_screen()
                print("\n所有项目已粘贴完毕！")
                listener.stop()
        else:
            listener.stop()


def run_manual_mode():
    """执行手动粘贴模式"""
    global listener, manual_paste_index, manual_page_start_index
    if not data_to_paste:
        print("没有可粘贴的数据。")
        return

    manual_paste_index = 0
    manual_page_start_index = 0

    redraw_manual_ui(manual_page_start_index, manual_paste_index, data_to_paste)

    listener = keyboard.Listener(on_press=on_press_manual)
    listener.start()
    listener.join()


# --- TAB 模式 ---
def run_tab_mode():
    """执行自动 TAB 粘贴模式"""
    delete_first_input = input("是否在每次粘贴前删除输入框内的原有内容? (y/n) [默认: n]: ").lower() or 'n'
    delete_first = delete_first_input == 'y'

    while True:
        tab_count_str = input("请输入每次粘贴后需要按下的 TAB 键次数 [默认: 1]: ") or '1'
        try:
            tab_count = int(tab_count_str)
            if tab_count >= 0:
                break
            else:
                print("请输入一个非负整数。")
        except ValueError:
            print("无效输入，请输入一个数字。")

    trigger_key_name = str(TRIGGER_KEY).split('.')[-1].upper()
    print("\n--- TAB 模式 ---")
    print(f"请将光标移动到起始位置，然后按一次 '{trigger_key_name}' 键开始全自动粘贴。")

    def on_press_start_auto(key):
        if key == TRIGGER_KEY:
            return False

    with keyboard.Listener(on_press=on_press_start_auto) as start_listener:
        start_listener.join()

    print("\n开始自动粘贴...")
    kb_controller = keyboard.Controller()
    modifier_key_paste = get_paste_shortcut()
    modifier_key_select = get_select_all_shortcut()

    for i, item in enumerate(data_to_paste):
        current_item = str(item)

        if delete_first:
            with kb_controller.pressed(modifier_key_select):
                kb_controller.press('a')
                kb_controller.release('a')
            time.sleep(KEY_PRESS_DELAY)
            kb_controller.press(keyboard.Key.delete)
            kb_controller.release(keyboard.Key.delete)
            time.sleep(KEY_PRESS_DELAY)

        pyperclip.copy(current_item)
        time.sleep(KEY_PRESS_DELAY)
        with kb_controller.pressed(modifier_key_paste):
            kb_controller.press('v')
            kb_controller.release('v')

        print(f"  > 已粘贴: {current_item[:30]}")
        time.sleep(TAB_PRESS_DELAY)

        if i < len(data_to_paste) - 1:
            for _ in range(tab_count):
                kb_controller.press(keyboard.Key.tab)
                kb_controller.release(keyboard.Key.tab)
                time.sleep(KEY_PRESS_DELAY)
            time.sleep(TAB_PRESS_DELAY)

    print("\n所有项目已通过 TAB 模式粘贴完毕！")


def main():
    """主函数"""
    global data_to_paste
    clear_screen()
    print("欢迎使用 Excel 区域智能粘贴工具！")
    print("------------------------------------")

    df = get_data_from_clipboard()
    if df is None:
        input("\n按 Enter 键退出。")
        return

    display_data(df)
    data_to_paste = [item for sublist in df.values.tolist() for item in sublist if pd.notna(item)]

    while True:
        mode = input(
            "请选择操作模式:\n  1. 手动模式 (逐个粘贴)\n  2. TAB 模式 (自动跳格粘贴)\n请输入选项 [默认: 1]: ") or '1'
        if mode in ['1', '2']:
            break
        else:
            print("无效输入，请输入 1 或 2。")

    if mode == '1':
        run_manual_mode()
    elif mode == '2':
        run_tab_mode()

    print("\n程序执行完毕。")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n程序被用户中断。正在退出...")
    except Exception as e:
        print(f"\n发生未知错误: {e}")
    finally:
        input("按 Enter 键退出。")