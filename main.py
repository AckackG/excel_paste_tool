import pandas as pd
import pyperclip
import time
import sys
import os
from pynput import keyboard

# --- 全局配置常量 ---
# 推荐使用 F1-F12 功能键, 'f9' 表示 F9 键
TRIGGER_KEY = keyboard.Key.f9

# 模拟按键之间的延迟时间（秒），防止目标程序反应不过来
KEY_PRESS_DELAY = 0.1
TAB_PRESS_DELAY = 0.2

# --- 全局状态变量 ---
# 用于手动模式下追踪当前粘贴项的索引
manual_paste_index = 0
# 用于存储待粘贴的数据列表
data_to_paste = []
# 键盘监听器对象
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
        # 使用 pandas 读取，它可以很好地处理 Excel 的制表符分隔格式
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
    # 对于 pynput, 特殊键需要是 Key 枚举或 a-z 字符
    # 'cmd' 是 mac 的 command 键
    key_map = {'win32': keyboard.Key.ctrl, 'darwin': keyboard.Key.cmd}
    # 默认使用 ctrl
    return key_map.get(sys.platform, keyboard.Key.ctrl)


def get_select_all_shortcut():
    """根据操作系统返回正确的全选快捷键组合"""
    key_map = {'win32': keyboard.Key.ctrl, 'darwin': keyboard.Key.cmd}
    return key_map.get(sys.platform, keyboard.Key.ctrl)


# --- 模式实现 ---

# --- 手动模式 ---
def on_press_manual(key):
    """手动模式下的按键监听回调函数"""
    global manual_paste_index, listener

    if key == TRIGGER_KEY:
        if manual_paste_index < len(data_to_paste):
            current_item = str(data_to_paste[manual_paste_index])

            # 1. 将当前项放入剪贴板
            pyperclip.copy(current_item)
            time.sleep(KEY_PRESS_DELAY)  # 等待剪贴板稳定

            # 2. 模拟粘贴
            kb_controller = keyboard.Controller()
            modifier_key = get_paste_shortcut()
            with kb_controller.pressed(modifier_key):
                kb_controller.press('v')
                kb_controller.release('v')

            manual_paste_index += 1

            # 3. 更新提示
            if manual_paste_index < len(data_to_paste):
                next_item = str(data_to_paste[manual_paste_index])
                # 使用 \r 实现原地更新，避免刷屏
                sys.stdout.write(f"\r准备粘贴: [{next_item[:20]}]... 按 F9 继续。")
                sys.stdout.flush()
            else:
                print("\n\n所有项目已粘贴完毕！")
                listener.stop()  # 停止监听
        else:
            listener.stop()


def run_manual_mode():
    """执行手动粘贴模式"""
    global listener
    print("--- 手动模式 ---")
    print(f"请将光标移动到目标位置，每按一次 '{str(TRIGGER_KEY).split('.')[-1].upper()}' 键将粘贴一个项目。")

    if not data_to_paste:
        print("没有可粘贴的数据。")
        return

    first_item = str(data_to_paste[0])
    print(f"准备粘贴: [{first_item[:20]}]...")

    # 创建并启动监听器
    listener = keyboard.Listener(on_press=on_press_manual)
    listener.start()
    listener.join()  # 阻塞主线程直到监听器停止


# --- TAB 模式 ---
def run_tab_mode():
    """执行自动 TAB 粘贴模式"""
    # 1. 获取用户参数
    delete_first_input = input("是否在每次粘贴前删除输入框内的原有内容? (y/n) [默认: n]: ").lower()
    delete_first = delete_first_input == 'y'

    while True:
        try:
            tab_count = int(input("请输入每次粘贴后需要按下的 TAB 键次数 (例如: 1): "))
            if tab_count >= 0:
                break
            else:
                print("请输入一个非负整数。")
        except ValueError:
            print("无效输入，请输入一个数字。")

    print("\n--- TAB 模式 ---")
    print(f"请将光标移动到起始位置，然后按一次 '{str(TRIGGER_KEY).split('.')[-1].upper()}' 键开始全自动粘贴。")

    # 使用一个简单的监听器来触发整个流程
    def on_press_start_auto(key):
        if key == TRIGGER_KEY:
            return False  # 返回 False 来停止监听器

    with keyboard.Listener(on_press=on_press_start_auto) as start_listener:
        start_listener.join()

    # 监听器已停止，开始执行自动化粘贴
    print("\n开始自动粘贴...")
    kb_controller = keyboard.Controller()
    modifier_key_paste = get_paste_shortcut()
    modifier_key_select = get_select_all_shortcut()

    for i, item in enumerate(data_to_paste):
        current_item = str(item)

        # 1. (可选) 删除原有内容
        if delete_first:
            with kb_controller.pressed(modifier_key_select):
                kb_controller.press('a')
                kb_controller.release('a')
            time.sleep(KEY_PRESS_DELAY)
            kb_controller.press(keyboard.Key.delete)
            kb_controller.release(keyboard.Key.delete)
            time.sleep(KEY_PRESS_DELAY)

        # 2. 粘贴
        pyperclip.copy(current_item)
        time.sleep(KEY_PRESS_DELAY)
        with kb_controller.pressed(modifier_key_paste):
            kb_controller.press('v')
            kb_controller.release('v')

        print(f"  > 已粘贴: {current_item[:30]}")
        time.sleep(TAB_PRESS_DELAY)

        # 3. 按 TAB，如果不是最后一项
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

    # 1. 读取数据
    df = get_data_from_clipboard()
    if df is None:
        input("\n按 Enter 键退出。")
        return

    # 2. 展示数据
    display_data(df)

    # 将 DataFrame 展平为列表，按行优先
    data_to_paste = [item for sublist in df.values.tolist() for item in sublist]

    # 3. 选择模式
    while True:
        mode = input("请选择操作模式:\n  1. 手动模式 (逐个粘贴)\n  2. TAB 模式 (自动跳格粘贴)\n请输入选项 (1/2): ")
        if mode in ['1', '2']:
            break
        else:
            print("无效输入，请输入 1 或 2。")

    # 4. 执行对应模式
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