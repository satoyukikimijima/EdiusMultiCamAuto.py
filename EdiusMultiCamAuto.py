import pyautogui
import pywinauto
import time
import re
import tkinter as tk
from tkinter import simpledialog
from pywinauto import Desktop

# 設定ファイルの保存先
SETTINGS_FILE = "settings.json"

def load_settings():
    """設定を読み込む。ファイルがなければデフォルト値を返す"""
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            pass
    return {"max_clips": "10", "cam_total": "3", "waiting_str":"0.1"}

def save_settings(max_clips, cam_total):
    """設定をファイルに保存する"""
    settings = {"max_clips": max_clips, "cam_total": cam_total}
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=4)

def get_cam_number(text):
    """
    ファイル名からCAM番号を抽出する
    例: "2025_1212 CAM 2 01.mp4" -> 2
    """
    # 「CAM」の後の数字を検索
    match = re.search(r'CAM\s*(\d+)', text)
    if match:
        return match.group(1)
    return None
def main():
    # --- 1. ユーザー入力セクション ---
    root = tk.Tk()
    root.withdraw()
    
    # カット数の入力
    count_str = simpledialog.askstring("EDIUS Auto", "何カット繰り返しますか？", initialvalue="10")
    if not count_str: return
    max_clips = int(count_str)

    # カメラ台数の入力
    cam_total_str = simpledialog.askstring("EDIUS Auto", "カメラは何台（何トラック）ですか？", initialvalue="5")
    if not cam_total_str: return
    cam_total = int(cam_total_str)
    
    # トラック移動回数の計算 (カメラ台数x2 + 1)
    up_presses = cam_total*2 + 1

    # プロパティ画面のまち時間入力
    waiting_str = simpledialog.askstring("EDIUS Auto", "プロパティの待ち時間(秒)？", initialvalue="0.1")
    if not waiting_str: return
    waiting_time = float(waiting_str)
    sleep_time = waiting_time

    try:
        # 2. EDIUSメインウィンドウに接続
        app = pywinauto.Application(backend="win32").connect(class_name="CtsGuiClass.Frame", title_re=".*EDIUS.*")
        #edius = app.top_window()
        edius = app.window(class_name="CtsGuiClass.Frame", title_re=".*EDIUS.*", found_index=0)
        edius.set_focus()

        for i in range(max_clips):
            #print(f"[{i+1}/{max_clips}] 処理中... (カメラ台数: {cam_total}台)")

            # --- プロパティ読み取り ---
            pyautogui.hotkey('alt', 'enter')
            
            prop_win = None
            for _ in range(10):
                prop_win = Desktop(backend="win32").window(title="プロパティ", class_name="#32770")
                if prop_win.exists(timeout=0.1):
                    break
                time.sleep(0.1)

            if prop_win and prop_win.exists():
                prop_win.set_focus()
                pyautogui.press('left')
                time.sleep(sleep_time)
                
                #try:
                cam_id = prop_win.Edit1.window_text()
                #except:
            		# 取得に失敗した場合は予備手段（クリップボード経由など）も検討可能
            	#print("テキスト取得失敗")
            	#return

        # 5. カメラ番号を判定
                cam_num = get_cam_number(cam_id)
                
                pyautogui.press('esc')
                time.sleep(0.1)
            else:
                print(f"エラー: {i+1}カット目で停止しました。")
                break

            # 3. カメラ切り替え
            if cam_num:
                print(f"  -> CAM {cam_num} を検出")

                pyautogui.press(f'num{cam_num}')
                time.sleep(0.1)

            # --- 4. 【修正ポイント】計算した回数分だけトラックを上に移動 ---
            pyautogui.keyDown('alt')
            pyautogui.press('right')
            # ユーザーが入力したカメラ台数+1 の回数分連打する
            pyautogui.press('up', presses=up_presses, interval=0.01)
            pyautogui.keyUp('alt')
            time.sleep(0.5)
            # 5. 次のクリップへ
            pyautogui.press('s')
            time.sleep(0.3)

        print(f"すべての処理（{max_clips}カット）が完了しました。")

    except Exception as e:
        print(f"エラーが発生しました: {e}")

if __name__ == "__main__":
    main()