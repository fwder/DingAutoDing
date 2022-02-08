import time
import win32gui
import win32api
import win32con
import pyautogui
import cv2
import numpy
import pytesseract
from PIL import Image, ImageGrab

hwnd_title = {}


def get_all_hwnd(hwnd, mouse):
    if (win32gui.IsWindow(hwnd) and
            win32gui.IsWindowEnabled(hwnd) and
            win32gui.IsWindowVisible(hwnd)):
        hwnd_title.update({hwnd: win32gui.GetWindowText(hwnd)})


def get_all_child_window(parent):
    if not parent:
        return
    hwndChildList = []
    win32gui.EnumChildWindows(
        parent, lambda hwnd, param: param.append(hwnd), hwndChildList)
    return hwndChildList


# 查找所有窗口标题和句柄
win32gui.EnumWindows(get_all_hwnd, 0)
ding_list = []
summ = 0

# 遍历一遍所有的钉钉并记录，之后不再开启
print("成功获取到钉钉客户端句柄：")
for h, t in hwnd_title.items():
    if t == '钉钉':
        print(h)
        summ += 1
        ding_list.append(h)

if summ == 0:
    exit(0)

while True:

    try:
        # 在前面先用ding_list[0]钉钉句柄检测
        # 将窗口调到前台
        win = ding_list[0]  # 获取当前钉钉句柄
        win32gui.ShowWindow(win, win32con.SW_SHOWNORMAL)
        win32gui.SetForegroundWindow(win)  # show window
        time.sleep(1)
        left, top, right, botton = win32gui.GetWindowRect(win)  # 获得窗口的位置，得到距离左上的位置以及框体的宽高
        print(str(left))
        print(str(top))
        print(str(right))
        print(str(botton))
        left += 480
        top += 108
        right -= 100
        botton -= 550
        if left < 0:
            left = 1
        if top < 0:
            top = 1
        if right < 0:
            right = 1
        if botton < 0:
            botton = 1
        # pic = pg.screenshot(region=(left, top, right, botton))
        pic = ImageGrab.grab((left, top, right, botton))
        img = cv2.cvtColor(numpy.array(pic), cv2.COLOR_RGB2BGR)
        cv2.waitKey(100)
        image = Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
        # image.save("./result.jpg")
        string = pytesseract.image_to_string(image, lang='chi_sim')
        print("识别结果：" + string)
        if len(string) != 0:
            if string.find('[') != -1 or string.find(']') != -1:
                print("检测到有直播")
                time.sleep(1)
            else:
                print("没检查到直播")
                time.sleep(60)
                continue
        else:
            print("直接就没有识别出来文字")
            time.sleep(60)
            continue
    except:
        print("检查直播出错")
        continue

    # 开始遍历
    try:
        for ding in ding_list:
            # 将窗口调到前台
            win = ding  # 获取当前钉钉句柄
            win32gui.ShowWindow(win, win32con.SW_SHOWNORMAL)
            win32gui.SetForegroundWindow(win)  # show window
            # 实现窗口移动
            left, top, right, botton = win32gui.GetWindowRect(win)  # 获得窗口的位置，得到距离左上的位置以及框体的宽高
            move_x = left + 200  # 设置鼠标坐标x值。坐标增加偏移值，使得鼠标位于可以拖动的框体的位置上
            move_y = top + 20  # 设置鼠标坐标y值。坐标增加偏移值，使得鼠标位于可以拖动的框体的位置上
            win32api.SetCursorPos((move_x, move_y))  # 鼠标挪到窗口所在坐标
            time.sleep(2)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)  # 鼠标左键按下
            win32api.SetCursorPos(
                (600, 180))  # 鼠标左键按下的同时移动鼠标位置，实现拖动框体，这里是要移动到左上角，但是不能写（0,0），（0,0）+（x偏移值，y偏移值），确保框体的左上角在窗口的左上角
            time.sleep(1)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)  # 鼠标左键抬起
            time.sleep(1)
            left, top, right, botton = win32gui.GetWindowRect(win)  # 再次获得窗口的位置
            move_x = left + 600  # 设置鼠标坐标x值。坐标增加偏移值，使得鼠标位于可以拖动的框体的位置上
            move_y = top + 120  # 设置鼠标坐标y值。坐标增加偏移值，使得鼠标位于可以拖动的框体的位置上
            win32api.SetCursorPos((move_x, move_y))  # 鼠标挪到窗口所在坐标
            time.sleep(1)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)  # 鼠标左键按下
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)  # 鼠标左键抬起
            time.sleep(1)
            win32gui.CloseWindow(win)
            time.sleep(1)

        # 都完事之后把鼠标移动到左上角
        win32api.SetCursorPos((0, 0))
        time.sleep(60)
    except:
        print("进入直播出错")


