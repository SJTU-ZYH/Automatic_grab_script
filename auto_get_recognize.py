import pyautogui
import time
import xlrd
import pyperclip
import sys

# 定义鼠标事件
# duration类似于移动时间或移动速度，省略后则是瞬间移动到指定的位置
def Mouse(click_times, img_name, retry_times):
    if retry_times == 1:
        location = pyautogui.locateCenterOnScreen(img_name, confidence=0.9)
        if location is not None:
            pyautogui.click(location.x, location.y, clicks=click_times, duration=0.2, interval=0.2)

    elif retry_times == -1:
        while True:
            location = pyautogui.locateCenterOnScreen(img_name,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=click_times, duration=0.2, interval=0.2)
    elif retry_times > 1:
        i = 1
        while i < retry_times + 1:
            location = pyautogui.locateCenterOnScreen(img_name,confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=click_times, duration=0.2, interval=0.2)
                print("重复{}第{}次".format(img_name, i))
                i = i + 1


def Mouse_pos(click_times, x_pos, y_pos):
    pyautogui.click(x=x_pos, y=y_pos, clicks=click_times, interval= 0.01, duration= 0.01, button="left")  # move to 100, 200, then click the left mouse button.

# cell_value     1.0：左键单击
#                2.0：输入字符串
#                3.0：等待
#                4.0：热键

# cell_type      空：0
#                字符串：1
#                数字：2
#                日期：3
#                布尔：4
#                error：5

# 任务一：进行一轮抢课
def WorkFunction1(sheet):
    i = 1
    while i < sheet.nrows:
        # 取excel表格中第i行操作
        cmd_type = sheet.cell_value(i, 1)
        # 1：左键单击
        if cmd_type == 1.0:
            # 获取图片名称
            img_name = sheet.cell_value(i, 2)
            retry_times = 1
            if sheet.cell_type(i, 3) == 2 and sheet.cell_value(i, 3) != 0:
                retry_times = sheet.cell_value(i, 3)
            Mouse(1, img_name, retry_times)
            print("单击左键:{}  Done".format(img_name))

        # 2：输入字符串
        elif cmd_type == 2.0:
            string = sheet.cell_value(i, 2)
            pyperclip.copy(string)
            pyautogui.hotkey('ctrl','v')
            print("输入字符串:{}  Done".format(string))
        # 3：等待
        elif cmd_type == 3.0:
            wait_time = sheet.cell_value(i, 2)
            time.sleep(wait_time)
            print("等待 {} 秒  Done".format(wait_time))
        # 4：键盘热键
        elif cmd_type == 4.0:
            hotkey = sheet.cell_value(i, 2)
            # 防止刷新过快停留在原网页
            time.sleep(1)
            pyautogui.hotkey(hotkey)
            print("按下 {}  Done".format(hotkey))
            time.sleep(1)
        i = i + 1

# 任务二：蹲点等人退课
def WorkFunction2(sheet) :
    while True:
        WorkFunction1(sheet)
        time.sleep(2)


# 任务三：理发预约(只支持早八，循环5次刷新)
def WorkFunction3(sheet) :
    j = 0
    while j < 5:
        i = 1
        # 1和2必须要有
        # 3 8:00-9:00
        # 4 9:00-10:00
        # 5 10:00-11:00
        # 6 11:00-12:00
        # 7 14:00-15:00
        # 8 15:00-16:00
        # 9 16:00-17:00
        # free_time = [1, 2, 3, 4, 5, 6, 7, 8, 9]
        free_time = [1, 2, 7, 8, 9]
        # 读取操作类型
        while i < sheet.nrows:
            if i not in free_time:
                i = i + 1
                continue
            cmd_type = sheet.cell_value(i, 3)
            if cmd_type == 1.0:
                # 取每一个带点击目标的坐标
                x_pos = sheet.cell_value(i, 1)
                y_pos = sheet.cell_value(i, 2)
                print("第{}步：x={},y={} Done".format(i, x_pos, y_pos))
                Mouse_pos(1, x_pos, y_pos)
            elif cmd_type == 0.0:
                # 三秒防刷/等待加载
                sleep_time = sheet.cell_value(i, 4)
                time.sleep(sleep_time)
                sleep_time = ('%.2f' % sleep_time)
                print("等待 {} 秒 Done".format(sleep_time))
            elif cmd_type == 2.0:
                input_value = sheet.cell_value(i, 4)
                pyperclip.copy(input_value)
                pyautogui.hotkey('ctrl', 'v')
                print("输入: {} Done".format(input_value))
            i = i + 1
        j = j + 1
        time.sleep(0.5)


if __name__ == '__main__':
    start_time = time.time()
    file = sys.argv[1]
    type = file.split('_')[1]
    # print(type)
    # file = "info_class.xlsx"
    # file = "info_hair.xlsx"
    # 打开文件
    xr = xlrd.open_workbook(filename=file)
    # 通过索引顺序获取表单
    sheet = xr.sheet_by_index(0)
    print("------欢迎使用自动抓取脚本------")
    print("---------@Hang---------")
    if type == "class.xlsx":
        print("1.抢课一次")
        print("2.蹲点等人退课后抢指定课")
        choice = input(">>")
        start_time = time.time()

        if choice == "1":
            WorkFunction1(sheet)
        elif choice == "2":
            WorkFunction2(sheet)
        else:
            print("非法输入，退出")
    elif type == "hair.xlsx":
        print("3.理发预约")
        choice = input(">>")
        start_time = time.time()

        if choice == "3":
            WorkFunction3(sheet)
        else:
            print("非法输入，退出")
    end_time = time.time()
    time_consume = end_time - start_time
    time_consume = ('%.2f' % time_consume)
    print("耗时 {} 秒".format(time_consume))