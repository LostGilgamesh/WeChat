import win32gui, win32con, time, pyautogui, keyboard
from pywinauto import Application

"""
GitHub：LostGilgamesh
前台运行，会占用鼠标和部分显屏区域。适合不用电脑时启用。
可能会被腾讯登出账户，所以挂机一整天的可行性待确认。
多人点赞过的内容，暂无法精准忽略。
停止方法：关闭朋友圈窗口，按 Esc 或者 Alt + F4 
"""


def Connected_Application(class_name: str = None, window_name: str = None):
    """连接应用
        class_name (str, optional): 类名
        window_name (str, optional): 完整的窗口标题
    Returns:
        tuple(句柄, 应用, 主窗口)
    """
    hwnd = win32gui.FindWindow(class_name, window_name)  # 获取应用的窗口句柄
    app = Application(backend = 'uia').connect(handle = hwnd)  # 使用 uia 后端连接应用
    main_window = app.window(handle = hwnd)  # 获取主窗口
    return hwnd, app, main_window


def Click_Controls(hwnd: int, window_specification, title: str, control_type: str = "Button", area = None):
    """查找并点击控件
    hwnd (int): 窗口句柄
    window_specification (WindowSpecification): 窗口规格
    title (str): 标题
    control_type (str): 控件类型
    area (tuple): 可选的点击区域（left, top, right, bottom）
    """
    control = window_specification.child_window(title=title, control_type=control_type)  # 查找按钮
    if control.exists() and control.is_enabled():  # 如果按钮存在且可用
        win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)  # 激活并显示窗口
        
        if area:  # 如果指定了区域，则检查控件是否在该区域内
            c_rect = control.rectangle()  # 获取控件的矩形区域
            if c_rect.top > area.top + 40 and c_rect.bottom < area.bottom: control.set_focus().click_input()  # 将焦点设置到控件，并点击
        else: control.set_focus().click_input()
        

def Main(cycle_time: int = 600):
    """
    cycle_time (int, optional): 默认600秒刷新一次朋友圈。
    """
    end_time = time.time() + cycle_time
    wechat = Connected_Application('WeChatMainWndForPC', '微信')
    wechat_name = wechat[2].child_window(title = "导航", control_type = "ToolBar").children()[0].window_text()  # 查找工具栏控件，获取第一个子控件的文本
    Click_Controls(wechat[0],  wechat[2], '朋友圈')
    win32gui.ShowWindow(wechat[0], win32con.SW_HIDE)  # 隐藏窗口

    friends_circle = Connected_Application('SnsWnd', '朋友圈')
    while True:  # 根据刷新按钮来确定朋友圈是否打开，并获取使用滚动条时的鼠标停留位置
        if (r_rect := friends_circle[2].child_window(title='刷新', control_type="Button").rectangle()):
            x, y = (r_rect.left + r_rect.right) // 2, ((r_rect.top + r_rect.bottom) // 2) + 50  # 计算按钮中心坐标
            break

    area = friends_circle[2].rectangle()
    while True:
        list_control = friends_circle[2].child_window(title = "朋友圈", control_type = "List")  # 获取列表控件
        items = list_control.children(control_type = "ListItem")  # 获取所有列表项
        
        for text in (_.window_text() for _ in items):
            keywords = ('哭', '崩溃', '节哀')  # 朋友圈文本包含关键字则不点赞。
            if any(_ in text for _ in keywords): continue
            
            control = list_control.child_window(title = text, control_type = "ListItem", found_index = 0)  # 获取第一个对应的列表项控件
            if not control.child_window(title = wechat_name, control_type = "Text").exists():  # 根据微信昵称检查是否点赞过
                Click_Controls(friends_circle[0], control, '评论', area = area)
                time.sleep(0.5)
                Click_Controls(friends_circle[0], friends_circle[2], '赞')
                
            pyautogui.moveTo(x, y)
            pyautogui.scroll(-300)
            
        if end_time < time.time(): Click_Controls(friends_circle[0], friends_circle[2], '刷新')


if __name__ == '__main__': Main()
