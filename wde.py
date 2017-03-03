import pyHook
import pyautogui
import pythoncom


def onMouseEvent(event):
    global count

    if event.Position in ( (0,0), (0, -1), (-1,0), (-1, -1) ):
        count+=1
        if count == 25:
            pyautogui.hotkey("win", "tab")

    elif event.Position[0] > 20 or event.Position[1] > 20:
        count=0

    if event.WindowName is not None:

        if event.WindowName in title_names:
            if event.Wheel == 1:
                pyautogui.hotkey("ctrl", "win", "left")
            elif event.Wheel == -1:
                pyautogui.hotkey("ctrl", "win", "right")
        elif event.WindowName in title_names_for_volume:

            if event.Wheel == 1:
                pyautogui.hotkey("volumeup")
            elif event.Wheel == -1:
                pyautogui.hotkey("volumedown")
            elif event.Message == 520:
                pyautogui.hotkey("volumemute")

    return True

title_names=(
    "Представление задач",
    "Поиск в Windows",
    "Работающие приложения",
    "Пуск",
    "Task View",
    "Start"
)

title_names_for_volume=(
    "Пользовательская область уведомлений"
)

count=0
pyautogui.FAILSAFE = False

hm = pyHook.HookManager()

hm.MouseAll = onMouseEvent
hm.HookMouse()

if __name__ == '__main__':    
    pythoncom.PumpMessages()