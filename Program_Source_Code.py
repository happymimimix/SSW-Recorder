# Import required modules
import os
import threading
import time
import keyboard
import mouse
import playsound
import win32api
import win32con
import win32file
import win32gui
import win32pipe
import win32ui
import win32console
from numpy import frombuffer
import six

# Init variables
Flipping = False
Prev_Flipping = False
Unloading = False
MisrecognitionProtect = 0
ProtectionLevel = 2
Show_Watermark = True
Show_ControlPanel = True
Show_Statistics = True
Toolbox_Open = False
Prev_MouseClick = False
SH_Override0 = 0
SH_Override1 = 0
Timer_Reset = 0
Play_Reset = 0
Recording_Speed = 1
Maximum_Depth = 0
OffScreenDC = win32gui.CreateCompatibleDC(win32gui.GetWindowDC(win32gui.GetDesktopWindow()))
OffScreenBMP = win32gui.CreateCompatibleBitmap(win32gui.GetWindowDC(win32gui.GetDesktopWindow()), win32gui.GetClientRect(win32gui.GetDesktopWindow())[2], win32gui.GetClientRect(win32gui.GetDesktopWindow())[3])
win32gui.SelectObject(OffScreenDC, OffScreenBMP)

# Create pipes for communicating with cheat engine
PipeSSW = win32pipe.CreateNamedPipe(
    r'\\.\pipe\SSW_Recorder_Pipe_SSW',
    win32pipe.PIPE_ACCESS_OUTBOUND,
    win32pipe.PIPE_TYPE_MESSAGE | win32pipe.PIPE_WAIT,
    1, 65536, 65536,
    0,
    None
)

PipeOBS = win32pipe.CreateNamedPipe(
    r'\\.\pipe\SSW_Recorder_Pipe_OBS',
    win32pipe.PIPE_ACCESS_OUTBOUND,
    win32pipe.PIPE_TYPE_MESSAGE | win32pipe.PIPE_WAIT,
    1, 65536, 65536,
    0,
    None
)


def auto_start():
    # Play song after a short delay
    while not Prev_Flipping:
        time.sleep(1 / 128)
    time.sleep(1)
    keyboard.press_and_release('alt')
    keyboard.press_and_release('p')
    keyboard.press_and_release('s')


def send_signal(signal):
    # Send signal to cheat engine through pipes
    if not UI_Debug_Mode:
        win32file.WriteFile(PipeSSW, signal.encode())
        win32file.WriteFile(PipeOBS, signal.encode())


def Format_Time(sec):
    min = sec // 60
    hr = min // 60
    sec = sec % 60
    min = min % 60
    return format(hr, '02.0f') + ':' + format(min, '02.0f') + ':' + format(sec, '05.2f')


def get_pixel_color(hwnd, x, y):
    # Grab the color of a specific pixel
    hDC = win32gui.GetWindowDC(hwnd)
    try:
        color = win32gui.GetPixel(hDC, x, y)
    except:
        return [255, 0, 0]
    win32gui.ReleaseDC(0, hDC)
    r = color & 0xff
    g = (color >> 8) & 0xff
    b = (color >> 16) & 0xff
    return [r, g, b]


def darkest_pixel(hwnd, x1, y1, x2, y2):
    # Find the brightness of the darkest pixel within selected range
    w = x2 - x1
    h = y2 - y1
    wDC = win32gui.GetWindowDC(hwnd)
    dcObj = win32ui.CreateDCFromHandle(wDC)
    cDC = dcObj.CreateCompatibleDC()
    dataBitMap = win32ui.CreateBitmap()
    dataBitMap.CreateCompatibleBitmap(dcObj, w, h)
    cDC.SelectObject(dataBitMap)
    # Take screenshot
    cDC.BitBlt((0, 0), (w, h), dcObj, (x1, y1), win32con.SRCCOPY)
    signedIntsArray = dataBitMap.GetBitmapBits(True)
    img = frombuffer(signedIntsArray, dtype='uint8')
    img = img.astype('int32')
    # Generate a list of colors
    img.shape = (h, w, 4)
    dcObj.DeleteDC()
    cDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, wDC)
    win32gui.DeleteObject(dataBitMap.GetHandle())
    # Return the brightness of the darkest pixel
    return int(((img[..., 2] + img[..., 1] + img[..., 0]) / 3).min())


def InRange(X, Y, x1, y1, x2, y2):
    # Check if the specified position is in the range
    return True if min(x1, x2) <= X <= max(x1, x2) and min(y1, y2) <= Y <= max(y1, y2) else False


def Process_Mouse_Input(hwnd, X, Y, W, H):
    global Unloading
    global Show_ControlPanel
    global Show_Watermark
    global Show_Statistics
    global Toolbox_Open
    global SH_Override0
    global SH_Override1
    global Timer_Reset
    global Play_Reset
    global Recording_Speed
    global ProtectionLevel
    global start_time
    if win32gui.GetForegroundWindow() == MainWindowHWND:
        print("Mouse clicked!  @Pos: (" + str(X) + "," + str(Y) + ")", end='')
        if Unloading:
            if InRange(X, Y, 650, H - 230, 750, H - 180):
                print("  |Button: Cancel unload", end='')
                playsound.playsound(".\\Resources\\KDE_Click.wav")
                draw_rectangle(hwnd, 100, H - 300, 800, H - 150, get_pixel_color(hwnd, 2, 20))
                Unloading = False
            if InRange(X, Y, 500, H - 230, 600, H - 180):
                print("  |Button: Confirm unload", end='')
                playsound.playsound(".\\Resources\\KDE_Click.wav")
                draw_rectangle(hwnd, 100, H - 300, 800, H - 150, [63, 63, 63])
                draw_text(hwnd, 200, H - 200, 50, [0, 191, 191], "Plugin unload success! ", 'L', 'B')
                send_signal('1\n')
                send_signal('1\n')
                send_signal('1\n')
                send_signal('Finished\n')
                time.sleep(1 / 32)
                win32pipe.DisconnectNamedPipe(PipeSSW)
                win32pipe.DisconnectNamedPipe(PipeOBS)
                win32file.CloseHandle(PipeSSW)
                win32file.CloseHandle(PipeOBS)
                win32gui.DeleteObject(OffScreenBMP)
                win32gui.DeleteDC(OffScreenDC)
                print('\033[0m', end='')
                playsound.playsound(".\\Resources\\KDE_Logout.wav")
                raise SystemExit
        else:
            if Show_ControlPanel:
                if InRange(X, Y, 35, H - 42, 85, H - 25):
                    print("  |Button: Hide control panel", end='')
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                    draw_rectangle(hwnd, 300, H - 200, 15, H - 15, get_pixel_color(hwnd, 2, 20))
                    Show_ControlPanel = False
                    if Prev_Flipping:
                        draw_rectangle(hwnd, 310, H - 20, 830, H - 50, get_pixel_color(hwnd, 2, 20))
                        draw_rectangle(hwnd, 10, H - 20, 15, H - 50, [255, 0, 0])
                        draw_text(hwnd, 20, H - 20, 30, [255, 0, 0], "Page flipping, recording paused... ", 'L', 'B')
                    if Toolbox_Open:
                        draw_rectangle(hwnd, 315, H - 200, 575, H - 56, get_pixel_color(hwnd, 2, 20))
                if InRange(X, Y, 155, H - 42, 205, H - 25):
                    print("  |Button: Unload", end='')
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                    Unloading = True
                    threading.Thread(target=playsound.playsound, args=(".\\Resources\\KDE_Error.wav",)).start()
                if InRange(X, Y, 95, H - 42, 145, H - 25):
                    print("  |Button: Toolbox", end='')
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                    if Toolbox_Open:
                        draw_rectangle(hwnd, 315, H - 200, 575, H - 56, get_pixel_color(hwnd, 2, 20))
                    Toolbox_Open = not Toolbox_Open
                if InRange(X, Y, 215, H - 42, 285, H - 25):
                    print("  |Button: Reset timer", end='')
                    start_time = time.time()
                    Timer_Reset = 5
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                if InRange(X, Y, 215, H - 63, 285, H - 47):
                    print("  |Button: Play and reset", end='')
                    send_signal(str(Recording_Speed) + '\n')
                    start_time = time.time()
                    Play_Reset = 5
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                if InRange(X, Y, 175, H - 142, 215, H - 156):
                    print("  |Button: Play", end='')
                    send_signal(str(Recording_Speed) + '\n')
                    SH_Override1 = 5
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                if InRange(X, Y, 225, H - 142, 265, H - 156):
                    print("  |Button: Pause", end='')
                    send_signal('0\n')
                    SH_Override0 = 5
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                if InRange(X, Y, 125, H - 138, 159, H - 122):
                    print("  |Button: 1.0x recording speed", end='')
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                    Recording_Speed = 1
                    if not Prev_Flipping:
                        send_signal('1\n')
                if InRange(X, Y, 170, H - 138, 204, H - 122):
                    print("  |Button: 0.5x recording speed", end='')
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                    Recording_Speed = 2
                    if not Prev_Flipping:
                        send_signal('2\n')
                if InRange(X, Y, 215, H - 138, 249, H - 122):
                    print("  |Button: 0.2x recording speed", end='')
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                    Recording_Speed = 5
                    if not Prev_Flipping:
                        send_signal('5\n')
                if InRange(X, Y, 260, H - 138, 294, H - 122):
                    print("  |Button: 0.1x recording speed", end='')
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                    Recording_Speed = 10
                    if not Prev_Flipping:
                        send_signal('10\n')
                if InRange(X, Y, 35, H - 103, 65, H - 87):
                    print("  |Button: Missrecognition protection level 0", end='')
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                    ProtectionLevel = 0
                if InRange(X, Y, 75, H - 103, 105, H - 87):
                    print("  |Button: Missrecognition protection level 1", end='')
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                    ProtectionLevel = 1
                if InRange(X, Y, 115, H - 103, 145, H - 87):
                    print("  |Button: Missrecognition protection level 2", end='')
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                    ProtectionLevel = 2
                if InRange(X, Y, 155, H - 103, 185, H - 87):
                    print("  |Button: Missrecognition protection level 3", end='')
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                    ProtectionLevel = 3
                if InRange(X, Y, 195, H - 103, 225, H - 87):
                    print("  |Button: Missrecognition protection level 4", end='')
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                    ProtectionLevel = 4
                if InRange(X, Y, 235, H - 103, 265, H - 87):
                    print("  |Button: Missrecognition protection level 5", end='')
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                    ProtectionLevel = 5
                if InRange(X, Y, 130, H - 83, 166, H - 67):
                    print("  |Button: Enable watermark", end='')
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                    Show_Watermark = True
                if InRange(X, Y, 171, H - 83, 207, H - 67):
                    print("  |Button: Disable watermark", end='')
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                    draw_rectangle(hwnd, W - 500, 25, W - 15, 85, get_pixel_color(hwnd, 2, 20))
                    Show_Watermark = False
                if InRange(X, Y, 120, H - 63, 156, H - 47):
                    print("  |Button: Enable statistics", end='')
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                    Show_Statistics = True
                if InRange(X, Y, 161, H - 63, 197, H - 47):
                    print("  |Button: Disable statistics", end='')
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                    draw_rectangle(hwnd, W - 265, H - 250, W - 15, H - 15, get_pixel_color(hwnd, 2, 20))
                    Show_Statistics = False
            elif InRange(X, Y, 0, H - 23, 50, H - 6):
                print("  |Button: Show control panel", end='')
                playsound.playsound(".\\Resources\\KDE_Click.wav")
                draw_rectangle(hwnd, 0, H - 23, 50, H - 6, get_pixel_color(hwnd, 2, 20))
                Show_ControlPanel = True
                if Prev_Flipping:
                    draw_rectangle(hwnd, 10, H - 20, 530, H - 50, get_pixel_color(hwnd, 2, 20))
                    draw_rectangle(hwnd, 310, H - 20, 315, H - 50, [255, 0, 0])
                    draw_text(hwnd, 320, H - 20, 30, [255, 0, 0], "Page flipping, recording paused... ", 'L', 'B')
            if Toolbox_Open:
                if InRange(X, Y, 325 - (0 if Show_ControlPanel else 300), H - 160, 441 - (0 if Show_ControlPanel else 300), H - 140):
                    print("  |Button: Toolbox - MTC quick fix", end='')
                    draw_rectangle(hwnd, 325 - (0 if Show_ControlPanel else 300), H - 160, 441 - (0 if Show_ControlPanel else 300), H - 140, [127, 127, 247])
                    mouse.wait(mouse.LEFT, mouse.UP)
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                    keyboard.press_and_release('alt')
                    keyboard.press_and_release('p')
                    keyboard.press_and_release('t')
                    keyboard.press_and_release('alt')
                    keyboard.press_and_release('s')
                    keyboard.press_and_release('p')
                    keyboard.press_and_release('s')
                    keyboard.press_and_release('n')
                    keyboard.press_and_release('return')
                    keyboard.press_and_release('alt')
                    keyboard.press_and_release('s')
                    keyboard.press_and_release('p')
                    keyboard.press_and_release('s')
                    keyboard.press_and_release('s')
                    keyboard.press_and_release('return')
                    threading.Thread(target=auto_start, args=()).start()
                if InRange(X, Y, 449 - (0 if Show_ControlPanel else 300), H - 160, 565 - (0 if Show_ControlPanel else 300), H - 140):
                    print("  |Button: Toolbox - Start over", end='')
                    draw_rectangle(hwnd, 449 - (0 if Show_ControlPanel else 300), H - 160, 565 - (0 if Show_ControlPanel else 300), H - 140, [127, 127, 247])
                    mouse.wait(mouse.LEFT, mouse.UP)
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                    keyboard.press_and_release('alt')
                    keyboard.press_and_release('p')
                    keyboard.press_and_release('t')
                    keyboard.press_and_release('alt')
                    keyboard.press_and_release('s')
                    keyboard.press_and_release('p')
                    keyboard.press_and_release('s')
                    keyboard.press_and_release('n')
                    keyboard.press_and_release('return')
                    keyboard.press_and_release('home')
                    global Version
                    if int(Version[0]) > 1 and int(Version[0]) < 1:
                        print("Note: We think your version of Singer Song Writer might not require MTC='Slave' for speed hacking. \nIf that isn't the case, click 'MTC Quick Fix' in the toolbox to fix it. ")
                    else:
                        keyboard.press_and_release('alt')
                        keyboard.press_and_release('s')
                        keyboard.press_and_release('p')
                        keyboard.press_and_release('s')
                        keyboard.press_and_release('s')
                        keyboard.press_and_release('return')
                    threading.Thread(target=auto_start, args=()).start()
                    start_time = time.time()
                if InRange(X, Y, 325 - (0 if Show_ControlPanel else 300), H - 135, 441 - (0 if Show_ControlPanel else 300), H - 115):
                    print("  |Button: Toolbox - Playback controls", end='')
                    draw_rectangle(hwnd, 325 - (0 if Show_ControlPanel else 300), H - 135, 441 - (0 if Show_ControlPanel else 300), H - 115, [127, 127, 247])
                    mouse.wait(mouse.LEFT, mouse.UP)
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                    keyboard.press_and_release('alt')
                    keyboard.press_and_release('w')
                    keyboard.press_and_release('e')
                    keyboard.press_and_release('p')
                if InRange(X, Y, 449 - (0 if Show_ControlPanel else 300), H - 135, 565 - (0 if Show_ControlPanel else 300), H - 115):
                    print("  |Button: Toolbox - Score type settings", end='')
                    draw_rectangle(hwnd, 449 - (0 if Show_ControlPanel else 300), H - 135, 565 - (0 if Show_ControlPanel else 300), H - 115, [127, 127, 247])
                    mouse.wait(mouse.LEFT, mouse.UP)
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                    keyboard.press_and_release('alt')
                    keyboard.press_and_release('p')
                    keyboard.press_and_release('t')
                    keyboard.press_and_release('home')
                    keyboard.press_and_release('alt')
                    keyboard.press_and_release('s')
                    keyboard.press_and_release('s')
                    keyboard.press_and_release('s')
                    keyboard.press_and_release('return')
                if InRange(X, Y, 325 - (0 if Show_ControlPanel else 300), H - 110, 441 - (0 if Show_ControlPanel else 300), H - 90):
                    print("  |Button: Toolbox - Track merger", end='')
                    draw_rectangle(hwnd, 325 - (0 if Show_ControlPanel else 300), H - 110, 441 - (0 if Show_ControlPanel else 300), H - 90, [127, 127, 247])
                    mouse.wait(mouse.LEFT, mouse.UP)
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                    os.system(".\\Resources\\CJC_Toolbox.exe CJCAMM en-us")
                if InRange(X, Y, 449 - (0 if Show_ControlPanel else 300), H - 110, 565 - (0 if Show_ControlPanel else 300), H - 90):
                    print("  |Button: Toolbox - Performance monitor", end='')
                    draw_rectangle(hwnd, 449 - (0 if Show_ControlPanel else 300), H - 110, 565 - (0 if Show_ControlPanel else 300), H - 90, [127, 127, 247])
                    mouse.wait(mouse.LEFT, mouse.UP)
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                    keyboard.press_and_release('alt')
                    keyboard.press_and_release('h')
                    keyboard.press_and_release('p')
                if InRange(X, Y, 325 - (0 if Show_ControlPanel else 300), H - 85, 441 - (0 if Show_ControlPanel else 300), H - 65):
                    print("  |Button: Toolbox - Save as *.ss10 file", end='')
                    draw_rectangle(hwnd, 325 - (0 if Show_ControlPanel else 300), H - 85, 441 - (0 if Show_ControlPanel else 300), H - 65, [127, 127, 247])
                    mouse.wait(mouse.LEFT, mouse.UP)
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                    keyboard.press_and_release('alt')
                    keyboard.press_and_release('p')
                    keyboard.press_and_release('t')
                    keyboard.press_and_release('alt')
                    keyboard.press_and_release('f')
                    keyboard.press_and_release('a')
                if InRange(X, Y, 449 - (0 if Show_ControlPanel else 300), H - 85, 565 - (0 if Show_ControlPanel else 300), H - 65):
                    print("  |Button: Toolbox - Equal width bar", end='')
                    draw_rectangle(hwnd, 449 - (0 if Show_ControlPanel else 300), H - 85, 565 - (0 if Show_ControlPanel else 300), H - 65, [127, 127, 247])
                    mouse.wait(mouse.LEFT, mouse.UP)
                    playsound.playsound(".\\Resources\\KDE_Click.wav")
                    keyboard.press_and_release('alt')
                    keyboard.press_and_release('d')
                    keyboard.press_and_release('g')
                    keyboard.press_and_release('s')
            print()
    else:
        print("Window in background, mouse clicks are ignored. ")


def Main_Program(hwnd, W, H, CX, CY, Clicked):
    global Unloading
    global Show_Statistics
    if Unloading:
        draw_rectangle_offscreen(100, H - 300, 800, H - 150, [63, 63, 63])
        draw_text_offscreen(120, H - 280, 30, [255, 255, 255], "Are you sure you want to unload this plugin? ", 'L', 'T')
        draw_rectangle_offscreen(650, H - 210, 750, H - 170, [191, 0, 0])
        draw_rectangle_offscreen(500, H - 210, 600, H - 170, [0, 191, 0])
        draw_text_offscreen(550, H - 175, 24, [255, 255, 255], "YES", 'M', 'B')
        draw_text_offscreen(700, H - 175, 24, [255, 255, 255], "NO", 'M', 'B')
        copy_offscreen_content(hwnd, 100, H - 300, 800, H - 150)
    else:
        global Flipping
        global Prev_Flipping
        global MisrecognitionProtect
        global SH_Override0
        global SH_Override1
        global Timer_Reset
        global Play_Reset
        # Detect page flips
        Brightness = darkest_pixel(hwnd, 2, H - 3, W - 4, H - 2)
        if Brightness > 127 and not (Prev_Flipping and get_pixel_color(hwnd, (312 if Show_ControlPanel else 12), H - 30) != [255, 0, 0]):
            # Flip detected
            Flipping = True
        else:
            Flipping = False

        # Check if pageflip is being detected for a few times in a row, prevents misrecognition.
        if Flipping and not Prev_Flipping:
            if MisrecognitionProtect < ProtectionLevel:
                MisrecognitionProtect += 1
            else:
                draw_rectangle(hwnd, (310 if Show_ControlPanel else 10), H - 20, (315 if Show_ControlPanel else 15), H - 50, [255, 0, 0])
                draw_text(hwnd, (320 if Show_ControlPanel else 20), H - 20, 30, [255, 0, 0], "Page flipping, recording paused... ", 'L', 'B')
                send_signal('0\n')
                MisrecognitionProtect = -1
        if not Flipping and Prev_Flipping:
            draw_rectangle(hwnd, (310 if Show_ControlPanel else 10), H - 20, (830 if Show_ControlPanel else 530), H - 50, get_pixel_color(hwnd, 2, 20))
            send_signal(str(Recording_Speed) + '\n')
            MisrecognitionProtect = 0
        if not Flipping or MisrecognitionProtect == -1:
            Prev_Flipping = Flipping
            MisrecognitionProtect = 0

        # Draw UI
        draw_rectangle(hwnd, 0, H - 5, W, H - 3, ([191, 0, 0] if Flipping else [0, 191, 191]))
        draw_rectangle(hwnd, 0, H - 2, W, H, ([191, 0, 0] if Flipping else [0, 191, 191]))
        draw_rectangle(hwnd, 0, H - 5, 2, H, ([191, 0, 0] if Flipping else [0, 191, 191]))
        draw_rectangle(hwnd, W - 2, H - 5, W, H, ([191, 0, 0] if Flipping else [0, 191, 191]))
        if Show_Watermark:
            draw_text(hwnd, W - 20, 25, 18, [255, 0, 0], "SSW Recorder v1.2.1 | Program written by: happy_mimimix", 'R', 'T')
            draw_text(hwnd, W - 20, 45, 18, [255, 191, 0], "http://github.com/sudo-000/SSW-Recorder", 'R', 'T')
            draw_text(hwnd, W - 20, 65, 18, [0, 191, 0], "Plugin successfully loaded", 'R', 'T')
        if Show_ControlPanel:
            draw_rectangle_offscreen(15, H - 200, 300, H - 15, [255, 191, 0])
            draw_text_offscreen(30, H - 190, 20, [0, 127, 255], "SSW Recorder Control Panel", 'L', 'T')
            draw_text_offscreen(30, H - 156, 12, [127, 0, 0], "Auto speedhack override: ", 'L', 'T')
            draw_text_offscreen(30, H - 138, 12, [127, 0, 0], "Recording speed: ", 'L', 'T')
            draw_text_offscreen(30, H - 120, 12, [127, 0, 0], "Misrecognition protection level: ", 'L', 'T')
            draw_text_offscreen(30, H - 85, 12, [127, 0, 0], "Show watermark: ", 'L', 'T')
            draw_text_offscreen(30, H - 65, 12, [127, 0, 0], "Show statistics: ", 'L', 'T')
            if SH_Override1 > 0:
                draw_rectangle_offscreen(175, H - 142, 215, H - 156, [127, 127, 247])
                SH_Override1 -= 1
            else:
                draw_rectangle_offscreen(175, H - 142, 215, H - 156, [247, 247, 247])
            draw_text_offscreen(195, H - 144, 12, [15, 15, 15], "Play", 'M', 'B')
            if SH_Override0 > 0:
                draw_rectangle_offscreen(225, H - 142, 265, H - 156, [127, 127, 247])
                SH_Override0 -= 1
            else:
                draw_rectangle_offscreen(225, H - 142, 265, H - 156, [247, 247, 247])
            draw_text_offscreen(245, H - 144, 12, [15, 15, 15], "Pause", 'M', 'B')
            draw_rectangle_offscreen(128, H - 138, 162, H - 122, ([127, 127, 247] if Recording_Speed == 1 else [247, 247, 247]))
            draw_text_offscreen(145, H - 125, 12, [15, 15, 15], "1.0x", 'M', 'B')
            draw_rectangle_offscreen(171, H - 138, 205, H - 122, ([127, 127, 247] if Recording_Speed == 2 else [247, 247, 247]))
            draw_text_offscreen(106 + 85 - 3, H - 125, 12, [15, 15, 15], "0.5x", 'M', 'B')
            draw_rectangle_offscreen(214, H - 138, 248, H - 122, ([127, 127, 247] if Recording_Speed == 5 else [247, 247, 247]))
            draw_text_offscreen(231, H - 125, 12, [15, 15, 15], "0.2x", 'M', 'B')
            draw_rectangle_offscreen(257, H - 138, 291, H - 122, ([127, 127, 247] if Recording_Speed == 10 else [247, 247, 247]))
            draw_text_offscreen(274, H - 125, 12, [15, 15, 15], "0.1x", 'M', 'B')
            draw_rectangle_offscreen(35, H - 103, 65, H - 87, ([127, 127, 247] if ProtectionLevel == 0 else [247, 247, 247]))
            draw_text_offscreen(50, H - 90, 12, [15, 15, 15], "0", 'M', 'B')
            draw_rectangle_offscreen(75, H - 103, 105, H - 87, ([127, 127, 247] if ProtectionLevel == 1 else [247, 247, 247]))
            draw_text_offscreen(90, H - 90, 12, [15, 15, 15], "1", 'M', 'B')
            draw_rectangle_offscreen(115, H - 103, 145, H - 87, ([127, 127, 247] if ProtectionLevel == 2 else [247, 247, 247]))
            draw_text_offscreen(130, H - 90, 12, [15, 15, 15], "2", 'M', 'B')
            draw_rectangle_offscreen(155, H - 103, 185, H - 87, ([127, 127, 247] if ProtectionLevel == 3 else [247, 247, 247]))
            draw_text_offscreen(170, H - 90, 12, [15, 15, 15], "3", 'M', 'B')
            draw_rectangle_offscreen(195, H - 103, 225, H - 87, ([127, 127, 247] if ProtectionLevel == 4 else [247, 247, 247]))
            draw_text_offscreen(210, H - 90, 12, [15, 15, 15], "4", 'M', 'B')
            draw_rectangle_offscreen(235, H - 103, 265, H - 87, ([127, 127, 247] if ProtectionLevel == 5 else [247, 247, 247]))
            draw_text_offscreen(250, H - 90, 12, [15, 15, 15], "5", 'M', 'B')
            draw_rectangle_offscreen(130, H - 83, 166, H - 67, ([127, 247, 127] if Show_Watermark else [247, 247, 247]))
            draw_text_offscreen(148, H - 70, 12, [15, 15, 15], "YES", 'M', 'B')
            draw_rectangle_offscreen(171, H - 83, 207, H - 67, ([247, 247, 247] if Show_Watermark else [247, 127, 127]))
            draw_text_offscreen(189, H - 70, 12, [15, 15, 15], "NO", 'M', 'B')
            draw_rectangle_offscreen(120, H - 63, 156, H - 47, ([127, 247, 127] if Show_Statistics else [247, 247, 247]))
            draw_text_offscreen(138, H - 50, 12, [15, 15, 15], "YES", 'M', 'B')
            draw_rectangle_offscreen(161, H - 63, 197, H - 47, ([247, 247, 247] if Show_Statistics else [247, 127, 127]))
            draw_text_offscreen(179, H - 50, 12, [15, 15, 15], "NO", 'M', 'B')
            draw_rectangle_offscreen(35, H - 42, 85, H - 25, [247, 247, 247])
            draw_text_offscreen(60, H - 28, 12, [15, 15, 15], "Hide", 'M', 'B')
            draw_rectangle_offscreen(95, H - 42, 145, H - 25, ([127, 127, 247] if Toolbox_Open else [247, 247, 247]))
            draw_text_offscreen(120, H - 28, 12, [15, 15, 15], "Toolbox", 'M', 'B')
            draw_rectangle_offscreen(155, H - 42, 205, H - 25, [247, 127, 127])
            draw_text_offscreen(180, H - 28, 12, [15, 15, 15], "Unload", 'M', 'B')
            if Timer_Reset > 0:
                draw_rectangle_offscreen(215, H - 42, 285, H - 25, [127, 127, 247])
                Timer_Reset -= 1
            else:
                draw_rectangle_offscreen(215, H - 42, 285, H - 25, [247, 247, 247])
            draw_text_offscreen(250, H - 28, 12, [15, 15, 15], "Reset timer", 'M', 'B')
            if Play_Reset > 0:
                draw_rectangle_offscreen(215, H - 63, 285, H - 47, [127, 127, 247])
                Play_Reset -= 1
            else:
                draw_rectangle_offscreen(215, H - 63, 285, H - 47, [247, 247, 247])
            draw_text_offscreen(250, H - 50, 12, [15, 15, 15], "Play + Reset", 'M', 'B')
            copy_offscreen_content(hwnd, 15, H - 200, 300, H - 15)
        else:
            draw_rectangle_offscreen(0, H - 23, 50, H - 6, [255, 191, 0])
            draw_text_offscreen(25, H - 8, 12, [15, 15, 15], "Show", 'M', 'B')
            copy_offscreen_content(hwnd, 0, H - 23, 50, H - 6)
        if Toolbox_Open:
            draw_rectangle_offscreen(315 - (0 if Show_ControlPanel else 300), H - 200, 575 - (0 if Show_ControlPanel else 300), H - 56, [0, 191, 255])
            draw_text_offscreen(330 - (0 if Show_ControlPanel else 300), H - 190, 20, [255, 255, 0], "Toolbox", 'L', 'T')
            draw_rectangle_offscreen(325 - (0 if Show_ControlPanel else 300), H - 160, 441 - (0 if Show_ControlPanel else 300), H - 140, [247, 247, 247])
            draw_text_offscreen(383 - (0 if Show_ControlPanel else 300), H - 150, 12, [15, 15, 15], "MTC Quick Fix", 'M', 'M')
            draw_rectangle_offscreen(449 - (0 if Show_ControlPanel else 300), H - 160, 565 - (0 if Show_ControlPanel else 300), H - 140, [247, 247, 247])
            draw_text_offscreen(507 - (0 if Show_ControlPanel else 300), H - 150, 12, [15, 15, 15], "Start Over", 'M', 'M')
            draw_rectangle_offscreen(325 - (0 if Show_ControlPanel else 300), H - 135, 441 - (0 if Show_ControlPanel else 300), H - 115, [247, 247, 247])
            draw_text_offscreen(383 - (0 if Show_ControlPanel else 300), H - 125, 12, [15, 15, 15], "Playback Controls", 'M', 'M')
            draw_rectangle_offscreen(449 - (0 if Show_ControlPanel else 300), H - 135, 565 - (0 if Show_ControlPanel else 300), H - 115, [247, 247, 247])
            draw_text_offscreen(507 - (0 if Show_ControlPanel else 300), H - 125, 12, [15, 15, 15], "Score Type Options", 'M', 'M')
            draw_rectangle_offscreen(325 - (0 if Show_ControlPanel else 300), H - 110, 441 - (0 if Show_ControlPanel else 300), H - 90, [247, 247, 247])
            draw_text_offscreen(383 - (0 if Show_ControlPanel else 300), H - 100, 12, [15, 15, 15], "Track Merger", 'M', 'M')
            draw_rectangle_offscreen(449 - (0 if Show_ControlPanel else 300), H - 110, 565 - (0 if Show_ControlPanel else 300), H - 90, [247, 247, 247])
            draw_text_offscreen(507 - (0 if Show_ControlPanel else 300), H - 100, 12, [15, 15, 15], "Performance Monitor", 'M', 'M')
            draw_rectangle_offscreen(325 - (0 if Show_ControlPanel else 300), H - 85, 441 - (0 if Show_ControlPanel else 300), H - 65, [247, 247, 247])
            draw_text_offscreen(383 - (0 if Show_ControlPanel else 300), H - 75, 12, [15, 15, 15], "Save As *.ss10 File", 'M', 'M')
            draw_rectangle_offscreen(449 - (0 if Show_ControlPanel else 300), H - 85, 565 - (0 if Show_ControlPanel else 300), H - 65, [247, 247, 247])
            draw_text_offscreen(507 - (0 if Show_ControlPanel else 300), H - 75, 12, [15, 15, 15], "Equal Width Bar", 'M', 'M')
            copy_offscreen_content(hwnd, 315 - (0 if Show_ControlPanel else 300), H - 200, 575 - (0 if Show_ControlPanel else 300), H - 56)
        if Show_Statistics:
            draw_rectangle_offscreen(W - 265, H - 250, W - 15, H - 15, [63, 63, 63])
            draw_text_offscreen(W - 250, H - 240, 20, [191, 191, 255], "Statistics: ", 'L', 'T')
            draw_text_offscreen(W - 250, H - 200, 12, [191, 191, 191], "Song: ", 'L', 'T')
            draw_text_offscreen(W - 30, H - 200, 12, [191, 191, 191], File[0:22].strip() + '...' if File.__len__() > 25 else File.strip(), 'R', 'T')
            draw_text_offscreen(W - 250, H - 180, 12, [191, 191, 191], "Time lapsed: ", 'L', 'T')
            draw_text_offscreen(W - 30, H - 180, 12, [191, 191, 191], Format_Time(time.time() - start_time), 'R', 'T')
            draw_text_offscreen(W - 250, H - 160, 12, [191, 191, 191], "Dialogue size: ", 'L', 'T')
            draw_text_offscreen(W - 30, H - 160, 12, [191, 191, 191], "(" + str(W) + "," + str(H) + ")", 'R', 'T')
            draw_text_offscreen(W - 250, H - 140, 12, [191, 191, 191], "Cursor position: ", 'L', 'T')
            draw_text_offscreen(W - 30, H - 140, 12, ([255, 0, 0] if Clicked else [191, 191, 191]), "(" + str(CX) + "," + str(CY) + ")", 'R', 'T')
            draw_text_offscreen(W - 250, H - 120, 12, [191, 191, 191], "Page flip: ", 'L', 'T')
            draw_text_offscreen(W - 30, H - 120, 12, ([0, 191, 0] if Prev_Flipping else [191, 0, 0]), str(Prev_Flipping), 'R', 'T')
            draw_text_offscreen(W - 250, H - 100, 12, [191, 191, 191], "Brightness of the darkest: ", 'L', 'T')
            draw_text_offscreen(W - 30, H - 100, 12, [191, 191, 191], str(Brightness), 'R', 'T')
            draw_text_offscreen(W - 250, H - 80, 12, [191, 191, 191], "Misrecognition protector state: ", 'L', 'T')
            draw_text_offscreen(W - 30, H - 80, 12, [191, 191, 191], str(MisrecognitionProtect), 'R', 'T')
            draw_text_offscreen(W - 250, H - 60, 12, [191, 191, 191], "App Version: ", 'L', 'T')
            draw_text_offscreen(W - 30, H - 60, 12, [191, 191, 191], Version, 'R', 'T')
            draw_text_offscreen(W - 250, H - 40, 12, [191, 191, 191], "Compatibility: ", 'L', 'T')
            if Version.__contains__("Professional") or Version.__contains__("Standard"):
                draw_text_offscreen(W - 30, H - 40, 12, [0, 191, 0], "Fully Compatible", 'R', 'T')
            elif Version.__contains__("Lite"):
                draw_text_offscreen(W - 30, H - 40, 12, [191, 0, 0], "Incompatible", 'R', 'T')
            else:
                draw_text_offscreen(W - 30, H - 40, 12, [191, 191, 0], "Unknown", 'R', 'T')
            copy_offscreen_content(hwnd, W - 265, H - 250, W - 15, H - 15)


def draw_rectangle(hwnd, x1, y1, x2, y2, color):
    # Draw rectangle with win32 GDI
    hdc = win32gui.GetWindowDC(hwnd)
    brush = win32gui.CreateSolidBrush(win32api.RGB(color[0], color[1], color[2]))
    win32gui.SelectObject(hdc, brush)
    win32gui.PatBlt(hdc, x1, y1, x2 - x1, y2 - y1, win32con.PATCOPY)
    win32gui.DeleteObject(brush)
    win32gui.ReleaseDC(hwnd, hdc)


def draw_text(hwnd, x, y, size, color, content, alignH, alignV):
    # Draw text with win32 GDI
    hdc = win32gui.GetWindowDC(hwnd)
    lf = win32gui.LOGFONT()
    lf.lfFaceName = "Tahoma"
    lf.lfHeight = -size
    font = win32gui.CreateFontIndirect(lf)
    win32gui.SelectObject(hdc, font)
    win32gui.SetTextColor(hdc, win32api.RGB(color[0], color[1], color[2]))
    win32gui.SetBkMode(hdc, win32con.TRANSPARENT)
    w, h = win32gui.GetTextExtentPoint32(hdc, content)
    # Align text
    if alignH == 'L':
        alignH2 = win32con.DT_LEFT
        x1 = x
        x2 = x + w
    elif alignH == 'R':
        alignH2 = win32con.DT_RIGHT
        x1 = x - w
        x2 = x
    elif alignH == 'M':
        alignH2 = win32con.DT_CENTER
        x1 = x - w // 2
        x2 = x + w // 2
    else:
        raise "Invalid alignH type, should be 'L'(Left), 'M'(Middle), or 'R'(Right). "
    if alignV == 'T':
        alignV2 = win32con.DT_TOP
        y1 = y
        y2 = y + h
    elif alignV == 'B':
        alignV2 = win32con.DT_BOTTOM
        y1 = y - h
        y2 = y
    elif alignV == 'M':
        alignV2 = win32con.DT_VCENTER
        y1 = y - h // 2
        y2 = y + h // 2
    else:
        raise "Invalid alignV type, should be 'T'(Top), 'M'(Middle), or 'B'(Bottom). "
    win32gui.DrawText(hdc, content, -1, (x1, y1, x2, y2), alignH2 | alignV2 | win32con.DT_SINGLELINE)
    win32gui.DeleteObject(font)
    win32gui.ReleaseDC(hwnd, hdc)


def draw_rectangle_offscreen(x1, y1, x2, y2, color):
    global OffScreenDC
    global OffScreenBMP
    # Draw rectangle with win32 GDI
    brush = win32gui.CreateSolidBrush(win32api.RGB(color[0], color[1], color[2]))
    win32gui.SelectObject(OffScreenDC, brush)
    win32gui.PatBlt(OffScreenDC, x1, y1, x2 - x1, y2 - y1, win32con.PATCOPY)
    win32gui.DeleteObject(brush)


def draw_text_offscreen(x, y, size, color, content, alignH, alignV):
    global OffScreenDC
    global OffScreenBMP
    # Draw text with win32 GDI
    lf = win32gui.LOGFONT()
    lf.lfFaceName = "Tahoma"
    lf.lfHeight = -size
    font = win32gui.CreateFontIndirect(lf)
    win32gui.SelectObject(OffScreenDC, font)
    win32gui.SetTextColor(OffScreenDC, win32api.RGB(color[0], color[1], color[2]))
    win32gui.SetBkMode(OffScreenDC, win32con.TRANSPARENT)
    w, h = win32gui.GetTextExtentPoint32(OffScreenDC, content)
    # Align text
    if alignH == 'L':
        alignH2 = win32con.DT_LEFT
        x1 = x
        x2 = x + w
    elif alignH == 'R':
        alignH2 = win32con.DT_RIGHT
        x1 = x - w
        x2 = x
    elif alignH == 'M':
        alignH2 = win32con.DT_CENTER
        x1 = x - w // 2
        x2 = x + w // 2
    else:
        raise "Invalid alignH type, should be 'L'(Left), 'M'(Middle), or 'R'(Right). "
    if alignV == 'T':
        alignV2 = win32con.DT_TOP
        y1 = y
        y2 = y + h
    elif alignV == 'B':
        alignV2 = win32con.DT_BOTTOM
        y1 = y - h
        y2 = y
    elif alignV == 'M':
        alignV2 = win32con.DT_VCENTER
        y1 = y - h // 2
        y2 = y + h // 2
    else:
        raise "Invalid alignV type, should be 'T'(Top), 'M'(Middle), or 'B'(Bottom). "
    win32gui.DrawText(OffScreenDC, content, -1, (x1, y1, x2, y2), alignH2 | alignV2 | win32con.DT_SINGLELINE)
    win32gui.DeleteObject(font)


def copy_offscreen_content(hwnd, x1, y1, x2, y2):
    global OffScreenDC
    global OffScreenBMP
    # Copy from buffer
    hdc = win32gui.GetWindowDC(hwnd)
    win32gui.BitBlt(hdc, x1, y1, x2 - x1, y2 - y1, OffScreenDC, x1, y1, win32con.SRCCOPY)
    win32gui.ReleaseDC(hwnd, hdc)


def to_front(hwnd):
    # Bring targeted window to front
    win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
    win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
    mouse.move(win32gui.GetWindowRect(hwnd)[0] + 10, win32gui.GetWindowRect(hwnd)[1] + 10)
    mouse.click(mouse.RIGHT)
    mouse.click(mouse.RIGHT)
    keyboard.press('alt')
    keyboard.release('alt')


def enum_child_windows(hwnd):
    # Find all child windows
    child_windows = []
    win32gui.EnumChildWindows(hwnd, lambda hwnd, lparam: lparam.append(hwnd), child_windows)
    return child_windows


def tree_search(hwnd, depth=0):
    # Search for main window
    global Maximum_Depth
    window_text = win32gui.GetWindowText(hwnd)
    window_class = win32gui.GetClassName(hwnd)
    Maximum_Depth = max(depth, Maximum_Depth)
    if depth == 0:
        if not window_text.startswith('Singer Song Writer'):
            return
        else:
            global Version
            global File
            global MainWindowHWND
            Version = window_text.lstrip('Singer Song Writer ')
            if Version.__contains__(' - '):
                File = Version.split('[')[1].split(']')[0]
                File = File.split('\\')[-1]
                Version = Version.split(' - ')[0]
            else:
                File = 'Null'
            MainWindowHWND = hwnd
            print("\033[34mSinger Song Writer main window handle found! ")
    elif depth == 1:
        if window_class != 'MDIClient':
            return
        else:
            print("\033[34mMDIClient handle found! ")
    elif depth == 2:
        if not (":" in window_text) and not ("Track(" in window_text):
            return
        else:
            global ScoreMainClass
            ScoreMainClass = str(win32gui.GetClassName(hwnd)).split(':')
            print("\033[34mScore editing window handle found! ")
    elif depth == 3:
        if not window_class.startswith('AfxMDIFrame'):
            return
        else:
            global WindowType
            WindowType = str(window_class).lstrip('AfxMDIFrame')
            print("\033[34mAfxMDIFrame" + WindowType + " window handle found! ")
    elif depth == 4:
        if window_class != ('AfxMDIFrame' + WindowType):
            return
        else:
            print("\033[34mAfxMDIFrame" + WindowType + " window handle found! ")
    elif depth == 5:
        if window_class != ('AfxWnd' + WindowType):
            return
        else:
            child = enum_child_windows(hwnd)
            for child_hwnd in child:
                if win32gui.GetClassName(child_hwnd) == ('AfxMDIFrame' + WindowType) or enum_child_windows(child_hwnd):
                    return
            print("\033[34mAfxWnd" + WindowType + " window handle found! ")
    elif depth == 6:
        if window_class == 'ScrollBar' or window_class == 'Button' or str(window_class).split(':')[0:-2] != ScoreMainClass[0:-2]:
            return
        else:
            print("\033[34mScore editing window handle found! ")
            if not UI_Debug_Mode:
                print("\033[95m\nGreat! You've successfully loaded your MIDI file in Singer Song Writer and opened the score editing window. ")
                print("There are a few more steps we need to complete before recording your video. ")
                print("First, do you have Cheat Engine installed on your device? ")
                print("If yes, press 'Z' to continue. If not, press 'X' to install it. ")
                playsound.playsound(".\\Resources\\KDE_Click.wav")
                while (not keyboard.is_pressed('z') and not keyboard.is_pressed('x')) or not win32gui.GetForegroundWindow() == win32console.GetConsoleWindow():
                    time.sleep(1 / 128)
                if keyboard.is_pressed('x') and win32gui.GetForegroundWindow() == win32console.GetConsoleWindow():
                    os.system("EXPLORER .\\Resources\\CE74.exe")
                while keyboard.is_pressed('z') or keyboard.is_pressed('x') or not win32gui.GetForegroundWindow() == win32console.GetConsoleWindow():
                    time.sleep(1 / 128)
                print("\033[95m\nNext, do you have OBS studio installed on your device? ")
                print("If yes, press 'Z' to continue. If not, press 'X' to install it. ")
                playsound.playsound(".\\Resources\\KDE_Click.wav")
                while (not keyboard.is_pressed('z') and not keyboard.is_pressed('x')) or not win32gui.GetForegroundWindow() == win32console.GetConsoleWindow():
                    time.sleep(1 / 128)
                if keyboard.is_pressed('x') and win32gui.GetForegroundWindow() == win32console.GetConsoleWindow():
                    os.system("EXPLORER .\\Resources\\OBS30.exe")
                while keyboard.is_pressed('z') or keyboard.is_pressed('x') or not win32gui.GetForegroundWindow() == win32console.GetConsoleWindow():
                    time.sleep(1 / 128)
                print("\033[95m\nPlease setup a new application window capture in OBS studio via Sources -> Add -> Window capture. ")
                print("Select \"BitBlt\" as the capture method. ")
                print("Press 'Z' when done. ")
                playsound.playsound(".\\Resources\\KDE_Click.wav")
                while not keyboard.is_pressed('z') or not win32gui.GetForegroundWindow() == win32console.GetConsoleWindow():
                    time.sleep(1 / 128)
                while keyboard.is_pressed('z') or not win32gui.GetForegroundWindow() == win32console.GetConsoleWindow():
                    time.sleep(1 / 128)
                print("\033[95m\nNow, it's time to load my cheat tables into Cheat Engine! ")
                print("Please click 'Yes' when prompted. ")
                print("Note: If you are using a version of Singer Song Writer other than 10 Professional, don't forget to correct the process name in the cheat table. ")
                playsound.playsound(".\\Resources\\KDE_Click.wav")
                print("\033[34mWaiting for pipe connection from Cheat Engine...")
                os.system(".\\Resources\\SSW.CT")
                print("\033[92mConnected: 0/2")
                win32pipe.ConnectNamedPipe(PipeSSW, None)
                os.system(".\\Resources\\OBS.CT")
                print("\033[92mConnected: 1/2")
                playsound.playsound(".\\Resources\\KDE_Click.wav")
                win32pipe.ConnectNamedPipe(PipeOBS, None)
                print("\033[92mConnected: 2/2")
                playsound.playsound(".\\Resources\\KDE_Click.wav")
                print("\033[95m\n\nSome optional extra tips to make your video even better: ")
                print("\n1. You can adjust the spacing from the top margin by dragging the soprano symbol at the front of the score. ")
                print("This will make notes on all 128 keys visible on the screen. ")
                print("\n2. For better appearance, we strongly recommend merging all channels and tracks in your MIDI file into one. ")
                print("This can easily be done using my modified version of CJCAMM, which includes an option to change all MIDI messages to channel 1. You can find it in the toolbox. ")
                print("\n3. It's strongly recommended for you to render a perfect audio using Keppy's Midi Converter and add it to your video by video editing. ")
                print("The sound output in Singer Song Writer is terrible, so please do not use it in your video! ")
                print("\n4. Rendering the song arrangement window, the first window that appears when opening a MIDI, requires additional processing power and can cause the musical score display to lag or even crash. ")
                print("Since we can't close this window, the best way to reduce lag here is to add some blank tracks above and move the track with notes off-screen so Singer Song Writer won't render it. ")
                print("This can be done by right-clicking on any track in the song arrangement window and selecting 'トラックの追加(A)', followed by 'MIDI(M)', and then typing any number greater than 20 in '追加するトラック数'. ")
                print("New blank tracks will appear after the track you right-clicked on. ")
                print("Now, simply drag the track with notes after these blank tracks and scroll to the top. ")
                print("This should keep the track with notes off-screen and unloaded from the program's memory, preventing lag or crashes during playback. ")
                print("\nPress 'Z' when finished reading. ")
                playsound.playsound(".\\Resources\\KDE_Click.wav")
                while not keyboard.is_pressed('z') or not win32gui.GetForegroundWindow() == win32console.GetConsoleWindow():
                    time.sleep(1 / 128)
                while keyboard.is_pressed('z') or not win32gui.GetForegroundWindow() == win32console.GetConsoleWindow():
                    time.sleep(1 / 128)
            to_front(MainWindowHWND)
            to_front(MainWindowHWND)
            keyboard.press_and_release('alt')
            keyboard.press_and_release('p')
            keyboard.press_and_release('t')
            keyboard.press_and_release('alt')
            keyboard.press_and_release('s')
            keyboard.press_and_release('p')
            keyboard.press_and_release('s')
            keyboard.press_and_release('n')
            keyboard.press_and_release('return')
            keyboard.press_and_release('home')
            global Version
            if int(Version[0]) > 1 and int(Version[0]) < 1:
                print("Note: We think your version of Singer Song Writer might not require MTC='Slave' for speed hacking. \nIf that isn't the case, click 'MTC Quick Fix' in the toolbox to fix it. ")
            else:
                keyboard.press_and_release('alt')
                keyboard.press_and_release('s')
                keyboard.press_and_release('p')
                keyboard.press_and_release('s')
                keyboard.press_and_release('s')
                keyboard.press_and_release('return')
            threading.Thread(target=auto_start, args=()).start()
            threading.Thread(target=playsound.playsound, args=(".\\Resources\\KDE_Startup.wav",)).start()
            print("\033[92m\nCongratulation: \nSSW Recorder plugin has successfully loaded ! \033[95m")
            print("\033[95m\nIf you did everything correctly, you should now be able to see the SSW Recorder watermark in the top-right corner. ")
            print("The watermark can be removed by changing an option in the control panel, but it would make me happy if you could credit my program in your video. ")
            print("To begin recording, click on the following buttons in order: ")
            print("1. The 'Start Recording' button in OBS Studio")
            print("2. The 'Start Over' button in the toolbox")
            print("3. The 'Play' button in SSW Recorder control panel")
            print("To stop recording when finished, click on the following buttons in order: ")
            print("1. The 'Start Over' button in the toolbox")
            print("2. The 'Stop Recording' button in OBS Studio")
            print("3. The 'Unload' button in SSW Recorder control panel")
            print("\nHave fun making a perfect no-lag musical score video for your black midi! \n\033[32m")
            playsound.playsound(".\\Resources\\KDE_Click.wav")
            global start_time
            global Prev_MouseClick
            start_time = time.time()
            while True:
                ClintRECT = win32gui.GetClientRect(hwnd)
                ToScreen = win32gui.ClientToScreen(hwnd, (0, 0))
                Main_Program(hwnd, ClintRECT[2], ClintRECT[3], mouse.get_position()[0] - ToScreen[0], mouse.get_position()[1] - ToScreen[1], mouse.is_pressed(mouse.LEFT))
                if not mouse.is_pressed(mouse.LEFT):
                    Prev_MouseClick = False
                if mouse.is_pressed(mouse.LEFT) and (not Prev_MouseClick):
                    Process_Mouse_Input(hwnd, mouse.get_position()[0] - ToScreen[0], mouse.get_position()[1] - ToScreen[1], ClintRECT[2], ClintRECT[3])
                    Prev_MouseClick = True
                time.sleep(1 / 128)
    child_windows = enum_child_windows(hwnd)
    if not child_windows:
        return
    for child_hwnd in child_windows:
        tree_search(child_hwnd, depth + 1)


os.chdir(os.path.dirname(os.path.realpath(__file__)))
os.system("chcp 65001")
os.system("title SSW Recorder v1.3.0")
os.system("color f0")
playsound.playsound(".\\Resources\\KDE_Beep.wav")
print("\033[2J\033[H\033[1mPress 'Shift' to enter UI debug mode...")
for i in range(128):
    keyboard.is_pressed('shift')
    time.sleep(1 / 128)
if keyboard.is_pressed('shift') and win32gui.GetForegroundWindow() == win32console.GetConsoleWindow():
    playsound.playsound(".\\Resources\\KDE_Click.wav")
    UI_Debug_Mode = True
    print("\033[91mNotice: ")
    print("You are now in UI debug mode, OBS and Cheat Engine won't be connected. ")
    print("This mode is designed for testing the software UI more efficiently without the need to get OBS and Cheat Engine ready everytime the program restarts. ")
    print("You should not be here usually, please restart this app while making sure the 'Shift' key is not being pressed while the app is launching. ")
    print("If you know what you are doing, press 'Z' to proceed. ")
    playsound.playsound(".\\Resources\\KDE_Error.wav")
    while keyboard.is_pressed('shift') or not win32gui.GetForegroundWindow() == win32console.GetConsoleWindow():
        time.sleep(1 / 128)
    while not keyboard.is_pressed('z') or not win32gui.GetForegroundWindow() == win32console.GetConsoleWindow():
        time.sleep(1 / 128)
    while keyboard.is_pressed('z') or not win32gui.GetForegroundWindow() == win32console.GetConsoleWindow():
        time.sleep(1 / 128)
    playsound.playsound(".\\Resources\\KDE_Click.wav")
else:
    UI_Debug_Mode = False
    print("\033[2J\033[H\033[1mPress 'Ctrl' to extract source code...")
    for i in range(128):
        keyboard.is_pressed('ctrl')
        time.sleep(1 / 128)
    if keyboard.is_pressed('ctrl') and win32gui.GetForegroundWindow() == win32console.GetConsoleWindow():
        playsound.playsound(".\\Resources\\KDE_Click.wav")
        os.system("color 0F")
        os.system("md \"%userprofile%\\Desktop\\SSW Recorder Source Code\"")
        os.system("xcopy .\\Program_Source_Code.py \"%userprofile%\\Desktop\\SSW Recorder Source Code\" /c /v /r /y /k /i /g")
        os.system("xcopy .\\Compile_EXE.cmd \"%userprofile%\\Desktop\\SSW Recorder Source Code\" /c /v /r /y /k /i /g")
        os.system("xcopy .\\Environment_Setup.cmd \"%userprofile%\\Desktop\\SSW Recorder Source Code\" /c /v /r /y /k /i /g")
        os.system("xcopy .\\Resources \"%userprofile%\\Desktop\\SSW Recorder Source Code\\.\\Resources\" /c /e /v /r /y /k /i /g")
        print("\n\033[92mInfo: Source code extracted to \"%userprofile%\\Desktop\\SSW Recorder Source Code\"")
        playsound.playsound(".\\Resources\\KDE_Error.wav")
        time.sleep(3)
        raise SystemExit
while True:
    print("\033[2J\033[H", end='')
    child_windows = enum_child_windows(win32gui.GetDesktopWindow())
    for child_hwnd in child_windows:
        tree_search(child_hwnd)
    if Maximum_Depth == 0:
        print("\033[91m\nCRITICAL ERROR: \nFailed to obtain the handle of the main window in Singer Song Writer! ")
        print("\033[95m\nTroubleshooting Assistant: ")
        print("It seems that Singer Song Writer is not currently running. Please launch the program before using this plugin. ")
        print("If you haven't installed Singer Song Writer yet, press 'X' to install the program. ")
        print("Press 'Z' to try searching for the handle again. ")
        playsound.playsound(".\\Resources\\KDE_Error.wav")
        while (not keyboard.is_pressed('z') and not keyboard.is_pressed('x')) or not win32gui.GetForegroundWindow() == win32console.GetConsoleWindow():
            time.sleep(1 / 128)
        threading.Thread(target=playsound.playsound, args=(".\\Resources\\KDE_Click.wav",)).start()
        if keyboard.is_pressed('x') and win32gui.GetForegroundWindow() == win32console.GetConsoleWindow():
            os.system("EXPLORER .\\Resources\\SSW10PRO.iso")
        while keyboard.is_pressed('z') or keyboard.is_pressed('x') or not win32gui.GetForegroundWindow() == win32console.GetConsoleWindow():
            time.sleep(1 / 128)
    elif Maximum_Depth == 1:
        print("\033[91m\nCRITICAL ERROR: \nFailed to obtain the handle of the score editing window in Singer Song Writer! ")
        print("\033[95m\nTroubleshooting Assistant: ")
        print("It seems that you've already opened Singer Song Writer but have not yet loaded any MIDI files for it to play. ")
        print("If you're having trouble reading Japanese and can't identify the \"Open\" button, press 'X' on your keyboard, and we will open the dialog for you.")
        print("For better appearance, we strongly recommend merging all channels and tracks in your MIDI file into one. ")
        print("This can easily be done using my modified version of CJCAMM, which includes an option to change all MIDI messages to channel 1. Press 'M' to open it. ")
        print("Additionally, don't forget to uncheck 'MIDファイル読み込み時に楽譜の調整を行う' at the bottom of the dialog for better results. ")
        print("Press 'Z' to try searching for the handle again. ")
        playsound.playsound(".\\Resources\\KDE_Error.wav")
        while (not keyboard.is_pressed('z') and not keyboard.is_pressed('x') and not keyboard.is_pressed('m')) or not win32gui.GetForegroundWindow() == win32console.GetConsoleWindow():
            time.sleep(1 / 128)
        threading.Thread(target=playsound.playsound, args=(".\\Resources\\KDE_Click.wav",)).start()
        if keyboard.is_pressed('x') and win32gui.GetForegroundWindow() == win32console.GetConsoleWindow():
            while keyboard.is_pressed('x'):
                time.sleep(1 / 128)
            to_front(MainWindowHWND)
            to_front(MainWindowHWND)
            for i in range(128):
                keyboard.press_and_release('esc')
                time.sleep(1 / 128)
            for i in range(32):
                keyboard.press('ctrl')
                keyboard.press('o')
                keyboard.release('ctrl')
                keyboard.release('o')
                time.sleep(1 / 128)
            for i in range(4):
                keyboard.press_and_release('tab')
            keyboard.press_and_release('space')
            keyboard.press('shift')
            for i in range(4):
                keyboard.press_and_release('tab')
            keyboard.release('shift')
        elif keyboard.is_pressed('m') and win32gui.GetForegroundWindow() == win32console.GetConsoleWindow():
            os.system(".\\Resources\\CJC_Toolbox.exe CJCAMM en-us")
        while keyboard.is_pressed('z') or keyboard.is_pressed('x') or keyboard.is_pressed('c') or not win32gui.GetForegroundWindow() == win32console.GetConsoleWindow():
            time.sleep(1 / 128)
    elif Maximum_Depth == 2:
        print("\033[91m\nCRITICAL ERROR: \nFailed to obtain the handle of the score editing window in Singer Song Writer! ")
        print("\033[95m\nTroubleshooting Assistant: ")
        print("It seems that you have already loaded a MIDI file in Singer Song Writer but have not yet opened the score editing window. ")
        print("If you're having trouble reading Japanese and can't identify the \"Open Score Editing Window\" button, press 'X' on your keyboard, and we will open the it for you. ")
        print("For better appearance, we strongly recommend merging all channels and tracks in your MIDI file into one. ")
        print("This can easily be done using my modified version of CJCAMM, which includes an option to change all MIDI messages to channel 1. Press 'M' to open it. ")
        print("Press 'Z' to try searching for the handle again. \n")
        playsound.playsound(".\\Resources\\KDE_Error.wav")
        while (not keyboard.is_pressed('z') and not keyboard.is_pressed('x') and not keyboard.is_pressed('m')) or not win32gui.GetForegroundWindow() == win32console.GetConsoleWindow():
            time.sleep(1 / 128)
        threading.Thread(target=playsound.playsound, args=(".\\Resources\\KDE_Click.wav",)).start()
        if keyboard.is_pressed('x') and win32gui.GetForegroundWindow() == win32console.GetConsoleWindow():
            while keyboard.is_pressed('x'):
                time.sleep(1 / 128)
            to_front(MainWindowHWND)
            to_front(MainWindowHWND)
            keyboard.press_and_release('alt')
            keyboard.press_and_release('w')
            keyboard.press_and_release('e')
            keyboard.press_and_release('c')
        elif keyboard.is_pressed('m') and win32gui.GetForegroundWindow() == win32console.GetConsoleWindow():
            os.system(".\\Resources\\CJC_Toolbox.exe CJCAMM en-us")
        while keyboard.is_pressed('z') or keyboard.is_pressed('x') or keyboard.is_pressed('m') or not win32gui.GetForegroundWindow() == win32console.GetConsoleWindow():
            time.sleep(1 / 128)
