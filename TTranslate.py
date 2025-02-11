import json
import numpy as np
import win32api
import win32con
import win32gui
import ctypes
import pygetwindow as gw
import warnings
from paddleocr import PaddleOCR
from mtranslate import translate
from difflib import SequenceMatcher


writenText = []
isCleared = True
totalFixedY = []
templateName = input("Enter the name of the JSON files you want to read: ")
if templateName == "":
    selectedTemplate = None
else:
    with open(f"{templateName}.json", "r") as json_file:
        selectedTemplate = json.load(json_file)
        print("Data read successfully:")
        print(selectedTemplate)
        print(type(selectedTemplate[0][0][0]))

useCombineVertical = input("Use Vertical combine Y? (y/n): ")
if useCombineVertical == "y":
    useCombineVertical = True
else:
    useCombineVertical = False
useBestFontSize = input("Use Best Font size? (y/n): ")
if useBestFontSize == "y":
    useBestFontSize = True
else:
    useBestFontSize = False
warnings.filterwarnings("ignore", category=FutureWarning)
ocr = PaddleOCR(
    use_angle_cls=False,
    show_log=False,
    lang="en",
)

windows = gw.getWindowsWithTitle("")


def is_window_focused(window):
    focused_window = gw.getActiveWindow()
    return focused_window == window


for i, window in enumerate(windows):
    print(f"{i}: {window.title}")

index = int(input("Select window number: "))
selected_window = windows[index]


wc = win32gui.WNDCLASS()
wc.lpfnWndProc = win32gui.DefWindowProc
wc.hInstance = win32api.GetModuleHandle(None)
wc.lpszClassName = "TransparentWindow"

class_atom = win32gui.RegisterClass(wc)

screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)

hwnd = win32gui.CreateWindowEx(
    win32con.WS_EX_LAYERED | win32con.WS_EX_TRANSPARENT | win32con.WS_EX_TOPMOST,
    class_atom,
    "TransparentWindow",
    win32con.WS_POPUP,
    0,
    0,
    screen_width,
    screen_height,
    None,
    None,
    wc.hInstance,
    None,
)

win32gui.SetLayeredWindowAttributes(hwnd, 0x000000, 255, win32con.LWA_COLORKEY)
win32gui.ShowWindow(hwnd, win32con.SW_SHOW)

hdc = win32gui.GetDC(hwnd)


def calculateFixedY():
    return np.bincount(totalFixedY).argmax()


def calculateFixedSize(sizeValue):
    global totalFixedSize
    totalFixedSize.pop(0)
    totalFixedSize.append(sizeValue)
    return np.bincount(totalFixedSize).argmax()


def calculateBestFontSize(hdc, text, width, height, min_font_size=15, max_font_size=50):
    gdi32 = ctypes.WinDLL("gdi32")

    best_font_size = min_font_size
    low = min_font_size
    high = max_font_size

    while low <= high:
        mid_font_size = (low + high) // 2

        class LOGFONT(ctypes.Structure):
            _fields_ = [
                ("lfHeight", ctypes.c_long),
                ("lfWidth", ctypes.c_long),
                ("lfEscapement", ctypes.c_long),
                ("lfOrientation", ctypes.c_long),
                ("lfWeight", ctypes.c_long),
                ("lfItalic", ctypes.c_byte),
                ("lfUnderline", ctypes.c_byte),
                ("lfStrikeOut", ctypes.c_byte),
                ("lfCharSet", ctypes.c_byte),
                ("lfOutPrecision", ctypes.c_byte),
                ("lfClipPrecision", ctypes.c_byte),
                ("lfQuality", ctypes.c_byte),
                ("lfPitchAndFamily", ctypes.c_byte),
                ("lfFaceName", ctypes.c_wchar * 32),
            ]

        logfont = LOGFONT()
        logfont.lfHeight = -mid_font_size
        logfont.lfWeight = win32con.FW_NORMAL
        logfont.lfCharSet = win32con.DEFAULT_CHARSET
        logfont.lfFaceName = "Arial"

        hfont = gdi32.CreateFontIndirectW(ctypes.byref(logfont))
        old_font = win32gui.SelectObject(hdc, hfont)

        text_size = win32gui.GetTextExtentPoint32(hdc, text)

        win32gui.SelectObject(hdc, old_font)
        gdi32.DeleteObject(hfont)

        if text_size[0] <= width and text_size[1] <= height:
            best_font_size = mid_font_size
            low = mid_font_size + 1
        else:
            high = mid_font_size - 1
    return best_font_size


def drawRectangleWithText(x, y, width, height, color, text):
    brush = win32gui.CreateSolidBrush(color)
    rect = (x, y, x + width, y + height)
    with open("writeText.txt", "a") as writeText:
        writeText.write(f"x : {x} {y} {width} {height} {text}\n")
    win32gui.FillRect(hdc, rect, brush)
    gdi32 = ctypes.WinDLL("gdi32")
    user32 = ctypes.WinDLL("user32")

    class LOGFONT(ctypes.Structure):
        _fields_ = [
            ("lfHeight", ctypes.c_long),
            ("lfWidth", ctypes.c_long),
            ("lfEscapement", ctypes.c_long),
            ("lfOrientation", ctypes.c_long),
            ("lfWeight", ctypes.c_long),
            ("lfItalic", ctypes.c_byte),
            ("lfUnderline", ctypes.c_byte),
            ("lfStrikeOut", ctypes.c_byte),
            ("lfCharSet", ctypes.c_byte),
            ("lfOutPrecision", ctypes.c_byte),
            ("lfClipPrecision", ctypes.c_byte),
            ("lfQuality", ctypes.c_byte),
            ("lfPitchAndFamily", ctypes.c_byte),
            ("lfFaceName", ctypes.c_wchar * 32),
        ]

    lf = LOGFONT()
    text = translate(text, "tr")
    fontSize = 0
    if useBestFontSize == True:
        fontSize = calculateBestFontSize(hdc, text, width, height)
    else:
        fontSize = max(int(height / 3.5), 22)
    lf.lfHeight = -fontSize
    lf.lfWeight = win32con.FW_NORMAL
    lf.lfCharSet = win32con.DEFAULT_CHARSET
    lf.lfFaceName = "Arial"

    hfont = gdi32.CreateFontIndirectW(ctypes.byref(lf))

    old_font = win32gui.SelectObject(hdc, hfont)

    gdi32.SetBkMode(hdc, win32con.TRANSPARENT)

    gdi32.SetTextColor(hdc, 0x2F2F2F)

    win32gui.DrawTextW(
        hdc,
        text,
        -1,
        rect,
        win32con.DT_WORDBREAK | win32con.DT_CENTER | win32con.DT_VCENTER,
    )

    win32gui.SelectObject(hdc, old_font)
    win32gui.DeleteObject(hfont)
    win32gui.DeleteObject(brush)


def clearRectangle(x, y, width, height):
    brush = win32gui.CreateSolidBrush(0x000000)
    x, y, width, height = (
        max(0, int(np.floor(x)) - 5),
        max(0, int(np.floor(y)) - 5),
        int(np.ceil(width)) + 12,
        int(np.ceil(height)) + 12,
    )
    rect = (x, y, x + width, y + height)
    win32gui.FillRect(hdc, rect, brush)
    win32gui.DeleteObject(brush)


def clearRectangleAll():
    global writenText
    brush = win32gui.CreateSolidBrush(0x000000)
    rect = (0, 0, screen_width, screen_height)
    win32gui.FillRect(hdc, rect, brush)
    win32gui.DeleteObject(brush)
    writenText = []


def getScreenshot():
    hwnd = win32gui.FindWindow(None, selected_window.title)
    if hwnd == 0:
        raise Exception("Pencere bulunamadı!")

    rect = win32gui.GetWindowRect(hwnd)
    width = rect[2] - rect[0]
    height = rect[3] - rect[1]

    class BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [
            ("biSize", ctypes.c_uint32),
            ("biWidth", ctypes.c_int32),
            ("biHeight", ctypes.c_int32),
            ("biPlanes", ctypes.c_uint16),
            ("biBitCount", ctypes.c_uint16),
            ("biCompression", ctypes.c_uint32),
            ("biSizeImage", ctypes.c_uint32),
            ("biXPelsPerMeter", ctypes.c_int32),
            ("biYPelsPerMeter", ctypes.c_int32),
            ("biClrUsed", ctypes.c_uint32),
            ("biClrImportant", ctypes.c_uint32),
        ]

    class BITMAPINFO(ctypes.Structure):
        _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", ctypes.c_uint32 * 0)]

    user32 = ctypes.windll.user32
    gdi32 = ctypes.windll.gdi32

    hdc_window = user32.GetWindowDC(hwnd)
    hdc_mem = gdi32.CreateCompatibleDC(hdc_window)
    hbm = gdi32.CreateCompatibleBitmap(hdc_window, width, height)
    old_hbm = gdi32.SelectObject(hdc_mem, hbm)

    gdi32.BitBlt(hdc_mem, 0, 0, width, height, hdc_window, 0, 0, win32con.SRCCOPY)

    bmp_info = BITMAPINFO()
    bmp_info.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    bmp_info.bmiHeader.biWidth = width
    bmp_info.bmiHeader.biHeight = -height
    bmp_info.bmiHeader.biPlanes = 1
    bmp_info.bmiHeader.biBitCount = 24
    bmp_info.bmiHeader.biCompression = 0
    bmp_info.bmiHeader.biSizeImage = width * height * 3

    buffer = ctypes.create_string_buffer(width * height * 3)
    gdi32.GetDIBits(hdc_mem, hbm, 0, height, buffer, bmp_info, win32con.DIB_RGB_COLORS)

    image_data = np.frombuffer(buffer, dtype=np.uint8).reshape((height, width, 3))

    gdi32.SelectObject(hdc_mem, old_hbm)
    gdi32.DeleteObject(hbm)
    gdi32.DeleteDC(hdc_mem)
    user32.ReleaseDC(hwnd, hdc_window)

    return image_data


def readText(image):
    global isCleared
    global totalFixedY
    result = ocr.ocr(image)
    if len(result) != None and result[0] != None:
        isCleared = False
        counter = 0
        while counter < len(result[0]):
            result[0][counter][1] = list(result[0][counter][1])
            if float(result[0][counter][1][1]) < 0.6 and (
                result[0][counter][0][2][1] - result[0][counter][0][0][1] < 20
                or result[0][counter][0][2][1] - result[0][counter][0][0][1] > 80
            ):
                result[0].pop(counter)
                continue
            if selectedTemplate != None:
                isFit = False
                for template in selectedTemplate:
                    if (
                        result[0][counter][0][0][0] >= template[0][0]
                        and result[0][counter][0][0][1] >= template[0][1]
                        and result[0][counter][0][1][0] <= template[1][0]
                        and result[0][counter][0][1][1] <= template[1][1]
                    ):
                        isFit = True
                        break
                if isFit == False:
                    result[0].pop(counter)
                    continue
            if (
                useCombineVertical == True
                and counter != 0
                and result[0][counter][1][0][0] != "-"
            ):
                if (
                    result[0][counter - 1][0][2][1] + 30 >= result[0][counter][0][0][1]
                    and result[0][counter - 1][0][2][1] + 10
                    < result[0][counter][0][0][1]
                ):
                    result[0][counter - 1][1][0] += " " + result[0][counter][1][0]
                    result[0][counter - 1][0][2][1] = result[0][counter][0][2][1]
                    result[0][counter - 1][0][3][1] = result[0][counter][0][2][1]
                    minX = min(
                        result[0][counter - 1][0][0][0], result[0][counter][0][0][0]
                    )
                    maxX = max(
                        result[0][counter - 1][0][1][0], result[0][counter][0][1][0]
                    )
                    result[0][counter - 1][0][0][0] = minX
                    result[0][counter - 1][0][3][0] = minX
                    result[0][counter - 1][0][1][0] = maxX
                    result[0][counter - 1][0][2][0] = maxX
                    result[0].pop(counter)
                    continue
            result[0][counter][1].pop(1)
            counter += 1
    elif isCleared == False:
        clearRectangleAll()
        isCleared = True
        return None
    return result


def calculateAndWriteDatas(results):
    global writenText
    counter = 0
    while counter < len(writenText):
        resultCounter = 0
        isFound = False
        while resultCounter < len(results[0]):
            trueCounter = 0
            matcher = SequenceMatcher(
                None, writenText[counter][1][0], results[0][resultCounter][1][0]
            )
            if matcher.ratio() > 0.8:
                trueCounter += 2
                if (
                    abs(
                        writenText[counter][0][0][1]
                        - results[0][resultCounter][0][0][1]
                    )
                    <= 20
                ):
                    trueCounter += 1
                elif (
                    abs(
                        writenText[counter][0][0][0]
                        - results[0][resultCounter][0][0][0]
                    )
                    >= 50
                ):
                    trueCounter += 1
                if trueCounter >= 3:
                    isFound = True
                    results[0].pop(resultCounter)
                    break
            resultCounter += 1
        if isFound == False:
            deletableText = writenText.pop(counter)
            clearRectangle(
                deletableText[0][0][0],
                deletableText[0][0][1],
                deletableText[0][1][0] - deletableText[0][0][0],
                deletableText[0][2][1] - deletableText[0][0][1],
            )
        else:
            counter += 1
    if results[0] == None:
        return
    for bbox, text in results[0]:
        top_left = bbox[0]
        bottom_right = bbox[2]

        x, y, width, height = (
            int(np.floor(top_left[0])),
            int(np.floor(top_left[1])),
            int(np.ceil(bottom_right[0] - top_left[0])),
            int(np.ceil(bottom_right[1] - top_left[1])),
        )

        drawRectangleWithText(
            x,
            y,
            width,
            height,
            0xFFFFFF,
            text[0],
        )
        writenText.append([bbox, text])


while True:
    if is_window_focused(selected_window):
        image = getScreenshot()
        readedtext = readText(image)
        if readedtext != None:
            calculateAndWriteDatas(readedtext)
    else:
        if isCleared == False:
            clearRectangleAll()
            isCleared = True
        allApps = gw.getAllTitles()
        if selected_window.title not in allApps:
            print("Window is closed")
