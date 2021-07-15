import os
import sys
import cv2
import time
import d3dshot
import win32gui
import win32con
import datetime
import win32process
import numpy as np

from imutils.object_detection import non_max_suppression
from pytessy.pytessy import PyTessy

net = cv2.dnn.readNet("fetd.pb")
ocr = PyTessy()
p1hwnd = None

layerNames = [
    "feature_fusion/Conv_7/Sigmoid",
    "feature_fusion/concat_3"]

def detect(image):
    orig = image.copy()
    if len(image.shape) == 2:
        image = cv2.cvtColor(image, cv2.COLOR_GRAY2RGB)
    (H, W) = image.shape[:2]
    (newW, newH) = (320, 320)
    rW = W / float(newW)
    rH = H / float(newH)

    image = cv2.resize(image, (newW, newH))
    (H, W) = image.shape[:2]

    blob = cv2.dnn.blobFromImage(image, 1.0, (W, H),
    	(123.68, 116.78, 103.94), swapRB=True, crop=False)
    net.setInput(blob)    
    (scores, geometry) = net.forward(layerNames)

    (numRows, numCols) = scores.shape[2:4]
    rects = []
    confidences = []
    for y in range(0, numRows):
        scoresData = scores[0, 0, y]
        xData0 = geometry[0, 0, y]
        xData1 = geometry[0, 1, y]
        xData2 = geometry[0, 2, y]
        xData3 = geometry[0, 3, y]
        anglesData = geometry[0, 4, y]
        for x in range(0, numCols):
            if scoresData[x] < 0.2:
                continue
            (offsetX, offsetY) = (x * 4.0, y * 4.0)
            angle = anglesData[x]
            cos = np.cos(angle)
            sin = np.sin(angle)
            h = xData0[x] + xData2[x]
            w = xData1[x] + xData3[x]
            endX = int(offsetX + (cos * xData1[x]) + (sin * xData2[x]))
            endY = int(offsetY - (sin * xData1[x]) + (cos * xData2[x]))
            startX = int(endX - w)
            startY = int(endY - h)
            rects.append((startX, startY, endX, endY))
            confidences.append(scoresData[x])
                        
    boxes = non_max_suppression(np.array(rects), probs=confidences)

    text = []

    for (startX, startY, endX, endY) in boxes:
        startX = int(startX * rW)
        startY = int(startY * rH)
        endX = int(endX * rW)
        endY = int(endY * rH)
        roi = orig[startY-10:endY+10, startX-10:endX+10]
        txt = None
        try:
            roi = cv2.cvtColor(roi, cv2.COLOR_BGR2RGB)
            txt = ocr.read(roi.tobytes(), roi.shape[1], roi.shape[0], 3)
        except:
            pass
        if txt:
            text.append(txt)
    return " ".join(text)   

def callback(hwnd, extra):
    global p1hwnd
    if p1hwnd:
        return
    rect = win32gui.GetWindowRect(hwnd)
    x = rect[0]
    y = rect[1]
    w = rect[2] - x
    h = rect[3] - y
    cname = win32gui.GetClassName(hwnd)
    if cname != "UnityWndClass":
        return
    title =  win32gui.GetWindowText(hwnd)
    if title.startswith("POPULATION: ONE"):
        p1hwnd = hwnd

cstrs = ["SE", "SW", "NW", "NE", "REMAINING", "SPAWNING"]

def main():
    status = None
    while not p1hwnd:
        win32gui.EnumWindows(callback, None)
        if not p1hwnd:
            print("p1 window not found")
            time.sleep(5.0)
    print("p1 window found")
    try:
        TId, PId = win32process.GetWindowThreadProcessId(p1hwnd)
        win32gui.ShowWindow(p1hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(p1hwnd)
        win32gui.SetWindowPos(p1hwnd, win32con.HWND_TOP, 435, 9, 998, 839, win32con.SWP_SHOWWINDOW)
    except Exception:
        print("failed to set size")
        time.sleep(5.0)
    print("p1 window sized")
    d = d3dshot.create(capture_output="numpy")
    d.capture()
    print("capturing")
    while win32gui.IsWindow(p1hwnd):
        try:
            img = d.screenshot(region=(435+300,9+600,435+700,9+800))
            text = detect(img)
            in_game = any([sub in text for sub in cstrs])
            if status == in_game:
                count += 1
            else:
                status = in_game
                count = 0
            if count == 5:
                print(f"status change: in_game={status} {datetime.datetime.now()}")
            time.sleep(1.0)
        except:
            d.stop()
            raise
            break
    d.stop()

if __name__ == '__main__':
    main()
