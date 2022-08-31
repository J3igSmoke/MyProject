import cv2
import numpy as np
import pyautogui


slika = pyautogui.screenshot()
slika.save('zajeta slika1.jpg')
img = cv2.imread("zajeta slika1.jpg");
gimg = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
template = cv2.imread("mario.PNG", cv2.IMREAD_GRAYSCALE);

w, h = template.shape[::-1]


razultat = cv2.matchTemplate(gimg, template, cv2.TM_CCOEFF_NORMED)


loc = np.where(razultat >= 0.84) #Natačnost

for pt in zip(*loc[::-1]):
    cv2.rectangle(img, pt, (pt[0] + w, pt[1] + h), (50,255,20), 1)

cv2.imshow('res.png', img)

cv2.waitKey(0)
