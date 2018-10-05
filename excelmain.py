#Author Kushagra Kalash (TheSherlockHomie)
#Made for Win 8.1 running on 1440x900 displays, with all graphical flairs enabled
import win32api
import pyautogui
import pyperclip
import time
meX = 202     #position of excel icon on taskbar
meY = 879
neX = 300     #position of new excel window(destination) on taskbar icon expand
neY = 786
oeX = 142     #position of old excel window(source) on taskbar icon expand
oeY = 775
lbarx = 248   #leftmost corner of formula bar in excel 2013
lbary = 166
rbarx = 1408  #rightmost corner of formula bar in excel 2013
rbary = 164
trng='k'
key = 0       #needed this in previous version of code. ignore for now
for i in range(36):   #there are 36 entries
    pyautogui.moveTo(meX, meY, 0.1)          #go to main excel icon
    pyautogui.click()
    pyautogui.moveTo(oeX, oeY, 0.2)          #go to old window
    pyautogui.click()                         #click
    pyautogui.moveTo(rbarx,rbary,0.1)        #go to right corner of formula bar
    pyautogui.click()
    pyautogui.keyDown('right')               #go to the rightmost of formula bar
    time.sleep(1.5)
    pyautogui.keyUp('right')
    pyautogui.keyDown('shiftleft')           #select stuff
    pyautogui.keyDown('left')
    pyautogui.keyUp('left')
    pyautogui.keyDown('left')
    time.sleep(7)
    pyautogui.keyUp('left')
    pyautogui.keyUp('shiftleft')              #finish selecting stuff
    pyautogui.keyDown('ctrl')                #copy it
    pyautogui.keyDown('c')
    pyautogui.keyUp('c')
    pyautogui.keyUp('ctrl')
    time.sleep(1.1)
    pyautogui.typewrite('\n')                 #go to next row in same column
    trng = pyperclip.paste()                  #get copied string
    pyautogui.moveTo(meX, meY, 0.2)           #go to main excel icon
    pyautogui.click()
    pyautogui.moveTo(neX, neY, 0.2)           #go to new window
    pyautogui.click()                         #click
    pyautogui.typewrite(trng)                 #write copied stuff
    pyautogui.typewrite('\n')                 #go to next row
    time.sleep(2.1)
    trng='k'
    
