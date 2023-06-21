import pyautogui
import pyperclip
import time
import re
    
# Copy table data from excel
pyautogui.click(160, 20)
pyautogui.hotkey('alt', 'f3')
time.sleep(0.3)
pyautogui.press('delete')
pyautogui.typewrite('D8:E30')
pyautogui.press('enter')
pyautogui.hotkey('ctrl', 'c')
time.sleep(0.3)

# Save table data as 2 arrays
clipboard_text = pyperclip.paste()

lines = clipboard_text.split('\n')

prodcode = []
cost = []
Listofprod = 2

for line in lines:
    
    parts = line.split('\t')
    
    variable = parts[0]
    value = parts[1]
    
    prodcode.append(variable)
    cost.append(value)

# Print the prodcode and cost
for variable, value in zip(prodcode, cost):
    print(f"prodcode: {variable}, cost: {value}")

# WCD to denticon code conversions
for i in range(len(prodcode)):
    if prodcode[i] == '8750':
        prodcode[i] = '8680'
    elif prodcode[i] == '8750A':
        prodcode[i] = '8680A'
    elif prodcode[i] == '0553':
        prodcode[i] = 'z0053'
    elif prodcode[i] == '553':
        prodcode[i] = 'z0053'
    elif prodcode[i] == '0052':
        prodcode[i] = 'z0052'
    elif prodcode[i] == '0038':
        prodcode[i] = 'z0038'
    elif prodcode[i] == '0000':
        prodcode[i] = '8999'
    elif prodcode[i] == '0':
        prodcode[i] = '8999'


for i in range(0, 22):

    if prodcode[i].strip() == "":
        pass

# 8670 code case
    elif prodcode[i].startswith("8670"):
        
        if '*' in prodcode[i]:
        # getting number of prod contracts in ()
            match = re.search(r'\((\d+)\*', prodcode[i])
            number = match.group(1)
            contracts86 = int(number)
            match = re.search(r'\*(\d+)\)', prodcode[i])
            number = match.group(1)
            contractprice = int(number)


            for i in range(0, contracts86):
                
                pyperclip.copy(8670)

                # get rid of the last digit so production won't auto fill
                # for z codes
                cliptemp = pyperclip.paste()
                takelastdigit = cliptemp[:-1]
                pyperclip.copy(takelastdigit)

                # click code entry and enter
                pyautogui.click(550, 773)
                pyautogui.press('delete')
                pyautogui.hotkey('ctrl', 'v')
                pyautogui.press('enter')
                time.sleep(0.3)

                # click code description box
                pyautogui.click(1100, 750)

                # hotkey 1 pending tx
                pyautogui.hotkey('alt', 'a')
                pyautogui.hotkey('alt', 'a')
                time.sleep(1.5)

                # click text box
                pyautogui.click(1200, 400)
                pyautogui.click(1100, 400)

                # tab to correct tx
                for _ in range(Listofprod):
                    pyautogui.press('tab')
                time.sleep(0.3)
                pyautogui.press('enter')

                pyperclip.copy(contractprice)
                time.sleep(4)

                # click on fee input box
                pyautogui.click(260, 610)
                pyautogui.click(260, 610)
                pyautogui.click(260, 610)
                pyautogui.hotkey('ctrl', 'v')
                pyautogui.press('tab')
                time.sleep(0.3)

                # in case of popup
                pyautogui.press('space')
                
                # click to uncheck Re-Estimate at Posting
                pyautogui.click(1110, 613)

                # click advanced
                pyautogui.click(960, 613)
                time.sleep(0.3)

                # 0 out INS production
                pyautogui.press('tab')
                pyautogui.typewrite('0')
                pyautogui.press('tab')
                pyautogui.typewrite('0')
                pyautogui.press('tab')

                # zeroing deductable popup
                pyautogui.press('space')

                # tab to save button and enter
                pyautogui.press('tab')
                pyautogui.press('tab')
                pyautogui.press('tab')
                pyautogui.press('tab')
                pyautogui.press('enter')
                time.sleep(5)

                # Contract finished
                numcont = Listofprod / 2

                print("Number of contracts done:", numcont)

                Listofprod+=2

                
        elif '(' in prodcode[i] and ')' in prodcode[i]:
            
        # getting number of prod contracts in ()
            match = re.search(r'\((\d+)\)', prodcode[i])
            if match:
                number = match.group(1)
                contracts0 = int(number)
                for i in range(0, contracts0):
                    
                    pyperclip.copy(8670)

                    # get rid of the last digit so production won't auto fill
                    cliptemp = pyperclip.paste()
                    takelastdigit = cliptemp[:-1]
                    pyperclip.copy(takelastdigit)

                    # click code entry and enter
                    pyautogui.click(550, 773)
                    pyautogui.press('delete')
                    pyautogui.hotkey('ctrl', 'v')
                    pyautogui.press('enter')
                    time.sleep(0.3)

                    # click code description box
                    pyautogui.click(1100, 750)

                    # hotkey 1 pending tx
                    pyautogui.hotkey('alt', 'a')
                    pyautogui.hotkey('alt', 'a')
                    time.sleep(2)

                    # click text box
                    pyautogui.click(1200, 400)
                    pyautogui.click(1100, 400)

                    # tab to correct tx
                    for _ in range(Listofprod):
                        pyautogui.press('tab')
                    time.sleep(0.3)
                    pyautogui.press('enter')

                    pyperclip.copy(0)
                    time.sleep(4)

                    # click on fee input box
                    pyautogui.click(260, 610)
                    pyautogui.click(260, 610)
                    pyautogui.click(260, 610)
                    pyautogui.hotkey('ctrl', 'v')
                    pyautogui.press('tab')

                    # in case of popup
                    pyautogui.press('space')
                    
                    # click to uncheck Re-Estimate at Posting
                    pyautogui.click(1110, 613)

                    # click advanced
                    pyautogui.click(960, 613)
                    
                    time.sleep(0.3)

                    # 0 out ins production
                    pyautogui.press('tab')
                    pyautogui.typewrite('0')
                    pyautogui.press('tab')
                    pyautogui.typewrite('0')
                    pyautogui.press('tab')
                    pyautogui.press('space')

                    # tab to save button and enter
                    pyautogui.press('tab')
                    pyautogui.press('tab')
                    pyautogui.press('tab')
                    pyautogui.press('tab')
                    pyautogui.press('enter')
                    time.sleep(5)

                    numcont = Listofprod / 2

                    print("Number of contracts done:", numcont)

                    # Contract finished
                    Listofprod+=2

        elif prodcode[i] == '8670':

            pyperclip.copy(prodcode[i])

            # get rid of the last digit so production won't auto fill
            cliptemp = pyperclip.paste()
            takelastdigit = cliptemp[:-1]
            pyperclip.copy(takelastdigit)

            # click code entry and enter
            pyautogui.click(550, 773)
            pyautogui.press('delete')
            pyautogui.hotkey('ctrl', 'v')
            pyautogui.press('enter')
            time.sleep(0.3)

            # click code description box
            pyautogui.click(1100, 750)

            # hotkey 1 pending tx
            pyautogui.hotkey('alt', 'a')
            pyautogui.hotkey('alt', 'a')
            time.sleep(2)

            # click text box
            pyautogui.click(1200, 400)
            pyautogui.click(1100, 400)
            # tab to correct tx
            for _ in range(Listofprod):
                pyautogui.press('tab')
            time.sleep(0.3)
            pyautogui.press('enter')

            pyperclip.copy(cost[i])
            time.sleep(4)

            # click on fee input box
            pyautogui.click(260, 610)
            pyautogui.click(260, 610)
            pyautogui.click(260, 610)
            pyautogui.hotkey('ctrl', 'v')
            pyautogui.press('tab')

            # in case of popup
            pyautogui.press('space')
            
            # click Re-Estimate at Posting
            pyautogui.click(1110, 613)

            # click advanced
            pyautogui.click(960, 613)
            time.sleep(0.3)

            # 0 out ins production
            pyautogui.press('tab')
            pyautogui.typewrite('0')
            pyautogui.press('tab')
            pyautogui.typewrite('0')
            pyautogui.press('tab')
            pyautogui.press('space')

            # tab to save button and enter
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('enter')
            time.sleep(5)

            # Contract finished

            numcont = Listofprod / 2

            print("Number of contracts done:", numcont)

            Listofprod+=2
            
    # Other codes
    else:
        
        pyperclip.copy(prodcode[i])
        if 'A' in prodcode[i]:
            descy = 773
        elif 'B' in prodcode[i]:
            descy = 790
        elif '38' in prodcode[i]:
            descy = 829
        else:
            descy = 753

        # get rid of the last digit so production won't auto fill
        cliptemp = pyperclip.paste()
        takelastdigit = cliptemp[:-1]
        pyperclip.copy(takelastdigit)

        # click code entry and enter
        pyautogui.click(550, 773)
        pyautogui.press('delete')
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
        time.sleep(0.3)

        # click code description box
        pyautogui.click(1100, descy)

        # hotkey 1 pending tx
        pyautogui.hotkey('alt', 'a')
        pyautogui.hotkey('alt', 'a')
        time.sleep(2)

        # click text box
        pyautogui.click(1200, 400)
        pyautogui.click(1100, 400)

        # tab to correct tx
        for _ in range(Listofprod):
            pyautogui.press('tab')
        time.sleep(0.3)
        pyautogui.press('enter')

        pyperclip.copy(cost[i])
        time.sleep(4)

        # click on fee input box
        pyautogui.click(260, 610)
        pyautogui.click(260, 610)
        pyautogui.click(260, 610)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('tab')

        # in case of popup
        pyautogui.press('space')
        
        # click Re-Estimate at Posting
        pyautogui.click(1110, 613)

        # click advanced
        pyautogui.click(960, 613)
        time.sleep(0.3)

        # 0 out ins production
        pyautogui.press('tab')
        pyautogui.typewrite('0')
        pyautogui.press('tab')
        pyautogui.typewrite('0')
        pyautogui.press('tab')
        pyautogui.press('space')

        # tab to save button and enter
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('enter')
        time.sleep(5)

        # Contract finished

        numcont = Listofprod / 2

        print("Number of contracts done:", numcont)

        Listofprod+=2

        

