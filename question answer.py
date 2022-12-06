from PIL import ImageGrab
from pytesseract import pytesseract
from pynput.keyboard import Listener, KeyCode
from time import sleep
import pyautogui as pg
import xlwings as xw
import threading

tesseract_path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
pytesseract.tesseract_cmd = tesseract_path

exit_key = KeyCode(char='|')


def get_screen_info():
    question = ImageGrab.grab(bbox=(145, 205, 1825, 470))
    # Extract text from image
    question_text = pytesseract.image_to_string(question)

    print(question_text.strip())
    sleep(0.2)

    button1 = ImageGrab.grab(bbox=(160, 485, 970, 715))
    button1text = pytesseract.image_to_string(button1)
    print(button1text.strip())
    sleep(0.2)

    button2 = ImageGrab.grab(bbox=(995, 485, 1815, 715))
    button2text = pytesseract.image_to_string(button2)
    print(button2text.strip())
    sleep(0.2)

    button3 = ImageGrab.grab(bbox=(155, 735, 975, 965))
    button3text = pytesseract.image_to_string(button3)
    print(button3text.strip())
    sleep(0.2)

    button4 = ImageGrab.grab(bbox=(995, 735, 1815, 960))
    button4text = pytesseract.image_to_string(button4)
    print(button4text.strip())
    sleep(0.2)

    # show all the screenshots
    """question.show()
    button1.show()
    button2.show()
    button3.show()
    button4.show()"""

    ans = calculate_ans(question_text.strip(), button1text.strip(), button2text.strip(), button3text.strip(),
                        button4text.strip())
    print("Answer: " + ans)

    buttons = {button1text.strip(): button1,
               button2text.strip(): button2,
               button3text.strip(): button3,
               button4text.strip(): button4}
    for i in buttons:
        if ans == i:
            coords = pg.locateOnScreen(buttons[i])
            pg.moveTo(coords)
            pg.click()
        else:
            pass

    press_continue_button()


def calculate_ans(question, b1, b2, b3, b4):
    print(f"Question: {question}\nA: {b1}\nB: {b2}\nC: {b3}\nD: {b4}")
    buttons = [b1, b2, b3, b4]
    key = xw.Book("answer_key.xlsx").sheets['Sheet1']
    lat_to_eng = key.range("A1:B30").value
    answer = "None"
    for i in lat_to_eng:
        if str(i[0]).lower() == question.lower():
            for b in buttons:
                if b.lower() == i[1].lower():
                    answer = i[1]
                    break
        else:
            pass

    return answer


def press_continue_button():
    try:
        sleep(0.2)  # cooldown
        continue_button = pg.center(pg.locateOnScreen('continue.png'))
        pg.moveTo(continue_button)
        pg.click(continue_button)
    except TypeError:
        continue_button = pg.center(pg.locateOnScreen('continue2.png'))
        pg.moveTo(continue_button)
        pg.click(continue_button)


def on_press(key):
    if key == exit_key:
        thread.exit()
        listener.stop()
    else:
        pass


class Run(threading.Thread):
    def __init__(self):
        super(Run, self).__init__()
        self.running = True

    def run(self):
        while self.running:
            try:
                get_screen_info()
                print('--- cycle completed ---')
            except Exception as err:
                print(err)

    def exit(self):
        self.running = False


if __name__ == '__main__':
    sleep(3.5)
    thread = Run()
    thread.start()

    with Listener(on_press=on_press) as listener:
        listener.join()
