from PIL import ImageGrab
from pytesseract import pytesseract
from pynput.keyboard import Listener, KeyCode
from time import sleep, time
import pyautogui as pg
import xlwings as xw
import threading

tesseract_path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
pytesseract.tesseract_cmd = tesseract_path

exit_key = KeyCode(char=')')
start_stop_key = KeyCode(char='*')


def load_answers(filepath):
    key = xw.Book(filepath).sheets['Sheet1']
    lat_to_eng = key.range("A1:B32").value
    result_dict = {}
    for i in lat_to_eng:
        result_dict[i[0]] = i[1]

    return result_dict


def get_screen_info():
    question = ImageGrab.grab(bbox=(145, 205, 1825, 470))
    # Extract text from image
    question_text = pytesseract.image_to_string(question)

    print(question_text.strip())

    button1 = ImageGrab.grab(bbox=(160, 485, 970, 715))
    button1text = pytesseract.image_to_string(button1)
    print(button1text.strip())

    button2 = ImageGrab.grab(bbox=(995, 485, 1815, 715))
    button2text = pytesseract.image_to_string(button2)
    print(button2text.strip())

    button3 = ImageGrab.grab(bbox=(155, 735, 975, 965))
    button3text = pytesseract.image_to_string(button3)
    print(button3text.strip())

    button4 = ImageGrab.grab(bbox=(995, 735, 1815, 960))
    button4text = pytesseract.image_to_string(button4)
    print(button4text.strip())

    # show all the screenshots, only do this for troubleshooting
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
    answer = None
    a = qa_key[question]
    for button in buttons:
        if button == a:
            answer = button
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


class Run(threading.Thread):
    def __init__(self):
        super(Run, self).__init__()
        self.running = True
        self.program_run = True

    def stop(self):
        self.running = False

    def restart(self):
        self.running = True

    def run(self):
        while self.program_run:
            while self.running:
                try:
                    start = time()
                    get_screen_info()
                    end = time()
                    print(f'--- cycle completed in {end - start} seconds---')
                except Exception as err:
                    print(err)

    def exit(self):
        self.running = False
        self.program_run = False


if __name__ == '__main__':
    sleep(3.5)
    qa_key = load_answers("answer_key.xlsx")
    thread = Run()
    thread.start()


    def on_press(key):
        if key == exit_key:
            thread.exit()
            listener.stop()
        elif key == start_stop_key:
            if thread.running:
                thread.stop()
            else:
                thread.restart()
        else:
            pass


    with Listener(on_press=on_press) as listener:
        listener.join()
