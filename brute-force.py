#
# Created by Berke Akyıldız on 21/June/2019
#
import multiprocessing
import sys
import time

import win32com.client
import zipfile2
from itertools import product

formats = "Word, Excel, PDF, WinZip, 7-Zip, RAR5"

testzip = "C:\\Users\\MONSTER\\Desktop\\radiohead-2016-nasty-little-man-SebastianEdge-billboard-1548-650.zip"
testdocx = "C:\\Users\\MONSTER\\Desktop\\test.docx"
testpdf = "C:\\Users\\MONSTER\\Desktop\\test.docx"

lowercase = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u',
             'v', 'w', 'x', 'y', 'z']
uppercase = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U',
             'V', 'W', 'X', 'Y', 'Z']
numbers = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']


def attackZip(workerNum, fileName, passFile):
    file = zipfile2.ZipFile(fileName)
    for attempt in passFile:
        try:
            password = "Worker-no: " + str(workerNum) + "    Password found: " + attempt
            file.extractall(pwd=attempt.encode("utf8"))
            print(password)
            f()
        except:
            print("Worker-no: " + str(workerNum) + "    Not matched with : " + attempt)


def createPasswordPool(*args):
    wordCount = 1
    alphabet = lowercase
    if len(args) == 1 and isinstance(args[0], int):
        wordCount = args[0]

    if len(args) == 2 and isinstance(args[1], bool):
        wordCount = args[0]
        boolUpperCase = args[1]
        if boolUpperCase:
            alphabet = alphabet + uppercase

    if len(args) == 3 and isinstance(args[2], bool):
        wordCount = args[0]
        boolUpperCase = args[1]
        if boolUpperCase:
            alphabet = alphabet + uppercase
        boolNumber = args[2]
        if boolNumber:
            alphabet = alphabet + numbers

    if len(args) == 4 and isinstance(args[3], str):
        wordCount = args[0]
        boolUpperCase = args[1]
        if boolUpperCase:
            alphabet = alphabet + uppercase
        boolNumber = args[2]
        if boolNumber:
            alphabet = alphabet + numbers
        chars = args[3]
        for char in chars:
            alphabet.append(char)

    pool = open("pool.txt", "w+")
    for length in range(1, wordCount + 1):
        to_attempt = product(alphabet, repeat=length)
        for attempt in to_attempt:
            attempt = ''.join(attempt)
            pool.write(attempt + "\n")
    return pool.name


def attackDocx(workerNum, document_path, passArray):
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = False

    for attempt in passArray:
        try:
            doc = word.Documents.Open(document_path, False, True, None, attempt)
            password = "Worker-no: " + str(workerNum) + "    Password found: " + attempt
            print(password)
        except ValueError:
            print("Worker-no: " + str(workerNum) + "    Not matched with : " + attempt)


def worker(workerNum, fileName, passArray, type):
    if type == "docx":
        attackDocx(workerNum, fileName, passArray)
    elif type == "zip":
        attackZip(workerNum, fileName, passArray)


def main():
    pool = createPasswordPool(4)

    jobs = []
    keepGoing = multiprocessing.Event()
    file = open(pool).read().splitlines()

    division = len(file) / 8
    w0 = int(1 * division)
    w1 = int(2 * division)
    w2 = int(3 * division)
    w3 = int(4 * division)
    w4 = int(5 * division)
    w5 = int(6 * division)
    w6 = int(7 * division)
    w7 = int(8 * division)

    worker0 = file[:w0]
    worker1 = file[w0:w1]
    worker2 = file[w1:w2]
    worker3 = file[w2:w3]
    worker4 = file[w3:w4]
    worker5 = file[w4:w5]
    worker6 = file[w5:w6]
    worker7 = file[w6:w7]

    # 8 Workers For 8 Cpu
    p0 = multiprocessing.Process(target=worker, args=(0, testzip, worker0, "zip"))
    jobs.append(p0)
    p0.start()

    p1 = multiprocessing.Process(target=worker, args=(1, testzip, worker1, "zip"))
    jobs.append(p1)
    p1.start()

    p2 = multiprocessing.Process(target=worker, args=(2, testzip, worker2, "zip"))
    jobs.append(p2)
    p2.start()

    p3 = multiprocessing.Process(target=worker, args=(3, testzip, worker3, "zip"))
    jobs.append(p3)
    p3.start()

    p4 = multiprocessing.Process(target=worker, args=(4, testzip, worker4, "zip"))
    jobs.append(p4)
    p4.start()

    p5 = multiprocessing.Process(target=worker, args=(5, testzip, worker5, "zip"))
    jobs.append(p5)
    p5.start()

    p6 = multiprocessing.Process(target=worker, args=(6, testzip, worker6, "zip"))
    jobs.append(p6)
    p6.start()

    p7 = multiprocessing.Process(target=worker, args=(7, testzip, worker7, "zip"))
    jobs.append(p7)
    p7.start()

    # timeout = 36000
    #
    # while timeout > 0:
    #     if not boolKeepSearching:
    #         break
    #     time.sleep(1)
    #     timeout -= 1
    while True:
        if not keepGoing.is_set():
            print("Exiting all child processess..")
            for i in jobs:
                # Terminate each process
                i.terminate()
            # Terminating main process
            sys.exit(1)
        time.sleep(2)


if __name__ == "__main__":
    main()
