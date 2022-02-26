import re
import time
import xlwt
from os import path
from glob import glob
from termcolor import colored
from pdfminer.high_level import extract_text


all_emails = []


def extractor(filename):
    text = extract_text(filename)
    emails = re.findall(r'[\w.+-]+@[\w-]+\.[\w.-]+', text)
    # 去重
    for i in emails:
        if i not in all_emails:
            all_emails.append(i)


def writeToExcel():
    wb = xlwt.Workbook()
    ws = wb.add_sheet('emails')
    row = 0
    for e in all_emails:
        ws.write(row, 0, e)
        row += 1
    if len(all_emails) < 1:
        return printError('No email accounts retrieved')
    current_time = time.strftime("%Y%m%d %H%M%S", time.localtime())
    wb.save('./output/'+current_time+'.xls')


def printError(msg):
    print(colored(msg, 'red'))


def printInfo(msg):
    print(colored(msg, 'yellow'))


def printSuccess(msg):
    print(colored(msg, 'green'))


def main():
    files = glob(path.join('./input', "*.pdf"))
    total = len(files)
    if total < 1:
        return printError('No PDF files found.')

    index = 0
    for file in files:
        index += 1
        printInfo('Extracting [%s/%s] %s' % (index, total, file))
        extractor(file)
    writeToExcel()
    printSuccess('complete, %s emails found.' % len(all_emails))


if __name__ == '__main__':
    main()
    a = input('Press any key to exit.')
    if a:
        exit(0)
