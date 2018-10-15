#Program to find the most used phrases in a text file
#Author : Mehul Suresh-Kumar
## TODO:
# Write Documentation
# Remove spaces from delete.txt

import re, string, os, sys
import os
from sys import path
import sys
from time import sleep
from collections import Counter


def query_yes_no(question, default="no"):
    """Ask a yes/no question via input() and return their answer.

    "question" is a string that is presented to the user.
    "default" is the presumed answer if the user just hits <Enter>.
        It must be "yes" (the default), "no" or None (meaning
        an answer is required of the user).

    The "answer" return value is True for "yes" or False for "no".
    """
    valid = {"yes": True, "y": True, "ye": True,
             "no": False, "n": False}
    if default is None:
        prompt = " [y/n] "
    elif default == "yes":
        prompt = " [Y/n] "
    elif default == "no":
        prompt = " [y/n] "
    else:
        raise ValueError("invalid default answer: '%s'" % default)

    while True:
        sys.stdout.write(question + prompt)
        choice = input().lower()
        if default is not None and choice == '':
            return valid[default]
        elif choice in valid:
            return valid[choice]
        else:
            sys.stdout.write("Please respond with 'yes' or 'no' "
                             "(or 'y' or 'n').\n")
repeat = True
while repeat:
    try:
        bigtxt = open('Input.txt',encoding="latin-1")
    except IOError as e:
        print ("I/O error({0}): {1}".format(e.errno, e.strerror))
        input()
        sys.exit()
    except ValueError:
        print ("Encoding Error: Could not convert data to an integer.")
        input()
        sys.exit()
    except:
        print ("Unexpected error:", sys.exc_info()[0])
        raise
        input()
        sys.exit()
    regex = re.compile('[%s]' % re.escape(string.punctuation))
    def ngrams(text, n=2):
        return zip(*[text[i:] for i in range(n)])
    phrase_count = int(input("Please enter the number of words in a phrase : "))
    common_count = int(input("Enter the number of top phrases you would like to see : "))
    ngram_counts = Counter()
    with open("Delete.txt") as f:
        content = f.readlines()
    content = [x.strip() for x in content]
    q = query_yes_no("Would you like to see the Advanced options ? :")
    if q:
        sd = int(input("Enter the number of words you would like to soft delete : "))
        sd_a = []
        for x in range(sd):
            sd_a.append(input("Delete word number "+ str(x+1) +" : ")+" ")
            print(sd_a)
        hr = int(input("Enter the number of words you would like to hard replace :"))
        hr_a = []
        for x in range(hr):
            hr_a.append([" "+input("Replacement word number "+ str(x+1) +" : ")+" "," "+input("Replace with : ")+" "])

        print("Note: Soft replace is used to replace words that were not replaced by the Hard replace. This typically happens to words in the beginning or end of sentences")
        sr = int(input("Enter the number of words you would like to soft replace :"))
        sr_a = []
        for x in range(sr):
            sr_a.append([input("Replacement word number "+ str(x+1) +" : "),input("Replace with : ")+" "])
    for bigtxt in bigtxt:
        bigtxt = regex.sub(' ', bigtxt)
        bigtxt = bigtxt.lower()
        if q:
            for x in sd_a:
                bigtxt = bigtxt.replace(x.lower()," ")
                pass
            for i,j in hr_a:
                bigtxt = bigtxt.replace(i.lower(),j)
                pass
            for i,j in sr_a:
                bigtxt = bigtxt.replace(i.lower(),j)
                pass
            pass
        for i in content:
            if i != "":
                bigtxt = bigtxt.replace(" "+i.lower()+" "," ")
                pass
            pass
        ngram_counts.update(Counter(ngrams(bigtxt.split(), phrase_count)))
        # ngram_counts = Counter(ngrams(bigtxt[3:].split(), phrase_count))
        pass
    a = ngram_counts.most_common(common_count)


    import xlwt
    from tempfile import TemporaryFile
    book = xlwt.Workbook()
    sheet = book.add_sheet('sheet1')
    try:
        os.remove("Output.xls")
    except:
        pass
    for i, l in enumerate(a):
        for j, col in enumerate(l):
            if j==0:
                sheet.write(i, j, ' '.join(map(str,col)))
                pass
            else:
                sheet.write(i, j, col)


    while True:
        try:
            name = "Output.xls"
            book.save(name)
            book.save(TemporaryFile())
        except:
            print("The excel file is already open please close the file")
            sleep(10)
            continue
        break
    try:
        os.system('start Output.xls "%s\\file.xls"' % (path[0], ))
        pass
    except Exception as e:
        print("Excel File is already open")
    print("Program Completed")
    repeat = query_yes_no("Would you like to run this program again?")
    pass
