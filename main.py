# import math (I guess I don't need it yet)
# from openpyxl.workbook import Workbook (do I need it ?)
from openpyxl import load_workbook

wb = load_workbook('exlist.xlsx')
# print(wb.sheetnames)
wsratings = wb['rating']
wsscores = wb['tablo']
# p1row= wstable["B"]

RatingA = wsratings["B1"].value
RatingB = wsratings["B2"].value

Pawin = 1 / (1 + 10 ** ((RatingB - RatingA) / 400))
Pbwin = 1 / (1 + 10 ** ((RatingA - RatingB) / 400))


def rating_calc():
    global RatingA
    global RatingB
    RatingA = RatingA + 32 * (float(ScoreA) - Pawin)
    RatingB = RatingB + 32 * (float(ScoreB) - Pbwin)


def write_newrating():
    wsratings["C1"] = RatingA
    wsratings["C2"] = RatingB
    wb.save('exlist.xlsx')


def printall():
    # print("B=", ScoreB)
    # print("A=", ScoreA)
    print("A=", RatingA)
    print("B=", RatingB)


ScoreA = wsscores["B5"].value
# ScoreA = input("Enter the Score for A:")
ScoreB = 5

if float(ScoreA) == 1:
    ScoreB = 0
    rating_calc()
    write_newrating()
    printall()

elif float(ScoreA) == 0:
    ScoreB = 1
    rating_calc()
    write_newrating()
    printall()
elif float(ScoreA) == 0.5:
    ScoreB = 0.5
    rating_calc()
    write_newrating()
    printall()

else:
    print("error: Wrong score, please input 1, 0.5, 0 for win, draw, lose respectively")

