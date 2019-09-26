#!/usr/bin/env python3

'''计算广工研究生的课程成绩积分。

硕士研究生课程成绩积分计算公式为：课程成绩积分＝40×[∑（单科成绩×学分）/∑学分]/100 ，包括全部学位课程和选修课程，即 0.4*加权平均成绩。

使用方法：在学生教务平台上导出成绩单excel文件，将其路径作为参数运行本脚本。

使用案例：
    ./calculate_score.py ~/Downloads/GDUT_ZWCJD.xls
'''

import xlrd
import sys

def cal_weighted_average_score(scores, credits):
    assert len(scores) == len(credits) 

    total_credits = sum(credits)
    total_scores = sum([score * credit for score, credit in zip(scores, credits)])

    return total_scores / total_credits

def is_valid_score(score):
    if isinstance(score, str):
        try:
            score = float(score)
        except:
            return False
    
    if score >= 0 and score <= 100:
        return True
    else:
        return False

def is_valid_credit(credit):
    if isinstance(credit, str):
        try:
            credit = float(credit)
        except:
            return False
    
    if credit >= 0 and credit <= 5:
        return True
    else:
        return False

def load_scores_and_credits(xls_path):
    wb = xlrd.open_workbook(xls_path)
    sheet = wb.sheet_by_index(0) 
 
    scores = [float(cell.value) for cell in sheet.col(13) if is_valid_score(cell.value)]
    credits = [float(cell.value) for cell in sheet.col(9) if is_valid_credit(cell.value)] 

    return scores, credits


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: %s <path to GDUT_ZWCJD.xls>" % (sys.argv[0]))
        sys.exit(1)

    xls_path = sys.argv[1]

    scores, credits = load_scores_and_credits(xls_path)
    assert len(scores) == len(credits) 

    print("You have %d courses in total:" % len(scores))
    print("\n#course_index: course_score, course_credit\n")
    for index, (score, credit) in enumerate(zip(scores, credits)):
        print("#%d: %.2f, %d" % (index+1, score, credit))

    weighted_average_score = cal_weighted_average_score(scores, credits)

    print("\nYour weighted average score is: %.2f" % weighted_average_score)
    print("Your course score point is: %.3f" % (weighted_average_score * 0.4))
