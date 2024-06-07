'''import package and module'''
import numpy as np
from numpy import *
import string
import re
import pandas as pd
import time
from spire.pdf import *
from spire.pdf.common import *
import pdfplumber
import os
from pathlib import Path

# 定义四个信息存储字符串矩阵
BasePanel = np.zeros(shape=(3, 7)).astype(np.str_)
BasePanelTitle = np.array(
        ['Year', 'Competition', 'Discipline', 'Round', 'Place', 'Date', 'Time'])
BasePanel[0, :] = BasePanelTitle
JudgePanel = np.zeros(shape=(1, 8)).astype(np.str_)
JudgePanelTitle = np.array(
        ['Year', 'Competition', 'Discipline', 'Round',  'Role',
         'JudgeGender', 'JudgeName', 'JudgeNation'])
JudgePanel[0, :] = JudgePanelTitle
TESPanel = np.zeros(shape=(1, 20)).astype(np.str_)
TESPanelTitle = np.array(
        ['Year', 'Competition', 'Discipline',  'Round', 'Rank', 'SkaterName', 'SkaterNation', 'StartingNumber',
         'TSS', 'TES', 'PCS', 'DED', 'ElementOrder', 'ExecutedElements', 'Info', 'BV', 'GOE', 'ScoreofPanel','Score', 'Role'])
TESPanel[0, :] = TESPanelTitle
PCSPanel = np.zeros(shape=(1, 17)).astype(np.str_)
PCSPanelTitle = np.array(
        ['Year', 'Competition', 'Discipline',  'Round', 'Rank', 'SkaterName', 'SkaterNation', 'StartingNumber',
         'TSS', 'TES', 'PCS', 'DED', 'PC', 'Factor', 'PScoreofPanel','PScore', 'Role'])
PCSPanel[0, :] = PCSPanelTitle

# 读取所有文件（拆分由ABBYY根据目录拆分，手动命名为标准格式）
file_list = os.listdir('D:\Research\FigureSkating\Result')
file_num = len(file_list)
runflag = 1
for i in range(file_num):
    path = str(file_list[i])
    filetypeflag = Path(path).suffix
    if filetypeflag == '.pdf':
        basestr = Path(path).stem
        basetempstr = basestr.split("_", 1)
        yeartemp = basestr[0:4]
        competitiontemp = basetempstr[0]
        disciplinetemp = basetempstr[1]

        # 开始文档信息读取
        with pdfplumber.open(path) as pdf:
            #从文档名字获取比赛基本信息（举办国/冰场信息需要手动从网页爬取匹配），一个文档对应一场比赛中的一个单项
            BasePanel[1, 0] = str(yeartemp)
            BasePanel[2, 0] = str(yeartemp)
            BasePanel[1, 1] = competitiontemp
            BasePanel[2, 1] = competitiontemp
            BasePanel[1, 2] = disciplinetemp
            BasePanel[2, 2] = disciplinetemp
            total_pages = len(pdf.pages)
            for pages in range(total_pages):
                page = pdf.pages[pages]
                # 提取表格数据（格式默认为有换行的list，根据换行符拆分行，统一由空格分隔单元
                input_str = page.extract_text()
                input_list = input_str.split('\n')
                new_list = [item.split(' ') for item in input_list]
                df = pd.DataFrame(new_list)
                matrix = df.to_numpy()
                rows, cols = matrix.shape
                matrix[matrix == None] = '99999' #用99999来标识一行的末尾
                # 根据标志词词判断该页面类型：Judges Panel or Result Panel
                ## Judges Panel
                if 'Technical' in matrix and 'Specialist' in matrix:
                    temp1 = matrix
                    rows, cols = temp1.shape
                    ## 判断是哪个项目/哪个轮次
                    if 'SHORT' in temp1 or 'RHYTHM' in temp1 or 'FREE' in temp1:
                        if 'SHORT' in temp1 or 'RHYTHM' in temp1:
                            roundtemp = 'SP'
                        if 'FREE' in temp1:
                            roundtemp = 'FS'
                    else:
                        roundtemp = 'both'
                    ## 识别裁判身份种类（Technical Specialist有时候没有Assistant前缀）
                    for r in range(0, rows - 1, 1):
                        ttemp = temp1[r, :].tolist()
                        try:
                            endindex = ttemp.index('99999')
                        except ValueError:
                            endindex = cols
                        try:
                            prefixindex = ttemp.index('Ms.')
                        except ValueError:
                            try:
                                prefixindex = ttemp.index('Mr.')
                            except ValueError:
                                pass
                        ## 用于标识有裁判有效性息的行
                        addflag = 0
                        if temp1[r, 0] == 'Technical' or temp1[r, 0] == 'Data' or temp1[r, 0] == 'Replay' or temp1[
                            r, 0] == 'Judge' or temp1[r, 0] == 'Assistant' or temp1[r, 0] == 'Referee':
                            roletemp = temp1[r, 0]
                            for rr in range(1, prefixindex, 1):
                                roletemp = roletemp + temp1[r, rr]
                            gendertemp = temp1[r, prefixindex]
                            nametemp = temp1[r, prefixindex + 1]
                            for rr in range(prefixindex + 2, endindex, 1):
                                nametemp = nametemp + temp1[r, rr]
                            nationtemp = temp1[r, endindex - 1]
                            addflag = 1
                        ## 确定JudgePanel轮次，一些比如GPF裁判会同时判两场
                        if addflag == 1:
                            if roundtemp == 'SP':
                                JudgePanelAdd = [BasePanel[1, 0], BasePanel[1, 1], BasePanel[1, 2], roundtemp, roletemp,
                                                 gendertemp, nametemp, nationtemp]
                            if roundtemp == 'FS':
                                JudgePanelAdd = [BasePanel[2, 0], BasePanel[2, 1], BasePanel[2, 2], roundtemp, roletemp,
                                                 gendertemp, nametemp, nationtemp]
                            if roundtemp == 'both':
                                JudgePanelAdd = [
                                    [BasePanel[1, 0], BasePanel[1, 1], BasePanel[1, 2], 'SP', roletemp, gendertemp,
                                     nametemp, nationtemp],
                                    [BasePanel[2, 0], BasePanel[2, 1], BasePanel[2, 2], 'FS', roletemp, gendertemp,
                                     nametemp, nationtemp]]
                            JudgePanel = np.row_stack((JudgePanel, JudgePanelAdd))
                # 处理小分表，分TES和PCS
                if 'Deductions' in matrix:
                    temp2 = matrix
                    row = 0
                    ## 早年的比赛表头是文字（ISU...）而非图片，有一些'JUDGES DETAILS PERSKATER'是单列的，需要把这些行删去
                    while True:
                        new_rows, new_cols = temp2.shape
                        if row > new_rows or row == new_rows:
                            break
                        else:
                            if temp2[row, 0] == 'ISU' or temp2[row, 0] == 'JUDGES':
                                temp2 = np.delete(temp2, row, axis=0)
                            else:
                                row = row + 1
                    ## 判断项目和轮次
                    if 'SHORT' in temp2[0, :] or 'RHYTHM' in temp2[0, :]:
                        roundtemp = 'SP'
                    if 'FREE' in temp2[0, :]:
                        roundtemp = 'FS'

                    ## 处理杂质行，只留下打分panel
                    ### 处理逻辑：只留数字开头的（人名列/技术动作）和PCS5/3大项
                    row = 0
                    while True:
                        new_rows, new_cols = temp2.shape
                        if row > new_rows or row == new_rows:
                            break
                        else:
                            if temp2[row, 0].isdigit() != 1 and temp2[row, 0] != 'Skating' and temp2[
                                row, 0] != 'Transitions' and temp2[
                                row, 0] != 'Performance' and temp2[row, 0] != 'Composition' and temp2[
                                row, 0] != 'Interpretation' and \
                                    temp2[row, 0] != 'Presentation':
                                temp2 = np.delete(temp2, row, axis=0)
                            else:
                                row = row + 1
                    matrix_rows, matrix_cols = temp2.shape
                    ### 2019年识别行数错位，处理成正确顺序（需要手动发现改正）
                    for r in range(matrix_rows):
                        if temp2[r, 2].isdigit() != 1 and temp2[r, 2].find('.') == -1 and temp2[r, 3].find(
                                '.') == -1 and temp2[
                            r - 1, 0] == '1':
                            temp2[(r - 1, r), :] = temp2[(r, r - 1), :]
                    ### 处理符号、识别导致的未对齐问题
                    for r in range(matrix_rows):
                        #### 抓符号
                        if temp2[r, 2] == '!' or temp2[r, 2] == 'e' or temp2[r, 2] == 'q' or temp2[r, 2] == '<' or \
                                temp2[
                                    r, 2] == '<<' or temp2[r, 2] == 'F' or temp2[r, 2] == '*' or temp2[r, 2] == '>':
                            #### 抓符号，后半跳跃（后续处理需要单独识别，OCR把后半跳跃×识别为x）
                            if temp2[r, 4].find('.') == -1:
                                infotemp = temp2[r, 2] + temp2[r, 4]
                                temp2[r, 2] = temp2[r, 3]
                                temp2[r, 3] = temp2[r, 5]
                                for rr in range(4, matrix_cols - 2):
                                    temp2[r, rr] = temp2[r, rr + 2]
                                temp2[r, matrix_cols - 2] = infotemp
                                temp2[r, matrix_cols - 1] = infotemp
                            #### 抓符号，非后半跳跃
                            else:
                                infotemp = temp2[r, 2]
                                for rr in range(2, matrix_cols - 1):
                                    temp2[r, rr] = temp2[r, rr + 1]
                                temp2[r, matrix_cols - 1] = infotemp
                        else:
                            #### 不抓符号，后半跳跃
                            if temp2[r, 3] == 'x':
                                infotemp = temp2[r, 3]
                                for rr in range(3, matrix_cols - 1):
                                    temp2[r, rr] = temp2[r, rr + 1]
                                temp2[r, matrix_cols - 1] = infotemp
                        # 调整被空格分段识别成多格的PCS项（五大项和新三大项）
                        if temp2[r, 0] == 'Skating':
                            for rr in range(1, matrix_cols - 1):
                                temp2[r, rr] = temp2[r, rr + 1]
                            temp2[r, matrix_cols - 1] = '99999'
                        if temp2[r, 0] == 'Interpretation':
                            for rr in range(1, matrix_cols - 3):
                                temp2[r, rr] = temp2[r, rr + 3]
                            temp2[r, matrix_cols - 3] = '99999'
                            temp2[r, matrix_cols - 2] = '99999'
                            temp2[r, matrix_cols - 1] = '99999'

                    ## 分割小分表，整理TES和PCS表
                    matrix_rows, matrix_cols = temp2.shape
                    titleindex = 0
                    for r in range(1, matrix_rows, 1):
                        ### 判断行光标位置：更换titleindex/选择更新TESPanel/更新PCSPanel
                        #### TES
                        if temp2[r, 0].isdigit() and temp2[r - 1, 0].isdigit(): #识别包含技术动作打分信息的行（本行和前一行都以数字开头
                            ##### 处理resultpanel(表头基础信息)
                            ttemp = temp2[titleindex, :].tolist()
                            try:
                                endindex = ttemp.index('99999')
                            except ValueError:
                                endindex = matrix_cols - 1
                            ranktemp = temp2[titleindex, 0]
                            skaternationtemp = temp2[titleindex, endindex - 6]
                            startingnumbertemp = temp2[titleindex, endindex - 5]
                            tsstemp = temp2[titleindex, endindex - 4]
                            testmep = temp2[titleindex, endindex - 3]
                            pcstemp = temp2[titleindex, endindex - 2]
                            dedtemp = temp2[titleindex, endindex - 1]
                            skaternametemp = temp2[titleindex, 1]
                            for rr in range(2, endindex - 6):
                                skaternametemp = skaternametemp + temp2[titleindex, rr]
                            ##### 处理每一行TES
                            elementordertemp = temp2[r, 0]
                            executedelementstemp = temp2[r, 1]
                            ttempr = temp2[r, :].tolist()
                            try:
                                endindexr = ttempr.index('99999')
                            except ValueError:
                                endindexr = matrix_cols - 1 #用于标识裁判数量
                            TESPanelAdd = np.zeros(shape=(endindexr - 5, 20)).astype(np.str_)
                            for rr in range(4, endindexr - 1): #每个裁判对每个技术动作的打分自成一行
                                if temp2[r, matrix_cols - 1] != '99999':
                                    infotemp = temp2[r, matrix_cols - 1]
                                else:
                                    infotemp = ' '
                                bvtemp = temp2[r, 2]
                                goetemp = temp2[r, 3]
                                scoreofpaneltemp = temp2[r, endindexr - 1]
                                scoretemp = temp2[r, rr]
                                roletemp = 'J' + str(rr - 3)
                                TESPanelCurrentAdd = [BasePanel[1, 0], BasePanel[1, 1], BasePanel[1, 2], roundtemp,
                                                      ranktemp,
                                                      skaternametemp, skaternationtemp, startingnumbertemp,
                                                      tsstemp, testmep, pcstemp, dedtemp, elementordertemp,
                                                      executedelementstemp,
                                                      infotemp, bvtemp, goetemp, scoreofpaneltemp, scoretemp, roletemp]
                                TESPanelAdd[rr - 4, :] = TESPanelCurrentAdd
                            TESPanel = np.row_stack((TESPanel, TESPanelAdd))

                        #### PCS
                        if temp2[r, 0] == 'Composition' or temp2[r, 0] == 'Presentation' or temp2[r, 0] == 'Skating' or \
                                temp2[
                                    r, 0] == 'Transitions' or temp2[r, 0] == 'Interpretation' or temp2[
                            r, 0] == 'Performance': #识别包含节目内容分打分信息的行（P分五大/三大项首词）
                            ttempr = temp2[r, :].tolist()
                            try:
                                endindexr = ttempr.index('99999')
                            except ValueError:
                                endindexr = matrix_cols - 1
                            PCSPanelAdd = np.zeros(shape=(endindexr - 3, 17)).astype(np.str_)
                            for rr in range(2, endindexr - 1): #每个裁判对每个节目内容分项的打分自成一行
                                pctemp = temp2[r, 0]
                                factortemp = temp2[r, 1]
                                pscoreofpaneltemp = temp2[r, endindexr - 1]
                                pscoretemp = temp2[r, rr]
                                roletemp = 'J' + str(rr - 1)
                                PCSPanelCurrentAdd = [BasePanel[1, 0], BasePanel[1, 1], BasePanel[1, 2], roundtemp,
                                                      ranktemp,
                                                      skaternametemp, skaternationtemp, startingnumbertemp,
                                                      tsstemp, testmep, pcstemp, dedtemp, pctemp, factortemp,
                                                      pscoreofpaneltemp, pscoretemp, roletemp]
                                PCSPanelAdd[rr - 2, :] = PCSPanelCurrentAdd
                            PCSPanel = np.row_stack((PCSPanel, PCSPanelAdd))

                        #### 更新titleindex，一页可能有一/二/三个选手的小分表
                        if temp2[r, 0].isdigit() and temp2[r + 1, 0].isdigit() and temp2[r - 1, 0].isdigit() != 1: #识别换下一个选手
                            titleindex = r

        print(runflag) #跑的时候可以看看进度到哪里了，我好急
        runflag = runflag + 1



np.savetxt('JudgePanel.csv', JudgePanel, delimiter = ',', fmt='%s')
np.savetxt('TESPanel.csv', TESPanel, delimiter = ',', fmt='%s')
np.savetxt('PCSPanel.csv', PCSPanel, delimiter = ',', fmt='%s')





