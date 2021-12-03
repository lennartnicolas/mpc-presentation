import xlwings as xw
import numpy as np

if __name__ == "__main__":

    # Set workbooks and sheets
    worksheet = '../data/MusEQ Score Worksheet.xlsx'
    survey    = '../data/MPC Survey Responses.xlsx'
    wb = xw.Book(worksheet)
    sb = xw.Book(survey)
    ws = wb.sheets[0]
    ss = sb.sheets[0]

    # Get scores and convert daily and perform values and shrink idx
    scores  = ss.range('G2','N60').options(np.array, ndim=2).value
    sscores = []
    for row in scores:
        dm      = np.ceil(np.mean([row[0],row[1]]))
        pm      = np.ceil(np.mean([row[3],row[4]]))
        nrow    = np.array([dm,row[2],pm,row[5],row[6],row[7]])
        sscores.append(nrow)
    nscores = np.array(sscores)

    # MusEQ mappings
    daily_idx = np.array([1,11,14,19,34,35])
    emoti_idx = np.array([8,15,16,27,28,29,31,32])
    perfo_idx = np.array([2,6,7,20,22,23,24])
    consu_idx = np.array([10,12,18,21,25,26])
    respo_idx = np.array([3,4,5,17])
    prefe_idx = np.array([9,13,30,33])


    # Fill in the values and return the score
    musEQ = np.array([])
    for row in nscores:
        for idx, val in enumerate(row):
            if idx == 0:
                for ix in daily_idx:
                    ws.range('J' + str(ix + 3)).value = val
            if idx == 1:
                for ix in emoti_idx:
                    ws.range('J' + str(ix + 3)).value = val
            if idx == 2:
                for ix in perfo_idx:
                    ws.range('J' + str(ix + 3)).value = val
            if idx == 3:
                for ix in consu_idx:
                    ws.range('J' + str(ix + 3)).value = val
            if idx == 4:
                for ix in respo_idx:
                    ws.range('J' + str(ix + 3)).value = val
            if idx == 5:
                for ix in prefe_idx:
                    ws.range('J' + str(ix + 3)).value = val

        score = ws.range('F54').value
        musEQ = np.append(musEQ, score)
        print(score)

    # Save values in new Excelsheet
    scb = xw.Book()
    scs = scb.sheets[0]
    scs.range('A1').value = musEQ.reshape(musEQ.shape + (1,))
    scb.save('musEQscores.xlsx')
