import xlwings as xw
import numpy as np

if __name__ == "__main__":

    survey = '../data/MPC Survey Responses.xlsx'
    sb = xw.Book(survey)
    ss = sb.sheets[0]

    vScores = ss.range('T2', 'AB60').options(np.array, ndim=2).value

    rec_idx = np.array([1,3,9])
    bel_idx = np.array([2,5,6,7])
    reh_idx = np.array([4,8])

    vividScores = []
    for row in vScores:
        rec_mean = np.mean([row[idx - 1] for idx in rec_idx])
        bel_mean = np.mean([row[idx - 1] for idx in bel_idx])
        reh_mean = np.mean([row[idx - 1] for idx in reh_idx])
        viv_vals = np.array([rec_mean, bel_mean, reh_mean])
        vividScores.append(viv_vals)

    nvivScores = np.array(vividScores)

    wb = xw.Book()
    ws = wb.sheets[0]

    ws.range('A1').value = nvivScores
    wb.save('vividScores.xlsx')
