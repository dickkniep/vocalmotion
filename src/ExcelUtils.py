from decimal import Decimal

def find_value(value, ws, min_row: int = 1, max_row: int = None, min_col: int = 1, max_col: int = None):
    """
    Finds value in a Excel worksheet

    :param value:
    :param ws:
    :return:
    """
    if max_row is None:
        max_row = ws.max_row
    if max_col is None:
        max_col = ws.max_column

    if not isinstance(value, list):
        vlist = [value]
    else:
        vlist = value

    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        vfound = [0]*len(vlist)
        for cell in row:
            for idx, v in enumerate(vlist):
                if (not isinstance(v, float) and isinstance(cell.value, float) and v == round(Decimal(cell.value), 2)) \
                        or cell.value and v == cell.value:
                    vfound[idx] = 1
            if sum(vfound) == len(vlist):
                return cell
    return

def find_empty_row(ws, startrow: int = 2):
    idx = 0
    for idx, row in enumerate(ws.rows):
        if idx >= startrow and not row[0].value:
            return idx
    else:
        return idx + 1

def addlist(ws, inputlist, startpos=None, blankrow = 0):  # This function is to input an array/list from a input start point; len(startpos) must equal 2, where startpos = [1,1] is cell 1A. inputlist must be a two dimensional array; if you wish to input a single row then this can be done where len(inputlist) == 1, e.g. inputlist = [[1,2,3,4]]
    if startpos is None:
        startpos = [find_empty_row(ws) + 1 + blankrow, 1]
    if not isinstance(inputlist[0], list):
        inputlist = [inputlist]

    rownr = 0
    for rownr, row in enumerate(inputlist):  # For every row in inputlist
        for colnr, value in enumerate(row):  # For every cell in row
            ws.cell(row=startpos[0] + rownr, column=startpos[1] + colnr).value = value  # Set value for current cell
    return startpos[0] + rownr
