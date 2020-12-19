
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
        for cell in row:
            for v in vlist:
                if not cell.value or v != cell.value:
                    break
            else:
                return cell