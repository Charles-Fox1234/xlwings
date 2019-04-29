from xlsxwriter.utility import xl_col_to_name
import xlwings as xw
import pandas as pd

_book = 'test_named_ranges.xlsm'


def kill_broken_names(wb:xw.main.Book):
    """ Deletes all of the named ranges in a book with broken refernces. """
    names = wb.names
    for name in names:
        if '#REF!' in name.refers_to:
            del wb.names[name.name]
            

def name_used_range(wb:xw.main.Book,sheet_name:str=None):
    """ 
    When sheet_name is specified, this function names the used range of the 
    sheet to be "ur_" + the name of the sheet.
    
    When the sheet_name is not specified, this function is applied to all 
    sheets of the book.
    
    No name is applied to blank sheets.
    """
    if sheet_name is None:
        for sheet in wb.sheets:
            name_used_range(wb,sheet_name=sheet.name)
    else:
        rng = used_range(wb.sheets[sheet_name])[0]
        
        if rng is None: return False
        
        name = 'ur_'+sheet_name
        if name in wb.names:
            wb.names[name].refers_to = rng
        else:
            wb.names.add(name,rng)
            
        return True
    
            
def used_range(sht:xw.main.Sheet):
    
    """
        This function finds the used range of a sheet.  It returns the 
        range in three ways:
        
        It returns the range as a string with the sheet name.
        It returns the range as a string witout the sheet name.
        It returns a row,column pair for the bottom right corner of the sheet.
    """
    
    row = last_row(sht)
    if row == 0: return None,None,None

    column,col_letter = last_column(sht)
    partial = "a1:"+col_letter+str(row)
    
    full = "='" + sht.name + "'!" + partial
    
    return full,partial,(row,column)
    
def last_row(sht:xw.main.Sheet):
    """ Returns the row of the lowest non-empty cell in a sheet. """
    row_cell = sht.api.Cells.Find(What="*",
                   After=sht.api.Cells(1, 1),
                   LookAt=xw.constants.LookAt.xlPart,
                   LookIn=xw.constants.FindLookIn.xlFormulas,
                   SearchOrder=xw.constants.SearchOrder.xlByRows,
                SearchDirection=xw.constants.SearchDirection.xlPrevious,
                       MatchCase=False)
    
    if row_cell is None: return 0
    
    return row_cell.Row
        

def last_column(sht):
    """ Returns the row of the rightmost non-empty cell in a sheet. """
    column_cell = sht.api.Cells.Find(What="*",
                      After=sht.api.Cells(1, 1),
                      LookAt=xw.constants.LookAt.xlPart,
                      LookIn=xw.constants.FindLookIn.xlFormulas,
                      SearchOrder=xw.constants.SearchOrder.xlByColumns,
                      SearchDirection=xw.constants.SearchDirection.xlPrevious,
                      MatchCase=False)
    
    c = column_cell.Column
    return c, xl_col_to_name(c-1)



if __name__ == "__main__":
    
    try:
        wb = xw.books[_book]
    except:
        wb = xw.Book(_book)
    
    sht = wb.sheets('namer')
    kill_broken_names(wb)
    name_used_range(wb)
