sub offset()
    cells.interior.colorindex = xlNone
    sheet1.range("A1:A16").offset(0,16).interior.colorindex=1
end sub