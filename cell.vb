sub cell()
    dim icell as integer
    for icell=1 to 100
        sheet2.cells(icell, 1).value = icell
    next
end sub