Attribute VB_Name = "Module1"
Sub SimulateData()
    '# This macro simulated transcations data which is used to create customer profiles in
    '# the 'Customer Profiles' sheet.
    '#
    '# This macro first fills the Transaction ID column with numbers from 9121300 to
    '# 9121300 + Number of Transactions specified by user.
    '#
    '# The macro then fills the Old CustomerID column with numbers from
    '# 1 to the number in cell M5 and these numbers are sampled from a gamma distibution
    '# The data is autofilled and generated to match the number of transactions
    '# specified in cell M4. The Transaction Amount column as well as the Transaction Date column
    '# are also autofilled to match the number of transactions which are specified.

    Dim i As Integer
    Dim j As Integer
    
    '# i is number of distinct customers and j is number of distinct transactions
    i = Range("M5").Value + 1
    j = Range("M4").Value
    
    '# first delete old data from a previous simulation
    Rows(CStr(j + 1) & ":" & CStr(j + 1)).Select
    Rows(ActiveCell.Row & ":" & Rows.Count).Delete
    
    '# fill in the transactions ID column first for unique ids
    Range("A2").Value = 9121300
    Range("A2:A3").Select
    Selection.AutoFill Destination:=Range("A2:" & "A" & CStr(j)), Type:=xlFillDefault
    
    '# generate random numbers between zero to one to randomize the customer ids
    Range("B2").Value = "=RAND()"
    Range("B2:B3").Select
    Selection.AutoFill Destination:=Range("B2:B" & CStr(i)), Type:=xlFillDefault
    
    '# create i unique old Customer ids to make sure there are i distinct customers
    Range("C2").Value = 1
    Range("C3").Value = 2
    Range("C2:C3").Select
    Selection.AutoFill Destination:=Range("C2:C" & CStr(i)), Type:=xlFillDefault
    
    '# fill in more old customer ids to match the number of distinct transactions, this is
    '# where repeats are introduced
    Range("C" & CStr(i + 1)).Select
    Range("C" & CStr(i + 1)).Value = "=ROUND(GAMMA.INV(RAND(),R11C10,R12C10)*100,0)+1"
    Range("C" & CStr(i + 1)).Select
    Selection.AutoFill Destination:=Range("C" & CStr(i + 1) & ":C" & CStr(j)), Type:=xlFillDefault
    
    '# create random transactions amounts drawn from the gamma distibution with specified parameters
    '# given in cells j6 and j7.
    Range("F2").Value = "=(GAMMA.INV(RAND(),$J$6,$J$7)*100)+3"
    Range("F2").Select
    Selection.AutoFill Destination:=Range("F2:" & "F" & CStr(j)), Type:=xlFillDefault
    
    '# create random transactions dates between January 1st, 2017 and Feburary 1st, 2019
    '# this is based off random intergers drawn from the normal distribution
    Range("E2").Value = "=RANDBETWEEN(Date(2017,1,1),Date(2019,2,1))"
    Range("E2").Select
    Selection.AutoFill Destination:=Range("E2:" & "E" & CStr(j)), Type:=xlFillDefault
    
    Range("M6").Select '# click off data range
    
End Sub

Sub FixData()
'# This macro copies the table of simulated data and then pastes the values of the data in
'# the same place as the old data. This removes the formulas from the cells and this makes
'# it so the data becomes fixed so that the data does not keep generating each time something
'# is changed on the spreadsheet

    Columns("A:F").Select '# select entire table of data
    
    '# copy table and paste values only
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Application.CutCopyMode = False
    
    Range("K10").Select '# click off of range
End Sub

