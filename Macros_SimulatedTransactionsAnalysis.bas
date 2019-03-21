Attribute VB_Name = "Module1"
Sub Populate_Other_Columns()
Attribute Populate_Other_Columns.VB_ProcData.VB_Invoke_Func = " \n14"
'
' This macro fills the Old CustomerID column with numbers from
' 1 to the number in cell M5 and these numbers are sampled from a gamma distibution
' The data is autofilled and generated to match the number of transactions
' specified in cell M4. The Transaction Amount column as well as the Transaction Date column
' are also autofilled to match the number of transactions which are specified.

    Dim i As Integer 'fix data type
    Dim j As Integer
    
    i = Range("M5").Value + 1 'number of customers
    j = Range("M4").Value  'number of transactions
    
    
    Range("C2:C3").Select
    Selection.AutoFill Destination:=Range("C2:C" & CStr(i)), Type:=xlFillDefault 'populate cells below
    Range("C" & CStr(i + 1)).Select
    Range("C" & CStr(i + 1)).Value = "=ROUND(GAMMA.INV(RAND(),R11C10,R12C10)*100,0)+1" 'sample from gamma
    Range("C" & CStr(i + 1)).Select
    Selection.AutoFill Destination:=Range("C" & CStr(i + 1) & ":C" & CStr(j)), Type:=xlFillDefault
    
    Range("F2").Select 'Transaction Amount column
    Selection.AutoFill Destination:=Range("F2:" & "F" & CStr(j)), Type:=xlFillDefault
    
    Range("E2").Select 'Transaction Date column
    Selection.AutoFill Destination:=Range("E2:" & "E" & CStr(j)), Type:=xlFillDefault
    
    Range("C2:C3").Select 'get back to where you started
 
    
End Sub
Sub Populate_TransactionIds()
Attribute Populate_TransactionIds.VB_ProcData.VB_Invoke_Func = " \n14"
    ' This macro fills the Transaction ID column with numbers from 9121300 to
    ' 9121300 + Number of Transactions specified by user
    
    Dim j As Integer
    
    j = Range("M4").Value 'number of transactions is specified in the cell M4
    
    Rows(CStr(j + 1) & ":" & CStr(j + 1)).Select 'the next two lines delete any rows below which could be the result of previous macro runs
    Rows(ActiveCell.Row & ":" & Rows.Count).Delete
    
    Range("A2").Value = 9121300  'starting point for transactionID, this number is an arbitrary fixed large number
    Range("A2:A3").Select
    Selection.AutoFill Destination:=Range("A2:" & "A" & CStr(j)), Type:=xlFillDefault 'populate cells below
 
End Sub
