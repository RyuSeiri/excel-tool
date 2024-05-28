Option Explicit

Sub AddCustomMenu()
    Dim menu As Object
    Dim flag As Boolean
    
    For Each menu In MenuBars(xlWorksheet).Menus
        If menu.Caption = "CustomMenu(&S)" Then
            flag = False
            Exit For
        End If
        flag = True
    Next
    If flag Then
        MenuBars(xlWorksheet).Menus.Add Caption:="CustomMenu(&S)"
        With MenuBars(xlWorksheet).Menus("CustomMenu")
            Call .MenuItems.Add("SetFocusToHome(&H)", "SetFocusToHome")
            Call .MenuItems.Add("SetFocusToA1(&A)", "SetFocusToA1")
        End With
    End If
End Sub

Sub SetFocusToHome()
    Call setFocus(True)
    MsgBox ("Foucs set to Home")
End Sub

Sub SetFocusToA1()
    Call setFocus(False)
    MsgBox ("Foucs set to A1")
End Sub

Sub setFocus(homeFlg As Boolean)
    Dim book As Workbook
    Set book = ActiveWorkbook
    Dim sheetCount As Integer
    For sheetCount = book.Sheets.Count To 1 Step -1
        book.Sheets(sheetCount).Select
        book.Sheets(sheetCount).Range("A1").Select
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        If homeFlg Then
            SendKeys "^{HOME}", Wait:=True
        End If
    Next sheetCount
End Sub