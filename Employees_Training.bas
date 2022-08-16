Attribute VB_Name = "Module1"
Option Explicit

Sub Reset()

    Dim iRow As Long
    
    iRow = [Counta(Database!A:A)] 'Identifying last row
    
    With frmForm
        
        .txtName.Value = ""
        .txtID.Value = ""
        .txtEmail.Value = ""
        .txtDepartment.Value = ""
        .txtCourse.Value = ""
        .txtStart.Value = ""
        .txtEnd.Value = ""
          
        .lstDatabase.ColumnCount = 9
        .lstDatabase.ColumnHeads = True
        .lstDatabase.ColumnWidths = "20,70,40,70,40,70,30,30,30"
        
        If iRow > 1 Then
            .lstDatabase.RowSource = "Database!A2:I" & iRow
        Else
            .lstDatabase.RowSource = "Database!A2:I2"
        End If
        
       
    End With
    

End Sub

Sub Submit()

    Dim sh As Worksheet
    Dim iRow As Long
    
    Set sh = ThisWorkbook.Sheets("Database")
    iRow = [Counta(Database!A:A)] + 1
    With sh
    
        .Cells(iRow, 1) = iRow - 1
        .Cells(iRow, 2) = frmForm.txtName.Value
        .Cells(iRow, 3) = frmForm.txtID.Value
        .Cells(iRow, 4) = frmForm.txtEmail.Value
        .Cells(iRow, 5) = frmForm.txtDepartment.Value
        .Cells(iRow, 6) = frmForm.txtCourse.Value
        .Cells(iRow, 7) = frmForm.txtStart.Value
        .Cells(iRow, 8) = frmForm.txtEnd.Value

    End With
    
End Sub

Sub Show_Form()

End Sub
