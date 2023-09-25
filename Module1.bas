Attribute VB_Name = "Module1"
Sub btm1_Click()

    Sheets("ID").range("b5:g50").clear
    '�f�[�^����
    
    If sheets("ID").range("c2") = "" Then
        sheets("DB").range("a1").autofilter
        Exit Sub
    End If
    
    '�f�[�^�L���`�F�b�N
    If worksheetfunction.countif(Sheets("DB").range("a:a"), Sheets("ID").range("c2")) = 0 Then
    
        MsgBox "No Data"
    Else
    '�f�[�^���o
        Sheets("DB").range("a1").autofilter 1, Sheets("ID").range("c2")
        With Sheets("DB").Sort
            .SortFields.clear
            .SortFields.Add2 Key:=range("e1")
            .SetRange range("a2:f50")
            .Apply
        End With
        
    
        With Sheets("DB").range("a1").currentregion
            .resize(.rows.count - 1).offset(1, 1).copy Sheets("ID").range("b5")
        End With
    End If
    
    Sheets("DB").range("a1").autofilter
    
End Sub

