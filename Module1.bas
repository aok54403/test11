Attribute VB_Name = "Module1"
'Private Sub openform1()
'    UserForm1.Show
'End Sub




Sub btm1_Click()

    Sheets("ID").range("b5:g50").clear
    'データ検索
    
    
  '  If sheets("ID").range("c2") = "" Then
  '      sheets("DB").range("a1").autofilter
   '     Exit Sub
   ' End If
    
    
    
    'データ有無チェック
    If worksheetfunction.countif(Sheets("DB").range("a:a"), Sheets("ID").range("c2")) = 0 Then
    '    sheets("DB").range("a1").autofilter
    '    Exit Sub
    
        MsgBox "No Data"
        Exit Sub
    Else
    'データ抽出
        Sheets("DB").range("a1").autofilter 1, Sheets("ID").range("c2")
        With Sheets("DB").Sort
            .SortFields.clear
            .SortFields.Add2 Key:=range("e1")
            .SetRange range("a2:f50")
            .Apply
        End With
        
        Sheets("ID").range("b5") = Sheets("DB").Cells(rows.Count, 2).End(xlUp)
        Sheets("ID").range("c5") = Sheets("DB").Cells(rows.Count, 3).End(xlUp)
        Sheets("ID").range("d5") = Sheets("DB").Cells(rows.Count, 4).End(xlUp)
        Sheets("ID").range("e5") = Sheets("DB").Cells(rows.Count, 5).End(xlUp)
        Sheets("ID").range("f5") = Sheets("DB").Cells(rows.Count, 6).End(xlUp)

        Sheets("ID").range("e5") = Format(Sheets("ID").range("e5"), "yyyy/mm/dd")
        Sheets("ID").range("f5") = Format(Sheets("ID").range("f5"), "yyyy/mm/dd")
        
'        With Sheets("DB").range("a1").currentregion
'            .resize(.rows.count - 1).offset(1, 1).copy Sheets("ID").range("b5")
'        End With
    
    End If
    
    Sheets("DB").range("a1").autofilter
    
End Sub

