Attribute VB_Name = "Module1"




Sub btm1_Click()

    Sheets("ID").Range("a5:g50").clear
    
    '全件表示
    If Sheets("ID").Range("c2") = "9999" Then
        With Sheets("DB").Sort
            .SortFields.clear
            .SortFields.Add2 Key:=Range("a1")
            .SortFields.Add2 Key:=Range("e1")
            .SetRange Range("a2:f50")
            .Apply
        End With
    
        With Sheets("DB").Range("a1").currentregion
            .resize(.rows.Count - 1).offset(1, 0).copy Sheets("ID").Range("a5")
        End With
    Else
    
    'データ有無チェック
        If worksheetfunction.countif(Sheets("DB").Range("a:a"), Sheets("ID").Range("c2")) = 0 Then
    '    sheets("DB").range("a1").autofilter
    '    Exit Sub
    
            MsgBox "No Data"
            Exit Sub
        Else
    'データ抽出
    '    Sheets("DB").range("a1").autofilter 1, Sheets("ID").range("c2")
            With Sheets("DB").Sort
                .SortFields.clear
                .SortFields.Add2 Key:=Range("a1")
                .SortFields.Add2 Key:=Range("e1")
                .SetRange Range("a2:f50")
                .Apply
            End With

    'データ抽出
            Sheets("DB").Range("a1").autofilter 1, Sheets("ID").Range("c2")
        
            Sheets("ID").Range("b5") = Sheets("DB").Cells(rows.Count, 2).End(xlUp)
            Sheets("ID").Range("c5") = Sheets("DB").Cells(rows.Count, 3).End(xlUp)
            Sheets("ID").Range("d5") = Sheets("DB").Cells(rows.Count, 4).End(xlUp)
            Sheets("ID").Range("e5") = Sheets("DB").Cells(rows.Count, 5).End(xlUp)
            Sheets("ID").Range("f5") = Sheets("DB").Cells(rows.Count, 6).End(xlUp)

            Sheets("ID").Range("e5") = Format(Sheets("ID").Range("e5"), "yyyy/mm/dd")
            Sheets("ID").Range("f5") = Format(Sheets("ID").Range("f5"), "yyyy/mm/dd")

        End If
    End If
    
    Sheets("DB").Range("a1").autofilter
    
End Sub

