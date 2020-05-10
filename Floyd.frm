Option Explicit

Private Sub CommandButton1_Click()

    If matrix = "" Then
        MsgBox "輸入矩陣欄位不可空白!", , "資料錯誤"
        Exit Sub
    End If
    
    Dim ipmatrix
    Dim c_r, c_c, o_r, o_c
    Dim d_D, d_N
    Dim i, j, k
    
    Set ipmatrix = Range(matrix)
    
    c_r = ipmatrix.Rows.Count
    c_c = ipmatrix.Columns.Count
    o_r = ipmatrix.Row
    o_c = ipmatrix.Column
    
    If c_r <> c_c Then
        MsgBox "行數與列數不相等。", , "行列錯誤"
        Exit Sub
    End If
    
    ReDim d_D(c_r, c_c)
    ReDim d_N(c_r, c_c)
    
    Cells(o_r + c_r + 1, o_c) = "Distance matrix"
    Cells(o_r + c_r + 1, o_c + c_c + 1) = "Node matrix"
    
    For i = 1 To c_r
        For j = 1 To c_c
            If i = j Then
                d_D(i, j) = "-"
                d_N(i, j) = "-"
            Else
                d_D(i, j) = ipmatrix(i, j)
                d_N(i, j) = j
            End If
            Cells(o_r + (i - 1), o_c + c_c + j) = d_N(i, j)
        Next j
    Next i
    
    For k = 1 To c_r
        For i = 1 To c_r
            For j = 1 To c_c
                If i = j Then
                    d_D(i, j) = "-"
                ElseIf i = k Or j = k Then
                    d_D(i, j) = d_D(i, j)
                ElseIf d_D(i, j) > d_D(i, k) + d_D(k, j) Then
                    d_D(i, j) = d_D(i, k) + d_D(k, j)
                    d_N(i, j) = k
                    If k = 1 Then
                        Cells(o_r + (i - 1), o_c + (j - 1)).Interior.Color = RGB(0, 255, 0)
                        Cells(o_r + (i - 1), (o_c + c_c) + j).Interior.Color = RGB(0, 255, 0)
                    Else
                        Cells((o_r + c_r + 1) + (k - 2) * (c_r + 1) + i, o_c + (j - 1)).Interior.Color = RGB(0, 255, 0)
                        Cells((o_r + c_r + 1) + (k - 2) * (c_r + 1) + i, o_c + c_c + j).Interior.Color = RGB(0, 255, 0)
                    End If
                End If
                Cells(o_r + k * ((c_r) + 1) + i, o_c + (j - 1)) = d_D(i, j)
                Cells(o_r + k * ((c_r) + 1) + i, o_c + c_c + j) = d_N(i, j)
                
                If k = 1 Then
                    Cells(o_r + (i - 1), o_c).Interior.Color = RGB(255, 150, 0)
                    Cells(o_r, o_c + (j - 1)).Interior.Color = RGB(255, 150, 0)
                Else
                    Cells((o_r + c_r + 1) + (k - 2) * (c_r + 1) + k, o_c + (j - 1)).Interior.Color = RGB(255, 150, 0)
                    Cells((o_r + c_r + 1) + (k - 2) * (c_r + 1) + i, o_c + (k - 1)).Interior.Color = RGB(255, 150, 0)
                End If
                
            Next j
        Next i
    Next k

End Sub

Private Sub CommandButton2_Click()

    Unload Me
    
End Sub
