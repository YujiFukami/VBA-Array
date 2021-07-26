Attribute VB_Name = "ModArray"
Option Explicit

'配列の処理関係のプロシージャ

Function SortArrayByNetFramework(InputArray, Optional InputOrder As OrderType = xlAscending)
'一次元配列をNet.Frameworkを使って昇順にする
'20210726

    Dim DataList, I&, Msg$
    Set DataList = CreateObject("System.Collections.ArrayList")
    
    For I = LBound(InputArray, 1) To UBound(InputArray, 1)
        DataList.Add InputArray(I)
    Next I
    
    Dim Output
    ReDim Output(LBound(InputArray, 1) To UBound(InputArray, 1))
    
    If InputOrder = xlAscending Then
        DataList.Sort
    Else
        DataList.Reverse
    End If
    
    For I = 0 To DataList.Count - 1
        Output(I + LBound(InputArray, 1)) = DataList(I)
    Next I
    Set DataList = Nothing
    
    SortArrayByNetFramework = Output
    
End Function
Sub ADAGADA()
    Dim TmpList
    TmpList = Range("B20").CurrentRegion.Value
    Dim SortList
    
    SortList = SortArray2D(TmpList, 2)
    Call DPH(SortList)
    
End Sub
Function SortArray2D(InputArray2D, Optional SortCol%, Optional InputOrder As OrderType = xlAscending)
'指定の2次元配列を、指定列を基準に並び替える
        
    '指定列を1次元配列で抽出
    Dim KijunArray1D
    Dim MinRow&, MaxRow&
    MinRow = LBound(InputArray2D, 1)
    MaxRow = UBound(InputArray2D, 1)
    ReDim KijunArray1D(MinRow To MaxRow)
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    For I = MinRow To MaxRow
        KijunArray1D(I) = InputArray2D(I, SortCol)
    Next I
    
    '並び替え
    Dim Output
    Output = SortArray2Dby1D(InputArray2D, KijunArray1D, InputOrder)
    SortArray2D = Output
    
End Function
Function SortArray2Dby1D(InputArray2D, ByVal KijunArray1D, Optional InputOrder As OrderType = xlAscending)
'指定の2次元配列を、指定1次元配列を基準に並び替える

    '入力値のチェック
    Dim Dummy%
    On Error Resume Next
    Dummy = UBound(KijunArray1D, 2)
    On Error GoTo 0
    If Dummy <> 0 Then
        MsgBox ("基準配列は1次元配列を入力してください")
        Stop
        End
    End If
    
    Dummy = 0
    On Error Resume Next
    Dummy = UBound(InputArray2D, 2)
    On Error GoTo 0
    If Dummy = 0 Then
        MsgBox ("並び替え対象の配列は2次元配列を入力してください")
        Stop
        End
    End If
    
    Dim MinRow&, MaxRow&, MinCol&, MaxCol&
    MinRow = LBound(InputArray2D, 1)
    MaxRow = UBound(InputArray2D, 1)
    MinCol = LBound(InputArray2D, 2)
    MaxCol = UBound(InputArray2D, 2)
    If MinRow <> LBound(KijunArray1D, 1) Or MaxRow <> UBound(KijunArray1D, 1) Then
        MsgBox ("並び替え対象の配列と、基準配列の開始、終了要素番号を一致させてください")
        Stop
        End
    End If
    
    '基準配列に文字列が含まれている場合はISOコードに変換
    Dim StrAruNaraTrue As Boolean
    StrAruNaraTrue = False
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    Dim Tmp, TmpStr$
    For I = MinRow To MaxRow
        Tmp = KijunArray1D(I)
        If VarType(Tmp) = vbString Then
            StrAruNaraTrue = True
            Exit For
        End If
    Next I
    
    If StrAruNaraTrue Then
        For I = MinRow To MaxRow
            TmpStr = KijunArray1D(I)
            KijunArray1D(I) = ConvStrToISO(TmpStr)
        Next I
    End If
    
    '基準配列を正規化して、(1～要素数)の間の数値にする
    Dim Count&, MinNum#, MaxNum#
    Count = MaxRow - MinRow + 1
    MinNum = WorksheetFunction.Min(KijunArray1D)
    MaxNum = WorksheetFunction.Max(KijunArray1D)
    For I = MinRow To MaxRow
        KijunArray1D(I) = (KijunArray1D(I) - MinNum) / (MaxNum - MinNum) '(0～1)の間で正規化
        KijunArray1D(I) = (Count - 1) * KijunArray1D(I) + 1 '(1～要素数)の間
    Next I
    
    '並び替え(1,2,3の配列を作ってクイックソートで並び替えて、対象の配列を並び替え後の1,2,3で入れ替える)
    Dim TmpArray, Array123, TmpNum&
    ReDim Array123(MinRow To MaxRow)
    For I = MinRow To MaxRow
        If InputOrder = xlAscending Then
            '昇順
            Array123(I) = I - MinRow + 1
        Else
            '降順
            Array123(MaxRow - I + 1) = I - MinRow + 1
        End If
    Next I
    
    Call DPH(Array123)
    Call DPH(KijunArray1D)
    Call SortArrayQuick(KijunArray1D, Array123)
        
    ReDim TmpArray(MinRow To MaxRow, MinCol To MaxCol)
    For I = MinRow To MaxRow
        TmpNum = Array123(I)
        For J = MinCol To MaxCol
            TmpArray(I, J) = InputArray2D(TmpNum, J)
        Next J
    Next I
    
    '出力
    SortArray2Dby1D = TmpArray

End Function
Sub ADAGADAE()

    Dim TestArray, Array123
    TestArray = Array(1.2, 6, 3, 1, 3, 7, 8, 3)
    Array123 = Array(1, 2, 3, 4, 5, 6, 7, 8)
    Call SortArrayQuick(TestArray, Array123)
    
    Call DPH(TestArray)
    Call DPH(Array123)
    
End Sub
Sub SortArrayQuick(KijunArray, Array123, Optional StartNum%, Optional EndNum%)
'クイックソートで配列を並び替える

    If StartNum = 0 Then
        StartNum = LBound(KijunArray, 1)
    End If
    
    If EndNum = 0 Then
        EndNum = UBound(KijunArray, 1)
    End If
    
    Dim Tmp#, Counter#, I&, J&
    Counter = KijunArray((StartNum + EndNum) \ 2)
    I = StartNum - 1
    J = EndNum + 1
    
    '並び替え対象の配列の処理
    Dim Col&, MinCol&, MaxCol&
    Dim Tmp2
    
    Do
        Do
            I = I + 1
        Loop While KijunArray(I) < Counter
        
        Do
            J = J - 1
        Loop While KijunArray(J) > Counter
        
        If I >= J Then Exit Do
        Tmp = KijunArray(J)
        KijunArray(J) = KijunArray(I)
        KijunArray(I) = Tmp
        
        Tmp2 = Array123(J)
        Array123(J) = Array123(I)
        Array123(I) = Tmp2
    
    Loop
    If I - StartNum > 1 Then
        Call SortArrayQuick(KijunArray, Array123, StartNum, I - 1) '再帰呼び出し
    End If
    If EndNum - J > 1 Then
        Call SortArrayQuick(KijunArray, Array123, J + 1, EndNum) '再帰呼び出し
    End If
End Sub
Function ConvStrToISO(InputStr$)
'文字列を並び替え用にISOコードに変換
'20210726

    Dim Mojiretu As String
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    Dim UniCode
    
    Dim UniMax&
    UniMax = 65536
    
    Dim StartKeta%, Kurai#
    StartKeta = 20 '←←←←←←←←←←←←←←←←←←←←←←←
    Kurai = Exp(1) '←←←←←←←←←←←←←←←←←←←←←←←
    
    Dim Output#
    
    If InputStr = "" Then
        Output = 0
    Else
        N = Len(InputStr)
        ReDim UniCode(1 To N)
        
        Output = 0
        For I = 1 To N
            UniCode(I) = Abs(AscW(Mid(InputStr, I, 1)))
            Output = Output + ((Kurai ^ StartKeta) / (UniMax) ^ (I - 1)) * UniCode(I)
            
        Next I
    End If
    
    ConvStrToISO = Output

End Function
