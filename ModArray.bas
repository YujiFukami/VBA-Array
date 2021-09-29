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

Sub TestSortArray2D()
    Dim TmpList
    TmpList = Range("B20").CurrentRegion.Value
    Dim SortList
    
    SortList = SortArray2D(TmpList, 2)
    Call DPH(SortList)
    
End Sub

Function SortArray2D(InputArray2D, Optional SortCol&, Optional InputOrder As OrderType = xlAscending)
'指定の2次元配列を、指定列を基準に並び替える
'配列は文字列を含んでいてもよい
'20210726

'InputArray2D・・・並び替え対象の2次元配列
'SortCol     ・・・並び替えの基準で指定する列番号
'InputOrder  ・・・xlAscending→昇順, xlDescending→降順

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

Private Function SortArray2Dby1D(InputArray2D, ByVal KijunArray1D, Optional InputOrder As OrderType = xlAscending)
'指定の2次元配列を、指定1次元配列を基準に並び替える
'配列は文字列を含んでいてもよい
'20210726
'20210917修正
                                    
'InputArray2D・・・並び替え対象の2次元配列
'KijunArray1D・・・並び替えの基準となる配列
'InputOrder  ・・・xlAscending→昇順, xlDescending→降順

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
    
    '基準配列を正規化して、(1〜要素数)の間の数値にする
    Dim Count&, MinNum#, MaxNum#
    Count = MaxRow - MinRow + 1
    MinNum = WorksheetFunction.Min(KijunArray1D)
    MaxNum = WorksheetFunction.Max(KijunArray1D)
    For I = MinRow To MaxRow
        KijunArray1D(I) = (KijunArray1D(I) - MinNum) / (MaxNum - MinNum) '(0〜1)の間で正規化
        KijunArray1D(I) = (Count - 1) * KijunArray1D(I) + 1 '(1〜要素数)の間
    Next I
    
    '並び替え(1,2,3の配列を作ってクイックソートで並び替えて、対象の配列を並び替え後の1,2,3で入れ替える)
    Dim TmpArray, Array123, TmpNum&
    ReDim Array123(MinRow To MaxRow)
    For I = MinRow To MaxRow
        Array123(I) = I - MinRow + 1
    Next I
        
    Call SortArrayQuick(KijunArray1D, Array123)
    
    ReDim TmpArray(MinRow To MaxRow, MinCol To MaxCol)
    For I = MinRow To MaxRow
    
        TmpNum = Array123(I)
        
        For J = MinCol To MaxCol
            If InputOrder = xlAscending Then '20210917修正
                TmpArray(I, J) = InputArray2D(TmpNum, J)
            Else
                TmpArray(MaxRow - I + 1, J) = InputArray2D(TmpNum, J)
            End If
                
        Next J
    Next I
    
    '出力
    SortArray2Dby1D = TmpArray

End Function

Sub SortArrayQuick(KijunArray, Array123, Optional StartNum%, Optional EndNum%)
'クイックソートで1次元配列を並び替える
'並び替え後の順番を出力するために配列「Array123」を同時に並び替える
'20210726

'KijunArray・・・並び替え対象の配列（1次元配列）
'Array123  ・・・「1,2,3」の値が入った1次元配列
'StartNum  ・・・再帰用の引数
'EndNum    ・・・再帰用の引数

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

Sub CheckArray1D(InputArray, Optional HairetuName$ = "配列")
'入力配列が1次元配列かどうかチェックする
'20210804

    Dim Dummy%
    On Error Resume Next
    Dummy = UBound(InputArray, 2)
    On Error GoTo 0
    If Dummy <> 0 Then
        MsgBox (HairetuName & "は1次元配列を入力してください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Sub CheckArray2D(InputArray, Optional HairetuName$ = "配列")
'入力配列が2次元配列かどうかチェックする
'20210804

    Dim Dummy2%, Dummy3%
    On Error Resume Next
    Dummy2 = UBound(InputArray, 2)
    Dummy3 = UBound(InputArray, 3)
    On Error GoTo 0
    If Dummy2 = 0 Or Dummy3 <> 0 Then
        MsgBox (HairetuName & "は2次元配列を入力してください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Sub CheckArray1DStart1(InputArray, Optional HairetuName$ = "配列")
'入力1次元配列の開始番号が1かどうかチェックする
'20210804

    If LBound(InputArray, 1) <> 1 Then
        MsgBox (HairetuName & "の開始要素番号は1にしてください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Sub CheckArray2DStart1(InputArray, Optional HairetuName$ = "配列")
'入力2次元配列の開始番号が1かどうかチェックする
'20210804

    If LBound(InputArray, 1) <> 1 Or LBound(InputArray, 2) <> 1 Then
        MsgBox (HairetuName & "の開始要素番号は1にしてください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub


Sub ClipCopyArray2D(Array2D)
'2次元配列を変数宣言用のテキストデータに変換して、クリップボードにコピーする
'20210805
    
    '引数チェック
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(Array2D, 1)
    M = UBound(Array2D, 2)
    
    Dim TmpValue
    Dim Output$
    
    Output = ""
    For I = 1 To N
        If I = 1 Then
            Output = Output & String(3, Chr(9)) & "Array(Array("
        Else
            Output = Output & String(3, Chr(9)) & "Array("
        End If
        
        For J = 1 To M
            TmpValue = Array2D(I, J)
            If IsNumeric(TmpValue) Then
                Output = Output & TmpValue
            Else
                Output = Output & """" & TmpValue & """"
            End If
            
            If J < M Then
                Output = Output & ","
            Else
                Output = Output & ")"
            End If
        Next J
        
        If I < N Then
            Output = Output & ", _" & vbLf
        Else
            Output = Output & ")"
        End If
    Next I
    
    Output = "Application.Transpose(Application.Transpose( _" & vbLf & Output & " _" & vbLf & String(3, Chr(9)) & "))"
    
    Call ClipboardCopy(Output, True)
    
End Sub

Sub ClipCopyArray1D(Array1D)
'1次元配列を変数宣言用のテキストデータに変換して、クリップボードにコピーする
'20210805
    
    '引数チェック
    Call CheckArray1D(Array1D, "Array1D")
    Call CheckArray1DStart1(Array1D, "Array1D")
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(Array1D, 1)
    
    Dim TmpValue
    Dim Output$
    
    Output = String(3, Chr(9)) & "Array("
    For I = 1 To N
        
        TmpValue = Array1D(I)
        If IsNumeric(TmpValue) Then
            Output = Output & TmpValue
        Else
            Output = Output & """" & TmpValue & """"
        End If
        
        If I < N Then
            Output = Output & ","
        Else
            Output = Output & ")"
        End If
        
    Next I
    
    Output = "Application.Transpose(Application.Transpose( _" & vbLf & Output & " _" & vbLf & String(3, Chr(9)) & "))"
    
    Call ClipboardCopy(Output, True)
    
End Sub

Sub MessageArray2D(Array2D)
'二次元配列をメッセージに表示する。
'20210824

    '引数チェック
    Call CheckArray2D(Array2D)
    Call CheckArray2DStart1(Array2D)
    
    '処理
    Dim Tmp$
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    For I = 1 To UBound(Array2D, 1)
        For J = 1 To UBound(Array2D, 2)
            If J = 1 Then
                Tmp = Tmp & Array2D(I, J)
            Else
                Tmp = Tmp & Chr(9) & Array2D(I, J)
            End If
        Next J
        
        If I <> UBound(Array2D, 1) Then
            Tmp = Tmp & vbLf
        End If
    Next I
        
    MsgBox (Tmp)
    
End Sub

Function ConvArray1Dto1N(InputArray1D)
'1次元配列を、(1,N)配列に変換する
'20210917

'InputArray1D・・・1次元配列

    '引数チェック
    Call CheckArray1D(InputArray1D, "InputArray1D")
    Call CheckArray1DStart1(InputArray1D, "InputArray1D")
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(InputArray1D, 1)
    Dim Output
    ReDim Output(1 To 1, 1 To N)
    For I = 1 To N
        Output(1, I) = InputArray1D(I)
    Next I
    
    '出力
    ConvArray1Dto1N = Output

End Function

Function DeleteRowArray(Array2D, DeleteRow&)
'二次元配列の指定行を消去した配列を出力する
'20210917

'引数
'Array2D  ・・・二次元配列
'DeleteRow・・・消去する行番号

    '引数チェック
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(Array2D, 1) '行数
    M = UBound(Array2D, 2) '列数
    
    If DeleteRow < 1 Then
        MsgBox ("削除する行番号は1以上の値を入れてください")
        Stop
        End
    ElseIf DeleteRow > N Then
        MsgBox ("削除する行番号は元の二次元配列の行数" & N & "以下の値を入れてください")
        Stop
        End
    End If
    
    '処理
    Dim Output
    ReDim Output(1 To N - 1, 1 To M)
    K = 0
    For I = 1 To N
        If I <> DeleteRow Then
            K = K + 1
            For J = 1 To M
                Output(K, J) = Array2D(I, J)
            Next J
        End If
    Next I
    
    '出力
    DeleteRowArray = Output

End Function

Function DeleteColArray(Array2D, DeleteCol&)
'二次元配列の指定列を消去した配列を出力する
'20210917

'引数
'Array2D  ・・・二次元配列
'DeleteCol・・・消去する列番号

    '引数チェック
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(Array2D, 1) '行数
    M = UBound(Array2D, 2) '列数

    If DeleteCol < 1 Then
        MsgBox ("削除する列番号は1以上の値を入れてください")
        Stop
        End
    ElseIf DeleteCol > M Then
        MsgBox ("削除する列番号は元の二次元配列の列数" & M & "以下の値を入れてください")
        Stop
        End
    End If
    
    '処理
    Dim Output
    ReDim Output(1 To N, 1 To M - 1)
    For I = 1 To N
        K = 0
        For J = 1 To M
            If J <> DeleteCol Then
                K = K + 1
                Output(I, K) = Array2D(I, J)
            End If
        Next J
    Next I
    
    '出力
    DeleteColArray = Output

End Function

Function ExtractRowArray(Array2D, TargetRow&)
'二次元配列の指定行を一次元配列で抽出する
'20210917

'引数
'Array2D  ・・・二次元配列
'TargetRow・・・抽出する対象の行番号


    '引数チェック
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(Array2D, 1) '行数
    M = UBound(Array2D, 2) '列数

    If TargetRow < 1 Then
        MsgBox ("抽出する行番号は1以上の値を入れてください")
        Stop
        End
    ElseIf TargetRow > N Then
        MsgBox ("抽出する行番号は元の二次元配列の行数" & N & "以下の値を入れてください")
        Stop
        End
    End If

    '処理
    Dim Output
    ReDim Output(1 To M)
    
    For I = 1 To M
        Output(I) = Array2D(TargetRow, I)
    Next I
    
    '出力
    ExtractRowArray = Output
    
End Function

Function ExtractColArray(Array2D, TargetCol&)
'二次元配列の指定列を一次元配列で抽出する
'20210917

'引数
'Array2D  ・・・二次元配列
'TargetCol・・・抽出する対象の列番号


    '引数チェック
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(Array2D, 1) '行数
    M = UBound(Array2D, 2) '列数
 
    If TargetCol < 1 Then
        MsgBox ("抽出する列番号は1以上の値を入れてください")
        Stop
        End
    ElseIf TargetCol > N Then
        MsgBox ("抽出する列番号は元の二次元配列の行数" & M & "以下の値を入れてください")
        Stop
        End
    End If
    
    '処理
    Dim Output
    ReDim Output(1 To N)
    
    For I = 1 To N
        Output(I) = Array2D(I, TargetCol)
    Next I
    
    '出力
    ExtractColArray = Output
    
End Function

Function ExtractArray(Array2D, StartRow&, StartCol&, EndRow&, EndCol&)
'二次元配列の指定範囲を配列として抽出する
'20210917

'引数
'Array2D ・・・二次元配列
'StartRow・・・抽出範囲の開始行番号
'StartCol・・・抽出範囲の開始列番号
'EndRow  ・・・抽出範囲の終了行番号
'EndCol  ・・・抽出範囲の終了列番号
                                   
    '引数チェック
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(Array2D, 1) '行数
    M = UBound(Array2D, 2) '列数
    
    If StartRow > EndRow Then
        MsgBox ("抽出範囲の開始行「StartRow」は、終了行「EndRow」以下でなければなりません")
        Stop
        End
    ElseIf StartCol > EndCol Then
        MsgBox ("抽出範囲の開始列「StartCol」は、終了列「EndCol」以下でなければなりません")
        Stop
        End
    ElseIf StartRow < 1 Then
        MsgBox ("抽出範囲の開始行「StartRow」は1以上の値を入れてください")
        Stop
        End
    ElseIf StartCol < 1 Then
        MsgBox ("抽出範囲の開始列「StartCol」は1以上の値を入れてください")
        Stop
        End
    ElseIf EndRow > N Then
        MsgBox ("抽出範囲の終了行「StartRow」は抽出元の二次元配列の行数" & N & "以下の値を入れてください")
        Stop
        End
    ElseIf EndCol > M Then
        MsgBox ("抽出範囲の終了列「StartCol」は抽出元の二次元配列の列数" & M & "以下の値を入れてください")
        Stop
        End
    End If
    
    '処理
    Dim Output
    ReDim Output(1 To EndRow - StartRow + 1, 1 To EndCol - StartCol + 1)
    
    For I = StartRow To EndRow
        For J = StartCol To EndCol
            Output(I - StartRow + 1, J - StartCol + 1) = Array2D(I, J)
        Next J
    Next I
    
    '出力
    ExtractArray = Output
    
End Function

Function UnionArray1D(UpperArray1D, LowerArray1D)
'一次元配列同士を結合して1つの配列とする。
'20210923

'UpperArray1D・・・上に結合する一次元配列
'LowerArray1D・・・下に結合する一次元配列

    '引数チェック
    Call CheckArray1D(UpperArray1D, "UpperArray1D")
    Call CheckArray1DStart1(UpperArray1D, "UpperArray1D")
    Call CheckArray1D(LowerArray1D, "LowerArray1D")
    Call CheckArray1DStart1(LowerArray1D, "LowerArray1D")
    
    '処理
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    Dim N1&, N2&
    N1 = UBound(UpperArray1D, 1)
    N2 = UBound(UpperArray1D, 2)
    Dim Output
    ReDim Output(1 To N1 + N2)
    For I = 1 To N1
        Output(I) = UpperArray1D(I)
    Next I
    For I = 1 To N2
        Output(N1 + I) = UpperArray2D(I)
    Next I
    
    '出力
    UnionArray1D = Output
    
End Function

Function UnionArrayUL(UpperArray2D, LowerArray2D)
'二次元配列同士を上下に結合して1つの配列とする。
'20210923

'UpperArray2D・・・上に結合する二次元配列
'LowerArray2D・・・下に結合する二次元配列

    '引数チェック
    Call CheckArray2D(UpperArray2D, "UpperArray2D")
    Call CheckArray2DStart1(UpperArray2D, "UpperArray2D")
    Call CheckArray2D(LowerArray2D, "LowerArray2D")
    Call CheckArray2DStart1(LowerArray2D, "LowerArray2D")
    
    If UBound(UpperArray2D, 2) <> UBound(LowerArray2D, 2) Then
        MsgBox ("UpperArray2DとLowerArray2Dの二次元要素数を合わせてください" & vbLf & _
                "UpperArray2Dの二次元要素数 = " & UBound(UpperArray2D, 2) & vbLf & _
                "LowerArray2Dの二次元要素数 = " & UBound(LowerArray2D, 2))
    End If
    
    '処理
    Dim Output
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    Dim N1&, N2&
    N1 = UBound(UpperArray2D, 1)
    N2 = UBound(LowerArray2D, 1)
    M = UBound(UpperArray2D, 2)
    
    ReDim Output(1 To N1 + N2, 1 To M)
    For I = 1 To N1
        For J = 1 To M
            Output(I, J) = UpperArray2D(I, J)
        Next J
    Next I
    
    For I = 1 To N2
        For J = 1 To M
            Output(N1 + I, J) = LowerArray2D(I, J)
        Next J
    Next I
    
    '出力
    UnionArrayUL = Output

End Function

Function UnionArrayLR(LeftArray2D, RightArray2D)
'二次元配列同士を左右に結合して1つの配列とする。
'20210923

'LeftArray2D ・・・上に結合する二次元配列
'RightArray2D・・・下に結合する二次元配列

    '引数チェック
    Call CheckArray2D(LeftArray2D, "LeftArray2D")
    Call CheckArray2DStart1(LeftArray2D, "LeftArray2D")
    Call CheckArray2D(RightArray2D, "RightArray2D")
    Call CheckArray2DStart1(RightArray2D, "RightArray2D")
    
    If UBound(LeftArray2D, 1) <> UBound(RightArray2D, 1) Then
        MsgBox ("LeftArray2DとRightArray2Dの一次元要素数を合わせてください" & vbLf & _
                "LeftArray2Dの一次元要素数 = " & UBound(LeftArray2D, 1) & vbLf & _
                "RightArray2Dの一次元要素数 = " & UBound(RightArray2D, 1))
    End If
    
    '処理
    Dim Output
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    Dim M1&, M2&
    M1 = UBound(LeftArray2D, 2)
    M2 = UBound(RightArray2D, 2)
    N = UBound(LeftArray2D, 1)
    
    ReDim Output(1 To N, 1 To M1 + M2)
    For I = 1 To N
        For J = 1 To M1
            Output(I, J) = LeftArray2D(I, J)
        Next J
    Next I
    
    For I = 1 To N
        For J = 1 To M2
            Output(I, M1 + J) = RightArray2D(I, J)
        Next J
    Next I
    
    '出力
    UnionArrayLR = Output

End Function


Function DimArray1DSameValue(Count&, Value)
'全て同じ値が入った一次元配列を定義する
'20210923

'Count・・・要素数(Long型)
'Value・・・同じ値を入れる値(オブジェクト型でも可能)
    
    '引数チェック
    If Count <= 0 Then
        MsgBox ("要素数Countは1以上の値を入れてください。" & vbLf & _
               "Count = " & Count)
        Stop
    End If
    
    '処理
    Dim Output
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    ReDim Output(1 To Count)
    For I = 1 To Count
        If IsObject(Value) Then
            Set Output(I) = Value
        Else
            Output(I) = Value
        End If
    Next I
    
    '出力
    DimArray1DSameValue = Output
    
End Function

Function DimArray2DSameValue(Count1&, Count2&, Value)
'全て同じ値が入った二次元配列を定義する
'20210923

'Count1・・・一次元要素数(Long型)
'Count2・・・二次元要素数(Long型)
'Value ・・・同じ値を入れる値(オブジェクト型でも可能)
    
    '引数チェック
    If Count1 <= 0 Then
        MsgBox ("一次元要素数Count1は1以上の値を入れてください。" & vbLf & _
               "Count1 = " & Count1)
        Stop
    End If
    
    If Count2 <= 0 Then
        MsgBox ("二次元要素数Count2は1以上の値を入れてください。" & vbLf & _
               "Count2 = " & Count2)
        Stop
    End If
    
    '処理
    Dim Output
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    ReDim Output(1 To Count1, 1 To Count2)
    For I = 1 To Count1
        For J = 1 To Count2
            If IsObject(Value) Then
                Set Output(I, J) = Value
            Else
                Output(I, J) = Value
            End If
        Next J
    Next I
    
    '出力
    DimArray2DSameValue = Output
    
End Function

Function FilterArray2D(Array2D, FilterStr$, TargetCol&)
'二次元配列を指定セルでフィルターした配列を出力する。
'20210929

'引数
'Array2D  ・・・二次元配列
'FilterStr・・・フィルターする文字（String型）
'TargetCol・・・フィルターする列（Long型）
    
    '引数チェック
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    'フィルター件数計算
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    Dim FilterCount&, TargetStr$
    N = UBound(Array2D, 1)
    M = UBound(Array2D, 2)
    K = 0
    For I = 1 To N
        TargetStr = Array2D(I, TargetCol)
        If TargetStr = FilterStr Then
            K = K + 1
        End If
    Next I
    
    FilterCount = K
    
    If K = 0 Then
        'フィルターで何もかからなかった場合はEmptyを返す
        FilterArray2D = Empty
        Exit Function
    End If
    
    '出力する配列の作成
    Dim Output
    ReDim Output(1 To FilterCount, 1 To M)
    
    K = 0
    For I = 1 To N
        TargetStr = Array2D(I, TargetCol)
        If TargetStr = FilterStr Then
            K = K + 1
            For J = 1 To M
                Output(K, J) = Array2D(I, J)
            Next J
        End If
    Next I
    
    '出力
    FilterArray2D = Output
    
End Function


