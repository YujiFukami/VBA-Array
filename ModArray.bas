Attribute VB_Name = "ModArray"
Option Explicit

'SortArrayByNetFramework         ・・・元場所：FukamiAddins3.ModArray    
'TestSortArray2D                 ・・・元場所：FukamiAddins3.ModArray    
'SortArray2D                     ・・・元場所：FukamiAddins3.ModArray    
'SortArray2Dby1D                 ・・・元場所：FukamiAddins3.ModArray    
'SortArrayQuick                  ・・・元場所：FukamiAddins3.ModArray    
'ConvStrToISO                    ・・・元場所：FukamiAddins3.ModArray    
'CheckArray1D                    ・・・元場所：FukamiAddins3.ModArray    
'CheckArray2D                    ・・・元場所：FukamiAddins3.ModArray    
'CheckArray1DStart1              ・・・元場所：FukamiAddins3.ModArray    
'CheckArray2DStart1              ・・・元場所：FukamiAddins3.ModArray    
'ClipCopyArray2D                 ・・・元場所：FukamiAddins3.ModArray    
'ClipCopyArray1D                 ・・・元場所：FukamiAddins3.ModArray    
'MessageArray2D                  ・・・元場所：FukamiAddins3.ModArray    
'ConvArray1Dto1N                 ・・・元場所：FukamiAddins3.ModArray    
'DeleteRowArray                  ・・・元場所：FukamiAddins3.ModArray    
'DeleteColArray                  ・・・元場所：FukamiAddins3.ModArray    
'ExtractRowArray                 ・・・元場所：FukamiAddins3.ModArray    
'ExtractColArray                 ・・・元場所：FukamiAddins3.ModArray    
'ExtractArray                    ・・・元場所：FukamiAddins3.ModArray    
'ExtractArray1D                  ・・・元場所：FukamiAddins3.ModArray    
'UnionArray1D                    ・・・元場所：FukamiAddins3.ModArray    
'UnionArrayLR1D                  ・・・元場所：FukamiAddins3.ModArray    
'UnionArrayUL                    ・・・元場所：FukamiAddins3.ModArray    
'UnionArrayLR                    ・・・元場所：FukamiAddins3.ModArray    
'DimArray1DSameValue             ・・・元場所：FukamiAddins3.ModArray    
'DimArray2DSameValue             ・・・元場所：FukamiAddins3.ModArray    
'FilterArray2D                   ・・・元場所：FukamiAddins3.ModArray    
'DimArray1DNumbers               ・・・元場所：FukamiAddins3.ModArray    
'DPH                             ・・・元場所：FukamiAddins3.ModImmediate
'DebugPrintHairetu               ・・・元場所：FukamiAddins3.ModImmediate
'文字列を指定バイト数文字数に省略・・・元場所：FukamiAddins3.ModImmediate
'文字列の各文字累計バイト数計算  ・・・元場所：FukamiAddins3.ModImmediate
'文字列分解                      ・・・元場所：FukamiAddins3.ModImmediate
'ClipboardCopy                   ・・・元場所：FukamiAddins3.ModClipboard

'宣言セクション※※※※※※※※※※※※※※※※※※※※※※※※※※※
'-----------------------------------
'元場所:FukamiAddins3.ModEnum.OrderType
Public Enum OrderType '昇順降順の列挙型
    xlAscending = 1
    xlDescending = 2
End Enum
'宣言セクション終了※※※※※※※※※※※※※※※※※※※※※※※※※※※

Function SortArrayByNetFramework(InputArray, Optional InputOrder As OrderType = xlAscending)
'一次元配列をNet.Frameworkを使って昇順にする
'20210726

    Dim DataList
    Dim I   As Long
    Dim Msg As String
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
    Dim SortList
    TmpList = Range("B20").CurrentRegion.Value
    SortList = SortArray2D(TmpList, 2)
    Call DPH(SortList)
    
End Sub

Function SortArray2D(InputArray2D, Optional SortCol As Long, Optional InputOrder As OrderType = xlAscending)
'指定の2次元配列を、指定列を基準に並び替える
'配列は文字列を含んでいてもよい
'20210726

'InputArray2D・・・並び替え対象の2次元配列
'SortCol     ・・・並び替えの基準で指定する列番号
'InputOrder  ・・・xlAscending→昇順, xlDescending→降順

    '指定列を1次元配列で抽出
    Dim KijunArray1D
    Dim MinRow      As Long
    Dim MaxRow      As Long
    Dim I           As Long
    MinRow = LBound(InputArray2D, 1)
    MaxRow = UBound(InputArray2D, 1)
    ReDim KijunArray1D(MinRow To MaxRow)
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
'配列は文字列を含んでいてもよい
'20210726
'20210917修正
'20211016修正 配列のの中身がオブジェクト変数でも対応
                                   
'InputArray2D・・・並び替え対象の2次元配列
'KijunArray1D・・・並び替えの基準となる配列
'InputOrder  ・・・xlAscending→昇順, xlDescending→降順

    '入力値のチェック
    Dim Dummy As Integer
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
    
    Dim MinRow As Long
    Dim MaxRow As Long
    Dim MinCol As Long
    Dim MaxCol As Long
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
    Dim StrAruNaraTrue As Boolean: StrAruNaraTrue = False
    Dim I      As Long
    Dim J      As Long
    Dim Tmp
    Dim TmpStr As String
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
    Dim Count  As Long
    Dim MinNum As Double
    Dim MaxNum As Double
    Count = MaxRow - MinRow + 1
    MinNum = WorksheetFunction.Min(KijunArray1D)
    MaxNum = WorksheetFunction.Max(KijunArray1D)
    
    Dim TmpArray
    If MinNum = MaxNum Then '20211016修正'最大と最小が一致するならそのまま返す
        TmpArray = InputArray2D
        GoTo EndEscape
    End If
    
    For I = MinRow To MaxRow
        KijunArray1D(I) = (KijunArray1D(I) - MinNum) / (MaxNum - MinNum) '(0～1)の間で正規化
        KijunArray1D(I) = (Count - 1) * KijunArray1D(I) + 1 '(1～要素数)の間
    Next I
    
    '並び替え(1,2,3の配列を作ってクイックソートで並び替えて、対象の配列を並び替え後の1,2,3で入れ替える)
    Dim Array123
    Dim TmpNum  As Long
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
                If IsObject(InputArray2D(TmpNum, J)) Then '20211016修正
                    Set TmpArray(I, J) = InputArray2D(TmpNum, J)
                Else
                    TmpArray(I, J) = InputArray2D(TmpNum, J)
                End If
            Else
                If IsObject(InputArray2D(TmpNum, J)) Then '20211016修正
                    Set TmpArray(MaxRow - I + 1, J) = InputArray2D(TmpNum, J)
                Else
                    TmpArray(MaxRow - I + 1, J) = InputArray2D(TmpNum, J)
                End If
            End If
                
        Next J
    Next I

EndEscape:

    '出力
    SortArray2Dby1D = TmpArray

End Function

Sub SortArrayQuick(KijunArray, Array123, Optional StartNum As Integer, Optional EndNum As Integer)
'クイックソートで1次元配列を並び替える
'並び替え後の順番を出力するために配列「Array123」を同時に並び替える
'20210726
'20211016修正 配列の中身がオブジェクト変数でも対応

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
    
    Dim Tmp     As Double
    Dim Counter As Double
    Dim I       As Long
    Dim J       As Long
    Counter = KijunArray((StartNum + EndNum) \ 2)
    I = StartNum - 1
    J = EndNum + 1
    
    '並び替え対象の配列の処理
    Dim Col    As Long
    Dim MinCol As Long
    Dim MaxCol As Long
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
        
        If IsObject(Array123(I)) Then '20211016修正
            Set Tmp2 = Array123(J)
            Set Array123(J) = Array123(I)
            Set Array123(I) = Tmp2
        Else
            Tmp2 = Array123(J)
            Array123(J) = Array123(I)
            Array123(I) = Tmp2
        End If
    Loop
    If I - StartNum > 1 Then
        Call SortArrayQuick(KijunArray, Array123, StartNum, I - 1) '再帰呼び出し
    End If
    If EndNum - J > 1 Then
        Call SortArrayQuick(KijunArray, Array123, J + 1, EndNum) '再帰呼び出し
    End If
End Sub

Function ConvStrToISO(InputStr As String)
'文字列を並び替え用にISOコードに変換
'20210726

    Dim Mojiretu As String
    Dim I        As Long
    Dim J        As Long
    Dim K        As Long
    Dim M        As Long
    Dim N        As Long
    Dim UniCode
    Dim UniMax   As Long
    UniMax = 65536
    
    Dim StartKeta As Integer
    Dim Kurai     As Double
    StartKeta = 20 '←←←←←←←←←←←←←←←←←←←←←←←
    Kurai = Exp(1) '←←←←←←←←←←←←←←←←←←←←←←←
    
    Dim Output As Double
    
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

Sub CheckArray1D(InputArray, Optional HairetuName As String = "配列")
'入力配列が1次元配列かどうかチェックする
'20210804

    Dim Dummy As Integer
    On Error Resume Next
    Dummy = UBound(InputArray, 2)
    On Error GoTo 0
    If Dummy <> 0 Then
        MsgBox (HairetuName & "は1次元配列を入力してください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Sub CheckArray2D(InputArray, Optional HairetuName As String = "配列")
'入力配列が2次元配列かどうかチェックする
'20210804

    Dim Dummy2 As Integer
    Dim Dummy3 As Integer
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

Sub CheckArray1DStart1(InputArray, Optional HairetuName As String = "配列")
'入力1次元配列の開始番号が1かどうかチェックする
'20210804

    If LBound(InputArray, 1) <> 1 Then
        MsgBox (HairetuName & "の開始要素番号は1にしてください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Sub CheckArray2DStart1(InputArray, Optional HairetuName As String = "配列")
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
'20211016 「"」を含む場合も対応

    '引数チェック
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I As Long
    Dim J As Long
    Dim M As Long
    Dim N As Long '数え上げ用(Long型)
    N = UBound(Array2D, 1)
    M = UBound(Array2D, 2)
    
    Dim TmpValue
    Dim Output As String
    
    Output = ""
    For I = 1 To N
        If I = 1 Then
            Output = Output & String(3, Chr(9)) & "Array(Array("
        Else
            Output = Output & String(3, Chr(9)) & "Array("
        End If
        
        For J = 1 To M
            TmpValue = Array2D(I, J)
            
            TmpValue = Replace(TmpValue, """", String(2, """")) '20211016
            
            If TmpValue = "" Then
                Output = Output & """" & """"
            ElseIf IsNumeric(TmpValue) Then
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
    
    Dim I As Long
    Dim N As Long
    N = UBound(Array1D, 1)
    
    Dim TmpValue
    Dim Output As String
    
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
    Dim Tmp As String
    Dim I   As Long
    Dim J   As Long
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
    
    Dim I As Long
    Dim N As Long
    N = UBound(InputArray1D, 1)
    Dim Output
    ReDim Output(1 To 1, 1 To N)
    For I = 1 To N
        Output(1, I) = InputArray1D(I)
    Next I
    
    '出力
    ConvArray1Dto1N = Output

End Function

Function DeleteRowArray(Array2D, DeleteRow As Long)
'二次元配列の指定行を消去した配列を出力する
'20210917

'引数
'Array2D  ・・・二次元配列
'DeleteRow・・・消去する行番号

    '引数チェック
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I As Long
    Dim J As Long
    Dim K As Long
    Dim M As Long
    Dim N As Long
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

Function DeleteColArray(Array2D, DeleteCol As Long)
'二次元配列の指定列を消去した配列を出力する
'20210917

'引数
'Array2D  ・・・二次元配列
'DeleteCol・・・消去する列番号

    '引数チェック
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I As Long
    Dim J As Long
    Dim K As Long
    Dim M As Long
    Dim N As Long
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

Function ExtractRowArray(Array2D, TargetRow As Long)
'二次元配列の指定行を一次元配列で抽出する
'20210917

'引数
'Array2D  ・・・二次元配列
'TargetRow・・・抽出する対象の行番号


    '引数チェック
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I As Long
    Dim M As Long
    Dim N As Long
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

Function ExtractColArray(Array2D, TargetCol As Long)
'二次元配列の指定列を一次元配列で抽出する
'20210917
'20211016修正 配列の中身がオブジェクト変数でも対応

'引数
'Array2D  ・・・二次元配列
'TargetCol・・・抽出する対象の列番号

    '引数チェック
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I As Long
    Dim M As Long
    Dim N As Long
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
        If IsObject(Array2D(I, TargetCol)) Then '20211016修正
            Set Output(I) = Array2D(I, TargetCol)
        Else
            Output(I) = Array2D(I, TargetCol)
        End If
    Next I
    
    '出力
    ExtractColArray = Output
    
End Function

Function ExtractArray(Array2D, StartRow As Long, StartCol As Long, EndRow As Long, EndCol As Long)
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
    
    Dim I As Long
    Dim J As Long
    Dim M As Long
    Dim N As Long
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

Function ExtractArray1D(Array1D, StartNum As Long, EndNum As Long)
'一次元配列の指定範囲を配列として抽出する
'20211009

'引数
'Array1D ・・・一次元配列
'StartNum・・・抽出範囲の開始番号
'EndNum  ・・・抽出範囲の終了番号
                                   
    '引数チェック
    Call CheckArray1D(Array1D, "Array1D")
    Call CheckArray1DStart1(Array1D, "Array1D")
    
    Dim I As Long
    Dim N As Long
    N = UBound(Array1D, 1) '要素数
    
    If StartNum > EndNum Then
        MsgBox ("抽出範囲の開始位置「StartNum」は、終了位置「EndNum」以下でなければなりません")
        Stop
        Exit Function
    ElseIf StartNum < 1 Then
        MsgBox ("抽出範囲の開始位置「StartNum」は1以上の値を入れてください")
        Stop
        Exit Function
    ElseIf EndNum > N Then
        MsgBox ("抽出範囲の終了行「EndNum」は抽出元の一次元配列の要素数" & N & "以下の値を入れてください")
        Stop
        Exit Function
    End If
    
    '処理
    Dim Output
    ReDim Output(1 To EndNum - StartNum + 1)
    
    For I = StartNum To EndNum
        Output(I - StartNum + 1) = Array1D(I)
    Next I
    
    '出力
    ExtractArray1D = Output
    
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
    Dim I  As Long
    Dim N1 As Long
    Dim N2 As Long
    N1 = UBound(UpperArray1D, 1)
    N2 = UBound(LowerArray1D, 1)
    Dim Output
    ReDim Output(1 To N1 + N2)
    For I = 1 To N1
        Output(I) = UpperArray1D(I)
    Next I
    For I = 1 To N2
        Output(N1 + I) = LowerArray1D(I)
    Next I
    
    '出力
    UnionArray1D = Output
    
End Function

Function UnionArrayLR1D(LeftArray1D, RightArray1D)
'一次元配列同士を左右に結合して1つの配列とする。
'20211016

'LeftArray1D ・・・左に結合する一次元配列
'RightArray1D・・・右に結合する一次元配列

    '引数チェック
    Call CheckArray1D(LeftArray1D, "LeftArray1D")
    Call CheckArray1DStart1(LeftArray1D, "LeftArray1D")
    Call CheckArray1D(RightArray1D, "RightArray1D")
    Call CheckArray1DStart1(RightArray1D, "RightArray1D")
    If UBound(LeftArray1D, 1) <> UBound(RightArray1D, 1) Then
        MsgBox ("LeftArray1DとRightArray1Dの要素数は揃えてください")
        Stop
        Exit Function
    End If
    
    '処理
    Dim I As Long
    Dim N As Long
    N = UBound(LeftArray1D, 1)
    Dim Output
    ReDim Output(1 To N, 1 To 2)
    For I = 1 To N
        If IsObject(LeftArray1D(I)) Then
            Set Output(I, 1) = LeftArray1D(I)
        Else
            Output(I, 1) = LeftArray1D(I)
        End If
        
        If IsObject(RightArray1D(I)) Then
            Set Output(I, 2) = RightArray1D(I)
        Else
            Output(I, 2) = RightArray1D(I)
        End If
    Next I
    
    '出力
    UnionArrayLR1D = Output
    
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
    Dim I  As Long
    Dim J  As Long
    Dim M  As Long
    Dim N1 As Long
    Dim N2 As Long
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
    Dim I     As Long
    Dim J     As Long
    Dim K     As Long
    Dim N     As Long
    Dim M1    As Long
    Dim M2    As Long
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

Function DimArray1DSameValue(Count As Long, Value)
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
    Dim I     As Long
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

Function DimArray2DSameValue(Count1 As Long, Count2 As Long, Value)
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
    Dim I     As Long
    Dim J     As Long
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

Function FilterArray2D(Array2D, FilterStr As String, TargetCol As Long)
'二次元配列を指定列でフィルターした配列を出力する。
'20210929

'引数
'Array2D  ・・・二次元配列
'FilterStr・・・フィルターする文字（String型）
'TargetCol・・・フィルターする列（Long型）
    
    '引数チェック
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    'フィルター件数計算
    Dim I           As Long
    Dim J           As Long
    Dim K           As Long
    Dim M           As Long
    Dim N           As Long
    Dim FilterCount As Long
    Dim TargetStr   As String
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

Function DimArray1DNumbers(StartNum As Long, EndNum As Long, Optional ByVal StepNum As Long = 1)
'連番の入った一次元配列を定義する
'20211018

'引数
'StartNum ・・・最初の番号/Long型
'EndNum　 ・・・最後の番号/Long型
'[Step]   ・・・連番の間隔/Long型/デフォルトは1
    
    '引数のチェック
    If StepNum = 0 Then
        MsgBox ("StepNumは0以外の整数を入力してください")
        Stop
        Exit Function
    End If
    
    '引数の修正
    If StartNum < EndNum And StepNum < 0 Then
        StepNum = -StepNum
    ElseIf StartNum > EndNum And StepNum > 0 Then
        StepNum = -StepNum
    End If
    
    '連番の作成
    Dim Output
    Dim I     As Long
    Dim K     As Long
    ReDim Output(1 To 1)
    K = 0
    For I = StartNum To EndNum Step StepNum
        K = K + 1
        ReDim Preserve Output(1 To K)
        Output(K) = I
    Next I
    
    '出力
    DimArray1DNumbers = Output
    
End Function

Private Sub DPH(ByVal Hairetu, Optional HyoujiMaxNagasa As Integer, Optional HairetuName As String)
    '20210428追加
    '入力高速化用に作成
    
    Call DebugPrintHairetu(Hairetu, HyoujiMaxNagasa, HairetuName)
End Sub

Private Sub DebugPrintHairetu(ByVal Hairetu, Optional HyoujiMaxNagasa As Integer, Optional HairetuName As String)
'20201023追加
'20211018 入力した配列がHairetu(1 to 1)の一次元配列の場合でも処理できるように修正

    '二次元配列をイミディエイトウィンドウに見やすく表示する
    
    Dim I       As Long
    Dim J       As Long
    Dim M       As Long
    Dim N       As Long
    Dim TateMin As Long
    Dim TateMax As Long
    Dim YokoMin As Long
    Dim YokoMax As Long

    Dim WithTableHairetu             'テーブル付配列…イミディエイトウィンドウに表示する際にインデックス番号を表示したテーブルを追加した配列
    Dim NagasaList
    Dim MaxNagasaList                '各文字の文字列長さを格納、各列での文字列長さの最大値を格納
    Dim NagasaOnajiList              '" "（半角スペース）を文字列に追加して各列で文字列長さを同じにした文字列を格納
    Dim OutputList                   'イミディエイトウィンドウに表示する文字列を格納
    Const SikiriMoji As String = "|" 'イミディエイトウィンドウに表示する時に各列の間に表示する「仕切り文字」
    
    '※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '入力引数の処理
    Dim Jigen1 As Long
    Dim Jigen2 As Long
    Dim Tmp
    On Error Resume Next
    Jigen2 = UBound(Hairetu, 2)
    On Error GoTo 0
    If Jigen2 = 0 Then '1次元配列は2次元配列にする
        Jigen1 = UBound(Hairetu, 1) '20211018 入力した配列がHairetu(1 to 1)の一次元配列の場合でも処理できるように修正
        If Jigen1 = 1 Then
            Tmp = Hairetu(Jigen1)
            ReDim Hairetu(1 To 1, 1 To 1)
            Hairetu(1, 1) = Tmp
        Else
            Hairetu = Application.Transpose(Hairetu)
        End If
    End If
    
    TateMin = LBound(Hairetu, 1) '配列の縦番号（インデックス）の最小
    TateMax = UBound(Hairetu, 1) '配列の縦番号（インデックス）の最大
    YokoMin = LBound(Hairetu, 2) '配列の横番号（インデックス）の最小
    YokoMax = UBound(Hairetu, 2) '配列の横番号（インデックス）の最大
    
    'テーブル付き配列の作成
    ReDim WithTableHairetu(1 To TateMax - TateMin + 1 + 1, 1 To YokoMax - YokoMin + 1 + 1) 'テーブル追加の分で"+1"する。
    '「TateMax -TateMin + 1」は入力した「Hairetu」の縦インデックス数
    '「YokoMax -YokoMin + 1」は入力した「Hairetu」の横インデックス数
    
    For I = 1 To TateMax - TateMin + 1
        WithTableHairetu(I + 1, 1) = TateMin + I - 1 '縦テーブル（Hairetuの縦インデックス番号）
        For J = 1 To YokoMax - YokoMin + 1
            WithTableHairetu(1, J + 1) = YokoMin + J - 1 '横テーブル（Hairetuの横インデックス番号）
            WithTableHairetu(I + 1, J + 1) = Hairetu(I - 1 + TateMin, J - 1 + YokoMin) 'Hairetuの中の値
        Next J
    Next I
    
    '※※※※※※※※※※※※※※※※※※※※※※※※※※※
    'イミディエイトウィンドウに表示するときに各列の幅を同じに整えるために
    '文字列長さとその各列の最大値を計算する。
    '以下では「Hairetu」は扱わず、「WithTableHairetu」を扱う。
    N = UBound(WithTableHairetu, 1) '「WithTableHairetu」の縦インデックス数（行数）
    M = UBound(WithTableHairetu, 2) '「WithTableHairetu」の横インデックス数（列数）
    ReDim NagasaList(1 To N, 1 To M)
    ReDim MaxNagasaList(1 To M)
    
    Dim TmpStr As String
    For J = 1 To M
        For I = 1 To N
        
            If J > 1 And HyoujiMaxNagasa <> 0 Then
                '最大表示長さが指定されている場合。
                '1列目のテーブルはそのままにする。
                TmpStr = WithTableHairetu(I, J)
                WithTableHairetu(I, J) = 文字列を指定バイト数文字数に省略(TmpStr, HyoujiMaxNagasa)
            End If
            
            NagasaList(I, J) = LenB(StrConv(WithTableHairetu(I, J), vbFromUnicode)) '全角と半角を区別して長さを計算する。
            MaxNagasaList(J) = WorksheetFunction.Max(MaxNagasaList(J), NagasaList(I, J))
            
        Next I
    Next J
    
    '※※※※※※※※※※※※※※※※※※※※※※※※※※※
    'イミディエイトウィンドウに表示するために" "(半角スペース)を追加して
    '文字列長さを同じにする。
    ReDim NagasaOnajiList(1 To N, 1 To M)
    Dim TmpMaxNagasa As Long
    
    For J = 1 To M
        TmpMaxNagasa = MaxNagasaList(J) 'その列の最大文字列長さ
        For I = 1 To N
            'Rept…指定文字列を指定個数連続してつなげた文字列を出力する。
            '（最大文字数-文字数）の分" "（半角スペース）を後ろにくっつける。
            NagasaOnajiList(I, J) = WithTableHairetu(I, J) & WorksheetFunction.Rept(" ", TmpMaxNagasa - NagasaList(I, J))
       
        Next I
    Next J
    
    '※※※※※※※※※※※※※※※※※※※※※※※※※※※
    'イミディエイトウィンドウに表示する文字列を作成
    ReDim OutputList(1 To N)
    For I = 1 To N
        For J = 1 To M
            If J = 1 Then
                OutputList(I) = NagasaOnajiList(I, J)
            Else
                OutputList(I) = OutputList(I) & SikiriMoji & NagasaOnajiList(I, J)
            End If
        Next J
    Next I
    
    ''※※※※※※※※※※※※※※※※※※※※※※※※※※※
    'イミディエイトウィンドウに表示
    Debug.Print HairetuName
    For I = 1 To N
        Debug.Print OutputList(I)
    Next I
    
End Sub

Private Function 文字列を指定バイト数文字数に省略(Mojiretu As String, ByteNum As Integer)
    '20201023追加
    '文字列を指定省略バイト文字数までの長さで省略する。
    '省略された文字列の最後の文字は"."に変更する。
    '例：Mojiretu = "魑魅魍魎" , ByteNum = 6 … 出力 = "魑魅.."
    '例：Mojiretu = "魑魅魍魎" , ByteNum = 7 … 出力 = "魑魅魍."
    '例：Mojiretu = "魑魅XX魎" , ByteNum = 6 … 出力 = "魑魅X."
    '例：Mojiretu = "魑魅XX魎" , ByteNum = 7 … 出力 = "魑魅XX."
    
    Dim OriginByte As Integer '入力した文字列「Mojiretu」のバイト文字数
    Dim Output                '出力する変数を格納
    
    '「Mojiretu」のバイト文字数計算
    OriginByte = LenB(StrConv(Mojiretu, vbFromUnicode))
    
    If OriginByte <= ByteNum Then
        '「Mojiretu」のバイト文字数計算が省略するバイト文字数以下なら
        '省略はしない
        Output = Mojiretu
    Else
    
        Dim RuikeiByteList, BunkaiMojiretu
        RuikeiByteList = 文字列の各文字累計バイト数計算(Mojiretu)
        BunkaiMojiretu = 文字列分解(Mojiretu)
        
        Dim AddMoji As String
        AddMoji = "."
        
        Dim I As Long, N As Long
        N = Len(Mojiretu)
        
        For I = 1 To N
            If RuikeiByteList(I) < ByteNum Then
                Output = Output & BunkaiMojiretu(I)
                
            ElseIf RuikeiByteList(I) = ByteNum Then
                If LenB(StrConv(BunkaiMojiretu(I), vbFromUnicode)) = 1 Then
                    '例：Mojiretu = "魑魅魍魎" , ByteNum = 6 ,RuikeiByteList(3) = 6
                    'Output = "魑魅.."
                    Output = Output & AddMoji
                Else
                    '例：Mojiretu = "魑魅XX魎" , ByteNum = 6 ,RuikeiByteList(4) = 6
                    'Output = "魑魅X."
                    Output = Output & AddMoji & AddMoji
                End If
                
                Exit For
                
            ElseIf RuikeiByteList(I) > ByteNum Then
                '例：Mojiretu = "魑魅魍魎" , ByteNum = 7 ,RuikeiByteList(4) = 8
                'Output = "魑魅魍."
                Output = Output & AddMoji
                Exit For
            End If
        Next I
        
    End If
        
    文字列を指定バイト数文字数に省略 = Output

    
End Function

Private Function 文字列の各文字累計バイト数計算(Mojiretu As String)
    '20201023追加

    '文字列を1文字ずつに分解して、各文字のバイト文字長を計算し、
    'その累計値を計算する。
    '例：Mojiretu="新型EKワゴン"
    '出力→Output = (2,4,5,6,7,10,12)
    
    Dim MojiKosu As Integer
    Dim I        As Long
    Dim TmpMoji  As String
    Dim Output
    MojiKosu = Len(Mojiretu)
    ReDim Output(1 To MojiKosu)
    
    For I = 1 To MojiKosu
        TmpMoji = Mid(Mojiretu, I, 1)
        If I = 1 Then
            Output(I) = LenB(StrConv(TmpMoji, vbFromUnicode))
        Else
            Output(I) = LenB(StrConv(TmpMoji, vbFromUnicode)) + Output(I - 1)
        End If
    Next I
    
    文字列の各文字累計バイト数計算 = Output
    
End Function

Private Function 文字列分解(Mojiretu As String)
    '20201023追加

    '文字列を1文字ずつ分解して配列に格納
    Dim I     As Long
    Dim N     As Long
    Dim Output
    
    N = Len(Mojiretu)
    ReDim Output(1 To N)
    For I = 1 To N
        Output(I) = Mid(Mojiretu, I, 1)
    Next I
    
    文字列分解 = Output
    
End Function

Private Sub ClipboardCopy(ByVal InputClipText, Optional MessageIrunaraTrue As Boolean = False)
'入力テキストをクリップボードに格納
'配列ならば列方向をTabわけ、行方向を改行する。
'20210719作成
    
    '入力した引数が配列か、配列の場合は1次元配列か、2次元配列か判定
    Dim HairetuHantei As Integer
    Dim Jigen1        As Integer
    Dim Jigen2        As Integer
    If IsArray(InputClipText) = False Then
        '入力引数が配列でない
        HairetuHantei = 0
    Else
        On Error Resume Next
        Jigen2 = UBound(InputClipText, 2)
        On Error GoTo 0
        
        If Jigen2 = 0 Then
            HairetuHantei = 1
        Else
            HairetuHantei = 2
        End If
    End If
    
    'クリップボードに格納用のテキスト変数を作成
    Dim Output As String
    Dim I      As Integer
    Dim J      As Integer
    Dim M      As Integer
    Dim N      As Integer
    
    If HairetuHantei = 0 Then '配列でない場合
        Output = InputClipText
    ElseIf HairetuHantei = 1 Then '1次元配列の場合
    
        If LBound(InputClipText, 1) <> 1 Then '最初の要素番号が1出ない場合は最初の要素番号を1にする
            InputClipText = Application.Transpose(Application.Transpose(InputClipText))
        End If
        
        N = UBound(InputClipText, 1)
        
        Output = ""
        For I = 1 To N
            If I = 1 Then
                Output = InputClipText(I)
            Else
                Output = Output & vbLf & InputClipText(I)
            End If
            
        Next I
    ElseIf HairetuHantei = 2 Then '2次元配列の場合
        
        If LBound(InputClipText, 1) <> 1 Or LBound(InputClipText, 2) <> 1 Then
            InputClipText = Application.Transpose(Application.Transpose(InputClipText))
        End If
        
        N = UBound(InputClipText, 1)
        M = UBound(InputClipText, 2)
        
        Output = ""
        
        For I = 1 To N
            For J = 1 To M
                If J < M Then
                    Output = Output & InputClipText(I, J) & Chr(9)
                Else
                    Output = Output & InputClipText(I, J)
                End If
                
            Next J
            
            If I < N Then
                Output = Output & Chr(10)
            End If
        Next I
    End If
    
    
    'クリップボードに格納'参考 https://www.ka-net.org/blog/?p=7537
    With CreateObject("Forms.TextBox.1")
        .MultiLine = True
        .Text = Output
        .SelStart = 0
        .SelLength = .TextLength
        .Copy
    End With

    '格納したテキスト変数をメッセージ表示
    If MessageIrunaraTrue Then
        MsgBox ("「" & Output & "」" & vbLf & _
                "をクリップボードにコピーしました。")
    End If
    
End Sub


