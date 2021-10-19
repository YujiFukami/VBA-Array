Attribute VB_Name = "ModArray"
Option Explicit

'�z��̏����֌W�̃v���V�[�W��

Function SortArrayByNetFramework(InputArray, Optional InputOrder As OrderType = xlAscending)
'�ꎟ���z���Net.Framework���g���ď����ɂ���
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
'�w���2�����z����A�w������ɕ��ёւ���
'�z��͕�������܂�ł��Ă��悢
'20210726

'InputArray2D�E�E�E���ёւ��Ώۂ�2�����z��
'SortCol     �E�E�E���ёւ��̊�Ŏw�肷���ԍ�
'InputOrder  �E�E�ExlAscending������, xlDescending���~��

    '�w����1�����z��Œ��o
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
    
    '���ёւ�
    Dim Output
    Output = SortArray2Dby1D(InputArray2D, KijunArray1D, InputOrder)
    SortArray2D = Output
    
End Function

Private Function SortArray2Dby1D(InputArray2D, ByVal KijunArray1D, Optional InputOrder As OrderType = xlAscending)
'�w���2�����z����A�w��1�����z�����ɕ��ёւ���
'�z��͕�������܂�ł��Ă��悢
'20210726
'20210917�C��
'20211016�C�� �z��̂̒��g���I�u�W�F�N�g�ϐ��ł��Ή�
                                   
'InputArray2D�E�E�E���ёւ��Ώۂ�2�����z��
'KijunArray1D�E�E�E���ёւ��̊�ƂȂ�z��
'InputOrder  �E�E�ExlAscending������, xlDescending���~��

    '���͒l�̃`�F�b�N
    Dim Dummy As Integer
    On Error Resume Next
    Dummy = UBound(KijunArray1D, 2)
    On Error GoTo 0
    If Dummy <> 0 Then
        MsgBox ("��z���1�����z�����͂��Ă�������")
        Stop
        End
    End If
    
    Dummy = 0
    On Error Resume Next
    Dummy = UBound(InputArray2D, 2)
    On Error GoTo 0
    If Dummy = 0 Then
        MsgBox ("���ёւ��Ώۂ̔z���2�����z�����͂��Ă�������")
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
        MsgBox ("���ёւ��Ώۂ̔z��ƁA��z��̊J�n�A�I���v�f�ԍ�����v�����Ă�������")
        Stop
        End
    End If
    
    '��z��ɕ����񂪊܂܂�Ă���ꍇ��ISO�R�[�h�ɕϊ�
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
    
    '��z��𐳋K�����āA(1�`�v�f��)�̊Ԃ̐��l�ɂ���
    Dim Count  As Long
    Dim MinNum As Double
    Dim MaxNum As Double
    Count = MaxRow - MinRow + 1
    MinNum = WorksheetFunction.Min(KijunArray1D)
    MaxNum = WorksheetFunction.Max(KijunArray1D)
    
    Dim TmpArray
    If MinNum = MaxNum Then '20211016�C��'�ő�ƍŏ�����v����Ȃ炻�̂܂ܕԂ�
        TmpArray = InputArray2D
        GoTo EndEscape
    End If
    
    For I = MinRow To MaxRow
        KijunArray1D(I) = (KijunArray1D(I) - MinNum) / (MaxNum - MinNum) '(0�`1)�̊ԂŐ��K��
        KijunArray1D(I) = (Count - 1) * KijunArray1D(I) + 1 '(1�`�v�f��)�̊�
    Next I
    
    '���ёւ�(1,2,3�̔z�������ăN�C�b�N�\�[�g�ŕ��ёւ��āA�Ώۂ̔z�����ёւ����1,2,3�œ���ւ���)
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
            If InputOrder = xlAscending Then '20210917�C��
                If IsObject(InputArray2D(TmpNum, J)) Then '20211016�C��
                    Set TmpArray(I, J) = InputArray2D(TmpNum, J)
                Else
                    TmpArray(I, J) = InputArray2D(TmpNum, J)
                End If
            Else
                If IsObject(InputArray2D(TmpNum, J)) Then '20211016�C��
                    Set TmpArray(MaxRow - I + 1, J) = InputArray2D(TmpNum, J)
                Else
                    TmpArray(MaxRow - I + 1, J) = InputArray2D(TmpNum, J)
                End If
            End If
                
        Next J
    Next I

EndEscape:

    '�o��
    SortArray2Dby1D = TmpArray

End Function

Sub SortArrayQuick(KijunArray, Array123, Optional StartNum As Integer, Optional EndNum As Integer)
'�N�C�b�N�\�[�g��1�����z�����ёւ���
'���ёւ���̏��Ԃ��o�͂��邽�߂ɔz��uArray123�v�𓯎��ɕ��ёւ���
'20210726
'20211016�C�� �z��̒��g���I�u�W�F�N�g�ϐ��ł��Ή�

'KijunArray�E�E�E���ёւ��Ώۂ̔z��i1�����z��j
'Array123  �E�E�E�u1,2,3�v�̒l��������1�����z��
'StartNum  �E�E�E�ċA�p�̈���
'EndNum    �E�E�E�ċA�p�̈���

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
    
    '���ёւ��Ώۂ̔z��̏���
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
        
        If IsObject(Array123(I)) Then '20211016�C��
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
        Call SortArrayQuick(KijunArray, Array123, StartNum, I - 1) '�ċA�Ăяo��
    End If
    If EndNum - J > 1 Then
        Call SortArrayQuick(KijunArray, Array123, J + 1, EndNum) '�ċA�Ăяo��
    End If
End Sub

Function ConvStrToISO(InputStr As String)
'���������ёւ��p��ISO�R�[�h�ɕϊ�
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
    StartKeta = 20 '����������������������������������������������
    Kurai = Exp(1) '����������������������������������������������
    
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

Sub CheckArray1D(InputArray, Optional HairetuName As String = "�z��")
'���͔z��1�����z�񂩂ǂ����`�F�b�N����
'20210804

    Dim Dummy As Integer
    On Error Resume Next
    Dummy = UBound(InputArray, 2)
    On Error GoTo 0
    If Dummy <> 0 Then
        MsgBox (HairetuName & "��1�����z�����͂��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Sub CheckArray2D(InputArray, Optional HairetuName As String = "�z��")
'���͔z��2�����z�񂩂ǂ����`�F�b�N����
'20210804

    Dim Dummy2 As Integer
    Dim Dummy3 As Integer
    On Error Resume Next
    Dummy2 = UBound(InputArray, 2)
    Dummy3 = UBound(InputArray, 3)
    On Error GoTo 0
    If Dummy2 = 0 Or Dummy3 <> 0 Then
        MsgBox (HairetuName & "��2�����z�����͂��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Sub CheckArray1DStart1(InputArray, Optional HairetuName As String = "�z��")
'����1�����z��̊J�n�ԍ���1���ǂ����`�F�b�N����
'20210804

    If LBound(InputArray, 1) <> 1 Then
        MsgBox (HairetuName & "�̊J�n�v�f�ԍ���1�ɂ��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Sub CheckArray2DStart1(InputArray, Optional HairetuName As String = "�z��")
'����2�����z��̊J�n�ԍ���1���ǂ����`�F�b�N����
'20210804

    If LBound(InputArray, 1) <> 1 Or LBound(InputArray, 2) <> 1 Then
        MsgBox (HairetuName & "�̊J�n�v�f�ԍ���1�ɂ��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub


Sub ClipCopyArray2D(Array2D)
'2�����z���ϐ��錾�p�̃e�L�X�g�f�[�^�ɕϊ����āA�N���b�v�{�[�h�ɃR�s�[����
'20210805
'20211016 �u"�v���܂ޏꍇ���Ή�

    '�����`�F�b�N
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I As Long
    Dim J As Long
    Dim M As Long
    Dim N As Long '�����グ�p(Long�^)
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
'1�����z���ϐ��錾�p�̃e�L�X�g�f�[�^�ɕϊ����āA�N���b�v�{�[�h�ɃR�s�[����
'20210805
    
    '�����`�F�b�N
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
'�񎟌��z������b�Z�[�W�ɕ\������B
'20210824

    '�����`�F�b�N
    Call CheckArray2D(Array2D)
    Call CheckArray2DStart1(Array2D)
    
    '����
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
'1�����z����A(1,N)�z��ɕϊ�����
'20210917

'InputArray1D�E�E�E1�����z��

    '�����`�F�b�N
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
    
    '�o��
    ConvArray1Dto1N = Output

End Function

Function DeleteRowArray(Array2D, DeleteRow As Long)
'�񎟌��z��̎w��s�����������z����o�͂���
'20210917

'����
'Array2D  �E�E�E�񎟌��z��
'DeleteRow�E�E�E��������s�ԍ�

    '�����`�F�b�N
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I As Long
    Dim J As Long
    Dim K As Long
    Dim M As Long
    Dim N As Long
    N = UBound(Array2D, 1) '�s��
    M = UBound(Array2D, 2) '��
    
    If DeleteRow < 1 Then
        MsgBox ("�폜����s�ԍ���1�ȏ�̒l�����Ă�������")
        Stop
        End
    ElseIf DeleteRow > N Then
        MsgBox ("�폜����s�ԍ��͌��̓񎟌��z��̍s��" & N & "�ȉ��̒l�����Ă�������")
        Stop
        End
    End If
    
    '����
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
    
    '�o��
    DeleteRowArray = Output

End Function

Function DeleteColArray(Array2D, DeleteCol As Long)
'�񎟌��z��̎w�������������z����o�͂���
'20210917

'����
'Array2D  �E�E�E�񎟌��z��
'DeleteCol�E�E�E���������ԍ�

    '�����`�F�b�N
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I As Long
    Dim J As Long
    Dim K As Long
    Dim M As Long
    Dim N As Long
    N = UBound(Array2D, 1) '�s��
    M = UBound(Array2D, 2) '��

    If DeleteCol < 1 Then
        MsgBox ("�폜�����ԍ���1�ȏ�̒l�����Ă�������")
        Stop
        End
    ElseIf DeleteCol > M Then
        MsgBox ("�폜�����ԍ��͌��̓񎟌��z��̗�" & M & "�ȉ��̒l�����Ă�������")
        Stop
        End
    End If
    
    '����
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
    
    '�o��
    DeleteColArray = Output

End Function

Function ExtractRowArray(Array2D, TargetRow As Long)
'�񎟌��z��̎w��s���ꎟ���z��Œ��o����
'20210917

'����
'Array2D  �E�E�E�񎟌��z��
'TargetRow�E�E�E���o����Ώۂ̍s�ԍ�


    '�����`�F�b�N
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I As Long
    Dim M As Long
    Dim N As Long
    N = UBound(Array2D, 1) '�s��
    M = UBound(Array2D, 2) '��

    If TargetRow < 1 Then
        MsgBox ("���o����s�ԍ���1�ȏ�̒l�����Ă�������")
        Stop
        End
    ElseIf TargetRow > N Then
        MsgBox ("���o����s�ԍ��͌��̓񎟌��z��̍s��" & N & "�ȉ��̒l�����Ă�������")
        Stop
        End
    End If

    '����
    Dim Output
    ReDim Output(1 To M)
    
    For I = 1 To M
        Output(I) = Array2D(TargetRow, I)
    Next I
    
    '�o��
    ExtractRowArray = Output
    
End Function

Function ExtractColArray(Array2D, TargetCol As Long)
'�񎟌��z��̎w�����ꎟ���z��Œ��o����
'20210917
'20211016�C�� �z��̒��g���I�u�W�F�N�g�ϐ��ł��Ή�

'����
'Array2D  �E�E�E�񎟌��z��
'TargetCol�E�E�E���o����Ώۂ̗�ԍ�

    '�����`�F�b�N
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I As Long
    Dim M As Long
    Dim N As Long
    N = UBound(Array2D, 1) '�s��
    M = UBound(Array2D, 2) '��
 
    If TargetCol < 1 Then
        MsgBox ("���o�����ԍ���1�ȏ�̒l�����Ă�������")
        Stop
        End
    ElseIf TargetCol > N Then
        MsgBox ("���o�����ԍ��͌��̓񎟌��z��̍s��" & M & "�ȉ��̒l�����Ă�������")
        Stop
        End
    End If
    
    '����
    Dim Output
    ReDim Output(1 To N)
    
    For I = 1 To N
        If IsObject(Array2D(I, TargetCol)) Then '20211016�C��
            Set Output(I) = Array2D(I, TargetCol)
        Else
            Output(I) = Array2D(I, TargetCol)
        End If
    Next I
    
    '�o��
    ExtractColArray = Output
    
End Function

Function ExtractArray(Array2D, StartRow As Long, StartCol As Long, EndRow As Long, EndCol As Long)
'�񎟌��z��̎w��͈͂�z��Ƃ��Ē��o����
'20210917

'����
'Array2D �E�E�E�񎟌��z��
'StartRow�E�E�E���o�͈͂̊J�n�s�ԍ�
'StartCol�E�E�E���o�͈͂̊J�n��ԍ�
'EndRow  �E�E�E���o�͈͂̏I���s�ԍ�
'EndCol  �E�E�E���o�͈͂̏I����ԍ�
                                   
    '�����`�F�b�N
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I As Long
    Dim J As Long
    Dim M As Long
    Dim N As Long
    N = UBound(Array2D, 1) '�s��
    M = UBound(Array2D, 2) '��
    
    If StartRow > EndRow Then
        MsgBox ("���o�͈͂̊J�n�s�uStartRow�v�́A�I���s�uEndRow�v�ȉ��łȂ���΂Ȃ�܂���")
        Stop
        End
    ElseIf StartCol > EndCol Then
        MsgBox ("���o�͈͂̊J�n��uStartCol�v�́A�I����uEndCol�v�ȉ��łȂ���΂Ȃ�܂���")
        Stop
        End
    ElseIf StartRow < 1 Then
        MsgBox ("���o�͈͂̊J�n�s�uStartRow�v��1�ȏ�̒l�����Ă�������")
        Stop
        End
    ElseIf StartCol < 1 Then
        MsgBox ("���o�͈͂̊J�n��uStartCol�v��1�ȏ�̒l�����Ă�������")
        Stop
        End
    ElseIf EndRow > N Then
        MsgBox ("���o�͈͂̏I���s�uStartRow�v�͒��o���̓񎟌��z��̍s��" & N & "�ȉ��̒l�����Ă�������")
        Stop
        End
    ElseIf EndCol > M Then
        MsgBox ("���o�͈͂̏I����uStartCol�v�͒��o���̓񎟌��z��̗�" & M & "�ȉ��̒l�����Ă�������")
        Stop
        End
    End If
    
    '����
    Dim Output
    ReDim Output(1 To EndRow - StartRow + 1, 1 To EndCol - StartCol + 1)
    
    For I = StartRow To EndRow
        For J = StartCol To EndCol
            Output(I - StartRow + 1, J - StartCol + 1) = Array2D(I, J)
        Next J
    Next I
    
    '�o��
    ExtractArray = Output
    
End Function

Function ExtractArray1D(Array1D, StartNum As Long, EndNum As Long)
'�ꎟ���z��̎w��͈͂�z��Ƃ��Ē��o����
'20211009

'����
'Array1D �E�E�E�ꎟ���z��
'StartNum�E�E�E���o�͈͂̊J�n�ԍ�
'EndNum  �E�E�E���o�͈͂̏I���ԍ�
                                   
    '�����`�F�b�N
    Call CheckArray1D(Array1D, "Array1D")
    Call CheckArray1DStart1(Array1D, "Array1D")
    
    Dim I As Long
    Dim N As Long
    N = UBound(Array1D, 1) '�v�f��
    
    If StartNum > EndNum Then
        MsgBox ("���o�͈͂̊J�n�ʒu�uStartNum�v�́A�I���ʒu�uEndNum�v�ȉ��łȂ���΂Ȃ�܂���")
        Stop
        Exit Function
    ElseIf StartNum < 1 Then
        MsgBox ("���o�͈͂̊J�n�ʒu�uStartNum�v��1�ȏ�̒l�����Ă�������")
        Stop
        Exit Function
    ElseIf EndNum > N Then
        MsgBox ("���o�͈͂̏I���s�uEndNum�v�͒��o���̈ꎟ���z��̗v�f��" & N & "�ȉ��̒l�����Ă�������")
        Stop
        Exit Function
    End If
    
    '����
    Dim Output
    ReDim Output(1 To EndNum - StartNum + 1)
    
    For I = StartNum To EndNum
        Output(I - StartNum + 1) = Array1D(I)
    Next I
    
    '�o��
    ExtractArray1D = Output
    
End Function

Function UnionArray1D(UpperArray1D, LowerArray1D)
'�ꎟ���z�񓯎m����������1�̔z��Ƃ���B
'20210923

'UpperArray1D�E�E�E��Ɍ�������ꎟ���z��
'LowerArray1D�E�E�E���Ɍ�������ꎟ���z��

    '�����`�F�b�N
    Call CheckArray1D(UpperArray1D, "UpperArray1D")
    Call CheckArray1DStart1(UpperArray1D, "UpperArray1D")
    Call CheckArray1D(LowerArray1D, "LowerArray1D")
    Call CheckArray1DStart1(LowerArray1D, "LowerArray1D")
    
    '����
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
    
    '�o��
    UnionArray1D = Output
    
End Function

Function UnionArrayLR1D(LeftArray1D, RightArray1D)
'�ꎟ���z�񓯎m�����E�Ɍ�������1�̔z��Ƃ���B
'20211016

'LeftArray1D �E�E�E���Ɍ�������ꎟ���z��
'RightArray1D�E�E�E�E�Ɍ�������ꎟ���z��

    '�����`�F�b�N
    Call CheckArray1D(LeftArray1D, "LeftArray1D")
    Call CheckArray1DStart1(LeftArray1D, "LeftArray1D")
    Call CheckArray1D(RightArray1D, "RightArray1D")
    Call CheckArray1DStart1(RightArray1D, "RightArray1D")
    If UBound(LeftArray1D, 1) <> UBound(RightArray1D, 1) Then
        MsgBox ("LeftArray1D��RightArray1D�̗v�f���͑����Ă�������")
        Stop
        Exit Function
    End If
    
    '����
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
    
    '�o��
    UnionArrayLR1D = Output
    
End Function

Function UnionArrayUL(UpperArray2D, LowerArray2D)
'�񎟌��z�񓯎m���㉺�Ɍ�������1�̔z��Ƃ���B
'20210923

'UpperArray2D�E�E�E��Ɍ�������񎟌��z��
'LowerArray2D�E�E�E���Ɍ�������񎟌��z��

    '�����`�F�b�N
    Call CheckArray2D(UpperArray2D, "UpperArray2D")
    Call CheckArray2DStart1(UpperArray2D, "UpperArray2D")
    Call CheckArray2D(LowerArray2D, "LowerArray2D")
    Call CheckArray2DStart1(LowerArray2D, "LowerArray2D")
    
    If UBound(UpperArray2D, 2) <> UBound(LowerArray2D, 2) Then
        MsgBox ("UpperArray2D��LowerArray2D�̓񎟌��v�f�������킹�Ă�������" & vbLf & _
                "UpperArray2D�̓񎟌��v�f�� = " & UBound(UpperArray2D, 2) & vbLf & _
                "LowerArray2D�̓񎟌��v�f�� = " & UBound(LowerArray2D, 2))
    End If
    
    '����
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
    
    '�o��
    UnionArrayUL = Output

End Function

Function UnionArrayLR(LeftArray2D, RightArray2D)
'�񎟌��z�񓯎m�����E�Ɍ�������1�̔z��Ƃ���B
'20210923

'LeftArray2D �E�E�E��Ɍ�������񎟌��z��
'RightArray2D�E�E�E���Ɍ�������񎟌��z��

    '�����`�F�b�N
    Call CheckArray2D(LeftArray2D, "LeftArray2D")
    Call CheckArray2DStart1(LeftArray2D, "LeftArray2D")
    Call CheckArray2D(RightArray2D, "RightArray2D")
    Call CheckArray2DStart1(RightArray2D, "RightArray2D")
    
    If UBound(LeftArray2D, 1) <> UBound(RightArray2D, 1) Then
        MsgBox ("LeftArray2D��RightArray2D�̈ꎟ���v�f�������킹�Ă�������" & vbLf & _
                "LeftArray2D�̈ꎟ���v�f�� = " & UBound(LeftArray2D, 1) & vbLf & _
                "RightArray2D�̈ꎟ���v�f�� = " & UBound(RightArray2D, 1))
    End If
    
    '����
    Dim Output
    Dim I     As Long
    Dim J     As Long
    Dim K     As Long
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
    
    '�o��
    UnionArrayLR = Output

End Function


Function DimArray1DSameValue(Count As Long, Value)
'�S�ē����l���������ꎟ���z����`����
'20210923

'Count�E�E�E�v�f��(Long�^)
'Value�E�E�E�����l������l(�I�u�W�F�N�g�^�ł��\)
    
    '�����`�F�b�N
    If Count <= 0 Then
        MsgBox ("�v�f��Count��1�ȏ�̒l�����Ă��������B" & vbLf & _
               "Count = " & Count)
        Stop
    End If
    
    '����
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
    
    '�o��
    DimArray1DSameValue = Output
    
End Function

Function DimArray2DSameValue(Count1 As Long, Count2 As Long, Value)
'�S�ē����l���������񎟌��z����`����
'20210923

'Count1�E�E�E�ꎟ���v�f��(Long�^)
'Count2�E�E�E�񎟌��v�f��(Long�^)
'Value �E�E�E�����l������l(�I�u�W�F�N�g�^�ł��\)
    
    '�����`�F�b�N
    If Count1 <= 0 Then
        MsgBox ("�ꎟ���v�f��Count1��1�ȏ�̒l�����Ă��������B" & vbLf & _
               "Count1 = " & Count1)
        Stop
    End If
    
    If Count2 <= 0 Then
        MsgBox ("�񎟌��v�f��Count2��1�ȏ�̒l�����Ă��������B" & vbLf & _
               "Count2 = " & Count2)
        Stop
    End If
    
    '����
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
    
    '�o��
    DimArray2DSameValue = Output
    
End Function

Function FilterArray2D(Array2D, FilterStr As String, TargetCol As Long)
'�񎟌��z����w���Ńt�B���^�[�����z����o�͂���B
'20210929

'����
'Array2D  �E�E�E�񎟌��z��
'FilterStr�E�E�E�t�B���^�[���镶���iString�^�j
'TargetCol�E�E�E�t�B���^�[�����iLong�^�j
    
    '�����`�F�b�N
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    '�t�B���^�[�����v�Z
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
        '�t�B���^�[�ŉ���������Ȃ������ꍇ��Empty��Ԃ�
        FilterArray2D = Empty
        Exit Function
    End If
    
    '�o�͂���z��̍쐬
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
    
    '�o��
    FilterArray2D = Output
    
End Function

Function DimArray1DNumbers(StartNum As Long, EndNum As Long, Optional ByVal StepNum As Long = 1)
'�A�Ԃ̓������ꎟ���z����`����
'20211018

'����
'StartNum �E�E�E�ŏ��̔ԍ�/Long�^
'EndNum�@ �E�E�E�Ō�̔ԍ�/Long�^
'[Step]   �E�E�E�A�Ԃ̊Ԋu/Long�^/�f�t�H���g��1
    
    '�����̃`�F�b�N
    If StepNum = 0 Then
        MsgBox ("StepNum��0�ȊO�̐�������͂��Ă�������")
        Stop
        Exit Function
    End If
    
    '�����̏C��
    If StartNum < EndNum And StepNum < 0 Then
        StepNum = -StepNum
    ElseIf StartNum > EndNum And StepNum > 0 Then
        StepNum = -StepNum
    End If
    
    '�A�Ԃ̍쐬
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
    
    '�o��
    DimArray1DNumbers = Output
    
End Function


