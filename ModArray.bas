Attribute VB_Name = "ModArray"
Option Explicit

'�z��̏����֌W�̃v���V�[�W��

Function SortArrayByNetFramework(InputArray, Optional InputOrder As OrderType = xlAscending)
'�ꎟ���z���Net.Framework���g���ď����ɂ���
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
'�w���2�����z����A�w������ɕ��ёւ���
        
    '�w����1�����z��Œ��o
    Dim KijunArray1D
    Dim MinRow&, MaxRow&
    MinRow = LBound(InputArray2D, 1)
    MaxRow = UBound(InputArray2D, 1)
    ReDim KijunArray1D(MinRow To MaxRow)
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    For I = MinRow To MaxRow
        KijunArray1D(I) = InputArray2D(I, SortCol)
    Next I
    
    '���ёւ�
    Dim Output
    Output = SortArray2Dby1D(InputArray2D, KijunArray1D, InputOrder)
    SortArray2D = Output
    
End Function
Function SortArray2Dby1D(InputArray2D, ByVal KijunArray1D, Optional InputOrder As OrderType = xlAscending)
'�w���2�����z����A�w��1�����z�����ɕ��ёւ���

    '���͒l�̃`�F�b�N
    Dim Dummy%
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
    
    Dim MinRow&, MaxRow&, MinCol&, MaxCol&
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
    Dim StrAruNaraTrue As Boolean
    StrAruNaraTrue = False
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
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
    
    '��z��𐳋K�����āA(1�`�v�f��)�̊Ԃ̐��l�ɂ���
    Dim Count&, MinNum#, MaxNum#
    Count = MaxRow - MinRow + 1
    MinNum = WorksheetFunction.Min(KijunArray1D)
    MaxNum = WorksheetFunction.Max(KijunArray1D)
    For I = MinRow To MaxRow
        KijunArray1D(I) = (KijunArray1D(I) - MinNum) / (MaxNum - MinNum) '(0�`1)�̊ԂŐ��K��
        KijunArray1D(I) = (Count - 1) * KijunArray1D(I) + 1 '(1�`�v�f��)�̊�
    Next I
    
    '���ёւ�(1,2,3�̔z�������ăN�C�b�N�\�[�g�ŕ��ёւ��āA�Ώۂ̔z�����ёւ����1,2,3�œ���ւ���)
    Dim TmpArray, Array123, TmpNum&
    ReDim Array123(MinRow To MaxRow)
    For I = MinRow To MaxRow
        If InputOrder = xlAscending Then
            '����
            Array123(I) = I - MinRow + 1
        Else
            '�~��
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
    
    '�o��
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
'�N�C�b�N�\�[�g�Ŕz�����ёւ���

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
    
    '���ёւ��Ώۂ̔z��̏���
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
        Call SortArrayQuick(KijunArray, Array123, StartNum, I - 1) '�ċA�Ăяo��
    End If
    If EndNum - J > 1 Then
        Call SortArrayQuick(KijunArray, Array123, J + 1, EndNum) '�ċA�Ăяo��
    End If
End Sub
Function ConvStrToISO(InputStr$)
'���������ёւ��p��ISO�R�[�h�ɕϊ�
'20210726

    Dim Mojiretu As String
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    Dim UniCode
    
    Dim UniMax&
    UniMax = 65536
    
    Dim StartKeta%, Kurai#
    StartKeta = 20 '����������������������������������������������
    Kurai = Exp(1) '����������������������������������������������
    
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