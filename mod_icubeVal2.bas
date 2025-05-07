Attribute VB_Name = "mod_icubeVal2"


Option Compare Database
Option Explicit

'�e�[�u��Icube�֊�{�H�����̂𕪊����ċL��
Public Sub mod_icube_Val2ALL()

'�e�[�u��Icube�֊�{�H�����̂̕����L��
    Call mod_icube_Val2copy1
'�e�[�u��Icube�֊�{�H�����̂���̊��L��
    Call mod_icube_Val2copy2
'�e�[�u��Icube�։���{�H���R�[�h�]��
    Call mod_icube_Val2copy3


End Sub


'�e�[�u��Icube�։���{�H���R�[�h�]��
Public Sub mod_icube_Val2copy3()
    On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim rsIcube As DAO.Recordset
    Dim rsRef As DAO.Recordset
    Dim sqlIcube As String, sqlRef As String
    Dim ��Ə� As String, Q As String, ���� As String, �J�z As String
    Dim matchCount As Long

    Set db = CurrentDb

    ' �Q�ƃe�[�u���S���擾
    sqlRef = "SELECT * FROM tb_����{�H��"
    Set rsRef = db.OpenRecordset(sqlRef, dbOpenSnapshot)

    ' �ꌏ�H������ = "�����H��" �̂ݑΏ�
    sqlIcube = "SELECT * FROM Icube_ WHERE �ꌏ�H������ = '�����H��'"
    Set rsIcube = db.OpenRecordset(sqlIcube, dbOpenDynaset)

    matchCount = 0

    Do While Not rsIcube.EOF
        Dim hitFound As Boolean
        hitFound = False

        ��Ə� = Trim(Nz(rsIcube!��{�H����_��Ə�, ""))
        Q = Trim(Nz(rsIcube!��{�H����_Q, ""))
        ���� = Trim(Nz(rsIcube!��{�H����_����, ""))
        �J�z = Trim(Nz(rsIcube!��{�H����_�J�z, ""))

        rsRef.MoveFirst
        Do While Not rsRef.EOF
            If Trim(Nz(rsRef!��{�H����_��Ə�, "")) = ��Ə� And _
               Trim(Nz(rsRef!��{�H����_Q, "")) = Q And _
               Trim(Nz(rsRef!��{�H����_����, "")) = ���� And _
               Trim(Nz(rsRef!��{�H����_�J�z, "")) = �J�z Then

                rsIcube.Edit
                rsIcube!����{�H���R�[�h = rsRef!����{�H���R�[�h
                rsIcube.Update
                matchCount = matchCount + 1
                hitFound = True
                Exit Do
            End If
            rsRef.MoveNext
        Loop

        '--- �f�o�b�O�p�o�́i���݂͖��������j
        'If Not hitFound Then
        '    Debug.Print "��v����: No=" & rsIcube!No & " �b��Ə�=[" & ��Ə� & "] Q=[" & Q & "] ����=[" & ���� & "] �J�z=[" & �J�z & "]"
        'End If

        rsIcube.MoveNext
    Loop

    'MsgBox "�X�V�����j���I" & vbCrLf & "��v���ē]�ʂ��ꂽ�����F" & matchCount, vbInformation

Exit_Handler:
    On Error Resume Next
    rsIcube.Close
    rsRef.Close
    Set rsIcube = Nothing
    Set rsRef = Nothing
    Set db = Nothing
    Exit Sub

Err_Handler:
    MsgBox "�G���[�����������ɂ�F" & Err.description, vbExclamation
    Resume Exit_Handler
End Sub




'�e�[�u��Icube�֊�{�H�����̂���̊��L��
Sub mod_icube_Val2copy2()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim n�N�x As Variant
    Dim �������� As String
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("Icube_", dbOpenDynaset)
    
    Do While Not rs.EOF
        �������� = Nz(rs!��{�H����_�N�x, "")
        
        If �������� <> "" And �������� <> "N/A" Then
            ' �S�p�𔼊p�ɕϊ����Ă��琔�l��
            n�N�x = val(StrConv(��������, vbNarrow)) - 12
        Else
            n�N�x = Null
        End If
        
        rs.Edit
        rs!��{�H����_�� = n�N�x
        rs.Update
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    'MsgBox "�l���S�� ?12 ���������A���������͂��j��?�I", vbInformation
End Sub



'�e�[�u��Icube�֊�{�H�����̂̕����L��
Sub mod_icube_Val2copy1()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim �������� As String
    Dim pos�N�x As Long, posRN As Long, ��Ə� As String
    Dim �N�x As String, Q As String, ���� As String, �J�z As String
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("Icube_", dbOpenDynaset)
    
    Do While Not rs.EOF
        If rs!�ꌏ�H������ = "�����H��" Then
            �������� = Nz(rs!��{�H������, "")
            
            '����Ə�
            If Left(��������, 3) = "���z��" Then
                ��Ə� = "���z��"
            Else
                posRN = InStr(��������, "�q�m")
                If posRN >= 3 Then
                    ��Ə� = Mid(��������, posRN - 2, 2)
                Else
                    ��Ə� = ""
                End If
            End If
            
            '���N�x
            pos�N�x = InStr(��������, "�N�x")
            If pos�N�x >= 3 Then
                �N�x = Mid(��������, pos�N�x - 2, 2)
            Else
                �N�x = ""
            End If
            
            '��Q
            If pos�N�x > 0 And Len(��������) >= pos�N�x + 2 Then
                Q = Mid(��������, pos�N�x + 2, 2)
            Else
                Q = ""
            End If
            
            '������
            If InStr(��������, "����") > 0 Then
                ���� = "����"
            ElseIf InStr(��������, "����") > 0 Then
                ���� = "����"
            Else
                ���� = ""
            End If
            
            '���J�z
            If InStr(��������, "�i�J�z�j") > 0 Then
                �J�z = "�i�J�z�j"
            Else
                �J�z = ""
            End If
            
        Else
            ' �ꌏ�H�����肪�����H���ȊO
            ��Ə� = "N/A"
            �N�x = "N/A"
            Q = "N/A"
            ���� = "N/A"
            �J�z = "N/A"
        End If
        
        ' �]��
        rs.Edit
        rs!��{�H����_��Ə� = ��Ə�
        rs!��{�H����_�N�x = �N�x
        rs!��{�H����_Q = Q
        rs!��{�H����_���� = ����
        rs!��{�H����_�J�z = �J�z
        rs.Update
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    'MsgBox "�]�ʏ������������܂����j��?�I", vbInformation
End Sub


