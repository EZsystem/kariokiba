Attribute VB_Name = "mod_icubeVal1"
Option Compare Database
Option Explicit

'�e�[�u���FIcube_�̋L�ڏ���
Public Sub mod_icube_All_1()
'�H�����̂���ꌏ�H������
    Call mod_icube_input1
'��{�R�[�h���u�����N�̎��A�H���R�[�h��]��
    Call mod_icube_copy1
'��{�H�����̂������ꍇ�ɍH�����[����]��
    Call mod_icube_copy2
'�}�ԍH���R�[�h�̋L��(�H���R�[�h�Ǝ}�Ԃ̘A��)
    Call mod_icube_merge1
'�󒍌v��N�����g���A�󒍔N���������L��
    Call mod_icube_dateCnv_1
'���H���}�Ԃ��g���A���H�N���������L��
    Call mod_icube_dateCnv_2
'�e�[�u��Icube�֎{�H�Ǌ��g�D�R�[�h�̋L��
    Call mod_icube_copy4
'�e�[�u��Icube�֎{�H�Ǌ��g�D���̋L��
    Call mod_icube_copy5
'�e�[�u��Icube�֊�{�H�����̂𕪊����ċL��
    Call mod_icube_Val2ALL


' �����������b�Z�[�W
    'MsgBox "�������������܂����B", vbInformation
End Sub



'�e�[�u��Icube�֎{�H�Ǌ��g�D�R�[�h�̋L��
Public Sub mod_icube_copy4()

    ' ��`
    Dim db As DAO.Database
    Dim rsTarget As DAO.Recordset  ' Icube_
    Dim rsError As DAO.Recordset   ' t_err��Ə�
    Dim rsCheck As DAO.Recordset   ' t_�x�X��Ə�_
    Dim strSQLTarget As String
    Dim dict As Object
    Dim key As String
    Dim isErrorExists As Boolean

    ' �f�[�^�x�[�X�̎Q��
    Set db = CurrentDb()

    ' �g�D�R�[�h�������쐬
    Set dict = CreateObject("Scripting.Dictionary")

    ' t_�x�X��Ə�_ ��ǂݍ���� Dictionary �ɕۑ�
    Set rsCheck = db.OpenRecordset("SELECT * FROM t_�x�X��Ə�_", dbOpenSnapshot)
    Do While Not rsCheck.EOF
        key = Trim(CStr(rsCheck!�g�D�R�[�h))
        If Not dict.Exists(key) Then
            If Not IsNull(rsCheck!�{�H�Ǌ��g�D�R�[�h) Then
                ' �K��������Ƃ��ēo�^�iCStr�œ���j
                dict.Add key, CStr(rsCheck!�{�H�Ǌ��g�D�R�[�h)
                'Debug.Print "�o�^�F" & key & " �� " & CStr(rsCheck!�{�H�Ǌ��g�D�R�[�h)
            End If
        End If
        rsCheck.MoveNext
    Loop
    rsCheck.Close
    Set rsCheck = Nothing

    ' Icube_ �̑S�f�[�^���擾�i���[�J���e�[�u���j
    strSQLTarget = "SELECT * FROM Icube_;"
    Set rsTarget = db.OpenRecordset(strSQLTarget, dbOpenDynaset)

    ' �G���[�L�^�p�e�[�u���i���[�J���j
    Set rsError = db.OpenRecordset("t_err��Ə�", dbOpenDynaset)

    ' ������
    isErrorExists = False

    ' Icube_ �̑S���R�[�h�����[�v
    If Not rsTarget.EOF Then
        rsTarget.MoveFirst
        Do While Not rsTarget.EOF
            Dim orgCode As String
            Dim valueToWrite As Variant

            ' �{�H�S���g�D�R�[�h�𕶎���Ŏ��o��
            If Not IsNull(rsTarget!�{�H�S���g�D�R�[�h) Then
                orgCode = Trim(CStr(rsTarget!�{�H�S���g�D�R�[�h))
            Else
                orgCode = ""
            End If

            ' �����ɃL�[������Ώ���
            If dict.Exists(orgCode) Then
                valueToWrite = dict(orgCode)

                ' �f�o�b�O���O�Ō^�m�F�i��肪����Ώo�́j
                'Debug.Print "dict(" & orgCode & ") �̌^: " & TypeName(valueToWrite)

                If Not IsNull(valueToWrite) Then
                    If Len(Trim(valueToWrite & "")) > 0 Then
                        rsTarget.Edit
                        rsTarget!�{�H�Ǌ��g�D�R�[�h = valueToWrite  ' ������Ƃ��ēo�^�ς�
                        rsTarget.Update
                    Else
                        Debug.Print "�󕶎��������ɂ�F" & orgCode
                    End If
                Else
                    Debug.Print "Null�l�������ɂ�F" & orgCode
                End If
            Else
                ' ��v���Ȃ��ꍇ�̓G���[�L�^
                rsError.AddNew
                rsError!�ǉ��H������ = rsTarget!�ǉ��H������
                rsError!�{�H�S���g�D�R�[�h = rsTarget!�{�H�S���g�D�R�[�h
                rsError!�{�H�S���g�D�� = rsTarget!�{�H�S���g�D��
                rsError.Update
                isErrorExists = True
                Debug.Print "�����Ɍ�����Ȃ������ɂ�F" & orgCode
            End If

            rsTarget.MoveNext
        Loop
    End If

    ' �N���[���A�b�v
    rsTarget.Close
    rsError.Close
    Set rsTarget = Nothing
    Set rsError = Nothing
    Set db = Nothing
    Set dict = Nothing

    ' �G���[����������ʒm
    If isErrorExists Then
        MsgBox "���o�^��Ə�������܂�", vbExclamation, "�G���["
    End If

End Sub


'�e�[�u���FIcube�̃��R�[�h�N���A
Public Sub mod_icubeClear1()
    On Error GoTo ErrHandler

    ' �f�[�^�x�[�X�I�u�W�F�N�g�̐錾
    Dim db As DAO.Database
    Dim sql As String
    
    ' �f�[�^�x�[�X���擾
    Set db = CurrentDb()
    
    ' ���R�[�h��S�폜����SQL
    sql = "DELETE * FROM Icube_"
    
    ' �N�G���̎��s
    db.Execute sql, dbFailOnError
    
    ' ���������̃��b�Z�[�W
    MsgBox "�e�[�u���wIcube_�x�̃��R�[�h��S�ăN���A���܂����I", vbInformation

ExitProcedure:
    ' �㏈��
    On Error Resume Next
    Set db = Nothing
    Exit Sub

ErrHandler:
    ' �G���[���̏���
    MsgBox "�G���[���������܂���: " & Err.description, vbCritical
    Resume ExitProcedure
End Sub



'�e�[�u��Icube�֎{�H�Ǌ��g�D���̋L��
Public Sub mod_icube_copy5()

    Dim db As DAO.Database
    Dim rsSource As DAO.Recordset
    Dim rsTarget As DAO.Recordset
    Dim strSQLSource As String
    Dim strSQLTarget As String
    Dim dict As Object

    ' �f�[�^�x�[�X
    Set db = CurrentDb()
    
    ' �f�[�^�擾SQL
    strSQLSource = "SELECT �{�H�Ǌ��g�D�R�[�h, �{�H�Ǌ��g�D�� FROM tb_�Ǌ���Ə�_RN���P�v��Ə�3;"
    strSQLTarget = "SELECT �{�H�Ǌ��g�D�R�[�h, �{�H�Ǌ��g�D�� FROM Icube_;"

    ' Dictionary �쐬
    Set dict = CreateObject("Scripting.Dictionary")

    ' �Q�ƌ����J���Ď����Ɋi�[
    Set rsSource = db.OpenRecordset(strSQLSource, dbOpenSnapshot)
    Do While Not rsSource.EOF
        dict(Trim(CStr(rsSource!�{�H�Ǌ��g�D�R�[�h))) = rsSource!�{�H�Ǌ��g�D��
        rsSource.MoveNext
    Loop
    rsSource.Close
    Set rsSource = Nothing

    ' �Q�Ɛ���J���čX�V����
    Set rsTarget = db.OpenRecordset(strSQLTarget, dbOpenDynaset)
    If Not rsTarget.EOF Then
        rsTarget.MoveFirst
        Do While Not rsTarget.EOF
            Dim key As String
            key = Trim(CStr(rsTarget!�{�H�Ǌ��g�D�R�[�h))

            rsTarget.Edit
            If dict.Exists(key) Then
                rsTarget!�{�H�Ǌ��g�D�� = dict(key)
                'Debug.Print "�X�V: " & key & " => " & dict(key)
            Else
                rsTarget!�{�H�Ǌ��g�D�� = Null
                Debug.Print "��v�Ȃ�: " & key
            End If
            rsTarget.Update

            rsTarget.MoveNext
        Loop
    End If

    rsTarget.Close
    Set rsTarget = Nothing
    Set db = Nothing
    Set dict = Nothing

    'MsgBox "Icube_ �ւ̓]�ʏ��������������ɂ�", vbInformation

End Sub

'=================================================
' �T�u���[�`���� : mod_icube_dateCnv_1
' ����   : Icube_ �̔N���e�L�X�g����e���t���ڂ��X�V����ɂ�
'=================================================
Public Sub mod_icube_dateCnv_1()
    On Error GoTo EH

    ' --- ������ ---
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Dim dateMath As New com_clsDateMath
    Dim rawText As String

    ' --- �f�[�^�擾 ---
    Set rs = db.OpenRecordset( _
        "SELECT No, [�f�[�^�N���i�󒍌v��N���j], �󒍔N�x, �󒍊�, ��Q, �󒍌�, �󒍌v���_���t�^ " & _
        "FROM Icube_", dbOpenDynaset)

    ' --- ���R�[�h�����݂���ꍇ�̂ݏ��� ---
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
        Do While Not rs.EOF
            rawText = Nz(rs![�f�[�^�N���i�󒍌v��N���j], "")

            ' --- ���t��������Z�b�g����Ǝ����ŉ�͂����ɂ� ---
            dateMath.rawValue = rawText

            If dateMath.IsValid Then
                rs.Edit
                rs!�󒍔N�x = dateMath.GetFiscalYear
                rs!�󒍊� = dateMath.GetPeriod
                rs!��Q = dateMath.GetQuarter
                rs!�󒍌� = dateMath.GetMonth
                rs!�󒍌v���_���t�^ = dateMath.GetDateValue
                rs.Update
            Else
                Debug.Print "�������ȔN���F" & rawText & " (No: " & rs!No & ")"
            End If

            rs.MoveNext
        Loop
    End If

    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Set dateMath = Nothing

    'MsgBox "Icube_ �e�[�u���̓��t�ϊ������������ɂ�I", vbInformation
    Exit Sub

' --- �G���[�n���h�����O ---
EH:
    MsgBox "�G���[�����������ɂ�F" & vbCrLf & Err.description, vbCritical
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Set dateMath = Nothing
End Sub


'=================================================
' �T�u���[�`���� : mod_icube_dateCnv_2
' ����   : �����N�����i�}�ԒP�ʁj���犮�H�N�x�E���EQ�E���E���t�^���ĎZ�o���čX�V����
'=================================================
Public Sub mod_icube_dateCnv_2()
    On Error GoTo Err_Handler

    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Dim dateMath As New com_clsDateMath
    Dim rawText As String

    Set rs = db.OpenRecordset( _
        "SELECT No, [�����N�����i�}�ԒP�ʁj], ���H�N�x, ���H��, ���HQ, ���H��, ���H��_���t�^ " & _
        "FROM Icube_", dbOpenDynaset)

    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
        Do While Not rs.EOF
            rawText = Nz(rs![�����N�����i�}�ԒP�ʁj], "")
            dateMath.rawValue = rawText

            If dateMath.IsValid Then
                rs.Edit
                rs!���H�N�x = dateMath.GetFiscalYear
                rs!���H�� = dateMath.GetPeriod
                rs!���HQ = dateMath.GetQuarter
                rs!���H�� = dateMath.GetMonth
                rs!���H��_���t�^ = dateMath.GetDateValue
                rs.Update
            Else
                Debug.Print "�������Ȋ����N�����F" & rawText & " (No: " & rs!No & ")"
            End If

            rs.MoveNext
        Loop
    End If

    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Set dateMath = Nothing

    'MsgBox "���H���f�[�^�̍X�V�����������ɂ�I", vbInformation
    Exit Sub

Err_Handler:
    MsgBox "�G���[�����������ɂ�F" & vbCrLf & Err.description, vbExclamation
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Set dateMath = Nothing
End Sub



'�}�ԍH���R�[�h�̋L��(�H���R�[�h�Ǝ}�Ԃ̘A��)
Public Sub mod_icube_merge1()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim tableName As String
    Dim primaryKeyField As String
    Dim field1 As String
    Dim field2 As String
    Dim targetField As String
    Dim combinedValue As String

    ' �����Ώۏ��
    tableName = "Icube_"
    primaryKeyField = "No"
    field1 = "�H���R�[�h"
    field2 = "�H���}��"
    targetField = "�}�ԍH���R�[�h"

    ' �f�[�^�x�[�X�ƃ��R�[�h�Z�b�g���擾
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT * FROM " & tableName, dbOpenDynaset)

    ' ���R�[�h�����[�v
    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            ' �A���t�B�[���h1�ƘA���t�B�[���h2�̒l���擾
            Dim value1 As Variant
            Dim value2 As Variant

            value1 = rs.Fields(field1).value
            value2 = rs.Fields(field2).value

            ' Null�l����������i�󕶎��ɒu���j
            If IsNull(value1) Then value1 = ""
            If IsNull(value2) Then value2 = ""

            ' �t�B�[���h�l��A��
            combinedValue = value1 & "-" & value2

            ' �A�����ʂ��L���t�B�[���h�ɍX�V
            rs.Edit
            rs.Fields(targetField).value = combinedValue
            rs.Update

            rs.MoveNext
        Loop
    End If

    ' �N���[���A�b�v
    rs.Close
    Set rs = Nothing
    Set db = Nothing

    'MsgBox "�t�B�[���h�l�̘A���������������܂����B", vbInformation
End Sub


'��{�R�[�h���u�����N�̎��A�H���R�[�h��]��
Public Sub mod_icube_copy1()
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim transferCount As Long
    transferCount = 0 ' �]�ʌ�����������

    ' �f�[�^�x�[�X���擾
    Set db = CurrentDb()

    ' �Ώۃ��R�[�h���擾
    strSQL = "SELECT No, �H���R�[�h, ��{�H���R�[�h FROM Icube_"
    Set rs = db.OpenRecordset(strSQL, dbOpenDynaset)

    ' ���R�[�h�����[�v
    Do While Not rs.EOF
        ' �]�ʐ�t�B�[���h���u�����N�܂���Null�̏ꍇ
        If IsNull(rs!��{�H���R�[�h) Or rs!��{�H���R�[�h = "N/A" Then
            rs.Edit
            rs!��{�H���R�[�h = rs!�H���R�[�h
            rs.Update
            transferCount = transferCount + 1 ' �]�ʌ������J�E���g
        End If
        rs.MoveNext
    Loop

    ' ���ʂ�\��
    'MsgBox "�]�ʏ������������܂����B�]�ʌ���: " & transferCount & " ��", vbInformation

CleanUp:
    ' ��n��
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub

ErrHandler:
    MsgBox "�G���[���������܂���: " & Err.description, vbExclamation
    Resume CleanUp
End Sub

'��{�H�����̂������ꍇ�ɍH�����[����]��
Public Sub mod_icube_copy2()
    ' ��`
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    
    ' �f�[�^�x�[�X�̎Q��
    Set db = CurrentDb
    
    ' �Ώۃ��R�[�h��SQL���`
    strSQL = "SELECT No, �H�����[��, ��{�H������ " & _
             "FROM Icube_ " & _
             "WHERE [��{�H������] IS NULL OR [��{�H������] = '' OR [��{�H������] = 'N/A';"
    
    ' ���R�[�h�Z�b�g���J��
    Set rs = db.OpenRecordset(strSQL, dbOpenDynaset)
    
    ' ���R�[�h�����݂���ꍇ
    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            ' �]�ʏ���
            If Not IsNull(rs!�H�����[��) And rs!�H�����[�� <> "" Then
                rs.Edit
                rs!��{�H������ = rs!�H�����[��
                rs.Update
            End If
            rs.MoveNext
        Loop
    End If
    
    ' �N���[���A�b�v
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    ' �����������b�Z�[�W
    'MsgBox "�H�����[���̓]�ʏ������������܂����B", vbInformation
End Sub



'��{�H�����̂���ꌏ�H������
Public Sub mod_icube_input1()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim targetTable As String
    Dim targetFieldCondition As String
    Dim targetFieldUpdate As String
    Dim conditionValues As Variant
    Dim updateValue1 As String
    Dim updateValue2 As String
    
    ' �����Ώۏ��
    targetTable = "Icube_"
    targetFieldCondition = "��{�H������"
    targetFieldUpdate = "�ꌏ�H������"
    
    ' �����l
    conditionValues = Array("�P�Q���H��", "�P�R���H��", "�P�p", "�Q�p", "�R�p", "�S�p")
    
    ' �L���l
    updateValue1 = "�����H��"
    updateValue2 = "�ꌏ�H��"
    
    ' �f�[�^�x�[�X�̎Q��
    Set db = CurrentDb()
    
    ' �Ώۃe�[�u���̃f�[�^���擾
    strSQL = "SELECT No, " & targetFieldCondition & ", " & targetFieldUpdate & " FROM " & targetTable
    Set rs = db.OpenRecordset(strSQL, dbOpenDynaset)
    
    ' ���R�[�h�̃��[�v����
    If Not rs.EOF Then
        Do While Not rs.EOF
            Dim currentCondition As String
            Dim shouldUpdate As Boolean
            Dim i As Integer
            
            currentCondition = Nz(rs.Fields(targetFieldCondition).value, "")
            shouldUpdate = False
            
            ' �����l�ɊY�����邩�m�F
            For i = LBound(conditionValues) To UBound(conditionValues)
                If InStr(1, currentCondition, conditionValues(i), vbTextCompare) > 0 Then
                    shouldUpdate = True
                    Exit For
                End If
            Next i
            
            ' �t�B�[���h�l���X�V
            If shouldUpdate Then
                rs.Edit
                rs.Fields(targetFieldUpdate).value = updateValue1
                rs.Update
            Else
                rs.Edit
                rs.Fields(targetFieldUpdate).value = updateValue2
                rs.Update
            End If
            
            rs.MoveNext
        Loop
    End If
    
    ' �N���[���A�b�v
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    'MsgBox "�������������܂����B", vbInformation
End Sub

'�G���[�l�u����
 ' �G���[�l�u�����Ώۃe�[�u���� "�G���[�l�u�����L��=TRUE"�����{�Ώ�
Public Sub mod_icube_err_1()
    Dim db As DAO.Database
    Dim rsTarget As DAO.Recordset
    Dim rsIcube As DAO.Recordset
    
    Dim clsCom As New cls_err
    Dim targetFieldName As String
    Dim fieldType As Integer
    Dim oldValue As Variant
    Dim newValue As Variant
    
    Set db = CurrentDb
    
    ' �G���[�l�u�����Ώۃe�[�u���� "�G���[�l�u�����L��=TRUE" �̃t�B�[���h���擾
    Set rsTarget = db.OpenRecordset( _
        "SELECT [�t�B�[���h��] " & _
        "FROM t_�G���[�l�u�����Ώ� " & _
        "WHERE [�G���[�l�u�����L��] = TRUE" _
    )
    
    If rsTarget.EOF Then
        MsgBox "�u�����Ώۂ̃t�B�[���h������܂���B", vbInformation
        GoTo CleanUp
    End If
    
    ' t_�G���[�l�u�����Ώ� �����R�[�h�P�ʂő���
    Do While Not rsTarget.EOF
        
        targetFieldName = rsTarget!�t�B�[���h��
        
        ' Icube_ �e�[�u���̊Y���t�B�[���h���擾
        ' ���K�v�ɉ����ĕK�v�ȃt�B�[���h��SELECT��Ŏw�肵�ĉ�����
        Set rsIcube = db.OpenRecordset( _
            "SELECT [No], [" & targetFieldName & "] " & _
            "FROM Icube_" _
        )
        
        ' Icube_�̑S���R�[�h�����[�v
        Do While Not rsIcube.EOF
            
            oldValue = rsIcube.Fields(targetFieldName).value
            fieldType = rsIcube.Fields(targetFieldName).Type
            newValue = clsCom.GetDefaultValue(fieldType, oldValue)
            
            ' �l���ύX�����ꍇ�����X�V
            If Nz(newValue, "") <> Nz(oldValue, "") Then
                rsIcube.Edit
                rsIcube.Fields(targetFieldName).value = newValue
                rsIcube.Update
            End If
            
            rsIcube.MoveNext
        Loop
        
        rsIcube.Close
        Set rsIcube = Nothing
        
        rsTarget.MoveNext
    Loop

CleanUp:
    If Not rsTarget Is Nothing Then
        rsTarget.Close
        Set rsTarget = Nothing
    End If
    
    If Not rsIcube Is Nothing Then
        rsIcube.Close
        Set rsIcube = Nothing
    End If

    Set db = Nothing
    
    MsgBox "�G���[�l�u�����������������܂����B", vbInformation
End Sub
