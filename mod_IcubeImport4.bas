Attribute VB_Name = "mod_IcubeImport4"
'-------------------------------------
' Module: mod_IcubeImport4
' �����@�F�N���[�j���O�e�[�u������Icube_�֓]�ʁi�^�ϊ��t���A�����̍\���j
' �쐬���F2025/04/30
' �X�V���F-
'-------------------------------------

Option Compare Database
Option Explicit


'=================================================
' ������ : Run_TransferToIcube
' ����   : ���e�[�u���itbl_Temp_Icube_Import�j����
'        : �{�e�[�u���iIcube_�j�֓]�ʂ��鏈�������s����
'        : �]�ʏ����E�^�ϊ����[���̓}�X�^�itbl_xl_IcubeColSetting�j�ɏ]��
'=================================================
Public Sub Run_TransferToIcube()
    Dim tempTable As String
    Dim settingTable As String

    ' --- �]�ʌ��e�[�u���i���j ---
    tempTable = "tbl_Temp_Icube_Import"

    ' --- �]�ʃ��[����`�e�[�u���i�}�X�^�j ---
    settingTable = "tbl_xl_IcubeColSetting"

    ' --- �]�ʏ����̌Ăяo���i�� �� �{�e�[�u�� Icube_�j ---
    Call TransferToIcube_StandardStyle(tempTable, settingTable)
End Sub




'=================================================
' ������ : TransferToIcube_StandardStyle
' ����   : ���e�[�u������{�e�[�u��Icube_�ֈ��S�ɓ]�ʁi�^�ϊ��E�X�L�b�v����j
' ����   : ���e�[�u�����A�}�X�^�e�[�u����
'=================================================
Public Sub TransferToIcube_StandardStyle( _
    ByVal tempTable As String, _
    ByVal settingTable As String)

    On Error GoTo EH

    Dim db As DAO.Database: Set db = CurrentDb
    Dim rsSource As DAO.Recordset: Set rsSource = db.OpenRecordset(tempTable, dbOpenSnapshot)
    Dim rsTarget As DAO.Recordset: Set rsTarget = db.OpenRecordset("Icube_", dbOpenDynaset)

    ' --- 1. �}�X�^�������֓ǂݍ��݁i�t�B�[���h�� �� �^���j ---
    Dim rsMap As DAO.Recordset
    Set rsMap = db.OpenRecordset( _
        "SELECT [�^�C�g����_�u������], [�f�[�^�^] " & _
        "FROM " & settingTable & " " & _
        "WHERE Nz([�捞�t���O], False) = True", dbOpenSnapshot)

    Dim fieldTypeMap As Object: Set fieldTypeMap = CreateObject("Scripting.Dictionary")
    Dim dataCleaner As New acc_clsDataCleaner

    Do Until rsMap.EOF
        Dim fname As String: fname = Trim(rsMap("�^�C�g����_�u������"))
        Dim jptype As String: jptype = rsMap("�f�[�^�^")
        If Not fieldTypeMap.Exists(fname) Then
            fieldTypeMap.Add fname, jptype
        End If
        rsMap.MoveNext
    Loop
    rsMap.Close

    '=================================================
    ' �X�L�b�v�����������ɓǂݍ��ށi�t�B�[���h�� �� �l�W���j
    '=================================================
    Dim rsSkip As DAO.Recordset
    Dim skipDict As Object: Set skipDict = CreateObject("Scripting.Dictionary")

    Set rsSkip = db.OpenRecordset("tbl_xl_IcubeRowSkip", dbOpenSnapshot)
    Do Until rsSkip.EOF
        Dim fld As String: fld = Trim(rsSkip("�Ώۃt�B�[���h��"))
        Dim val As String: val = Trim(rsSkip("�폜�Ώےl"))
        If Not skipDict.Exists(fld) Then
            skipDict.Add fld, CreateObject("Scripting.Dictionary")
        End If
        skipDict(fld)(val) = True
        rsSkip.MoveNext
    Loop
    rsSkip.Close

    '=================================================
    ' ���e�[�u������{�e�[�u����1�����]�ʁi�X�L�b�v�����l���j
    '=================================================
    Do Until rsSource.EOF
        ' --- �]�ʃX�L�b�v�����̔��� ---
        Dim shouldSkip As Boolean: shouldSkip = False
        Dim fldSkip As Variant
    For Each fldSkip In skipDict.Keys
        Dim fldNameSkip As String
        fldNameSkip = CStr(fldSkip)
    
        If FieldExists(rsSource, fldNameSkip) Then
            Dim sourceVal As String
            sourceVal = Trim(Nz(rsSource(fldNameSkip).value, ""))
            If skipDict(fldNameSkip).Exists(sourceVal) Then
                shouldSkip = True
                Exit For
            End If
        End If
    Next fldSkip

        If shouldSkip Then
            rsSource.MoveNext
            GoTo nextRecord
        End If

        ' --- ���R�[�h�ǉ��J�n ---
        rsTarget.AddNew

        Dim fldName As Variant
        For Each fldName In fieldTypeMap.Keys
            ' --- �t�B�[���h���݃`�F�b�N�i���^��j ---
            If Not FieldExists(rsSource, CStr(fldName)) Then
                Debug.Print "���]�ʌ��ɑ��݂��Ȃ��F" & fldName
                GoTo skipField
            End If
            If Not FieldExists(rsTarget, CStr(fldName)) Then
                Debug.Print "���]�ʐ�ɑ��݂��Ȃ��F" & fldName
                GoTo skipField
            End If

            ' --- �^�ϊ����� ---
            Dim raw As Variant: raw = rsSource(fldName).value
            Dim vbaType As String

            On Error Resume Next
            vbaType = dataCleaner.GetSupportedVBAType(fieldTypeMap(fldName))
            If Err.Number <> 0 Then
                Debug.Print "�����Ή��^ �� �X�L�b�v�F" & fldName & " (" & fieldTypeMap(fldName) & ")"
                Err.Clear
                GoTo skipField
            End If
            On Error GoTo EH

            Select Case vbaType
                Case "String":   val = dataCleaner.TextToString(raw)
                Case "Long":     val = dataCleaner.TextToLong(raw)
                Case "Integer":  val = dataCleaner.TextToInteger(raw)
                Case "Single":   val = dataCleaner.TextToSingle(raw)
                Case "Double":   val = dataCleaner.TextToDouble(raw)
                Case "Currency": val = dataCleaner.TextToCurrency(raw)
                Case "Date":     val = dataCleaner.TextToDate(raw)
                Case "Boolean":  val = dataCleaner.TextToBoolean(raw)
                Case Else:       val = dataCleaner.CleanText(raw)
            End Select

            rsTarget(fldName).value = val

skipField:
        Next fldName

        rsTarget.Update
nextRecord:
        rsSource.MoveNext
    Loop

    rsSource.Close
    rsTarget.Close
    'MsgBox "Icube_ �e�[�u���ւ̓]�ʂ����������ɂ�", vbInformation
    Exit Sub

EH:
    MsgBox "�y�]�ʃG���[�z�F" & Err.description, vbCritical
    Debug.Print "�y�]�ʃG���[�z�F" & Err.description
End Sub


Private Function FieldExists(rs As DAO.Recordset, fieldName As String) As Boolean
    On Error Resume Next
    Dim dummy As Variant: dummy = rs.Fields(fieldName).Name
    FieldExists = (Err.Number = 0)
    Err.Clear
End Function

