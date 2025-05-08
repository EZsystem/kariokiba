Attribute VB_Name = "mod_IcubeImport4"
'-------------------------------------
' Module: mod_IcubeImport4
' 説明　：クリーニングテーブルからIcube_へ転写（型変換付き、いつもの構成）
' 作成日：2025/04/30
' 更新日：-
'-------------------------------------

Option Compare Database
Option Explicit


'=================================================
' 処理名 : Run_TransferToIcube
' 説明   : 仮テーブル（tbl_Temp_Icube_Import）から
'        : 本テーブル（Icube_）へ転写する処理を実行する
'        : 転写条件・型変換ルールはマスタ（tbl_xl_IcubeColSetting）に従う
'=================================================
Public Sub Run_TransferToIcube()
    Dim tempTable As String
    Dim settingTable As String

    ' --- 転写元テーブル（仮） ---
    tempTable = "tbl_Temp_Icube_Import"

    ' --- 転写ルール定義テーブル（マスタ） ---
    settingTable = "tbl_xl_IcubeColSetting"

    ' --- 転写処理の呼び出し（仮 → 本テーブル Icube_） ---
    Call TransferToIcube_StandardStyle(tempTable, settingTable)
End Sub




'=================================================
' 処理名 : TransferToIcube_StandardStyle
' 説明   : 仮テーブルから本テーブルIcube_へ安全に転写（型変換・スキップあり）
' 引数   : 仮テーブル名、マスタテーブル名
'=================================================
Public Sub TransferToIcube_StandardStyle( _
    ByVal tempTable As String, _
    ByVal settingTable As String)

    On Error GoTo EH

    Dim db As DAO.Database: Set db = CurrentDb
    Dim rsSource As DAO.Recordset: Set rsSource = db.OpenRecordset(tempTable, dbOpenSnapshot)
    Dim rsTarget As DAO.Recordset: Set rsTarget = db.OpenRecordset("Icube_", dbOpenDynaset)

    ' --- 1. マスタを辞書へ読み込み（フィールド名 → 型名） ---
    Dim rsMap As DAO.Recordset
    Set rsMap = db.OpenRecordset( _
        "SELECT [タイトル名_置換え後], [データ型] " & _
        "FROM " & settingTable & " " & _
        "WHERE Nz([取込フラグ], False) = True", dbOpenSnapshot)

    Dim fieldTypeMap As Object: Set fieldTypeMap = CreateObject("Scripting.Dictionary")
    Dim dataCleaner As New acc_clsDataCleaner

    Do Until rsMap.EOF
        Dim fname As String: fname = Trim(rsMap("タイトル名_置換え後"))
        Dim jptype As String: jptype = rsMap("データ型")
        If Not fieldTypeMap.Exists(fname) Then
            fieldTypeMap.Add fname, jptype
        End If
        rsMap.MoveNext
    Loop
    rsMap.Close

    '=================================================
    ' スキップ条件を辞書に読み込む（フィールド名 → 値集合）
    '=================================================
    Dim rsSkip As DAO.Recordset
    Dim skipDict As Object: Set skipDict = CreateObject("Scripting.Dictionary")

    Set rsSkip = db.OpenRecordset("tbl_xl_IcubeRowSkip", dbOpenSnapshot)
    Do Until rsSkip.EOF
        Dim fld As String: fld = Trim(rsSkip("対象フィールド名"))
        Dim val As String: val = Trim(rsSkip("削除対象値"))
        If Not skipDict.Exists(fld) Then
            skipDict.Add fld, CreateObject("Scripting.Dictionary")
        End If
        skipDict(fld)(val) = True
        rsSkip.MoveNext
    Loop
    rsSkip.Close

    '=================================================
    ' 仮テーブルから本テーブルへ1件ずつ転写（スキップ条件考慮）
    '=================================================
    Do Until rsSource.EOF
        ' --- 転写スキップ条件の判定 ---
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

        ' --- レコード追加開始 ---
        rsTarget.AddNew

        Dim fldName As Variant
        For Each fldName In fieldTypeMap.Keys
            ' --- フィールド存在チェック（元／先） ---
            If Not FieldExists(rsSource, CStr(fldName)) Then
                Debug.Print "※転写元に存在しない：" & fldName
                GoTo skipField
            End If
            If Not FieldExists(rsTarget, CStr(fldName)) Then
                Debug.Print "※転写先に存在しない：" & fldName
                GoTo skipField
            End If

            ' --- 型変換処理 ---
            Dim raw As Variant: raw = rsSource(fldName).value
            Dim vbaType As String

            On Error Resume Next
            vbaType = dataCleaner.GetSupportedVBAType(fieldTypeMap(fldName))
            If Err.Number <> 0 Then
                Debug.Print "※未対応型 → スキップ：" & fldName & " (" & fieldTypeMap(fldName) & ")"
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
    'MsgBox "Icube_ テーブルへの転写が完了したにゃ", vbInformation
    Exit Sub

EH:
    MsgBox "【転写エラー】：" & Err.description, vbCritical
    Debug.Print "【転写エラー】：" & Err.description
End Sub


Private Function FieldExists(rs As DAO.Recordset, fieldName As String) As Boolean
    On Error Resume Next
    Dim dummy As Variant: dummy = rs.Fields(fieldName).Name
    FieldExists = (Err.Number = 0)
    Err.Clear
End Function

