Attribute VB_Name = "mod_icubeVal3"
'-------------------------------------
' Module: acc_mod_icubeVal3
' 説明  : 正誤表に基づいて用途大区分を修正するにゃ
' 作成日: 2025/05/07
'-------------------------------------
Option Compare Database
Option Explicit

Public Sub Correct_用途大区分()
    Dim db As DAO.Database
    Set db = CurrentDb
    
    Dim rsMain As DAO.Recordset
    Dim rsMap As DAO.Recordset
    
    Set rsMain = db.OpenRecordset("Icube_", dbOpenDynaset)
    Set rsMap = db.OpenRecordset("tbl_建物用途正誤表", dbOpenSnapshot)
    
    Dim cleaner As New acc_clsDataCleaner
    Dim 誤 As String
    Dim 正1 As String, 正2 As String
    
    Do While Not rsMain.EOF
        誤 = cleaner.CleanText(rsMain!用途大区分)
        
        rsMap.MoveFirst
        Do While Not rsMap.EOF
            If 誤 = cleaner.CleanText(rsMap!誤_用途大区分) Then
                正1 = cleaner.CleanText(rsMap!正_用途大区分)
                正2 = cleaner.CleanText(rsMap!正_用途大区分名)
                
                rsMain.Edit
                rsMain!s用途大区分 = 正1
                rsMain!s用途大区分名 = 正2
                rsMain.Update
                Exit Do
            End If
            rsMap.MoveNext
        Loop
        
        rsMain.MoveNext
    Loop
    
    rsMain.Close
    rsMap.Close
    MsgBox "転写処理が完了したにゃ", vbInformation
End Sub

