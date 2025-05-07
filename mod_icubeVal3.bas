Attribute VB_Name = "mod_icubeVal3"
'-------------------------------------
' Module: acc_mod_icubeVal3
' à–¾  : ³Œë•\‚ÉŠî‚Ã‚¢‚Ä—p“r‘å‹æ•ª‚ğC³‚·‚é‚É‚á
' ì¬“ú: 2025/05/07
'-------------------------------------
Option Compare Database
Option Explicit

Public Sub Correct_—p“r‘å‹æ•ª()
    Dim db As DAO.Database
    Set db = CurrentDb
    
    Dim rsMain As DAO.Recordset
    Dim rsMap As DAO.Recordset
    
    Set rsMain = db.OpenRecordset("Icube_", dbOpenDynaset)
    Set rsMap = db.OpenRecordset("tbl_Œš•¨—p“r³Œë•\", dbOpenSnapshot)
    
    Dim cleaner As New acc_clsDataCleaner
    Dim Œë As String
    Dim ³1 As String, ³2 As String
    
    Do While Not rsMain.EOF
        Œë = cleaner.CleanText(rsMain!—p“r‘å‹æ•ª)
        
        rsMap.MoveFirst
        Do While Not rsMap.EOF
            If Œë = cleaner.CleanText(rsMap!Œë_—p“r‘å‹æ•ª) Then
                ³1 = cleaner.CleanText(rsMap!³_—p“r‘å‹æ•ª)
                ³2 = cleaner.CleanText(rsMap!³_—p“r‘å‹æ•ª–¼)
                
                rsMain.Edit
                rsMain!s—p“r‘å‹æ•ª = ³1
                rsMain!s—p“r‘å‹æ•ª–¼ = ³2
                rsMain.Update
                Exit Do
            End If
            rsMap.MoveNext
        Loop
        
        rsMain.MoveNext
    Loop
    
    rsMain.Close
    rsMap.Close
    MsgBox "“]Êˆ—‚ªŠ®—¹‚µ‚½‚É‚á", vbInformation
End Sub

