Attribute VB_Name = "mod_icubeVal3"
'-------------------------------------
' Module: acc_mod_icubeVal3
' ����  : ����\�Ɋ�Â��ėp�r��敪���C������ɂ�
' �쐬��: 2025/05/07
'-------------------------------------
Option Compare Database
Option Explicit

Public Sub Correct_�p�r��敪()
    Dim db As DAO.Database
    Set db = CurrentDb
    
    Dim rsMain As DAO.Recordset
    Dim rsMap As DAO.Recordset
    
    Set rsMain = db.OpenRecordset("Icube_", dbOpenDynaset)
    Set rsMap = db.OpenRecordset("tbl_�����p�r����\", dbOpenSnapshot)
    
    Dim cleaner As New acc_clsDataCleaner
    Dim �� As String
    Dim ��1 As String, ��2 As String
    
    Do While Not rsMain.EOF
        �� = cleaner.CleanText(rsMain!�p�r��敪)
        
        rsMap.MoveFirst
        Do While Not rsMap.EOF
            If �� = cleaner.CleanText(rsMap!��_�p�r��敪) Then
                ��1 = cleaner.CleanText(rsMap!��_�p�r��敪)
                ��2 = cleaner.CleanText(rsMap!��_�p�r��敪��)
                
                rsMain.Edit
                rsMain!s�p�r��敪 = ��1
                rsMain!s�p�r��敪�� = ��2
                rsMain.Update
                Exit Do
            End If
            rsMap.MoveNext
        Loop
        
        rsMain.MoveNext
    Loop
    
    rsMain.Close
    rsMap.Close
    MsgBox "�]�ʏ��������������ɂ�", vbInformation
End Sub

