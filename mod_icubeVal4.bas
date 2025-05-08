Attribute VB_Name = "mod_icubeVal4"
'-------------------------------------
' Module: acc_mod_�H�����z�敪�]��
' ����  : �H�����i�����ɋ敪��]�ʂ���ɂ�
'-------------------------------------
Option Compare Database
Option Explicit

Public Sub assign_priceCategory()
    Dim db As DAO.Database
    Set db = CurrentDb

    Dim rsMain As DAO.Recordset
    Dim rsZone As DAO.Recordset

    Set rsMain = db.OpenRecordset("Icube_", dbOpenDynaset)
    Set rsZone = db.OpenRecordset("tbl_�H�����z�敪�\", dbOpenSnapshot)

    Dim cleaner As New acc_clsDataCleaner
    Dim targetVal As Currency
    Dim minVal As Currency, maxVal As Currency

    Do While Not rsMain.EOF
        targetVal = cleaner.CleanNumber(rsMain!�H�����i)

        rsZone.MoveFirst
        Do While Not rsZone.EOF
            minVal = cleaner.CleanNumber(rsZone!�ŏ����z)
            maxVal = cleaner.CleanNumber(rsZone!�ő���z)

            If targetVal >= minVal And targetVal <= maxVal Then
                rsMain.Edit
                rsMain!�H�����z�敪�R�[�h = rsZone!�H�����z�敪�R�[�h
                rsMain!�H�����z�敪�� = rsZone!�H�����z�敪��
                rsMain!�H�����z�}�C�i�X���� = rsZone!�H�����z�}�C�i�X����
                rsMain.Update
                Exit Do
            End If

            rsZone.MoveNext
        Loop

        rsMain.MoveNext
    Loop

    rsMain.Close
    rsZone.Close
    'MsgBox "�H�����z�敪�̓]�ʂ����������ɂ�", vbInformation
End Sub

