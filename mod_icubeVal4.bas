Attribute VB_Name = "mod_icubeVal4"
'-------------------------------------
' Module: acc_mod_工事金額区分転写
' 説明  : 工事価格を元に区分を転写するにゃ
'-------------------------------------
Option Compare Database
Option Explicit

Public Sub assign_priceCategory()
    Dim db As DAO.Database
    Set db = CurrentDb

    Dim rsMain As DAO.Recordset
    Dim rsZone As DAO.Recordset

    Set rsMain = db.OpenRecordset("Icube_", dbOpenDynaset)
    Set rsZone = db.OpenRecordset("tbl_工事金額区分表", dbOpenSnapshot)

    Dim cleaner As New acc_clsDataCleaner
    Dim targetVal As Currency
    Dim minVal As Currency, maxVal As Currency

    Do While Not rsMain.EOF
        targetVal = cleaner.CleanNumber(rsMain!工事価格)

        rsZone.MoveFirst
        Do While Not rsZone.EOF
            minVal = cleaner.CleanNumber(rsZone!最小金額)
            maxVal = cleaner.CleanNumber(rsZone!最大金額)

            If targetVal >= minVal And targetVal <= maxVal Then
                rsMain.Edit
                rsMain!工事金額区分コード = rsZone!工事金額区分コード
                rsMain!工事金額区分名 = rsZone!工事金額区分名
                rsMain!工事金額マイナス判定 = rsZone!工事金額マイナス判定
                rsMain.Update
                Exit Do
            End If

            rsZone.MoveNext
        Loop

        rsMain.MoveNext
    Loop

    rsMain.Close
    rsZone.Close
    'MsgBox "工事金額区分の転写が完了したにゃ", vbInformation
End Sub

