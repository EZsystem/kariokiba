Attribute VB_Name = "mod_icubeVal4"
'-------------------------------------
' Module: acc_mod_HŽ–‹àŠz‹æ•ª“]ŽÊ
' à–¾  : HŽ–‰¿Ši‚ðŒ³‚É‹æ•ª‚ð“]ŽÊ‚·‚é‚É‚á
'-------------------------------------
Option Compare Database
Option Explicit

Public Sub assign_priceCategory()
    Dim db As DAO.Database
    Set db = CurrentDb

    Dim rsMain As DAO.Recordset
    Dim rsZone As DAO.Recordset

    Set rsMain = db.OpenRecordset("Icube_", dbOpenDynaset)
    Set rsZone = db.OpenRecordset("tbl_HŽ–‹àŠz‹æ•ª•\", dbOpenSnapshot)

    Dim cleaner As New acc_clsDataCleaner
    Dim targetVal As Currency
    Dim minVal As Currency, maxVal As Currency

    Do While Not rsMain.EOF
        targetVal = cleaner.CleanNumber(rsMain!HŽ–‰¿Ši)

        rsZone.MoveFirst
        Do While Not rsZone.EOF
            minVal = cleaner.CleanNumber(rsZone!Å¬‹àŠz)
            maxVal = cleaner.CleanNumber(rsZone!Å‘å‹àŠz)

            If targetVal >= minVal And targetVal <= maxVal Then
                rsMain.Edit
                rsMain!HŽ–‹àŠz‹æ•ªƒR[ƒh = rsZone!HŽ–‹àŠz‹æ•ªƒR[ƒh
                rsMain!HŽ–‹àŠz‹æ•ª–¼ = rsZone!HŽ–‹àŠz‹æ•ª–¼
                rsMain!HŽ–‹àŠzƒ}ƒCƒiƒX”»’è = rsZone!HŽ–‹àŠzƒ}ƒCƒiƒX”»’è
                rsMain.Update
                Exit Do
            End If

            rsZone.MoveNext
        Loop

        rsMain.MoveNext
    Loop

    rsMain.Close
    rsZone.Close
    'MsgBox "HŽ–‹àŠz‹æ•ª‚Ì“]ŽÊ‚ªŠ®—¹‚µ‚½‚É‚á", vbInformation
End Sub

