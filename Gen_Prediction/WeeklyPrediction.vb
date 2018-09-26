Imports System.Data.SqlClient
Imports System.Globalization
Module WeeklyPrediciton
    Sub Main()
        WeeklyPredicton("vcdubai@gmail.com", "3")
    End Sub
    Sub WeeklyPredicton(ByVal UID As String, ByVal HID As String)
        Dim connection As SqlConnection = New SqlConnection("data source=49.50.103.132;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=False;User Id=sa;password=pSI)TA1t0K[)")
        connection.Open()
        Dim RowSet As New DataSet()
        Dim cmd As New SqlCommand($"SELECT TOP 1 RQUSERID, RQHID, RQSDATE, RQEDATE, RQLONGTIDUE, RQLONGTITUDEEW, RQLATITUDE, RQLATITUDENS, RQDST, TIMESTAMP
                               FROM HREQUEST WHERE HREQUEST.REQCAT = '7' AND HREQUEST.RQUNREAD = 'Y' AND RQUSERID = '" + "922@gmail.com" + "' AND RQHID = '" + "3" + "';", connection) '->->->->->" + UID + "' AND RQUSERID = '" + HID + "
        Dim da As New SqlDataAdapter(cmd)
        da.Fill(RowSet)
        Dim RQSDATE = CType(RowSet.Tables(0).Rows(0)(2), DateTime).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
        Dim RQEDATE = CType(RowSet.Tables(0).Rows(0)(3), DateTime).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
        Dim RQLONGTIDUE = RowSet.Tables(0).Rows(0)(4).ToString().Trim
        Dim RQLONGTITUDEEW = RowSet.Tables(0).Rows(0)(5).ToString().Trim
        Dim RQLATITUDE = RowSet.Tables(0).Rows(0)(6).ToString().Trim
        Dim RQLATITUDENS = RowSet.Tables(0).Rows(0)(7).ToString().Trim
        Dim RQDST = RowSet.Tables(0).Rows(0)(8).ToString().Trim
        Dim TIMESTAMP = RowSet.Tables(0).Rows(0)(9).ToString().Trim

    End Sub
End Module
