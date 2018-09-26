Imports System.Collections.Specialized
Imports System.Data.SqlClient
Imports System.Globalization

Module DailyPrediction

    Sub Main()
        Dim StrPlanet(13) As String
        Dim RowsData As New DataSet()
        Dim Key1Cusp As String = ""
        Dim Key2Cusp As String = ""
        Dim Key3Cusp As String = ""
        Dim Key4Cusp As String = ""
        Dim Key5Cusp As String = ""
        Dim Key6Cusp As String = ""

        'connection.Open()
        ''Dim cmd As New SqlCommand($"SELECT TOP 1 RQUSERID, RQHID, RQSDATE, RQLONGTIDUE, RQLONGTITUDEEW, RQLATITUDE, 
        '                       RQLATITUDENS, RQDST, TIMESTAMP
        '                       FROM HREQUEST
        '                       WHERE HREQUEST.REQCAT = '2'
        '                           AND HREQUEST.RQUNREAD = 'Y';", connection)
        'Dim da As New SqlDataAdapter(cmd)
        'da.Fill(RowsData)
        Dim UID = "vcdubai@gmail.com"
        'Dim UID = RowsData.Tables(0).Rows(0)(0).ToString().Trim
        Dim HID = "3"
        'Dim HID = RowsData.Tables(0).Rows(0)(1).ToString().Trim
        'Dim RQSDATE = CType(RowsData.Tables(0).Rows(0)(2), DateTime).ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture)
        'Dim RQLONGTIDUE = RowsData.Tables(0).Rows(0)(3).ToString().Trim
        'Dim RQLONGTITUDEEW = RowsData.Tables(0).Rows(0)(4).ToString().Trim
        'Dim RQLATITUDE = RowsData.Tables(0).Rows(0)(5).ToString().Trim
        'Dim RQLATITUDENS = RowsData.Tables(0).Rows(0)(6).ToString().Trim
        'Dim RQDST = RowsData.Tables(0).Rows(0)(7).ToString().Trim
        'Dim TIMESTAMP = RowsData.Tables(0).Rows(0)(8).ToString().Trim
        'Dim Year = RQSDATE.Split("-")(0)
        'Dim Month = RQSDATE.Split("-")(1)
        'Dim Day = RQSDATE.Split("-")(2).Substring(0, 2)
        'DatetimeC.Add("Year", "2018")
        'DatetimeC.Add("Month", "9")
        'DatetimeC.Add("Day", "24")
        'DatetimeC.Add("Hour", "00")
        'DatetimeC.Add("Min", "00")
        'DatetimeC.Add("Sec", "00")
        'DatetimeC.Add("mSec", "000")
        'DateTimeB = DatetimeC
        'Dim command As New SqlCommand($"select top 1 HPLACE FROM HMAIN WHERE HUSERID = '" + UID + "' AND HID = '" + HID + "';", connection)
        'Dim daPlace As New SqlDataAdapter(command)
        'Dim GetPlace As New DataSet()
        'daPlace.Fill(GetPlace)
        'Dim HPLACE = GetPlace.Tables(0).Rows(0)(0).ToString().Trim
        'Dim Place = HPLACE.Split("-")(0)
        'Dim State = HPLACE.Split("-")(1)
        'Dim Country = HPLACE.Split("-")(2)
        'Dim lonB = RQLONGTIDUE
        'Dim latB = RQLATITUDE
        'Dim eW = RQLONGTITUDEEW
        'Dim nS = RQLATITUDENS
        'Dim TzB = TIMESTAMP
        'Dim DstB = RQDST
        'PlaceDataT.Add("Place", "Dubai ")
        'PlaceDataT.Add("State", " ")
        'PlaceDataT.Add("Country", " U.A.E.")
        'PlaceDataT.Add("lonB", 55.333333333333336)
        'PlaceDataT.Add("latB", 25.151620635773071)
        'PlaceDataT.Add("eW", "E")
        'PlaceDataT.Add("nS", "N")
        'PlaceDataT.Add("TzB", 4)
        'PlaceDataT.Add("DstB", 0)
        'Dim Horo As New TASystem.TrueAstro
        'Horo.getTransitList(DatetimeC, PlaceDataT, StrPlanet)
        'Dim DayDasa As String = ""
        'PlaceDataB = PlaceDataT
        'DatetimeC.Add("Year", "2018")
        'DatetimeC.Add("Month", "9")
        'DatetimeC.Add("Day", "24")
        'DateTimeB.Add("Hour", "23")
        'DateTimeB.Add("Min", "59")
        'DateTimeB.Add("Sec", "59")
        'DateTimeB.Add("mSec", "999")
        'Horo.getBirthDasaFordateTime(DatetimeC, DateTimeB, PlaceDataB, DayDasa)
        'UpdateHPREDICTION(UID, HID, StrPlanet, DayDasa)
        For hh As Integer = 0 To 23
            For mm As Integer = 0 To 59
                For ss As Integer = 0 To 59
                    For mmm As Integer = 0 To 950 Step 50
                        Dim PlaceDataB As New NameValueCollection
                        Dim DatetimeC As New NameValueCollection
                        Dim PlaceDataT As New NameValueCollection
                        Dim DateTimeB As New NameValueCollection
                        DatetimeC.Add("Year", "2018")
                        DatetimeC.Add("Month", "9")
                        DatetimeC.Add("Day", "24")
                        DatetimeC.Add("Hour", hh.ToString("D2"))
                        DatetimeC.Add("Min", mm.ToString("D2"))
                        DatetimeC.Add("Sec", ss.ToString("D2"))
                        DatetimeC.Add("mSec", mmm.ToString("D3"))
                        DateTimeB = DatetimeC
                        PlaceDataT.Add("Place", "Dubai ")
                        PlaceDataT.Add("State", " ")
                        PlaceDataT.Add("Country", " U.A.E.")
                        PlaceDataT.Add("lonB", 55.333333333333336)
                        PlaceDataT.Add("latB", 25.151620635773071)
                        PlaceDataT.Add("eW", "E")
                        PlaceDataT.Add("nS", "N")
                        PlaceDataT.Add("TzB", 4)
                        PlaceDataT.Add("DstB", 0)
                        Dim Horo As New TASystem.TrueAstro
                        Horo.getTransitList(DatetimeC, PlaceDataT, StrPlanet)
                        Dim DayDasa As String = ""
                        PlaceDataB = PlaceDataT
                        DatetimeC.Add("Year", "2018")
                        DatetimeC.Add("Month", "9")
                        DatetimeC.Add("Day", "24")
                        DateTimeB.Add("Hour", hh.ToString("D2"))
                        DateTimeB.Add("Min", mm.ToString("D2"))
                        DateTimeB.Add("Sec", ss.ToString("D2"))
                        DateTimeB.Add("mSec", If(mmm = 950, (mmm + 49).ToString("D3"), (mmm + 50).ToString("D3")))
                        Horo.getBirthDasaFordateTime(DatetimeC, DateTimeB, PlaceDataB, DayDasa)
                        UpdateHPREDICTION(UID, HID, StrPlanet, DayDasa, (hh.ToString() + ":" + mm.ToString + ":" + ss.ToString() + ":" + mmm.ToString()).ToString())
                    Next
                Next
            Next
        Next
        Console.WriteLine(Key1Cusp)
        Console.WriteLine(Key2Cusp)
        Console.WriteLine(Key3Cusp)
        Console.WriteLine(Key4Cusp)
        Console.WriteLine(Key5Cusp)
        Console.WriteLine(Key6Cusp)

    End Sub
    Function GetPlanetLongName(ByVal Pl As String)
        Dim Planet As String = "Not_A_Planet"
        Select Case Pl
            Case "Ma"
                Planet = "Mars"
            Case "Ve"
                Planet = "Venus"
            Case "Sa"
                Planet = "Saturn"
            Case "Ju"
                Planet = "Jupiter"
            Case "Su"
                Planet = "Sun"
            Case "Mo"
                Planet = "Moon"
            Case "Me"
                Planet = "Mercury"
            Case "Ra"
                Planet = "T.Rahu"
            Case "Ke"
                Planet = "T.Ketu"
            Case "Ne"
                Planet = "Neptune"
            Case "Pl"
                Planet = "Pluto"
        End Select
        Return Planet
    End Function
    Sub UpdateHPREDICTION(ByVal UID As String, ByVal HID As String, ByVal StrPlanet() As String, ByVal DayDasa As String, ByVal time As String)
        Dim Key1Cusp As String = "::k1->"
        Dim Key2Cusp As String = "::k2->"
        Dim Key3Cusp As String = "::k3->"
        Dim Key4Cusp As String = "::k4->"
        Dim Key5Cusp As String = "::k5->"
        Dim Key6Cusp As String = "::k6->"
        Dim MatchKey1 As New DataSet()
        Dim MatchKey2 As New DataSet()
        Dim MatchKey3 As New DataSet()
        Dim MatchKey4 As New DataSet()
        Dim MatchKey5 As New DataSet()
        Dim MatchKey6 As New DataSet()

        Dim Key_1 = StrPlanet(0).Split("|")(3) + StrPlanet(0).Split("|")(4) + StrPlanet(0).Split("|")(5) + StrPlanet(0).Split("|")(6) + StrPlanet(0).Split("|")(7) + StrPlanet(0).Split("|")(8)
        Dim Key_2 As String = ""
        Dim Key_3 As String = ""
        Dim Key_4 As String = ""
        Dim Key_5 As String = ""
        Dim Key_6 As String = ""
        For i As Integer = 1 To 12
            If StrPlanet(i).Split("|")(0) = GetPlanetLongName(DayDasa.Split("-")(0)) Then
                Key_2 = DayDasa.Split("-")(0) + StrPlanet(i).Split("|")(3) + StrPlanet(i).Split("|")(4) + StrPlanet(i).Split("|")(5) + StrPlanet(i).Split("|")(6) + StrPlanet(i).Split("|")(7) + StrPlanet(i).Split("|")(8)
            End If
            If StrPlanet(i).Split("|")(0) = GetPlanetLongName(DayDasa.Split("-")(1)) Then
                Key_3 = DayDasa.Split("-")(1) + StrPlanet(i).Split("|")(3) + StrPlanet(i).Split("|")(4) + StrPlanet(i).Split("|")(5) + StrPlanet(i).Split("|")(6) + StrPlanet(i).Split("|")(7) + StrPlanet(i).Split("|")(8)
            End If
            If StrPlanet(i).Split("|")(0) = GetPlanetLongName(DayDasa.Split("-")(2)) Then
                Key_4 = DayDasa.Split("-")(2) + StrPlanet(i).Split("|")(3) + StrPlanet(i).Split("|")(4) + StrPlanet(i).Split("|")(5) + StrPlanet(i).Split("|")(6) + StrPlanet(i).Split("|")(7) + StrPlanet(i).Split("|")(8)
            End If
            If StrPlanet(i).Split("|")(0) = GetPlanetLongName(DayDasa.Split("-")(3)) Then
                Key_5 = DayDasa.Split("-")(3) + StrPlanet(i).Split("|")(3) + StrPlanet(i).Split("|")(4) + StrPlanet(i).Split("|")(5) + StrPlanet(i).Split("|")(6) + StrPlanet(i).Split("|")(7) + StrPlanet(i).Split("|")(8)
            End If
            If StrPlanet(i).Split("|")(0) = GetPlanetLongName(DayDasa.Split("-")(4)) Then
                Key_6 = DayDasa.Split("-")(4) + StrPlanet(i).Split("|")(3) + StrPlanet(i).Split("|")(4) + StrPlanet(i).Split("|")(5) + StrPlanet(i).Split("|")(6) + StrPlanet(i).Split("|")(7) + StrPlanet(i).Split("|")(8)
            End If
        Next
        Dim connection As SqlConnection = New SqlConnection("data source=49.50.103.132;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=False;User Id=sa;password=pSI)TA1t0K[)")
        Dim cmd_Key1 As New SqlCommand($"SELECT CUSPID FROM HEADLETTERS_ENGINE.DBO.MATCHFILE_ASC
                                    WHERE MATCHFILE_ASC.HID = '" + HID + "'
                                        AND MATCHFILE_ASC.UID = '" + UID + "'
                                            AND MATCHFILE_ASC.MKEY = '" + Key_1.ToUpper() + "';", connection)
        Dim daKey1 As New SqlDataAdapter(cmd_Key1)
        daKey1.Fill(MatchKey1)
        If MatchKey1.Tables(0).Rows.Count = 0 Then
            Exit Sub
        End If
        Dim cmd_Key2 As New SqlCommand($"SELECT CUSPID FROM HEADLETTERS_ENGINE.DBO.MATCHFILE
                                    WHERE MATCHFILE.HID = '" + HID + "'
                                        AND MATCHFILE.UID = '" + UID + "'
                                            AND MATCHFILE.MKEY = '" + Key_2.ToUpper() + "';", connection)
        Dim daKey2 As New SqlDataAdapter(cmd_Key2)
        daKey2.Fill(MatchKey2)

        Dim cmd_Key3 As New SqlCommand($"SELECT CUSPID FROM HEADLETTERS_ENGINE.DBO.MATCHFILE
                                    WHERE MATCHFILE.HID = '" + HID + "'
                                        AND MATCHFILE.UID = '" + UID + "'
                                            AND MATCHFILE.MKEY = '" + Key_3.ToUpper() + "';", connection)
        Dim daKey3 As New SqlDataAdapter(cmd_Key3)
        daKey3.Fill(MatchKey3)

        Dim cmd_Key4 As New SqlCommand($"SELECT CUSPID FROM HEADLETTERS_ENGINE.DBO.MATCHFILE
                                    WHERE MATCHFILE.HID = '" + HID + "'
                                        AND MATCHFILE.UID = '" + UID + "'
                                            AND MATCHFILE.MKEY = '" + Key_4.ToUpper() + "';", connection)
        Dim daKey4 As New SqlDataAdapter(cmd_Key4)
        daKey4.Fill(MatchKey4)

        Dim cmd_Key5 As New SqlCommand($"SELECT CUSPID FROM HEADLETTERS_ENGINE.DBO.MATCHFILE
                                    WHERE MATCHFILE.HID = '" + HID + "'
                                        AND MATCHFILE.UID = '" + UID + "'
                                            AND MATCHFILE.MKEY = '" + Key_5.ToUpper() + "';", connection)
        Dim daKey5 As New SqlDataAdapter(cmd_Key5)
        daKey5.Fill(MatchKey5)

        Dim cmd_Key6 As New SqlCommand($"SELECT CUSPID FROM HEADLETTERS_ENGINE.DBO.MATCHFILE
                                    WHERE MATCHFILE.HID = '" + HID + "'
                                        AND MATCHFILE.UID = '" + UID + "'
                                            AND MATCHFILE.MKEY = '" + Key_6.ToUpper() + "';", connection)

        Dim daKey6 As New SqlDataAdapter(cmd_Key6)
        daKey6.Fill(MatchKey6)

        If MatchKey2.Tables(0).Rows.Count < 2 And MatchKey3.Tables(0).Rows.Count < 2 And MatchKey4.Tables(0).Rows.Count < 2 And MatchKey5.Tables(0).Rows.Count < 2 And MatchKey6.Tables(0).Rows.Count < 2 Then
            Exit Sub
        End If


        If MatchKey1.Tables(0).Rows.Count <> 0 Then
            For i As Integer = 0 To MatchKey1.Tables(0).Rows.Count - 1
                Key1Cusp = Key1Cusp + MatchKey1.Tables(0).Rows(i)(0).ToString().Trim + "-"
            Next
        End If
        If MatchKey2.Tables(0).Rows.Count <> 0 Then
            For i As Integer = 0 To MatchKey2.Tables(0).Rows.Count - 1
                Key2Cusp = Key2Cusp + MatchKey2.Tables(0).Rows(i)(0).ToString().Trim + "-"
            Next
        End If
        If MatchKey3.Tables(0).Rows.Count <> 0 Then
            For i As Integer = 0 To MatchKey3.Tables(0).Rows.Count - 1
                Key3Cusp = Key3Cusp + MatchKey3.Tables(0).Rows(i)(0).ToString().Trim + "-"
            Next
        End If
        If MatchKey4.Tables(0).Rows.Count <> 0 Then
            For i As Integer = 0 To MatchKey4.Tables(0).Rows.Count - 1
                Key4Cusp = Key4Cusp + MatchKey4.Tables(0).Rows(i)(0).ToString().Trim + "-"
            Next
        End If
        If MatchKey5.Tables(0).Rows.Count <> 0 Then
            For i As Integer = 0 To MatchKey5.Tables(0).Rows.Count - 1
                Key5Cusp = Key5Cusp + MatchKey5.Tables(0).Rows(i)(0).ToString().Trim + "-"
            Next
        End If
        If MatchKey6.Tables(0).Rows.Count <> 0 Then
            For i As Integer = 0 To MatchKey6.Tables(0).Rows.Count - 1
                Key6Cusp = Key6Cusp + MatchKey6.Tables(0).Rows(i)(0).ToString().Trim + "-"
            Next
        End If
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = "data source=49.50.103.132;initial catalog=ASTROLOGYSOFTWARE_DB;integrated security=False;User Id=sa;password=pSI)TA1t0K[)"
        con.Open()
        cmd.Connection = con
        cmd.CommandText = $"INSERT INTO ASTROLOGYSOFTWARE_DB.DBO._HPREDICTION VALUES ('" + UID + "','" + HID + "','??',007,
                                    '" + Key1Cusp + Key2Cusp + Key3Cusp + Key4Cusp + Key5Cusp + Key6Cusp + "','09/24/2018 " + time + "','" + time + "', '" + DayDasa + "','" + "::Key_1 ->" + Key_1 + "::Key_2 ->" + Key_2 + "::Key_3 ->" + Key_3 + "::Key_4 ->" + Key_4 + "::Key_5 ->" + Key_5 + "::Key_6 ->" + Key_6 + "',
                                        'NA','NA','NA','NA','NA','NA','NA');"
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub
End Module