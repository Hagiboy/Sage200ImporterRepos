Option Strict Off
Option Explicit On

Imports MySql.Data.MySqlClient
Imports System.Data.OracleClient
Imports System.Data.SqlClient

Friend Class ClassImport

    Dim objdbConnZHDB02 As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))
    Dim objdbcommandZHDB02 As New MySqlCommand

    Dim objdbAccessConn As New OleDb.OleDbConnection
    Dim objOLEdbcmdLoc As New OleDb.OleDbCommand

    Friend Function FcKreditFill(intAccounting As Int16) As Int16

        Dim strIdentityName As String
        Dim strMDBName As String
        Dim strSQL As String
        Dim strSQLSub As String
        Dim strKRGTableType As String
        Dim objmysqlcomdwritehead As New MySqlCommand
        Dim intFcReturns As Int16
        Dim objmysqlcomdwritesub As New MySqlCommand
        Dim objdtLocKrediHead As New DataTable
        Dim objdtLocKrediSubs As New DataTable
        Dim strConnection As String
        Dim objRGMySQLConn As New MySqlConnection
        Dim objlocMySQLcmd As New MySqlCommand
        Dim strSQLToParse As String

        Try

            objmysqlcomdwritehead.Connection = objdbConnZHDB02
            objmysqlcomdwritesub.Connection = objdbConnZHDB02

            'Für den Save der Records
            strIdentityName = System.Security.Principal.WindowsIdentity.GetCurrent().Name
            strIdentityName = Strings.Replace(strIdentityName, "\", "/")

            strMDBName = FcReadFromSettingsII("Buchh_KRGTableMDB",
                                              intAccounting)

            strSQL = FcReadFromSettingsII("Buchh_SQLHeadKred",
                                          intAccounting)

            strKRGTableType = FcReadFromSettingsII("Buchh_KRGTableType",
                                                   intAccounting)

            If IsDBNull(strSQL) Then
                Return 1

            Else
                If strKRGTableType = "A" Then

                    'Access
                    Call FcInitAccessConnecation(objdbAccessConn,
                                                  strMDBName)
                    objdbAccessConn.Open()
                    objOLEdbcmdLoc.CommandText = strSQL
                    objOLEdbcmdLoc.Connection = objdbAccessConn
                    objdtLocKrediHead.Load(objOLEdbcmdLoc.ExecuteReader)
                    objdbAccessConn.Close()
                ElseIf strKRGTableType = "M" Then

                    strConnection = System.Configuration.ConfigurationManager.AppSettings(strMDBName)
                    objRGMySQLConn.ConnectionString = strConnection
                    objlocMySQLcmd.Connection = objRGMySQLConn
                    objlocMySQLcmd.CommandText = strSQL
                    objRGMySQLConn.Open()
                    objdtLocKrediHead.Load(objlocMySQLcmd.ExecuteReader)
                    objRGMySQLConn.Close()

                End If

                strSQLToParse = FcReadFromSettingsII("Buchh_SQLDetailKred",
                                                    intAccounting)

                intFcReturns = FcInitInsCmdKHeads(objmysqlcomdwritehead)

                For Each row As DataRow In objdtLocKrediHead.Rows

                    objmysqlcomdwritehead.Connection.Open()
                    objmysqlcomdwritehead.Parameters("@IdentityName").Value = strIdentityName
                    objmysqlcomdwritehead.Parameters("@ProcessID").Value = Process.GetCurrentProcess().Id
                    objmysqlcomdwritehead.Parameters("@intBuchhaltung").Value = intAccounting
                    objmysqlcomdwritehead.Parameters("@strKredRGNbr").Value = row("strKredRGNbr")
                    objmysqlcomdwritehead.Parameters("@intBuchungsart").Value = row("intBuchungsart")
                    objmysqlcomdwritehead.Parameters("@lngKredID").Value = row("lngKredID")
                    objmysqlcomdwritehead.Parameters("@strOPNr").Value = row("strOPNr")
                    objmysqlcomdwritehead.Parameters("@lngKredNbr").Value = row("lngKredNbr")
                    objmysqlcomdwritehead.Parameters("@lngKredKtoNbr").Value = row("lngKredKtoNbr")
                    objmysqlcomdwritehead.Parameters("@strKredCur").Value = row("strKredCur")
                    objmysqlcomdwritehead.Parameters("@lngKrediKST").Value = row("lngKrediKST")
                    objmysqlcomdwritehead.Parameters("@dblKredNetto").Value = row("dblKredNetto")
                    objmysqlcomdwritehead.Parameters("@dblKredMwSt").Value = row("dblKredMwSt")
                    objmysqlcomdwritehead.Parameters("@dblKredBrutto").Value = row("dblKredBrutto")
                    objmysqlcomdwritehead.Parameters("@lngKredIdentNbr").Value = row("lngKredIdentNbr")
                    objmysqlcomdwritehead.Parameters("@strKredText").Value = row("strKredText")
                    objmysqlcomdwritehead.Parameters("@strKredRef").Value = row("strKredRef")
                    objmysqlcomdwritehead.Parameters("@datKredRGDatum").Value = row("datKredRGDatum")
                    objmysqlcomdwritehead.Parameters("@datKredValDatum").Value = row("datKredValDatum")
                    objmysqlcomdwritehead.Parameters("@intPayType").Value = row("intPayType")
                    objmysqlcomdwritehead.Parameters("@strKrediBank").Value = row("strKrediBank")
                    objmysqlcomdwritehead.Parameters("@strKrediBankInt").Value = row("strKrediBankInt")
                    objmysqlcomdwritehead.Parameters("@strRGName").Value = row("strRGName")
                    objmysqlcomdwritehead.Parameters("@strRGBemerkung").Value = row("strRGBemerkung")
                    objmysqlcomdwritehead.Parameters("@intZKond").Value = row("intZKond")
                    If objdtLocKrediHead.Columns.Contains("datPGVFrom") Then
                        objmysqlcomdwritehead.Parameters("@datPGVFrom").Value = row("datPGVFrom")
                    End If
                    If objdtLocKrediHead.Columns.Contains("datPGVTo") Then
                        objmysqlcomdwritehead.Parameters("@datPGVTo").Value = row("datPGVTo")
                    End If

                    objmysqlcomdwritehead.ExecuteNonQuery()
                    objmysqlcomdwritehead.Connection.Close()

                    'Subs einlesen
                    strSQLSub = FcSQLParseKredi(strSQLToParse,
                                                row("lngKredID"),
                                                objdtLocKrediHead)

                    If strKRGTableType = "A" Then
                        objdbAccessConn.Open()
                        objOLEdbcmdLoc.CommandText = strSQLSub
                        objdtLocKrediSubs.Load(objOLEdbcmdLoc.ExecuteReader)
                        objdbAccessConn.Close()
                    ElseIf strKRGTableType = "M" Then
                        objlocMySQLcmd.CommandText = strSQLSub
                        objRGMySQLConn.Open()
                        objdtLocKrediSubs.Load(objlocMySQLcmd.ExecuteReader)
                        objRGMySQLConn.Close()
                    End If

                Next
                'Subs schreiben
                intFcReturns = FcInitInscmdKSubs(objmysqlcomdwritesub)
                For Each drsub As DataRow In objdtLocKrediSubs.Rows

                    objmysqlcomdwritesub.Connection.Open()
                    objmysqlcomdwritesub.Parameters("@IdentityName").Value = strIdentityName
                    objmysqlcomdwritesub.Parameters("@ProcessID").Value = Process.GetCurrentProcess().Id
                    objmysqlcomdwritesub.Parameters("@lngKredID").Value = drsub("lngKredID")
                    objmysqlcomdwritesub.Parameters("@lngKto").Value = drsub("lngKto")
                    objmysqlcomdwritesub.Parameters("@lngKST").Value = drsub("lngKST")
                    objmysqlcomdwritesub.Parameters("@dblNetto").Value = drsub("dblNetto")
                    objmysqlcomdwritesub.Parameters("@dblMwSt").Value = drsub("dblMwSt")
                    objmysqlcomdwritesub.Parameters("@dblBrutto").Value = drsub("dblBrutto")
                    objmysqlcomdwritesub.Parameters("@dblMwStSatz").Value = drsub("dblMwStSatz")
                    objmysqlcomdwritesub.Parameters("@strMwStKey").Value = drsub("strMwStKey")
                    objmysqlcomdwritesub.Parameters("@intSollHaben").Value = drsub("intSollHaben")
                    If objdtLocKrediSubs.Columns.Contains("strKredSubText") Then
                        objmysqlcomdwritesub.Parameters("@strKredSubText").Value = drsub("strKredSubText")
                    End If
                    objmysqlcomdwritesub.Parameters("@booRebilling").Value = drsub("booRebilling")
                    objmysqlcomdwritesub.ExecuteNonQuery()
                    objmysqlcomdwritesub.Connection.Close()

                Next

            End If
            Return 0


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Kreditoeren-Daten-Lesen Problem", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 1

        End Try

    End Function

    Friend Function FcDebitFill(intAccounting As Int16) As Int16

        Dim strIdentityName As String
        Dim strMDBName As String
        Dim strSQL As String
        Dim strSQLSub As String
        Dim strRGTableType As String
        Dim objdtLocDebiHead As New DataTable
        Dim objdtlocDebiSub As New DataTable
        Dim strConnection As String
        Dim objRGMySQLConn As New MySqlConnection
        Dim objlocMySQLcmd As New MySqlCommand
        Dim strSQLToParse As String
        Dim objmysqlcomdwritehead As New MySqlCommand
        Dim intFcReturns As Int16
        Dim strmysqlSaveSub As String
        Dim objmysqlcomdwritesub As New MySqlCommand


        Try

            objmysqlcomdwritehead.Connection = objdbConnZHDB02
            objmysqlcomdwritesub.Connection = objdbConnZHDB02

            'Für den Save der Records
            strIdentityName = System.Security.Principal.WindowsIdentity.GetCurrent().Name
            strIdentityName = Strings.Replace(strIdentityName, "\", "/")

            strMDBName = FcReadFromSettingsII("Buchh_RGTableMDB",
                                                    intAccounting)

            strSQL = FcReadFromSettingsII("Buchh_SQLHead",
                                             intAccounting)

            strRGTableType = FcReadFromSettingsII("Buchh_RGTableType",
                                                     intAccounting)

            If strRGTableType = "A" Then

                'Access
                Call FcInitAccessConnecation(objdbAccessConn,
                                              strMDBName)
                objdbAccessConn.Open()
                objOLEdbcmdLoc.CommandText = strSQL
                objOLEdbcmdLoc.Connection = objdbAccessConn
                objdtLocDebiHead.Load(objOLEdbcmdLoc.ExecuteReader)
                objdbAccessConn.Close()
            ElseIf strRGTableType = "M" Then

                strConnection = System.Configuration.ConfigurationManager.AppSettings(strMDBName)
                objRGMySQLConn.ConnectionString = strConnection
                objlocMySQLcmd.Connection = objRGMySQLConn
                objlocMySQLcmd.CommandText = strSQL
                objRGMySQLConn.Open()
                objdtLocDebiHead.Load(objlocMySQLcmd.ExecuteReader)
                objRGMySQLConn.Close()


            End If

            strSQLToParse = FcReadFromSettingsII("Buchh_SQLDetail",
                                                    intAccounting)
            intFcReturns = FcInitInsCmdDHeads(objmysqlcomdwritehead)

            For Each row As DataRow In objdtLocDebiHead.Rows

                objmysqlcomdwritehead.Connection.Open()
                objmysqlcomdwritehead.Parameters("@IdentityName").Value = strIdentityName
                objmysqlcomdwritehead.Parameters("@ProcessID").Value = Process.GetCurrentProcess().Id
                objmysqlcomdwritehead.Parameters("@intBuchhaltung").Value = intAccounting
                objmysqlcomdwritehead.Parameters("@strDebRGNbr").Value = row("strDebRGNbr")
                objmysqlcomdwritehead.Parameters("@intBuchungsart").Value = row("intBuchungsart")
                objmysqlcomdwritehead.Parameters("@intRGArt").Value = row("intRGArt")
                If objdtLocDebiHead.Columns.Contains("strRGArt") Then
                    objmysqlcomdwritehead.Parameters("@strRGArt").Value = row("strRGArt")
                End If
                objmysqlcomdwritehead.Parameters("@strOPNr").Value = row("strOPNr")
                objmysqlcomdwritehead.Parameters("@lngDebNbr").Value = row("lngDebNbr")
                objmysqlcomdwritehead.Parameters("@lngDebKtoNbr").Value = row("lngDebKtoNbr")
                objmysqlcomdwritehead.Parameters("@strDebCur").Value = row("strDebCur")
                objmysqlcomdwritehead.Parameters("@lngDebiKST").Value = row("lngDebiKST")
                objmysqlcomdwritehead.Parameters("@dblDebNetto").Value = row("dblDebNetto")
                objmysqlcomdwritehead.Parameters("@dblDebMwSt").Value = row("dblDebMwSt")
                objmysqlcomdwritehead.Parameters("@dblDebBrutto").Value = row("dblDebBrutto")
                objmysqlcomdwritehead.Parameters("@lngDebIdentNbr").Value = row("lngDebIdentNbr")
                objmysqlcomdwritehead.Parameters("@strDebText").Value = row("strDebText")
                If objdtLocDebiHead.Columns.Contains("strDebReferenz") Then
                    objmysqlcomdwritehead.Parameters("@strDebreferenz").Value = row("strDebReferenz")
                End If
                objmysqlcomdwritehead.Parameters("@datDebRGDatum").Value = row("datDebRGDatum")
                objmysqlcomdwritehead.Parameters("@datDebValDatum").Value = row("datDebValDatum")
                If objdtLocDebiHead.Columns.Contains("datRGCreate") Then
                    objmysqlcomdwritehead.Parameters("@datRGCreate").Value = row("datRGCreate")
                End If
                If objdtLocDebiHead.Columns.Contains("intPayType") Then
                    objmysqlcomdwritehead.Parameters("@intPayType").Value = row("intPayType")
                End If
                objmysqlcomdwritehead.Parameters("@strDebiBank").Value = row("strDebiBank")
                objmysqlcomdwritehead.Parameters("@lngLinkedRG").Value = row("lngLinkedRG")
                objmysqlcomdwritehead.Parameters("@strRGName").Value = row("strRGName")
                If objdtLocDebiHead.Columns.Contains("strDebIdentNbr2") Then
                    objmysqlcomdwritehead.Parameters("@strDebIdentNbr2").Value = row("strDebIdentNbr2")
                End If
                If objdtLocDebiHead.Columns.Contains("booCrToInv") Then
                    objmysqlcomdwritehead.Parameters("@booCrToInv").Value = row("booCrToInv")
                End If
                If objdtLocDebiHead.Columns.Contains("datPGVFrom") Then
                    objmysqlcomdwritehead.Parameters("@datPGVFrom").Value = row("datPGVFrom")
                End If
                If objdtLocDebiHead.Columns.Contains("bdatPGVTo") Then
                    objmysqlcomdwritehead.Parameters("@datPGVTo").Value = row("datPGVTo")
                End If
                objmysqlcomdwritehead.ExecuteNonQuery()
                objmysqlcomdwritehead.Connection.Close()
                objdtLocDebiHead.AcceptChanges()

                'Subs einlesen
                strSQLSub = FcSQLParse(strSQLToParse,
                                                   row("strDebRGNbr"),
                                                   objdtLocDebiHead,
                                                   "D")
                If strRGTableType = "A" Then
                    objdbAccessConn.Open()
                    objOLEdbcmdLoc.CommandText = strSQLSub
                    objdtlocDebiSub.Load(objOLEdbcmdLoc.ExecuteReader)
                    objdbAccessConn.Close()
                ElseIf strRGTableType = "M" Then
                    objlocMySQLcmd.CommandText = strSQLSub
                    objRGMySQLConn.Open()
                    objdtlocDebiSub.Load(objlocMySQLcmd.ExecuteReader)
                    objRGMySQLConn.Close()
                End If

            Next
            'Subs schreiben
            intFcReturns = FcInitInscmdSubs(objmysqlcomdwritesub)
            For Each drsub As DataRow In objdtlocDebiSub.Rows

                objmysqlcomdwritesub.Connection.Open()
                objmysqlcomdwritesub.Parameters("@IdentityName").Value = strIdentityName
                objmysqlcomdwritesub.Parameters("@ProcessID").Value = Process.GetCurrentProcess().Id
                objmysqlcomdwritesub.Parameters("@strRGNr").Value = drsub("strRGNr")
                objmysqlcomdwritesub.Parameters("@lngKto").Value = drsub("lngKto")
                objmysqlcomdwritesub.Parameters("@lngKST").Value = drsub("lngKST")
                objmysqlcomdwritesub.Parameters("@dblNetto").Value = drsub("dblNetto")
                objmysqlcomdwritesub.Parameters("@dblMwSt").Value = drsub("dblMwSt")
                objmysqlcomdwritesub.Parameters("@dblBrutto").Value = drsub("dblBrutto")
                objmysqlcomdwritesub.Parameters("@dblMwStSatz").Value = drsub("dblMwStSatz")
                objmysqlcomdwritesub.Parameters("@strMwStKey").Value = drsub("strMwStKey")
                objmysqlcomdwritesub.Parameters("@intSollHaben").Value = drsub("intSollHaben")
                If objdtlocDebiSub.Columns.Contains("strArtikel") Then
                    objmysqlcomdwritesub.Parameters("@strArtikel").Value = drsub("strArtikel")
                End If
                objmysqlcomdwritesub.ExecuteNonQuery()
                objmysqlcomdwritesub.Connection.Close()

                objdtlocDebiSub.AcceptChanges()

            Next


            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Debitoren-Daten-Lesen Problem", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 1

        End Try


    End Function

    Friend Function FcSQLParseKredi(ByVal strSQLToParse As String,
                                           ByVal lngKredID As Int32,
                                           ByVal objdtKredi As DataTable) As String

        'Funktion setzt in eingelesenem SQL wieder Variablen ein
        Dim intPipePositionBegin, intPipePositionEnd As Integer
        Dim strWork, strField As String
        Dim RowKredi() As DataRow
        Dim strFieldType As String

        Try

            'Zuerst Datensatz in Kredii-Head suchen
            RowKredi = objdtKredi.Select("lngKredID=" + lngKredID.ToString)

            '| suchen
            If InStr(strSQLToParse, "|") > 0 Then
                'Vorkommen gefunden
                intPipePositionBegin = InStr(strSQLToParse, "|")
                intPipePositionEnd = InStr(intPipePositionBegin + 1, strSQLToParse, "|")
                Do Until intPipePositionBegin = 0
                    strField = Mid(strSQLToParse, intPipePositionBegin + 1, intPipePositionEnd - intPipePositionBegin - 1)
                    Select Case strField
                        Case "rsKredi.Fields(""KrediID"")"
                            strField = RowKredi(0).Item("lngKredID")
                            strFieldType = "V"
                        Case "rsKredi.Fields(""KrediRGNr"")"
                            strField = RowKredi(0).Item("strKredRGNbr")
                            strFieldType = "T"
                            'Case "rsDebiTemp.Fields([strRGArt])"
                            '    strField = rsDebiTemp.Fields("strRGArt")
                            'Case "rsDebiTemp.Fields([strRGName])"
                            '    strField = rsDebiTemp.Fields("strRGName")
                            'Case "rsDebiTemp.Fields([strDebIdentNbr2])"
                            '    strField = rsDebiTemp.Fields("strDebIdentNbr2")
                            'Case "rsDebi.Fields([RGBemerkung])"
                            '    strField = rsDebi.Fields("RGBemerkung")
                            'Case "rsDebi.Fields([JornalNr])"
                            '    strField = rsDebi.Fields("JornalNr")
                            'Case "rsDebiTemp.Fields([strRGBemerkung])"
                            '    strField = rsDebiTemp.Fields("strRGBemerkung")
                            'Case "rsDebiTemp.Fields(""strDebRGNbr"")"
                            '    strField = rsDebiTemp.Fields("strDebRGNbr")
                            'Case "rsDebiTemp.Fields([lngDebIdentNbr])"
                            '    strField = rsDebiTemp.Fields("lngDebIdentNbr")
                            'Case "rsDebiTemp.Fields([strDebText])"
                            '    strField = rsDebiTemp.Fields("strDebText")
                            'Case "KUNDENZEICHEN"
                            '    strField = fcGetKundenzeichen(rsDebiTemp.Fields("lngDebIdentNbr"))
                        Case Else
                            strField = "unknown field"
                    End Select
                    strSQLToParse = Left(strSQLToParse, intPipePositionBegin - 1) + IIf(strFieldType = "T", "'", "") + strField + IIf(strFieldType = "T", "'", "") + Right(strSQLToParse, Len(strSQLToParse) - intPipePositionEnd)
                    'Neuer Anfang suchen für evtl. weitere |
                    intPipePositionBegin = InStr(strSQLToParse, "|")
                    'intPipePositionBegin = InStr(intPipePositionEnd + 1, strSQLToParse, "|")
                    intPipePositionEnd = InStr(intPipePositionBegin + 1, strSQLToParse, "|")
                Loop
            End If

            Return strSQLToParse

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Fehler Parsing " + Err.Number.ToString)

        End Try

    End Function


    Friend Function FcSQLParse(ByVal strSQLToParse As String,
                                      ByVal strRGNbr As String,
                                      ByVal objdtBookings As DataTable,
                                      ByVal strDebiCredit As String) As String

        'Funktion setzt in eingelesenem SQL wieder Variablen ein
        Dim intPipePositionBegin, intPipePositionEnd As Integer
        Dim strWork, strField As String
        Dim RowBooking() As DataRow

        Try

            If strDebiCredit = "D" Then
                'Zuerst Datensatz in Debi-Head suchen
                RowBooking = objdtBookings.Select("strDebRGNbr='" + strRGNbr + "'")
            Else
                RowBooking = objdtBookings.Select("strKredRGNbr='" + strRGNbr + "'")
            End If
            '| suchen
            If InStr(strSQLToParse, "|") > 0 Then
                'Vorkommen gefunden
                intPipePositionBegin = InStr(strSQLToParse, "|")
                intPipePositionEnd = InStr(intPipePositionBegin + 1, strSQLToParse, "|")
                Do Until intPipePositionBegin = 0
                    strField = Mid(strSQLToParse, intPipePositionBegin + 1, intPipePositionEnd - intPipePositionBegin - 1)
                    Select Case strField
                        Case "rsDebi.Fields(""RGNr"")"
                            strField = RowBooking(0).Item("strDebRGNbr")
                        Case "rsKrediTemp.Fields([strKredRGNbr])"
                            strField = RowBooking(0).Item("strKredRGNbr")
                        Case "rsDebiTemp.Fields([strDebPKBez])"
                            strField = RowBooking(0).Item("strDebBez")
                        Case "rsKrediTemp.Fields([strKredPKBez])"
                            strField = RowBooking(0).Item("strKredBez")
                        Case "rsDebiTemp.Fields([lngDebIdentNbr])"
                            strField = RowBooking(0).Item("lngDebIdentNbr")
                        Case "rsKrediTemp.Fields([lngKredIdentNbr])"
                            strField = RowBooking(0).Item("lngKredIdentNbr")
                        Case "rsDebiTemp.Fields([strRGArt])"
                            strField = RowBooking(0).Item("strRGArt")
                        Case "rsDebiTemp.Fields([strRGName])"
                            strField = RowBooking(0).Item("strRGName")
                        Case "rsDebiTemp.Fields([strDebIdentNbr2])"
                            strField = RowBooking(0).Item("strDebIdentNbr2")
                        'Case "rsDebi.Fields([RGBemerkung])"
                        '    strField = rsDebi.Fields("RGBemerkung")
                        'Case "rsDebi.Fields([JornalNr])"
                        '    strField = rsDebi.Fields("JornalNr")
                        'Case "rsDebiTemp.Fields([strRGBemerkung])"
                        '    strField = rsDebiTemp.Fields("strRGBemerkung")
                        'Case "rsDebiTemp.Fields(""strDebRGNbr"")"
                        '    strField = rsDebiTemp.Fields("strDebRGNbr")
                        'Case "rsDebiTemp.Fields([lngDebIdentNbr])"
                        '    strField = rsDebiTemp.Fields("lngDebIdentNbr")
                        Case "rsDebiTemp.Fields([strDebText])"
                            strField = RowBooking(0).Item("strDebText")
                        Case "KUNDENZEICHEN"
                            strField = FcGetKundenzeichen(RowBooking(0).Item("lngDebIdentNbr"))
                        Case Else
                            strField = "unknown field"
                    End Select
                    strSQLToParse = Left(strSQLToParse, intPipePositionBegin - 1) & strField & Right(strSQLToParse, Len(strSQLToParse) - intPipePositionEnd)
                    'Neuer Anfang suchen für evtl. weitere |
                    intPipePositionBegin = InStr(strSQLToParse, "|")
                    'intPipePositionBegin = InStr(intPipePositionEnd + 1, strSQLToParse, "|")
                    intPipePositionEnd = InStr(intPipePositionBegin + 1, strSQLToParse, "|")
                Loop
            End If

            'Debug.Print("Parsed " + strRGNbr)
            Return strSQLToParse

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Parsing " + Err.Number.ToString)

        Finally
            RowBooking = Nothing
            'Application.DoEvents()

        End Try


    End Function

    Friend Function FcGetKundenzeichen(ByVal lngJournalNr As Int32) As String
        'ByRef objOracleCon As OracleConnection,
        'ByRef objOracleCmd As OracleCommand) As String

        Dim objdbConnCIS As New MySqlConnection
        Dim objdbCmdCIS As New MySqlCommand
        Dim objdtJournalKZ As New DataTable

        Try
            'Angaben einlesen
            objdbConnCIS.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringCIS")
            objdbConnCIS.Open()
            objdbCmdCIS.Connection = objdbConnCIS
            objdbCmdCIS.CommandText = "SELECT KUNDENZEICHEN FROM TAB_JOURNALSTAMM WHERE JORNALNR=" + lngJournalNr.ToString
            objdtJournalKZ.Load(objdbCmdCIS.ExecuteReader)

            If objdtJournalKZ.Rows.Count > 0 Then
                If Not IsDBNull(objdtJournalKZ.Rows(0).Item(0)) Then
                    Return objdtJournalKZ.Rows(0).Item(0)
                Else
                    Return "n/a"
                End If
            Else
                Return "new"
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Kundenzeichen holen " + Err.Number.ToString, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        Finally
            objdtJournalKZ.Clear()
            objdtJournalKZ = Nothing

            objdbConnCIS.Close()
            objdbConnCIS = Nothing

            objdbCmdCIS = Nothing


        End Try

    End Function

    Friend Function FcInitInscmdSubs(ByRef mysqlinscmd As MySqlCommand) As Int16

        'Debitoren - Head
        Dim inscmdFields As String
        Dim inscmdValues As String

        Try

            inscmdFields = "IdentityName"
            inscmdValues = "@IdentityName"
            inscmdFields += ", ProcessID"
            inscmdValues += ", @ProcessID"
            inscmdFields += ", strRGNr"
            inscmdValues += ", @strRGNr"
            inscmdFields += ", lngKto"
            inscmdValues += ", @lngKto"
            inscmdFields += ", lngKST"
            inscmdValues += ", @lngKST"
            inscmdFields += ", dblNetto"
            inscmdValues += ", @dblNetto"
            inscmdFields += ", dblMwSt"
            inscmdValues += ", @dblMwSt"
            inscmdFields += ", dblBrutto"
            inscmdValues += ", @dblBrutto"
            inscmdFields += ", dblMwStSatz"
            inscmdValues += ", @dblMwStSatz"
            inscmdFields += ", strMwStKey"
            inscmdValues += ", @strMwStKey"
            inscmdFields += ", intSollHaben"
            inscmdValues += ", @intSollHaben"
            inscmdFields += ", strArtikel"
            inscmdValues += ", @strArtikel"

            'Ins cmd DebiSub
            mysqlinscmd.CommandText = "INSERT INTO tbldebitorensub (" + inscmdFields + ") VALUES (" + inscmdValues + ")"
            mysqlinscmd.Parameters.Add("@IdentityName", MySqlDbType.String).SourceColumn = "IdentityName"
            mysqlinscmd.Parameters.Add("@ProcessID", MySqlDbType.Int16).SourceColumn = "ProcessID"
            mysqlinscmd.Parameters.Add("@strRGNr", MySqlDbType.String).SourceColumn = "strRGNr"
            mysqlinscmd.Parameters.Add("@lngKto", MySqlDbType.Int32).SourceColumn = "lngKto"
            mysqlinscmd.Parameters.Add("@lngKST", MySqlDbType.Int32).SourceColumn = "lngKST"
            mysqlinscmd.Parameters.Add("@dblNetto", MySqlDbType.Decimal).SourceColumn = "dblNetto"
            mysqlinscmd.Parameters.Add("@dblMwst", MySqlDbType.Decimal).SourceColumn = "dblMwSt"
            mysqlinscmd.Parameters.Add("@dblBrutto", MySqlDbType.Decimal).SourceColumn = "dblBrutto"
            mysqlinscmd.Parameters.Add("@dblMwStSatz", MySqlDbType.Double).SourceColumn = "dblMwStSatz"
            mysqlinscmd.Parameters.Add("@strMwStKey", MySqlDbType.String).SourceColumn = "strMwStKey"
            mysqlinscmd.Parameters.Add("@intSollHaben", MySqlDbType.Int16).SourceColumn = "intSollHaben"
            mysqlinscmd.Parameters.Add("@strArtikel", MySqlDbType.String).SourceColumn = "strArtikel"

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem SubCommand Init", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 1

        End Try


    End Function

    Friend Function FcInitInscmdKSubs(ByRef mysqlinscmd As MySqlCommand) As Int16

        'Debitoren - Head
        Dim inscmdFields As String
        Dim inscmdValues As String

        Try

            inscmdFields = "IdentityName"
            inscmdValues = "@IdentityName"
            inscmdFields += ", ProcessID"
            inscmdValues += ", @ProcessID"
            inscmdFields += ", lngKredID"
            inscmdValues += ", @lngKredID"
            inscmdFields += ", lngKto"
            inscmdValues += ", @lngKto"
            inscmdFields += ", lngKST"
            inscmdValues += ", @lngKST"
            inscmdFields += ", dblNetto"
            inscmdValues += ", @dblNetto"
            inscmdFields += ", dblMwSt"
            inscmdValues += ", @dblMwSt"
            inscmdFields += ", dblBrutto"
            inscmdValues += ", @dblBrutto"
            inscmdFields += ", dblMwStSatz"
            inscmdValues += ", @dblMwStSatz"
            inscmdFields += ", strMwStKey"
            inscmdValues += ", @strMwStKey"
            inscmdFields += ", intSollHaben"
            inscmdValues += ", @intSollHaben"
            inscmdFields += ", strKredSubText"
            inscmdValues += ", @strKredSubText"
            inscmdFields += ", booRebilling"
            inscmdValues += ", @booRebilling"


            'Ins cmd DebiSub
            mysqlinscmd.CommandText = "INSERT INTO tblkreditorensub (" + inscmdFields + ") VALUES (" + inscmdValues + ")"
            mysqlinscmd.Parameters.Add("@IdentityName", MySqlDbType.String).SourceColumn = "IdentityName"
            mysqlinscmd.Parameters.Add("@ProcessID", MySqlDbType.Int16).SourceColumn = "ProcessID"
            mysqlinscmd.Parameters.Add("@lngKredID", MySqlDbType.Int32).SourceColumn = "lngKredID"
            mysqlinscmd.Parameters.Add("@lngKto", MySqlDbType.Int32).SourceColumn = "lngKto"
            mysqlinscmd.Parameters.Add("@lngKST", MySqlDbType.Int32).SourceColumn = "lngKST"
            mysqlinscmd.Parameters.Add("@dblNetto", MySqlDbType.Decimal).SourceColumn = "dblNetto"
            mysqlinscmd.Parameters.Add("@dblMwst", MySqlDbType.Decimal).SourceColumn = "dblMwSt"
            mysqlinscmd.Parameters.Add("@dblBrutto", MySqlDbType.Decimal).SourceColumn = "dblBrutto"
            mysqlinscmd.Parameters.Add("@dblMwStSatz", MySqlDbType.Double).SourceColumn = "dblMwStSatz"
            mysqlinscmd.Parameters.Add("@strMwStKey", MySqlDbType.String).SourceColumn = "strMwStKey"
            mysqlinscmd.Parameters.Add("@intSollHaben", MySqlDbType.Int16).SourceColumn = "intSollHaben"
            mysqlinscmd.Parameters.Add("@strKredSubText", MySqlDbType.String).SourceColumn = "strKredSubText"
            mysqlinscmd.Parameters.Add("@booRebilling", MySqlDbType.Int16).SourceColumn = "booRebilling"

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem KSubCommand Init", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 1

        End Try


    End Function


    Friend Function FcInitInsCmdDHeads(ByRef mysqlinscmd As MySqlCommand) As Int16

        'Dim strIdentityName As String

        'Debitoren - Head
        Dim inscmdFields As String
        Dim inscmdValues As String

        Try

            inscmdFields = "IdentityName"
            inscmdValues = "@IdentityName"
            inscmdFields += ", ProcessID"
            inscmdValues += ", @ProcessID"
            inscmdFields += ", intBuchhaltung"
            inscmdValues += ", @intBuchhaltung"
            inscmdFields += ", strDebRGNbr"
            inscmdValues += ", @strDebRGNbr"
            inscmdFields += ", intBuchungsart"
            inscmdValues += ", @intBuchungsart"
            inscmdFields += ", intRGArt"
            inscmdValues += ", @intRGArt"
            inscmdFields += ", strRGArt"
            inscmdValues += ", @strRGArt"
            inscmdFields += ", strOPNr"
            inscmdValues += ", @strOPNr"
            inscmdFields += ", lngDebNbr"
            inscmdValues += ", @lngDebNbr"
            inscmdFields += ", lngDebKtoNbr"
            inscmdValues += ", @lngDebKtoNbr"
            inscmdFields += ", strDebCur"
            inscmdValues += ", @strDebcur"
            inscmdFields += ", lngDebiKST"
            inscmdValues += ", @lngDebiKST"
            inscmdFields += ", dblDebNetto"
            inscmdValues += ", @dblDebNetto"
            inscmdFields += ", dblDebMwSt"
            inscmdValues += ", @dblDebMwSt"
            inscmdFields += ", dblDebBrutto"
            inscmdValues += ", @dblDebBrutto"
            inscmdFields += ", lngDebIdentNbr"
            inscmdValues += ", @lngDebIdentNbr"
            inscmdFields += ", strDebText"
            inscmdValues += ", @strDebText"
            inscmdFields += ", strDebReferenz"
            inscmdValues += ", @strDebReferenz"
            inscmdFields += ", datDebRGDatum"
            inscmdValues += ", @datDebRGDatum"
            inscmdFields += ", datDebValDatum"
            inscmdValues += ", @datDebValDatum"
            inscmdFields += ", datRGCreate"
            inscmdValues += ", @datRGCreate"
            inscmdFields += ", intPayType"
            inscmdValues += ", @intPayType"
            inscmdFields += ", strDebiBank"
            inscmdValues += ", @strDebiBank"
            inscmdFields += ", lngLinkedRG"
            inscmdValues += ", @lngLinkedRG"
            inscmdFields += ", strRGName"
            inscmdValues += ", @strRGName"
            inscmdFields += ", strDebIdentNbr2"
            inscmdValues += ", @strDebIdentNbr2"
            inscmdFields += ", booCrToInv"
            inscmdValues += ", @booCrToInv"
            inscmdFields += ", datPGVFrom"
            inscmdValues += ", @datPGVFrom"
            inscmdFields += ", datPGVTo"
            inscmdValues += ", @datPGVTo"


            'Ins cmd DebiHead
            mysqlinscmd.CommandText = "INSERT INTO tbldebitorenjhead (" + inscmdFields + ") VALUES (" + inscmdValues + ")"
            mysqlinscmd.Parameters.Add("@IdentityName", MySqlDbType.String).SourceColumn = "IdentityName"
            mysqlinscmd.Parameters.Add("@ProcessID", MySqlDbType.Int16).SourceColumn = "ProcessID"
            mysqlinscmd.Parameters.Add("@intBuchhaltung", MySqlDbType.Int16).SourceColumn = "intBuchhaltung"
            mysqlinscmd.Parameters.Add("@strDebRGNbr", MySqlDbType.String).SourceColumn = "strDebRGNbr"
            mysqlinscmd.Parameters.Add("@intBuchungsart", MySqlDbType.Int16).SourceColumn = "intBuchungsart"
            mysqlinscmd.Parameters.Add("@intRGArt", MySqlDbType.Int16).SourceColumn = "intRGArt"
            mysqlinscmd.Parameters.Add("@strRGArt", MySqlDbType.String).SourceColumn = "strRGArt"
            mysqlinscmd.Parameters.Add("@strOPNr", MySqlDbType.String).SourceColumn = "strOPNr"
            mysqlinscmd.Parameters.Add("@lngDebNbr", MySqlDbType.Int32).SourceColumn = "lngDebNbr"
            mysqlinscmd.Parameters.Add("@lngDebKtoNbr", MySqlDbType.Int32).SourceColumn = "lngDebKtoNbr"
            mysqlinscmd.Parameters.Add("@strDebCur", MySqlDbType.String).SourceColumn = "strDebCur"
            mysqlinscmd.Parameters.Add("@lngDebiKST", MySqlDbType.Int32).SourceColumn = "lngDebiKST"
            mysqlinscmd.Parameters.Add("@dblDebNetto", MySqlDbType.Decimal).SourceColumn = "dblDebNetto"
            mysqlinscmd.Parameters.Add("@dblDebMwst", MySqlDbType.Decimal).SourceColumn = "dblDebMwSt"
            mysqlinscmd.Parameters.Add("@dblDebBrutto", MySqlDbType.Decimal).SourceColumn = "dblDebBrutto"
            mysqlinscmd.Parameters.Add("@strDebText", MySqlDbType.String).SourceColumn = "strDebText"
            mysqlinscmd.Parameters.Add("@lngDebIdentNbr", MySqlDbType.Int32).SourceColumn = "lngDebIdentNbr"
            mysqlinscmd.Parameters.Add("@strDebReferenz", MySqlDbType.String).SourceColumn = "strDebReferenz"
            mysqlinscmd.Parameters.Add("@datDebRGDatum", MySqlDbType.Date).SourceColumn = "datDebRGDatum"
            mysqlinscmd.Parameters.Add("@datDebValDatum", MySqlDbType.Date).SourceColumn = "datDebValDatum"
            mysqlinscmd.Parameters.Add("@datRGCreate", MySqlDbType.Date).SourceColumn = "datRGCreate"
            mysqlinscmd.Parameters.Add("@intPayType", MySqlDbType.Int16).SourceColumn = "intPayType"
            mysqlinscmd.Parameters.Add("@strDebiBank", MySqlDbType.String).SourceColumn = "strDebiBank"
            mysqlinscmd.Parameters.Add("@lngLinkedRG", MySqlDbType.Int32).SourceColumn = "lngLinkedRG"
            mysqlinscmd.Parameters.Add("@strRGName", MySqlDbType.String).SourceColumn = "strRGName"
            mysqlinscmd.Parameters.Add("@strDebIdentNbr2", MySqlDbType.String).SourceColumn = "strDebIdentNbr2"
            mysqlinscmd.Parameters.Add("@booCrToInv", MySqlDbType.Int16).SourceColumn = "booCrToInv"
            mysqlinscmd.Parameters.Add("@datPGVFrom", MySqlDbType.Date).SourceColumn = "datPGVFrom"
            mysqlinscmd.Parameters.Add("@datPGVTo", MySqlDbType.Date).SourceColumn = "datPGVTo"

            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem HeadCommand Init", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 1

        End Try

    End Function

    Friend Function FcInitInsCmdKHeads(ByRef mysqlinscmd As MySqlCommand) As Int16

        'Dim strIdentityName As String

        'Kreditoren - Head
        Dim inscmdFields As String
        Dim inscmdValues As String

        Try

            inscmdFields = "IdentityName"
            inscmdValues = "@IdentityName"
            inscmdFields += ", ProcessID"
            inscmdValues += ", @ProcessID"
            inscmdFields += ", intBuchhaltung"
            inscmdValues += ", @intBuchhaltung"
            inscmdFields += ", lngKredID"
            inscmdValues += ", @lngKredID"
            inscmdFields += ", strKredRGNbr"
            inscmdValues += ", @strKredRGNbr"
            inscmdFields += ", intBuchungsart"
            inscmdValues += ", @intBuchungsart"
            inscmdFields += ", strOPNr"
            inscmdValues += ", @strOPNr"
            inscmdFields += ", lngKredNbr"
            inscmdValues += ", @lngKredNbr"
            inscmdFields += ", lngKredKtoNbr"
            inscmdValues += ", @lngKredKtoNbr"
            inscmdFields += ", strKredCur"
            inscmdValues += ", @strKredCur"
            inscmdFields += ", lngKrediKST"
            inscmdValues += ", @lngKrediKST"
            inscmdFields += ", dblKredNetto"
            inscmdValues += ", @dblKredNetto"
            inscmdFields += ", dblKredMwSt"
            inscmdValues += ", @dblKredMwSt"
            inscmdFields += ", dblKredBrutto"
            inscmdValues += ", @dblKredBrutto"
            inscmdFields += ", lngKredIdentNbr"
            inscmdValues += ", @lngKredIdentNbr"
            inscmdFields += ", strKredText"
            inscmdValues += ", @strKredText"
            inscmdFields += ", strKredRef"
            inscmdValues += ", @strKredRef"
            inscmdFields += ", datKredRGDatum"
            inscmdValues += ", @datKredRGDatum"
            inscmdFields += ", datKredValDatum"
            inscmdValues += ", @datKredValDatum"
            inscmdFields += ", intPayType"
            inscmdValues += ", @intPayType"
            inscmdFields += ", strKrediBank"
            inscmdValues += ", @strKrediBank"
            inscmdFields += ", strKrediBankInt"
            inscmdValues += ", @strKrediBankInt"
            inscmdFields += ", strRGBemerkung"
            inscmdValues += ", @strRGBemerkung"
            inscmdFields += ", strRGName"
            inscmdValues += ", @strRGName"
            inscmdFields += ", intZKond"
            inscmdValues += ", @intZKond"
            inscmdFields += ", datPGVFrom"
            inscmdValues += ", @datPGVFrom"
            inscmdFields += ", datPGVTo"
            inscmdValues += ", @datPGVTo"



            'Ins cmd KrediiHead
            mysqlinscmd.CommandText = "INSERT INTO tblkreditorenhead (" + inscmdFields + ") VALUES (" + inscmdValues + ")"
            mysqlinscmd.Parameters.Add("@IdentityName", MySqlDbType.String).SourceColumn = "IdentityName"
            mysqlinscmd.Parameters.Add("@ProcessID", MySqlDbType.Int16).SourceColumn = "ProcessID"
            mysqlinscmd.Parameters.Add("@intBuchhaltung", MySqlDbType.Int16).SourceColumn = "intBuchhaltung"
            mysqlinscmd.Parameters.Add("@lngKredID", MySqlDbType.Int32).SourceColumn = "lngKredID"
            mysqlinscmd.Parameters.Add("@strKredRGNbr", MySqlDbType.String).SourceColumn = "strKredRGNbr"
            mysqlinscmd.Parameters.Add("@intBuchungsart", MySqlDbType.Int16).SourceColumn = "intBuchungsart"
            mysqlinscmd.Parameters.Add("@strOPNr", MySqlDbType.String).SourceColumn = "strOPNr"
            mysqlinscmd.Parameters.Add("@lngKredNbr", MySqlDbType.Int32).SourceColumn = "lngKredNbr"
            mysqlinscmd.Parameters.Add("@lngKredKtoNbr", MySqlDbType.Int32).SourceColumn = "lngKredKtoNbr"
            mysqlinscmd.Parameters.Add("@strKredCur", MySqlDbType.String).SourceColumn = "strKredCur"
            mysqlinscmd.Parameters.Add("@lngKrediKST", MySqlDbType.Int32).SourceColumn = "lngKrediKST"
            mysqlinscmd.Parameters.Add("@dblKredNetto", MySqlDbType.Decimal).SourceColumn = "dblKredNetto"
            mysqlinscmd.Parameters.Add("@dblKredMwSt", MySqlDbType.Decimal).SourceColumn = "dblKredMwSt"
            mysqlinscmd.Parameters.Add("@dblKredBrutto", MySqlDbType.Decimal).SourceColumn = "dblKredBrutto"
            mysqlinscmd.Parameters.Add("@strKredText", MySqlDbType.String).SourceColumn = "strKredText"
            mysqlinscmd.Parameters.Add("@lngKredIdentNbr", MySqlDbType.Int32).SourceColumn = "lngKredIdentNbr"
            mysqlinscmd.Parameters.Add("@strKredRef", MySqlDbType.String).SourceColumn = "strKredRef"
            mysqlinscmd.Parameters.Add("@datKredRGDatum", MySqlDbType.Date).SourceColumn = "datKredRGDatum"
            mysqlinscmd.Parameters.Add("@datKredValDatum", MySqlDbType.Date).SourceColumn = "datKredValDatum"
            mysqlinscmd.Parameters.Add("@intPayType", MySqlDbType.Int16).SourceColumn = "intPayType"
            mysqlinscmd.Parameters.Add("@strKrediBank", MySqlDbType.String).SourceColumn = "strKrediBank"
            mysqlinscmd.Parameters.Add("@strKrediBankInt", MySqlDbType.String).SourceColumn = "strKrediBankInt"
            mysqlinscmd.Parameters.Add("@strRGName", MySqlDbType.String).SourceColumn = "strRGName"
            mysqlinscmd.Parameters.Add("@strRGBemerkung", MySqlDbType.String).SourceColumn = "strRGBemerkung"
            mysqlinscmd.Parameters.Add("@intZKond", MySqlDbType.Int16).SourceColumn = "intZKond"
            mysqlinscmd.Parameters.Add("@datPGVFrom", MySqlDbType.Date).SourceColumn = "datPGVFrom"
            mysqlinscmd.Parameters.Add("@datPGVTo", MySqlDbType.Date).SourceColumn = "datPGVTo"

            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem KHeadCommand Init", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 1

        End Try

    End Function


    Friend Function FcReadFromSettingsII(ByVal strField As String,
                                              ByVal intMandant As Int16) As String

        Dim objdbconn As New MySqlConnection
        Dim objlocdtSetting As New DataTable("tbllocSettings")
        Dim objlocMySQLcmd As New MySqlCommand

        Try

            objlocMySQLcmd.CommandText = "SELECT t_sage_buchhaltungen." + strField + " FROM t_sage_buchhaltungen WHERE Buchh_Nr=" + intMandant.ToString
            'Debug.Print(objlocMySQLcmd.CommandText)
            objdbconn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString")
            objdbconn.Open()
            objlocMySQLcmd.Connection = objdbconn
            objlocdtSetting.Load(objlocMySQLcmd.ExecuteReader)
            objdbconn.Close()
            'Debug.Print("Records" + objlocdtSetting.Rows.Count.ToString)
            'Debug.Print("Return " + objlocdtSetting.Rows(0).Item(0).ToString)
            Return objlocdtSetting.Rows(0).Item(0).ToString


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Einstellung lesen")

        Finally
            objlocdtSetting = Nothing
            objlocMySQLcmd = Nothing
            objdbconn = Nothing

        End Try

    End Function

    Friend Function FcInitAccessConnecation(ByRef objaccesscon As OleDb.OleDbConnection,
                                                   ByVal strMDBName As String) As Int16

        'Access - Connection soll initialisiert werden
        '0 = ok, 1 = nicht ok

        Dim dbProvider, dbSource, dbPathAndFile As String

        Try

            dbProvider = "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
            dbSource = "Data Source="
            dbPathAndFile = "\\sdlc.mssag.ch\Apps\Backends\" + strMDBName + ";Jet OLEDB:System Database=\\sdlc.mssag.ch\Apps\Backends\Workbench.mdw;User ID=HagerR;Persist Security Info=False;"
            objaccesscon.ConnectionString = dbProvider + dbSource + dbPathAndFile
            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return 1

        End Try


    End Function


End Class
