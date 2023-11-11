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

            strRGTableType = Main.FcReadFromSettingsII("Buchh_RGTableType",
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
                objmysqlcomdwritehead.Parameters("@strRGArt").Value = row("strRGArt")
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
                objmysqlcomdwritehead.Parameters("@strDebreferenz").Value = row("strDebReferenz")
                objmysqlcomdwritehead.Parameters("@datDebRGDatum").Value = row("datDebRGDatum")
                objmysqlcomdwritehead.Parameters("@datDebValDatum").Value = row("datDebValDatum")
                objmysqlcomdwritehead.Parameters("@datRGCreate").Value = row("datRGCreate")
                objmysqlcomdwritehead.Parameters("@intPayType").Value = row("intPayType")
                objmysqlcomdwritehead.Parameters("@strDebiBank").Value = row("strDebiBank")
                objmysqlcomdwritehead.Parameters("@lngLinkedRG").Value = row("lngLinkedRG")
                objmysqlcomdwritehead.Parameters("@strRGName").Value = row("strRGName")
                objmysqlcomdwritehead.Parameters("@strDebIdentNbr2").Value = row("strDebIdentNbr2")
                If objdtLocDebiHead.Columns.Contains("booCrToInv") Then
                    objmysqlcomdwritehead.Parameters("@booCrToInv").Value = row("booCrToInv")
                End If
                objmysqlcomdwritehead.ExecuteNonQuery()
                objmysqlcomdwritehead.Connection.Close()

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
                objmysqlcomdwritesub.Parameters("@strArtikel").Value = drsub("strArtikel")
                objmysqlcomdwritesub.ExecuteNonQuery()
                objmysqlcomdwritesub.Connection.Close()

            Next


            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Debitoren-Daten-Lesen Problem", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 1

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
            Application.DoEvents()

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

            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem HeadCommand Init", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
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
