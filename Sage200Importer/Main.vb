Option Strict Off
Option Explicit On

Imports MySql.Data.MySqlClient
Imports System.Data.OracleClient
'Imports System.Data.OleDb

Friend NotInheritable Class Main

    Public Shared Function tblDebitorenHead() As DataTable
        Dim DT As DataTable
        'Dim myNewRow As DataRow
        DT = New DataTable("tblDebitorenHead")
        Dim strDebRGNbr As DataColumn = New DataColumn("strDebRGNbr")
        strDebRGNbr.DataType = System.Type.[GetType]("System.String")
        strDebRGNbr.MaxLength = 50
        DT.Columns.Add(strDebRGNbr)
        DT.PrimaryKey = New DataColumn() {DT.Columns("strDebRGNbr")}
        Dim intBuchhaltung As DataColumn = New DataColumn("intBuchhaltung")
        intBuchhaltung.DataType = System.Type.[GetType]("System.Int32")
        DT.Columns.Add(intBuchhaltung)
        Dim booDebBook As DataColumn = New DataColumn("booDebBook")
        booDebBook.DataType = System.Type.[GetType]("System.Boolean")
        DT.Columns.Add(booDebBook)
        Dim intBuchungsart As DataColumn = New DataColumn("intBuchungsart")
        intBuchungsart.DataType = System.Type.[GetType]("System.Int16")
        DT.Columns.Add(intBuchungsart)
        Dim intRGArt As DataColumn = New DataColumn("intRGArt")
        intRGArt.DataType = System.Type.[GetType]("System.Int32")
        DT.Columns.Add(intRGArt)
        Dim strRGArt As DataColumn = New DataColumn("strRGArt")
        strRGArt.DataType = System.Type.[GetType]("System.String")
        strRGArt.MaxLength = 50
        DT.Columns.Add(strRGArt)
        Dim lngLinkedRG As DataColumn = New DataColumn("lngLinkedRG")
        lngLinkedRG.DataType = System.Type.[GetType]("System.Int32")
        DT.Columns.Add(lngLinkedRG)
        Dim booLinked As DataColumn = New DataColumn("booLinked")
        booLinked.DataType = System.Type.[GetType]("System.Boolean")
        DT.Columns.Add(booLinked)
        Dim strRGName As DataColumn = New DataColumn("strRGName")
        strRGName.DataType = System.Type.[GetType]("System.String")
        strRGName.MaxLength = 50
        DT.Columns.Add(strRGName)
        Dim strOPNr As DataColumn = New DataColumn("strOPNr")
        strOPNr.DataType = System.Type.[GetType]("System.String")
        strOPNr.MaxLength = 13
        DT.Columns.Add(strOPNr)
        Dim lngDebPKNbr As DataColumn = New DataColumn("lngDebNbr")
        lngDebPKNbr.DataType = System.Type.[GetType]("System.Int32")
        DT.Columns.Add(lngDebPKNbr)
        Dim strDebPKBez As DataColumn = New DataColumn("strDebBez")
        strDebPKBez.DataType = System.Type.[GetType]("System.String")
        strDebPKBez.MaxLength = 50
        DT.Columns.Add(strDebPKBez)
        Dim lngDebKtoNbr As DataColumn = New DataColumn("lngDebKtoNbr")
        lngDebKtoNbr.DataType = System.Type.[GetType]("System.Int32")
        DT.Columns.Add(lngDebKtoNbr)
        Dim strDebKtoBez As DataColumn = New DataColumn("strDebKtoBez")
        strDebKtoBez.DataType = System.Type.[GetType]("System.String")
        strDebKtoBez.MaxLength = 50
        DT.Columns.Add(strDebKtoBez)
        Dim strDebCur As DataColumn = New DataColumn("strDebCur")
        strDebCur.DataType = System.Type.[GetType]("System.String")
        strDebCur.MaxLength = 3
        DT.Columns.Add(strDebCur)
        Dim dblDebNetto As DataColumn = New DataColumn("dblDebNetto")
        dblDebNetto.DataType = System.Type.[GetType]("System.Double")
        DT.Columns.Add(dblDebNetto)
        Dim dblDebMwSt As DataColumn = New DataColumn("dblDebMwSt")
        dblDebMwSt.DataType = System.Type.[GetType]("System.Double")
        DT.Columns.Add(dblDebMwSt)
        Dim dblDebBrutto As DataColumn = New DataColumn("dblDebBrutto")
        dblDebBrutto.DataType = System.Type.[GetType]("System.Double")
        DT.Columns.Add(dblDebBrutto)
        Dim intSubBookings As DataColumn = New DataColumn("intSubBookings")
        intSubBookings.DataType = System.Type.[GetType]("System.Int32")
        DT.Columns.Add(intSubBookings)
        Dim dblSumSubBookings As DataColumn = New DataColumn("dblSumSubBookings")
        dblSumSubBookings.DataType = System.Type.[GetType]("System.Double")
        DT.Columns.Add(dblSumSubBookings)
        Dim lngDebIdentNbr As DataColumn = New DataColumn("lngDebIdentNbr")
        lngDebIdentNbr.DataType = System.Type.[GetType]("System.Int32")
        DT.Columns.Add(lngDebIdentNbr)
        Dim strDebIdentNbr2 As DataColumn = New DataColumn("strDebIdentNbr2")
        strDebIdentNbr2.DataType = System.Type.[GetType]("System.String")
        strDebIdentNbr2.MaxLength = 50
        DT.Columns.Add(strDebIdentNbr2)
        Dim strDebText As DataColumn = New DataColumn("strDebText")
        strDebText.DataType = System.Type.[GetType]("System.String")
        strDebText.MaxLength = 50
        DT.Columns.Add(strDebText)
        Dim strRGBemerkung As DataColumn = New DataColumn("strRGBemerkung")
        strRGBemerkung.DataType = System.Type.[GetType]("System.String")
        strRGBemerkung.MaxLength = 50
        DT.Columns.Add(strRGBemerkung)
        Dim datDebRGDatum As DataColumn = New DataColumn("datDebRGDatum")
        datDebRGDatum.DataType = System.Type.[GetType]("System.DateTime")
        DT.Columns.Add(datDebRGDatum)
        Dim datDebValDatum As DataColumn = New DataColumn("datDebValDatum")
        datDebValDatum.DataType = System.Type.[GetType]("System.DateTime")
        DT.Columns.Add(datDebValDatum)
        Dim strDebiBank As DataColumn = New DataColumn("strDebiBank")
        strDebiBank.DataType = System.Type.[GetType]("System.String")
        strDebiBank.MaxLength = 5
        DT.Columns.Add(strDebiBank)
        Dim strDebRef As DataColumn = New DataColumn("strDebRef")
        strDebRef.DataType = System.Type.[GetType]("System.String")
        strDebRef.MaxLength = 27
        DT.Columns.Add(strDebRef)
        Dim strZahlBed As DataColumn = New DataColumn("strZahlBed")
        strZahlBed.DataType = System.Type.[GetType]("System.String")
        strZahlBed.MaxLength = 5
        DT.Columns.Add(strZahlBed)
        Dim strDebStatusBitLog As DataColumn = New DataColumn("strDebStatusBitLog")
        strDebStatusBitLog.DataType = System.Type.[GetType]("System.String")
        strDebStatusBitLog.MaxLength = 50
        DT.Columns.Add(strDebStatusBitLog)
        Dim strDebStatusText As DataColumn = New DataColumn("strDebStatusText")
        strDebStatusText.DataType = System.Type.[GetType]("System.String")
        strDebStatusText.MaxLength = 255
        DT.Columns.Add(strDebStatusText)
        Dim strDebBookStatus As DataColumn = New DataColumn("strDebBookStatus")
        strDebBookStatus.DataType = System.Type.[GetType]("System.String")
        strDebBookStatus.MaxLength = 50
        DT.Columns.Add(strDebBookStatus)
        Dim booBooked As DataColumn = New DataColumn("booBooked")
        booLinked.DataType = System.Type.[GetType]("System.Boolean")
        DT.Columns.Add(booBooked)
        Dim datBooked As DataColumn = New DataColumn("datBooked")
        datBooked.DataType = System.Type.[GetType]("System.DateTime")
        DT.Columns.Add(datBooked)
        Dim lngBelegNr As DataColumn = New DataColumn("lngBelegNr")
        lngBelegNr.DataType = System.Type.[GetType]("System.Int32")
        DT.Columns.Add(lngBelegNr)
        Return DT
    End Function

    Public Shared Function tblDebitorenSub() As DataTable
        Dim DT As DataTable
        DT = New DataTable("tblDebitorenSub")
        Dim lngID As DataColumn = New DataColumn("lngID")
        lngID.DataType = System.Type.[GetType]("System.Int32")
        lngID.AutoIncrement = True
        lngID.AutoIncrementSeed = 1
        lngID.AutoIncrementStep = 1
        DT.Columns.Add(lngID)
        Dim strRGNr As DataColumn = New DataColumn("strRGNr")
        strRGNr.DataType = System.Type.[GetType]("System.String")
        strRGNr.MaxLength = 50
        DT.Columns.Add(strRGNr)
        Dim intSollHaben As DataColumn = New DataColumn("intSollHaben")
        intSollHaben.DataType = System.Type.[GetType]("System.Int16")
        DT.Columns.Add(intSollHaben)
        Dim lngKto As DataColumn = New DataColumn("lngKto")
        lngKto.DataType = System.Type.[GetType]("System.Int32")
        DT.Columns.Add(lngKto)
        Dim strKtoBez As DataColumn = New DataColumn("strKtoBez")
        strKtoBez.DataType = System.Type.[GetType]("System.String")
        strKtoBez.MaxLength = 50
        DT.Columns.Add(strKtoBez)
        Dim lngKST As DataColumn = New DataColumn("lngKST")
        lngKST.DataType = System.Type.[GetType]("System.Int32")
        DT.Columns.Add(lngKST)
        Dim strKstBez As DataColumn = New DataColumn("strKstBez")
        strKstBez.DataType = System.Type.[GetType]("System.String")
        strKstBez.MaxLength = 50
        DT.Columns.Add(strKstBez)
        Dim dblNetto As DataColumn = New DataColumn("dblNetto")
        dblNetto.DataType = System.Type.[GetType]("System.Double")
        DT.Columns.Add(dblNetto)
        Dim dblMwSt As DataColumn = New DataColumn("dblMwSt")
        dblMwSt.DataType = System.Type.[GetType]("System.Double")
        DT.Columns.Add(dblMwSt)
        Dim dblBrutto As DataColumn = New DataColumn("dblBrutto")
        dblBrutto.DataType = System.Type.[GetType]("System.Double")
        DT.Columns.Add(dblBrutto)
        Dim lngMwStSatz As DataColumn = New DataColumn("lngMwStSatz")
        dblBrutto.DataType = System.Type.[GetType]("System.Double")
        DT.Columns.Add(lngMwStSatz)
        Dim strMwStKey As DataColumn = New DataColumn("strMwStKey")
        strMwStKey.DataType = System.Type.[GetType]("System.String")
        strMwStKey.MaxLength = 50
        DT.Columns.Add(strMwStKey)
        Dim strArtikel As DataColumn = New DataColumn("strArtikel")
        strArtikel.DataType = System.Type.[GetType]("System.String")
        strArtikel.MaxLength = 128
        DT.Columns.Add(strArtikel)
        Dim strDebSubText As DataColumn = New DataColumn("strDebSubText")
        strDebSubText.DataType = System.Type.[GetType]("System.String")
        strDebSubText.MaxLength = 50
        DT.Columns.Add(strDebSubText)
        Dim strStatusUBBitLog As DataColumn = New DataColumn("strStatusUBBitLog")
        strStatusUBBitLog.DataType = System.Type.[GetType]("System.String")
        strStatusUBBitLog.MaxLength = 50
        DT.Columns.Add(strStatusUBBitLog)
        Dim strStatusUBText As DataColumn = New DataColumn("strStatusUBText")
        strStatusUBText.DataType = System.Type.[GetType]("System.String")
        strStatusUBText.MaxLength = 255
        DT.Columns.Add(strStatusUBText)
        Dim strDebBookStatus As DataColumn = New DataColumn("strDebBookStatus")
        strDebBookStatus.DataType = System.Type.[GetType]("System.String")
        strDebBookStatus.MaxLength = 50
        DT.Columns.Add(strDebBookStatus)
        Return DT
    End Function


    Public Shared Function FcLoginSage(ByRef objdbconn As MySqlConnection, ByRef objFinanz As SBSXASLib.AXFinanz, ByRef objfiBuha As SBSXASLib.AXiFBhg, ByRef objdbBuha As SBSXASLib.AXiDbBhg, ByVal intAccounting As Int16) As Int16

        '0=ok, 1=Fibu nicht ok, 2=Debi nicht ok, 3=Debi nicht ok

        Dim booAccOk As Boolean
        Dim strMandant As String
        Dim b As Object
        b = Nothing

        objFinanz = Nothing
        objFinanz = New SBSXASLib.AXFinanz


        'On Error GoTo ErrorHandler

        'Loign
        Call objFinanz.ConnectSBSdb(System.Configuration.ConfigurationManager.AppSettings("OwnSageServer"), System.Configuration.ConfigurationManager.AppSettings("OwnSageDB"), System.Configuration.ConfigurationManager.AppSettings("OwnSageID"), System.Configuration.ConfigurationManager.AppSettings("OwnSagePsw"), "")

        strMandant = FcReadFromSettings(objdbconn, "Buchh200_Name", intAccounting)
        booAccOk = objFinanz.CheckMandant(strMandant)

        'Check Periode
        booAccOk = objFinanz.CheckPeriode(strMandant, "2020")

        'Open Mandantg
        objFinanz.OpenMandant(strMandant, "2020")

        If b = 0 Then GoTo isOk
        b = b - 200
        MsgBox("Mandant oder Periode falsch - Programm beendet", 0, "Fehler")
        objFinanz = Nothing
        End

isOk:
        'Finanz Buha öffnen
        objfiBuha = Nothing
        objfiBuha = New SBSXASLib.AXiFBhg
        objfiBuha = objFinanz.GetFibuObj
        'Debitor öffnen
        objdbBuha = Nothing
        objdbBuha = New SBSXASLib.AXiDbBhg
        objdbBuha = objFinanz.GetDebiObj
        'db = Main_Renamed.Finanz.GetDebiObj
        'UPGRADE_WARNING: Couldn't resolve default property of object s. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        's = db.ReadDebitor3(CInt(DebitorID.Text), WhgID.Text)
        's = DbBhg.ReadDebitor3(-1000, "")
        'Debug.Print("Angaben Debitor " + s)
        'UPGRADE_NOTE: Object db may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'db = Nothing

        'UPGRADE_WARNING: Couldn't resolve default property of object s. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'AusgabeText.Text = "ReadDebitor3:" & Chr(13) & Chr(10) & s



        'MsgBox("OpenMandant:" & Chr(13) & Chr(10) & "Funktionierte")
        Exit Function


ErrorHandler:

        b = Err.Number And 65535
        MsgBox("OpenMandant:" & Chr(13) & Chr(10) & "Error" & Chr(13) & Chr(10) & "Die Button auf dem Main wurden ausgeschaltet !!!" & Chr(13) & Chr(10) & "Error # " & Str(Err.Number) & " was generated by " & Err.Source & Chr(13) & Chr(10) & Err.Description & " Unsere Fehlernummer" & Str(b))
        Err.Clear()

    End Function

    Public Shared Function FcFillDebit(ByVal intAccounting As Integer, ByRef objdtHead As DataTable, ByRef objdtSub As DataTable, ByRef objdbconn As MySqlConnection, ByRef objdbAccessConn As OleDb.OleDbConnection) As Integer

        Dim strSQL As String
        Dim strSQLSub As String
        Dim objlocMySQLcmd As New MySqlCommand
        Dim objlocOLEdbcmd As New OleDb.OleDbCommand
        Dim objDTDebiHead As New DataTable
        Dim dbProvider, dbSource, dbPathAndFile As String
        Dim objdrSub As DataRow

        dbProvider = "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
        dbSource = "Data Source="
        dbPathAndFile = "\\sdlc.mssag.ch\Apps\Backends\Daten_Helpdata_Server.mdb;Jet OLEDB:System Database=\\sdlc.mssag.ch\Apps\Backends\Workbench.mdw;User ID=HagerR;"

        'Head Debitzoren löschen
        objdtHead.Clear()

        strSQL = FcReadFromSettings(objdbconn, "Buchh_SQLHead", intAccounting)

        Try

            'objlocMySQLcmd.CommandText = strSQL
            'Access
            objdbAccessConn.ConnectionString = dbProvider + dbSource + dbPathAndFile
            objlocOLEdbcmd.CommandText = strSQL
            objdbAccessConn.Open()
            objlocOLEdbcmd.Connection = objdbAccessConn
            objdtHead.Load(objlocOLEdbcmd.ExecuteReader)
            'objlocMySQLcmd.Connection = objdbconn
            'objDTDebiHead.Load(objlocMySQLcmd.ExecuteReader)
            'Durch die Records steppen und Sub-Tabelle füllen
            For Each row In objdtHead.Rows
                'Debug.Print(strSQLSub)
                If row("intBuchungsart") = 1 Then
                    objdrSub = objdtSub.NewRow()
                    objdrSub("strRGNr") = row("strDebRGNbr")
                    objdrSub("intSollHaben") = 2
                    objdrSub("lngKto") = row("lngDebKtoNbr")
                    objdrSub("dblBrutto") = row("dblDebBrutto")
                    objdrSub("dblNetto") = row("dblDebNetto")
                    objdrSub("dblMwSt") = row("dblDebMwSt")
                    objdrSub("strDebSubText") = row("Betrifft").ToString + " " + row("betrifft1").ToString
                    objdtSub.Rows.Add(objdrSub)
                End If
                strSQLSub = FcSQLParse(FcReadFromSettings(objdbconn, "Buchh_SQLDetail", intAccounting), row("strDebRGNbr"))
                objlocOLEdbcmd.CommandText = strSQLSub
                objdtSub.Load(objlocOLEdbcmd.ExecuteReader)

            Next


        Catch ex As Exception
            MessageBox.Show(ex.Message)

        End Try

    End Function

    Public Shared Function FcSQLParse(ByVal strSQLToParse As String, ByVal strRGNbr As String) As String

        'Funktion setzt in eingelesenem SQL wieder Variablen ein
        Dim intPipePositionBegin, intPipePositionEnd As Integer
        Dim strWork, strField As String

        '| suchen
        If InStr(strSQLToParse, "|") > 0 Then
            'Vorkommen gefunden
            intPipePositionBegin = InStr(strSQLToParse, "|")
            intPipePositionEnd = InStr(intPipePositionBegin + 1, strSQLToParse, "|")
            Do Until intPipePositionBegin = 0
                strField = Mid(strSQLToParse, intPipePositionBegin + 1, intPipePositionEnd - intPipePositionBegin - 1)
                Select Case strField
                    Case "rsDebi.Fields(""RGNr"")"
                        strField = strRGNbr
                        'Case "rsDebiTemp.Fields([strDebPKBez])"
                        '    strField = rsDebiTemp.Fields("strDebPKBez")
                        'Case "rsDebiTemp.Fields([lngDebIdentNbr])"
                        '    strField = rsDebiTemp.Fields("lngDebIdentNbr")
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
                strSQLToParse = Left(strSQLToParse, intPipePositionBegin - 1) & strField & Right(strSQLToParse, Len(strSQLToParse) - intPipePositionEnd)
                'Neuer Anfang suchen für evtl. weitere |
                intPipePositionBegin = InStr(strSQLToParse, "|")
                'intPipePositionBegin = InStr(intPipePositionEnd + 1, strSQLToParse, "|")
                intPipePositionEnd = InStr(intPipePositionBegin + 1, strSQLToParse, "|")
            Loop
        End If

        Return strSQLToParse


    End Function


    Public Shared Function FcReadFromSettings(ByRef objdbconn As MySqlConnection, ByVal strField As String, ByVal intMandant As Int16) As String

        Dim objlocdtSetting As New DataTable("tbllocSettings")
        Dim objlocMySQLcmd As New MySqlCommand

        Try

            objlocMySQLcmd.CommandText = "SELECT buchhaltungen." + strField + " FROM buchhaltungen WHERE Buchh_Nr=" + intMandant.ToString
            'Debug.Print(objlocMySQLcmd.CommandText)
            objlocMySQLcmd.Connection = objdbconn
            objlocdtSetting.Load(objlocMySQLcmd.ExecuteReader)
            'Debug.Print("Records" + objlocdtSetting.Rows.Count.ToString)


        Catch ex As Exception
            MessageBox.Show(ex.Message)

        End Try

        'Debug.Print("Return " + objlocdtSetting.Rows(0).Item(0).ToString)
        Return objlocdtSetting.Rows(0).Item(0).ToString

    End Function

    Public Shared Function FcCheckDebit(ByVal intAccounting As Integer,
                                        ByRef objdtDebits As DataTable,
                                        ByRef objFinanz As SBSXASLib.AXFinanz,
                                        ByRef objfiBuha As SBSXASLib.AXiFBhg,
                                        ByRef objdbBuha As SBSXASLib.AXiDbBhg,
                                        ByRef objdbconn As MySqlConnection,
                                        ByRef objsqlcommand As MySqlCommand,
                                        ByRef objOrdbconn As OracleClient.OracleConnection,
                                        ByRef objOrcommand As OracleClient.OracleCommand) As Integer

        'DebiBitLog 1=PK, 2=Konto, 3=Währung, 4=interne Bank, 5=OP Kopf, 6=RG-Datum, 7=Valuta Datum, 8=Subs, 9=OP doppelt
        Dim strBitLog As String
        Dim intReturnValue As Integer
        Dim strStatus As String

        Try

            For Each row In objdtDebits.Rows

                'Status-String erstellen
                intReturnValue = FcCheckDebitor(row("lngDebNbr"), row("intBuchungsart"), objdbBuha)
                strBitLog = Trim(intReturnValue.ToString)
                intReturnValue = FcCheckKonto(row("lngDebKtoNbr"), objfiBuha)
                strBitLog = strBitLog + Trim(intReturnValue.ToString)
                intReturnValue = FcCheckCurrency(row("strDebCur"), objfiBuha)
                strBitLog = strBitLog + Trim(intReturnValue.ToString)
                'intReturnValue = fcCheckIntBank()
                'Debug.Print("BitLog: " + strBitLog)
                'Status-String auswerten
                'Debitor
                If Left(strBitLog, 1) <> "0" Then
                    strStatus = "Deb"
                    intReturnValue = FcIsDebitorCreatable(objdbconn, objsqlcommand, objOrdbconn, objOrcommand, row("lngDebNbr"), intAccounting, objdbBuha)
                    If intReturnValue = 0 Then
                        strStatus = strStatus + " erstellt"
                    Else
                        strStatus = strStatus + " nicht erstellt."
                    End If
                Else
                    row("strDebBez") = FcReadDebitorName(objdbBuha, row("lngDebNbr"), row("strDebCur"))
                End If
                'Konto
                If Mid(strBitLog, 2, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Kto"
                    row("strDebKtoBez") = "n/a"
                Else
                    row("strDebKtoBez") = FcReadDebitorKName(objfiBuha, row("lngDebKtoNbr"))
                End If
                'Währung
                If Mid(strBitLog, 3, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Cur"
                End If
                'Status schreiben
                row("strDebStatusText") = strBitLog + ", " + strStatus
                'Init
                strBitLog = ""
                strStatus = ""
            Next


        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Function

    Public Shared Function FcReadDebitorKName(ByRef objfiBuha As SBSXASLib.AXiFBhg, ByVal lngDebKtoNbr As Long) As String

        Dim strDebitorKName As String
        Dim strDebitorKAr() As String

        strDebitorKName = objfiBuha.GetKontoInfo(lngDebKtoNbr)

        strDebitorKAr = Split(strDebitorKName, "{>}")

        Return strDebitorKAr(8)

    End Function

    Public Shared Function FcReadDebitorName(ByRef objDbBhg As SBSXASLib.AXiDbBhg, ByVal intDebiNbr As Int32, ByVal strCurrency As String) As String

        Dim strDebitorName As String
        Dim strDebitorAr() As String

        If strCurrency = "" Then

            strDebitorName = objDbBhg.ReadDebitor3(intDebiNbr * -1, strCurrency)

        Else

            strDebitorName = objDbBhg.ReadDebitor3(intDebiNbr, strCurrency)

        End If

        strDebitorAr = Split(strDebitorName, "{>}")

        Return strDebitorAr(0)

    End Function

    Public Shared Function FcIsDebitorCreatable(ByRef objdbconn As MySqlConnection, ByRef objsqlcommand As MySqlCommand, ByRef objOrdbconn As OracleClient.OracleConnection, ByRef objOrcommand As OracleClient.OracleCommand, ByVal lngDebiNbr As Long, ByVal intAccounting As Int32, ByRef objDbBhg As SBSXASLib.AXiDbBhg) As Int16

        'Return: 0=creatable und erstellt, 3=Sage - Suchtext nicht erfasst, 4=Betrieb nicht gefunden, 9=Nicht hinterlegt

        Dim strTableName, strTableType, strDebFieldName, strCompFieldName, strStreetFieldName, strZIPFieldName, strTownFieldName, strSageName, strDebiAccField As String
        Dim intCreatable As Int16
        Dim objdtDebitor As New DataTable

        strTableName = FcReadFromSettings(objdbconn, "Buchh_PKTable", intAccounting)
        strTableType = FcReadFromSettings(objdbconn, "Buchh_PKTableType", intAccounting)
        strDebFieldName = FcReadFromSettings(objdbconn, "Buchh_PKField", intAccounting)
        strCompFieldName = FcReadFromSettings(objdbconn, "Buchh_PKCompany", intAccounting)
        strStreetFieldName = FcReadFromSettings(objdbconn, "Buchh_PKStreet", intAccounting)
        strZIPFieldName = FcReadFromSettings(objdbconn, "Buchh_PKZIP", intAccounting)
        strTownFieldName = FcReadFromSettings(objdbconn, "Buchh_PKTown", intAccounting)
        strSageName = FcReadFromSettings(objdbconn, "Buchh_PKSageName", intAccounting)
        strDebiAccField = FcReadFromSettings(objdbconn, "Buchh_DPKAccount", intAccounting)

        If strTableName <> "" And strDebFieldName <> "" Then

            If strTableType = "O" Then 'Oracle
                'objOrdbconn.Open()
                objOrcommand.CommandText = "SELECT " + strDebFieldName + ", " + strCompFieldName + ", " + strStreetFieldName + ", " + strZIPFieldName + ", " + strTownFieldName + ", " + strSageName + ", " + strDebiAccField +
                                            " FROM " + strTableName + " WHERE " + strDebFieldName + "=" + lngDebiNbr.ToString
                objdtDebitor.Load(objOrcommand.ExecuteReader)
            Else
                'MySQL - Tabelle einlesen

            End If

            'Gefunden?
            If objdtDebitor.Rows.Count > 0 Then
                'Debug.Print("Gefunden, kann erstellt werden")

                intCreatable = FcCreateDebitor(objDbBhg,
                                               objdtDebitor.Rows(0).Item(strDebFieldName),
                                               objdtDebitor.Rows(0).Item(strCompFieldName),
                                               objdtDebitor.Rows(0).Item(strStreetFieldName),
                                               objdtDebitor.Rows(0).Item(strZIPFieldName),
                                               objdtDebitor.Rows(0).Item(strTownFieldName),
                                               objdtDebitor.Rows(0).Item(strDebiAccField))
                Return 0
            Else
                Return 4

            End If


        End If

    End Function

    Public Shared Function FcCreateDebitor(ByRef objDbBhg As SBSXASLib.AXiDbBhg, ByVal intDebitorNbr As Int32, ByVal strDebName As String, ByVal strDebStreet As String, ByVal strDebPLZ As String, ByVal strDebOrt As String, ByVal intDebSammelKto As Int32) As Int16

        Dim strDebCountry As String = "CH"
        Dim strDebCurrency As String = "CHF"
        Dim strDebSprachCode As String = "2055"
        Dim strDebSperren As String = "N"
        Dim intDebErlKto As Integer = 3200
        Dim shrDebZahlK As Short = 1
        Dim intDebToleranzNbr As Integer = 1
        Dim intDebMahnGroup As Integer = 1
        Dim strDebWerbung As String = "N"

        'Debitor erstellen, minimal - Angaben

        Try

            Call objDbBhg.SetCommonInfo2(intDebitorNbr, strDebName, "", strDebStreet, "", "", "", strDebCountry, strDebPLZ, strDebOrt, "", "", "", "", "", strDebCurrency, "", "", "", strDebSprachCode, "")
            Call objDbBhg.SetExtendedInfo2(strDebSperren, "", intDebSammelKto.ToString, intDebErlKto.ToString, "", "", "", shrDebZahlK.ToString, intDebToleranzNbr.ToString, intDebMahnGroup.ToString, "", "", strDebWerbung, "")
            Call objDbBhg.WriteDebitor3(0)

            Return 0
            'intDebAdrLaufN = DbBhg.GetAdressLaufnr()
            'intDebBankLaufNr = DbBhg.GetZahlungsverbLaufnr()

        Catch ex As Exception
            MessageBox.Show(ex.Message)

            Return 1

        End Try



    End Function

    Public Shared Function FcCheckCurrency(ByVal strCurrency As String, ByRef objfiBuha As SBSXASLib.AXiFBhg) As Integer

        Dim strReturn As String
        Dim booFoundCurrency As Boolean

        booFoundCurrency = False
        strReturn = ""

        Call objfiBuha.ReadWhg()

        strReturn = objfiBuha.GetWhgZeile()
        Do While strReturn <> "EOF"
            If Left(strReturn, 3) = strCurrency Then
                booFoundCurrency = True
            End If
            strReturn = objfiBuha.GetWhgZeile()
        Loop

        If booFoundCurrency Then
            Return 0
        Else
            Return 1
        End If


    End Function

    Public Shared Function FcCheckKonto(ByVal lngKtoNbr As Long, ByRef objfiBuha As SBSXASLib.AXiFBhg) As Integer

        Dim strReturn As String

        strReturn = objfiBuha.GetKontoInfo(lngKtoNbr.ToString)
        If strReturn = "EOF" Then
            Return 1
        Else
            Return 0
        End If

    End Function


    Public Shared Function FcCheckDebitor(ByVal lngDebitor As Long, ByVal intBuchungsart As Integer, ByRef objdbBuha As SBSXASLib.AXiDbBhg) As Integer

        Dim strReturn As String

        If intBuchungsart = 1 Then 'OP Buchung

            strReturn = objdbBuha.ReadDebitor3(lngDebitor * -1, "")
            If strReturn = "EOF" Then
                Return 1
            Else
                Return 0
            End If
        Else
            Return 0

        End If

    End Function
    Public Shared Function InsertDataTableColumnName(ByRef dtSouce As DataTable, ByRef dtResult As DataTable) As Boolean
        Dim rowResult As DataRow
        Dim Result As Boolean = True

        Try

            For Each rowSource As DataRow In dtSouce.Rows
                rowResult = dtResult.NewRow()

                For Each ColumnSource As DataColumn In rowSource.Table.Columns
                    Dim ColumnResult As DataColumnCollection = dtResult.Columns

                    If ColumnResult.Contains(ColumnSource.ColumnName) Then
                        rowResult(ColumnSource.ColumnName) = rowSource(ColumnSource.ColumnName)
                    End If
                Next
                dtResult.Rows.Add(rowResult)
            Next

            Return Result
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Result = False
            Return Result
        End Try
    End Function

    Public Shared Function FcSetBuchMode(ByRef objdbBuha As SBSXASLib.AXiDbBhg, ByVal strMode As String) As Int16

        objdbBuha.SetBuchMode(strMode)

        Return 0

    End Function

    Public Shared Function FcSetBelegKopf4(ByRef objdbBuha As SBSXASLib.AXiDbBhg,
                                           ByVal lngBelegNr As Long,
                                           ByVal strValutaDatum As String,
                                           ByVal lngDebitor As Long,
                                           ByVal strBelegTyp As String,
                                           ByVal strBelegDatum As String,
                                           ByVal strVerFallDatum As String,
                                           ByVal strBelegText As String,
                                           ByVal strReferenz As String,
                                           ByVal lngKondition As Long,
                                           ByVal strSachbearbeiter As String,
                                           ByVal strVerkaeufer As String,
                                           ByVal strMahnSperre As String,
                                           ByVal shrMahnstufe As Short,
                                           ByVal strBetraBrutto As String,
                                           ByVal strKurs As String,
                                           ByVal strBelegExt As String,
                                           ByVal strSKonto As String,
                                           ByVal strDebiCur As String,
                                           ByVal strSammelKonto As String,
                                           ByVal strVerzugsZ As String,
                                           ByVal strZusatzText As String,
                                           ByVal strEBankKonto As String,
                                           ByVal strIkoDebitor As String) As Integer


        'Zuerst prüfen ob Zwingende Werte angegeben worden sind

        'Ausführung
        objdbBuha.SetBelegKopf4(lngBelegNr, strValutaDatum, lngDebitor, strBelegTyp, strBelegDatum, strVerFallDatum, strBelegText, strReferenz, lngKondition, strSachbearbeiter, strVerkaeufer, strMahnSperre, shrMahnstufe, strBetraBrutto,
                                strKurs, strBelegExt, strSKonto, strDebiCur, strSammelKonto, strVerzugsZ, strZusatzText, strEBankKonto, strIkoDebitor)

        Return 0

    End Function

    Public Shared Function FcSetVerteilung(ByRef objdbBuha As SBSXASLib.AXiDbBhg,
                                           ByVal strGegenKonto As String,
                                           ByVal strFibuText As String,
                                           ByVal strNettoBetrag As String,
                                           ByVal strArraySteuer As String,
                                           ByVal strArrayKST As String,
                                           ByVal strArrayKSTE As String) As Integer

        'Prüfen ob Daten vollständig

        'Ausführung
        objdbBuha.SetVerteilung(strGegenKonto, strFibuText, strNettoBetrag, strArraySteuer, strArrayKST, strArrayKSTE)

        Return 0

    End Function

    Public Shared Function FcWriteBuchung(ByRef objdbBuha As SBSXASLib.AXiDbBhg) As Integer

        'Ausführung
        objdbBuha.WriteBuchung()

        Return 0

    End Function

End Class
