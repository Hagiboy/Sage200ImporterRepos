Option Strict Off
Option Explicit On

Imports MySql.Data.MySqlClient
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


    Public Shared Function fcLoginSage(ByRef objdbconn As MySqlConnection, ByRef objFinanz As SBSXASLib.AXFinanz, ByRef objfiBuha As SBSXASLib.AXiFBhg, ByRef objdbBuha As SBSXASLib.AXiDbBhg, ByVal intAccounting As Int16) As Int16

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

        strMandant = fcReadFromSettings(objdbconn, "Buchh200_Name", intAccounting)
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

    Public Shared Function fcFillDebit(ByVal intAccounting As Integer, ByRef objdtHead As DataTable, ByRef objdbconn As MySqlConnection, ByRef objdbAccessConn As OleDb.OleDbConnection) As Integer

        Dim strSQL As String
        Dim objlocMySQLcmd As New MySqlCommand
        Dim objlocOLEdbcmd As New OleDb.OleDbCommand
        Dim objDTDebiHead As New DataTable
        Dim dbProvider, dbSource, dbPathAndFile As String

        dbProvider = "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
        dbSource = "Data Source="
        dbPathAndFile = "\\sdlc.mssag.ch\Apps\Backends\Daten_Helpdata_Server.mdb;Jet OLEDB:System Database=\\sdlc.mssag.ch\Apps\Backends\Workbench.mdw;User ID=HagerR;"

        'Head Debitzoren löschen
        objdtHead.Clear()

        strSQL = fcReadFromSettings(objdbconn, "Buchh_SQLHead", intAccounting)

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


        Catch ex As Exception

        End Try

    End Function


    Public Shared Function fcReadFromSettings(ByRef objdbconn As MySqlConnection, ByVal strField As String, ByVal intMandant As Int16) As String

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

    Public Shared Function fcCheckDebit(ByVal intAccounting As Integer, ByRef objdtDebits As DataTable, ByRef objFinanz As SBSXASLib.AXFinanz, ByRef objfiBuha As SBSXASLib.AXiFBhg, ByRef objdbBuha As SBSXASLib.AXiDbBhg) As Integer

        'DebiBitLog 1=PK, 2=Konto, 3=Währung, 4=interne Bank, 5=OP Kopf, 6=RG-Datum, 7=Valuta Datum, 8=Subs, 9=OP doppelt
        Dim strBitLog As String
        Dim intReturnValue As Integer

        Try

            For Each row In objdtDebits.Rows

                intReturnValue = fcCheckDebitor(row("lngDebNbr"), row("intBuchungsart"), objdbBuha)
                strBitLog = Trim(intReturnValue.ToString)
                intReturnValue = fcCheckKonto(row("lngDebKtoNbr"), objfiBuha)
                strBitLog = strBitLog + Trim(intReturnValue.ToString)
                intReturnValue = fcCheckCurrency(row("strDebCur"), objfiBuha)
                strBitLog = strBitLog + Trim(intReturnValue.ToString)
                intReturnValue = fcCheckIntBank()
                Debug.Print("BitLog: " + strBitLog)
            Next

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Function

    Public Shared Function fcCheckCurrency(ByVal strCurrency As String, ByRef objfiBuha As SBSXASLib.AXiFBhg) As Integer

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

    Public Shared Function fcCheckKonto(ByVal lngKtoNbr As Long, ByRef objfiBuha As SBSXASLib.AXiFBhg) As Integer

        Dim strReturn As String

        strReturn = objfiBuha.GetKontoInfo(lngKtoNbr.ToString)
        If strReturn = "EOF" Then
            Return 1
        Else
            Return 0
        End If

    End Function


    Public Shared Function fcCheckDebitor(ByVal lngDebitor As Long, ByVal intBuchungsart As Integer, ByRef objdbBuha As SBSXASLib.AXiDbBhg) As Integer

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

    Public Shared Function fcSetBuchMode(ByRef objdbBuha As SBSXASLib.AXiDbBhg, ByVal strMode As String) As Int16

        objdbBuha.SetBuchMode(strMode)

        Return 0

    End Function

    Public Shared Function fcSetBelegKopf4(ByRef objdbBuha As SBSXASLib.AXiDbBhg,
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

    Public Shared Function fcSetVerteilung(ByRef objdbBuha As SBSXASLib.AXiDbBhg,
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

    Public Shared Function fcWriteBuchung(ByRef objdbBuha As SBSXASLib.AXiDbBhg) As Integer

        'Ausführung
        objdbBuha.WriteBuchung()

        Return 0

    End Function

End Class
