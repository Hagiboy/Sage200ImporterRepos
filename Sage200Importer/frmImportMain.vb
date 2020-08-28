Option Strict Off
Option Explicit On

Imports MySql.Data.MySqlClient
Imports System.Data.OracleClient
'Imports System.Data.OleDb


Friend Class frmImportMain

    Public Finanz As SBSXASLib.AXFinanz
    Public FBhg As SBSXASLib.AXiFBhg
    Public DbBhg As SBSXASLib.AXiDbBhg
    Public KrBhg As SBSXASLib.AXiKrBhg
    Public BsExt As SBSXASLib.AXiBSExt
    Public Adr As SBSXASLib.AXiAdr
    Public BeBu As SBSXASLib.AXiBeBu
    Public PIFin As SBSXASLib.AXiPlFin

    Public Methode As String
    Public DidOpenmandant As Boolean

    Public FELD_SEP As String
    Public REC_SEP As String
    Public KSTKTR_SEP As String
    Public FELD_SEP_OUT As String
    Public REC_SEP_OUT As String
    Public nID As String

    Public objdbConn As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString"))
    Public objdbAccessConn As New OleDb.OleDbConnection
    Public objdbcommand As New MySqlCommand
    Public objDABuchhaltungen As New MySqlDataAdapter("SELECT * FROM buchhaltungen WHERE NOT Buchh200_Name IS NULL", objdbConn)
    'Public objDACarsGrid As New MySqlDataAdapter("SELECT tblcars.idCar, tblunits.strUnit, tblplates.strPlate, tblcars.strVIN, tblmodelle.strModell FROM tblcars LEFT JOIN tblunits ON tblcars.refUnit = tblunits.idUnit LEFT JOIN tblplates ON tblcars.refPlate = tblplates.idPlate LEFT JOIN tblmodelle ON tblcars.refModell = tblmodelle.idModell", objdbConn)
    'Public objdtDebitor As New DataTable("tbliDebitor")
    Public objdtBuchhaltungen As New DataTable("tbliBuchhaltungen")
    Public objdtDebitorenHead As New DataTable("tbliDebiHead")
    Public objdtDebitorenHeadRead As New DataTable("tbliDebitorenHeadR")
    Public objdtDebitorenSub As New DataTable("tbliDebiSub")
    Public objOracleConn As New OracleConnection("Data Source=(DESCRIPTION=" _
                    + "(ADDRESS=(PROTOCOL=TCP)(HOST=10.0.0.29)(PORT=1521))" _
                    + "(CONNECT_DATA=(SERVICE_NAME=CISNEW)));" _
                    + "User Id=cis;Password=sugus;")
    Public objOracleCmd As New OracleCommand()

    Public Sub InitVar()

        PIFin = Nothing
        KrBhg = Nothing
        FBhg = Nothing
        DbBhg = Nothing
        BsExt = Nothing
        BeBu = Nothing
        Adr = Nothing
        Finanz = Nothing

        'Call Check_CheckStateChanged(Check, New System.EventArgs())

        FELD_SEP = "{<}"
        REC_SEP = "{>}"
        KSTKTR_SEP = "{-}"

        FELD_SEP_OUT = "{>}"
        REC_SEP_OUT = "{<}"

        'AXFinanzForm.rec1.Text = REC_SEP
        'AXFinanzForm.feld1.Text = FELD_SEP
        'AXFinanzForm.kst1.Text = KSTKTR_SEP

        'AXFinanzForm.rec2.Text = REC_SEP_OUT
        'AXFinanzForm.feld2.Text = FELD_SEP_OUT

        'lblVersion.Text = "SBSxas V-" & Version
    End Sub

    Private Sub butDebitoren_Click(sender As Object, e As EventArgs) Handles butDebitoren.Click

        Dim strIncrBelNbr As String = ""

        '        Dim booAccOk As Boolean
        '        Dim strMandant As String
        '        Dim b As Object
        '       Dim s As Object
        '       b = Nothing
        '       On Error GoTo ErrorHandler

        Me.Cursor = Cursors.WaitCursor

        objdtDebitorenHead.Clear()
        objdtDebitorenSub.Clear()
        objdtDebitorenHeadRead.Clear()

        Call InitVar()

        Call Main.FcLoginSage(objdbConn, Finanz, FBhg, DbBhg, PIFin, cmbBuha.SelectedValue)

        Call Main.FcFillDebit(cmbBuha.SelectedValue, objdtDebitorenHeadRead, objdtDebitorenSub, objdbConn, objdbAccessConn)

        'Call InitdgvDebitoren()
        Call Main.InsertDataTableColumnName(objdtDebitorenHeadRead, objdtDebitorenHead)

        'Grid neu aufbauen
        dgvDebitorenSub.Update()
        dgvDebitoren.Update()
        dgvDebitoren.Refresh()
        'dgvDebitoren.DataSource = objdtDebitorenHead
        'Debug.Print(objdtDebitorenHead.Rows.Count.ToString)
        'Call InitdgvDebitoren()

        Call Main.FcCheckDebit(cmbBuha.SelectedValue, objdtDebitorenHead, objdtDebitorenSub, Finanz, FBhg, DbBhg, PIFin, objdbConn, objdbcommand, objOracleConn, objOracleCmd)

        'strIncrBelNbr = DbBhg.IncrBelNbr
        'Debug.Print("Increment " + strIncrBelNbr)

        'Call 

        'Debug.Print("Gewählt " + cmbBuha.SelectedValue.ToString)

        'Vorübergehend
        'strMandant = "ZZ"

        'Finanz = Nothing
        'Finanz = New SBSXASLib.AXFinanz

        'Loign
        'Call Finanz.ConnectSBSdbNoPrompt("sage_Sage200", "Sage200", "sage200admin", "sage200", "")
        'Call Finanz.ConnectSBSdb("ZHAP03", "Sage200", "sage200admin", "sage200", "")

        'Check Mandant
        'booAccOk = Finanz.CheckMandant(strMandant)
        'Debug.Print("Ok " + booAccOk.ToString)

        'Check Access-Level
        'booAccOk = Finanz.CheckAccess(0, strMandant) 'admin
        'booAccOk = Finanz.CheckAccess(1, strMandant) 'hauptbuch
        'booAccOk = Finanz.CheckAccess(2, strMandant) 'debi
        'booAccOk = Finanz.CheckAccess(3, strMandant) 'kredi
        'booAccOk = Finanz.CheckAccess(4, strMandant) 'lohn
        'booAccOk = Finanz.CheckAccess(14, strMandant) 'darlehen

        'Check Periode
        'booAccOk = Finanz.CheckPeriode(strMandant, "2020")

        'Open Mandantg
        'Finanz.OpenMandant(strMandant, "2020")

        'If b = 0 Then GoTo isOk
        'b = b - 200
        'MsgBox("Mandant oder Periode falsch - Programm beendet", 0, "Fehler")
        'Finanz = Nothing
        'End

        'isOk:
        'Debitor öffnen und Konto überprüfen
        'Dim db As SBSXASLib.AXiDbBhg
        'DbBhg = Nothing
        'DbBhg = Finanz.GetDebiObj
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
        'in Cells ToolTip setzen
        Dim ToolTipAr() As DataRow
        For Each row In dgvDebitoren.Rows
            row.Cells(0).ToolTipText = objdtDebitorenSub.Columns("strRGNr").Caption + vbTab + objdtDebitorenSub.Columns("intSollHaben").Caption + vbTab + objdtDebitorenSub.Columns("lngKto").Caption + vbTab +
                objdtDebitorenSub.Columns("strKtoBez").Caption + vbTab + objdtDebitorenSub.Columns("lngKST").Caption + vbTab + objdtDebitorenSub.Columns("strKSTBez").Caption + vbTab + objdtDebitorenSub.Columns("dblNetto").Caption +
                vbTab + objdtDebitorenSub.Columns("dblMwSt").Caption + vbTab + objdtDebitorenSub.Columns("dblBrutto").Caption + vbTab + objdtDebitorenSub.Columns("lngMwStSatz").Caption +
                vbTab + objdtDebitorenSub.Columns("strDebSubText").Caption + "/ " + objdtDebitorenSub.Columns("strStatusUBText").Caption
            ToolTipAr = objdtDebitorenSub.Select("strRGNr='" + row.Cells(0).Value + "' AND intSollHaben<2")
            For Each ttrow In ToolTipAr
                row.Cells(0).ToolTipText = row.Cells(0).ToolTipText + vbCrLf + ttrow("strRGNr") + vbTab + ttrow("intSollHaben").ToString + vbTab + ttrow("lngKto").ToString + vbTab + ttrow("strKtoBez") + vbTab + ttrow("lngKST").ToString +
                    vbTab + ttrow("strKSTBez") + vbTab + ttrow("dblNetto").ToString + vbTab + ttrow("dblMwSt").ToString + vbTab + ttrow("dblBrutto").ToString + vbTab + ttrow("lngMwStSatz").ToString + vbTab + ttrow("strDebSubText") +
                    "/ " + ttrow("strStatusUBText")
            Next
        Next
        Me.Cursor = Cursors.Default
        Exit Sub

        'ErrorHandler:
        'UPGRADE_WARNING: Couldn't resolve default property of object b. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        b = Err.Number And 65535
        'UPGRADE_WARNING: Couldn't resolve default property of object b. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        MsgBox("OpenMandant:" & Chr(13) & Chr(10) & "Error" & Chr(13) & Chr(10) & "Die Button auf dem Main wurden ausgeschaltet !!!" & Chr(13) & Chr(10) & "Error # " & Str(Err.Number) & " was generated by " & Err.Source & Chr(13) & Chr(10) & Err.Description & " Unsere Fehlernummer" & Str(b))
        '       Err.Clear()


    End Sub

    Private Sub InitdgvDebitoren()

        dgvDebitoren.ShowCellToolTips = True
        dgvDebitoren.AllowUserToAddRows = False
        dgvDebitoren.AllowUserToDeleteRows = False
        dgvDebitoren.Columns("booDebBook").DisplayIndex = 0
        dgvDebitoren.Columns("booDebBook").HeaderText = "ok"
        dgvDebitoren.Columns("booDebBook").Width = 40
        dgvDebitoren.Columns("booDebBook").ValueType = System.Type.[GetType]("System.Boolean")
        dgvDebitoren.Columns("booDebBook").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvDebitoren.Columns("booDebBook").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvDebitoren.Columns("strDebRGNbr").DisplayIndex = 1
        dgvDebitoren.Columns("strDebRGNbr").HeaderText = "RG-Nr"
        dgvDebitoren.Columns("strDebRGNbr").Width = 60
        dgvDebitoren.Columns("strDebRGNbr").ReadOnly = True
        dgvDebitoren.Columns("lngDebNbr").DisplayIndex = 2
        dgvDebitoren.Columns("lngDebNbr").HeaderText = "Debitor"
        dgvDebitoren.Columns("lngDebNbr").Width = 60
        dgvDebitoren.Columns("strDebBez").DisplayIndex = 3
        dgvDebitoren.Columns("strDebBez").HeaderText = "Bezeichnung"
        dgvDebitoren.Columns("strDebBez").Width = 140
        dgvDebitoren.Columns("lngDebKtoNbr").DisplayIndex = 4
        dgvDebitoren.Columns("lngDebKtoNbr").HeaderText = "Konto"
        dgvDebitoren.Columns("lngDebKtoNbr").Width = 50
        dgvDebitoren.Columns("strDebKtoBez").DisplayIndex = 5
        dgvDebitoren.Columns("strDebKtoBez").HeaderText = "Bezeichnung"
        dgvDebitoren.Columns("strDebKtoBez").Width = 150
        dgvDebitoren.Columns("strDebCur").DisplayIndex = 6
        dgvDebitoren.Columns("strDebCur").HeaderText = "Währung"
        dgvDebitoren.Columns("strDebCur").Width = 60
        dgvDebitoren.Columns("dblDebNetto").DisplayIndex = 7
        dgvDebitoren.Columns("dblDebNetto").HeaderText = "Netto"
        dgvDebitoren.Columns("dblDebNetto").Width = 80
        dgvDebitoren.Columns("dblDebNetto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvDebitoren.Columns("dblDebNetto").ReadOnly = True
        dgvDebitoren.Columns("dblDebMwSt").DisplayIndex = 8
        dgvDebitoren.Columns("dblDebMwSt").HeaderText = "MwSt"
        dgvDebitoren.Columns("dblDebMwSt").Width = 70
        dgvDebitoren.Columns("dblDebMwSt").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvDebitoren.Columns("dblDebBrutto").DisplayIndex = 9
        dgvDebitoren.Columns("dblDebBrutto").HeaderText = "Brutto"
        dgvDebitoren.Columns("dblDebBrutto").Width = 80
        dgvDebitoren.Columns("dblDebBrutto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvDebitoren.Columns("intSubBookings").DisplayIndex = 10
        dgvDebitoren.Columns("intSubBookings").HeaderText = "Sub"
        dgvDebitoren.Columns("intSubBookings").Width = 50
        dgvDebitoren.Columns("intSubBookings").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvDebitoren.Columns("dblSumSubBookings").DisplayIndex = 11
        dgvDebitoren.Columns("dblSumSubBookings").HeaderText = "Sub-Summe"
        dgvDebitoren.Columns("dblSumSubBookings").Width = 80
        dgvDebitoren.Columns("dblSumSubBookings").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvDebitoren.Columns("lngDebIdentNbr").DisplayIndex = 12
        dgvDebitoren.Columns("lngDebIdentNbr").HeaderText = "Ident"
        dgvDebitoren.Columns("lngDebIdentNbr").Width = 80
        Dim cmbBuchungsart As New DataGridViewComboBoxColumn()
        Dim objdtBA As New DataTable("objidtBA")
        Dim objlocMySQLcmd As New MySqlCommand
        objlocMySQLcmd.CommandText = "SELECT * FROM tblBuchungsarten"
        objlocMySQLcmd.Connection = objdbConn
        objdtBA.Load(objlocMySQLcmd.ExecuteReader)
        cmbBuchungsart.DataSource = objdtBA
        cmbBuchungsart.DisplayMember = "strBuchungsart"
        cmbBuchungsart.ValueMember = "idBuchungsart"
        cmbBuchungsart.HeaderText = "BA"
        cmbBuchungsart.Name = "intBuchungsart"
        cmbBuchungsart.DataPropertyName = "intBuchungsart"
        cmbBuchungsart.DisplayIndex = 13
        cmbBuchungsart.Width = 60
        dgvDebitoren.Columns.Add(cmbBuchungsart)
        'dgvDebitoren.Columns("intBuchungsart").DisplayIndex = 13
        'dgvDebitoren.Columns("intBuchungsart").DisplayIndex = 13
        'dgvDebitoren.Columns("intBuchungsart").HeaderText = "BA"
        'dgvDebitoren.Columns("intBuchungsart").Width = 40
        dgvDebitoren.Columns("strOPNr").DisplayIndex = 14
        dgvDebitoren.Columns("strOPNr").HeaderText = "OP-Nr"
        dgvDebitoren.Columns("strOPNr").Width = 80
        dgvDebitoren.Columns("datDebRGDatum").DisplayIndex = 15
        dgvDebitoren.Columns("datDebRGDatum").HeaderText = "RG Datum"
        dgvDebitoren.Columns("datDebRGDatum").Width = 70
        dgvDebitoren.Columns("datDebValDatum").DisplayIndex = 16
        dgvDebitoren.Columns("datDebValDatum").HeaderText = "Val Datum"
        dgvDebitoren.Columns("datDebValDatum").Width = 70
        dgvDebitoren.Columns("strDebiBank").DisplayIndex = 17
        dgvDebitoren.Columns("strDebiBank").HeaderText = "Bank"
        dgvDebitoren.Columns("strDebiBank").Width = 60
        dgvDebitoren.Columns("strDebStatusText").DisplayIndex = 18
        dgvDebitoren.Columns("strDebStatusText").HeaderText = "Status"
        dgvDebitoren.Columns("strDebStatusText").Width = 200
        dgvDebitoren.Columns("intBuchhaltung").Visible = False
        dgvDebitoren.Columns("intBuchungsart").Visible = False
        dgvDebitoren.Columns("intRGArt").Visible = False
        dgvDebitoren.Columns("strRGArt").Visible = False
        dgvDebitoren.Columns("lngLinkedRG").Visible = False
        dgvDebitoren.Columns("booLinked").Visible = False
        dgvDebitoren.Columns("strRGName").Visible = False
        dgvDebitoren.Columns("strDebIdentnbr2").Visible = False
        dgvDebitoren.Columns("strDebText").Visible = False
        dgvDebitoren.Columns("strRGBemerkung").Visible = False
        dgvDebitoren.Columns("strDebRef").Visible = False
        dgvDebitoren.Columns("strZahlBed").Visible = False
        dgvDebitoren.Columns("strDebStatusBitLog").Visible = False
        dgvDebitoren.Columns("strDebBookStatus").Visible = False
        dgvDebitoren.Columns("booBooked").Visible = False
        dgvDebitoren.Columns("datBooked").Visible = False
        dgvDebitoren.Columns("lngBelegNr").Visible = False


    End Sub


    Private Sub frmImportMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'MySQL - Connection öffnen
        objdbConn.Open()
        'Oracle - Connection öffnen
        'objOracleConn.ConnectionString = strOraDB
        objOracleConn.Open()
        objOracleCmd.Connection = objOracleConn

        'Comboxen
        objdtBuchhaltungen.Clear()
        objDABuchhaltungen.Fill(objdtBuchhaltungen)
        'cmbMarken.Sorted = True
        cmbBuha.DataSource = objdtBuchhaltungen
        cmbBuha.DisplayMember = "Buchh_Bez"
        cmbBuha.ValueMember = "Buchh_Nr"

        'Tabelle Debi Head erstellen
        objdtDebitorenHead = Main.tblDebitorenHead()

        'Tabelle Debi Sub erstellen
        objdtDebitorenSub = Main.tblDebitorenSub()

        'Subbuchungen ausblenden, kann für Testzwecke aktiviert werden
        dgvDebitorenSub.DataSource = objdtDebitorenSub
        dgvDebitorenSub.Visible = False

        'DGV
        dgvDebitoren.DataSource = objdtDebitorenHead
        Call InitdgvDebitoren()


    End Sub

    Private Sub butImport_Click(sender As Object, e As EventArgs) Handles butImport.Click


        Dim intReturnValue As Int16
        Dim intDebBelegsNummer As Int32

        Dim intDebitorNbr As Int32
        Dim strBuchType As String
        Dim strBelegDatum As String
        Dim strValutaDatum As String
        Dim strVerfallDatum As String
        Dim strReferenz As String
        Dim intKondition As Int32
        Dim strSachBID As String = ""
        Dim strVerkID As String = ""
        Dim strMahnerlaubnis As String
        Dim sngAktuelleMahnstufe As Single
        Dim dblBetrag As Double
        Dim dblKurs As Double
        Dim strExtBelegNbr As String
        Dim strSkonto As String
        Dim strCurrency As String
        Dim strDebiText As String

        Dim intGegenKonto As Int32
        Dim strFibuText As String
        Dim dblNettoBetrag As Double
        Dim dblBebuBetrag As Double
        Dim strBeBuEintrag As String
        Dim strSteuerFeld As String

        Dim selDebiSub() As DataRow


        Try

            'Debitor erstellen, minimal - Angaben

            'Kopfbuchung
            For Each row In objdtDebitorenHead.Rows

                If IIf(IsDBNull(row("booDebBook")), False, row("booDebBook")) Then

                    'Test ob OP - Buchung
                    If row("intBuchungsart") = 1 Then

                        If IsDBNull(row("strOPNr")) Or row("strOPNr") = "" Then
                            'Zuerst Beleg-Nummerieungung aktivieren
                            DbBhg.IncrBelNbr = "J"
                            'Belegsnummer abholen
                            intDebBelegsNummer = DbBhg.GetNextBelNbr("R")
                        Else
                            'Beleg-Nummerierung abschalten
                            DbBhg.IncrBelNbr = "N"
                            intDebBelegsNummer = row("strOPNr")
                            'strExtBelegNbr = row("strOPNr")
                        End If

                        'Variablen zuweisen
                        intDebitorNbr = row("lngDebNbr")
                        strBuchType = "R"
                        strValutaDatum = Format(row("datDebValDatum"), "yyyyMMdd").ToString
                        strBelegDatum = Format(row("datDebRGDatum"), "yyyyMMdd").ToString
                        strVerfallDatum = ""
                        strReferenz = row("strDebRef")
                        strMahnerlaubnis = "" 'Format(row("datDebRGDatum"), "yyyyMMdd").ToString
                        dblBetrag = row("dblDebBrutto")
                        dblKurs = 1.0#
                        strDebiText = row("strDebText")
                        strCurrency = "CHF"

                        Call DbBhg.SetBelegKopf2(intDebBelegsNummer, strValutaDatum, intDebitorNbr, strBuchType, strBelegDatum, strVerfallDatum, strDebiText, strReferenz, intKondition, strSachBID, strVerkID, strMahnerlaubnis, sngAktuelleMahnstufe, dblBetrag.ToString, dblKurs.ToString, strExtBelegNbr, strSkonto, strCurrency)

                        selDebiSub = objdtDebitorenSub.Select("strRGNr='" + row("strDebRGNbr") + "' AND intSollHaben<>2")

                        For Each SubRow In selDebiSub

                            intGegenKonto = SubRow("lngKto")
                            strFibuText = SubRow("strDebSubText")
                            dblNettoBetrag = SubRow("dblNetto")
                            dblBebuBetrag = 1000.0#
                            strBeBuEintrag = SubRow("lngKST").ToString + "{<}" + SubRow("strDebSubText") + "{<}" + "CALCULATE" + "{>}"    '"PROD{<}BebuText{<}" + dblBebuBetrag.ToString + "{>}"
                            strSteuerFeld = Main.FcGetSteuerFeld(FBhg, SubRow("lngKto"), SubRow("strDebSubText"), SubRow("dblBrutto"), SubRow("strMwStKey"))     '"25{<}DEBI D Bhg Export MwSt{<}0{>}"

                            Call DbBhg.SetVerteilung(intGegenKonto.ToString, strFibuText, dblNettoBetrag.ToString, strSteuerFeld, strBeBuEintrag)

                            'Status Sub schreiben

                        Next


                        Call DbBhg.WriteBuchung()

                    Else

                        'Buchung nur in Fibu
                        'Prinzip Funktion WriteBuchung() anwenden mit allen Parametern

                    End If

                    'Status Head schreiben
                    row("strDebBookStatus") = row("strDebStatusBitLog")
                    row("booBooked") = True
                    row("datBooked") = Now()
                    row("lngBelegNr") = intDebBelegsNummer

                    'Status in File RG-Tabelle schreiben
                    intReturnValue = Main.FcWriteToRGTable(cmbBuha.SelectedValue, row("strDebRGNbr"), row("datBooked"), row("lngBelegNr"), objdbAccessConn)
                    If intReturnValue <> 0 Then
                        'Throw an exception
                    End If

                End If

            Next


        Catch ex As Exception
            MessageBox.Show(ex.Message)

        End Try

    End Sub

    Private Sub dgvDebitoren_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvDebitoren.CellValueChanged

        If e.ColumnIndex = 2 And e.RowIndex > 1 Then

            MsgBox("Geändert " + dgvDebitoren.Rows(e.RowIndex).Cells("strDebRGNbr").Value + ", " + dgvDebitoren.Rows(e.RowIndex).Cells("booDebBook").Value.ToString + Val(dgvDebitoren.Rows(e.RowIndex).Cells("strDebStatusBitLog").Value).ToString)


        End If

    End Sub

    'Private Sub dgvDebitoren_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvDebitoren.CellClick

    ''Nur auf booDebBook reagieren
    'If e.ColumnIndex = 2 Then
    '    'Verhindern das Buchungen mit Fehlern aktiviert werden können
    '    MsgBox("Aktueller Wert " + e.ColumnIndex.ToString + ", " + dgvDebitoren.Rows(e.RowIndex).Cells("booDebBook").Value.ToString)
    'End If

    'End Sub

    'Private Sub dgvDebitoren_MouseUp(sender As Object, e As MouseEventArgs) Handles dgvDebitoren.MouseUp

    '    MsgBox("up")

    'End Sub

    'Private Sub dgvDebitoren_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgvDebitoren.CellEndEdit

    '    If e.ColumnIndex = 2 Then

    '        MsgBox("Angaben " + e.ColumnIndex.ToString)

    '    End If

    'End Sub

    'Private Sub dgvDebitoren_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles dgvDebitoren.CellValidating

    '    If e.ColumnIndex = 2 Then

    '        MsgBox("Angaben " + e.ColumnIndex.ToString)

    '    End If


    'End Sub

    'Private Sub dgvDebitoren_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles dgvDebitoren.CellBeginEdit

    '    If e.ColumnIndex = 2 Then

    '        MsgBox("Angaben " + e.ColumnIndex.ToString)

    '    End If

    'End Sub


End Class
