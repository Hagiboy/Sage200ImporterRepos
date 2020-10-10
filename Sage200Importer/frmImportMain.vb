Option Strict Off
Option Explicit On

Imports MySql.Data.MySqlClient
Imports System.Data.OracleClient
Imports System.Data.SqlClient
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
    Public objdbConnZHDB02 As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))
    Public objdbMSSQLConn As New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("SQLConnectionString"))
    Public objdbAccessConn As New OleDb.OleDbConnection
    Public objdbcommand As New MySqlCommand
    Public objdbcommandZHDB02 As New MySqlCommand
    Public objdbSQLcommand As New SqlCommand
    Public objDABuchhaltungen As New MySqlDataAdapter("SELECT * FROM buchhaltungen WHERE NOT Buchh200_Name IS NULL", objdbConn)
    'Public objDACarsGrid As New MySqlDataAdapter("SELECT tblcars.idCar, tblunits.strUnit, tblplates.strPlate, tblcars.strVIN, tblmodelle.strModell FROM tblcars LEFT JOIN tblunits ON tblcars.refUnit = tblunits.idUnit LEFT JOIN tblplates ON tblcars.refPlate = tblplates.idPlate LEFT JOIN tblmodelle ON tblcars.refModell = tblmodelle.idModell", objdbConn)
    'Public objdtDebitor As New DataTable("tbliDebitor")
    Public objdtBuchhaltungen As New DataTable("tbliBuchhaltungen")
    Public objdtDebitorenHead As New DataTable("tbliDebiHead")
    Public objdtDebitorenHeadRead As New DataTable("tbliDebitorenHeadR")
    Public objdtDebitorenSub As New DataTable("tbliDebiSub")
    Public objdtInfo As New DataTable("tbliInfo")
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


        Me.Cursor = Cursors.WaitCursor

        objdtDebitorenHead.Clear()
        objdtDebitorenSub.Clear()
        objdtDebitorenHeadRead.Clear()
        objdtInfo.Clear()

        Call InitVar()

        Call Main.FcLoginSage(objdbConn, objdbMSSQLConn, objdbSQLcommand, Finanz, FBhg, DbBhg, PIFin, cmbBuha.SelectedValue, objdtInfo)

        Call Main.FcFillDebit(cmbBuha.SelectedValue, objdtDebitorenHeadRead, objdtDebitorenSub, objdbConn, objdbAccessConn)

        Call Main.InsertDataTableColumnName(objdtDebitorenHeadRead, objdtDebitorenHead)

        'Grid neu aufbauen
        dgvDebitorenSub.Update()
        dgvDebitoren.Update()
        dgvDebitoren.Refresh()

        Call Main.FcCheckDebit(cmbBuha.SelectedValue, objdtDebitorenHead, objdtDebitorenSub, Finanz, FBhg, DbBhg, PIFin, objdbConn, objdbConnZHDB02, objdbcommand, objdbcommandZHDB02, objOracleConn, objOracleCmd, objdtInfo)

        'Anzahl schreiben
        txtNumber.Text = objdtDebitorenHead.Rows.Count.ToString

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
        'Dim ToolTipAr() As DataRow
        'For Each row In dgvDebitoren.Rows
        '    row.Cells(0).ToolTipText = objdtDebitorenSub.Columns("strRGNr").Caption + vbTab + objdtDebitorenSub.Columns("intSollHaben").Caption + vbTab + objdtDebitorenSub.Columns("lngKto").Caption + vbTab +
        '        objdtDebitorenSub.Columns("strKtoBez").Caption + vbTab + objdtDebitorenSub.Columns("lngKST").Caption + vbTab + objdtDebitorenSub.Columns("strKSTBez").Caption + vbTab + objdtDebitorenSub.Columns("dblNetto").Caption +
        '        vbTab + objdtDebitorenSub.Columns("dblMwSt").Caption + vbTab + objdtDebitorenSub.Columns("dblBrutto").Caption + vbTab + objdtDebitorenSub.Columns("lngMwStSatz").Caption +
        '        vbTab + objdtDebitorenSub.Columns("strDebSubText").Caption + "/ " + objdtDebitorenSub.Columns("strStatusUBText").Caption
        '    ToolTipAr = objdtDebitorenSub.Select("strRGNr='" + row.Cells(0).Value + "' AND intSollHaben<2")
        '    For Each ttrow In ToolTipAr
        '        row.Cells(0).ToolTipText = row.Cells(0).ToolTipText + vbCrLf + ttrow("strRGNr") + vbTab + ttrow("intSollHaben").ToString + vbTab + ttrow("lngKto").ToString + vbTab + ttrow("strKtoBez") + vbTab + ttrow("lngKST").ToString +
        '            vbTab + ttrow("strKSTBez") + vbTab + ttrow("dblNetto").ToString + vbTab + ttrow("dblMwSt").ToString + vbTab + ttrow("dblBrutto").ToString + vbTab + ttrow("lngMwStSatz").ToString + vbTab + ttrow("strDebSubText") +
        '            "/ " + ttrow("strStatusUBText")
        '    Next
        'Next
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

        dgvDebitoren.ShowCellToolTips = False
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
        dgvDebitoren.Columns("dblDebNetto").DefaultCellStyle.Format = "N2"
        dgvDebitoren.Columns("dblDebNetto").ReadOnly = True
        dgvDebitoren.Columns("dblDebMwSt").DisplayIndex = 8
        dgvDebitoren.Columns("dblDebMwSt").HeaderText = "MwSt"
        dgvDebitoren.Columns("dblDebMwSt").Width = 70
        dgvDebitoren.Columns("dblDebMwSt").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvDebitoren.Columns("dblDebMwSt").DefaultCellStyle.Format = "N2"
        dgvDebitoren.Columns("dblDebBrutto").DisplayIndex = 9
        dgvDebitoren.Columns("dblDebBrutto").HeaderText = "Brutto"
        dgvDebitoren.Columns("dblDebBrutto").Width = 80
        dgvDebitoren.Columns("dblDebBrutto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvDebitoren.Columns("dblDebBrutto").DefaultCellStyle.Format = "N2"
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
        cmbBuchungsart.Width = 70
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

    Private Sub InitdgvDebitorenSub()

        dgvDebitorenSub.ShowCellToolTips = False
        dgvDebitorenSub.AllowUserToAddRows = False
        dgvDebitorenSub.AllowUserToDeleteRows = False
        dgvDebitorenSub.Columns("strRGNr").DisplayIndex = 0
        dgvDebitorenSub.Columns("strRGNr").Width = 60
        dgvDebitorenSub.Columns("strRGNr").HeaderText = "RG-Nr"
        dgvDebitorenSub.Columns("intSollHaben").Width = 30
        dgvDebitorenSub.Columns("intSollHaben").HeaderText = "S/H"
        dgvDebitorenSub.Columns("lngKto").Width = 50
        dgvDebitorenSub.Columns("lngKto").HeaderText = "Konto"
        dgvDebitorenSub.Columns("strKtoBez").HeaderText = "Bezeichnung"
        dgvDebitorenSub.Columns("lngKST").Width = 50
        dgvDebitorenSub.Columns("lngKST").HeaderText = "KST"
        dgvDebitorenSub.Columns("strKSTBez").Width = 60
        dgvDebitorenSub.Columns("strKSTBez").HeaderText = "Bezeichnung"
        dgvDebitorenSub.Columns("dblNetto").Width = 60
        dgvDebitorenSub.Columns("dblNetto").HeaderText = "Netto"
        dgvDebitorenSub.Columns("dblNetto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvDebitorenSub.Columns("dblNetto").DefaultCellStyle.Format = "N2"
        dgvDebitorenSub.Columns("dblMwSt").Width = 50
        dgvDebitorenSub.Columns("dblMwSt").HeaderText = "MwSt"
        dgvDebitorenSub.Columns("dblMwSt").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvDebitorenSub.Columns("dblMwSt").DefaultCellStyle.Format = "N2"
        dgvDebitorenSub.Columns("dblBrutto").Width = 60
        dgvDebitorenSub.Columns("dblBrutto").HeaderText = "Brutto"
        dgvDebitorenSub.Columns("dblBrutto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvDebitorenSub.Columns("dblBrutto").DefaultCellStyle.Format = "N2"
        dgvDebitorenSub.Columns("lngMwStSatz").Width = 50
        dgvDebitorenSub.Columns("lngMwStSatz").HeaderText = "MwStS"
        dgvDebitorenSub.Columns("strMwStKey").Width = 40
        dgvDebitorenSub.Columns("strMwStKey").HeaderText = "MwStK"
        dgvDebitorenSub.Columns("strStatusUBText").HeaderText = "Status"

        dgvDebitorenSub.Columns("lngID").Visible = False
        dgvDebitorenSub.Columns("strArtikel").Visible = False
        dgvDebitorenSub.Columns("strStatusUBBitLog").Visible = False
        dgvDebitorenSub.Columns("strDebSubText").Visible = False
        dgvDebitorenSub.Columns("strDebBookStatus").Visible = False

    End Sub


    Private Sub frmImportMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'MySQL - Connection öffnen
        'objdbConn.Open()
        'Oracle - Connection öffnen
        'objOracleConn.ConnectionString = strOraDB
        'objOracleConn.Open()
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

        'Info - Tabelle erstellen
        objdtInfo = Main.tblInfo()

        'Subbuchungen ausblenden, kann für Testzwecke aktiviert werden
        dgvDebitorenSub.DataSource = objdtDebitorenSub
        'dgvDebitorenSub.Visible = False

        'DGV - Info
        dgvInfo.DataSource = objdtInfo
        dgvInfo.AllowUserToAddRows = False
        dgvInfo.AllowUserToDeleteRows = False
        dgvInfo.Enabled = False
        dgvInfo.RowHeadersVisible = False
        dgvInfo.Columns("strInfoT").HeaderText = "Info"
        dgvInfo.Columns("strInfoT").Width = 100
        dgvInfo.Columns("strInfoV").HeaderText = "Wert"
        dgvInfo.Columns("strInfoV").Width = 250


        'DGV Debitoren
        dgvDebitoren.DataSource = objdtDebitorenHead
        objdbConn.Open()
        Call InitdgvDebitoren()
        Call InitdgvDebitorenSub()
        objdbConn.Close()

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
        Dim strExtBelegNbr As String = ""
        Dim strSkonto As String = ""
        Dim strCurrency As String
        Dim strDebiText As String

        Dim intGegenKonto As Int32
        Dim strFibuText As String
        Dim dblNettoBetrag As Double
        Dim dblBebuBetrag As Double
        Dim strBeBuEintrag As String
        Dim strSteuerFeld As String

        Dim intSollKonto As Int32
        Dim intHabenKonto As Int32
        Dim dblSollBetrag As Double
        Dim dblHabenBetrag As Double
        Dim strSteuerFeldSoll As String = ""
        Dim strSteuerFeldHaben As String = ""
        Dim strBeBuEintragSoll As String = ""
        Dim strBeBuEintragHaben As String = ""
        Dim strDebiTextSoll As String = ""
        Dim strDebiTextHaben As String = ""
        Dim dblKursSoll As Double = 0
        Dim dblKursHaben As Double = 0

        Dim selDebiSub() As DataRow
        Dim strSteuerInfo() As String

        Try

            'Debitor erstellen, minimal - Angaben

            Me.Cursor = Cursors.WaitCursor

            'Kopfbuchung
            For Each row In objdtDebitorenHead.Rows

                If IIf(IsDBNull(row("booDebBook")), False, row("booDebBook")) Then

                    'Test ob OP - Buchung
                    If row("intBuchungsart") = 1 Then

                        'Immer zuerst Belegs-Nummerierung aktivieren, falls vorhanden externe Nummer = OP - Nr. Rg
                        'Führt zu Problemen beim Ausbuchen des DTA - Files
                        'Resultat Besprechnung 17.09.20 mit Muhi/ Andy
                        'DbBhg.IncrBelNbr = "J"
                        'Belegsnummer abholen
                        'intDebBelegsNummer = DbBhg.GetNextBelNbr("R")

                        If IsDBNull(row("strOPNr")) Or row("strOPNr") = "" Then
                            'strExtBelegNbr = row("strOPNr")

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
                        strDebiText = row("strDebText")
                        strCurrency = row("strDebCur")
                        If strCurrency <> "CHF" Then 'Muss ergänzt werden => Was ist Leitwährung auf dem Konto
                            dblKurs = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg)
                        Else
                            dblKurs = 1.0#
                        End If

                        Call DbBhg.SetBelegKopf2(intDebBelegsNummer,
                                                 strValutaDatum,
                                                 intDebitorNbr,
                                                 strBuchType,
                                                 strBelegDatum,
                                                 strVerfallDatum,
                                                 strDebiText,
                                                 strReferenz,
                                                 intKondition,
                                                 strSachBID,
                                                 strVerkID,
                                                 strMahnerlaubnis,
                                                 sngAktuelleMahnstufe,
                                                 dblBetrag.ToString,
                                                 dblKurs.ToString,
                                                 strExtBelegNbr,
                                                 strSkonto,
                                                 strCurrency)

                        selDebiSub = objdtDebitorenSub.Select("strRGNr='" + row("strDebRGNbr") + "'")

                        For Each SubRow As DataRow In selDebiSub

                            intGegenKonto = SubRow("lngKto")
                            strFibuText = SubRow("strDebSubText")
                            dblNettoBetrag = SubRow("dblNetto")
                            'dblBebuBetrag = 1000.0#
                            strBeBuEintrag = SubRow("lngKST").ToString + "{<}" + SubRow("strDebSubText") + "{<}" + "CALCULATE" + "{>}"    '"PROD{<}BebuText{<}" + dblBebuBetrag.ToString + "{>}"
                            strSteuerFeld = Main.FcGetSteuerFeld(FBhg, SubRow("lngKto"), SubRow("strDebSubText"), SubRow("dblBrutto"), SubRow("strMwStKey"))     '"25{<}DEBI D Bhg Export MwSt{<}0{>}"
                            'strSteuerInfo = Split(FBhg.GetKontoInfo(intGegenKonto.ToString), "{>}")
                            'Debug.Print("Konto-Info: " + strSteuerInfo(26))


                            Call DbBhg.SetVerteilung(intGegenKonto.ToString, strFibuText, dblNettoBetrag.ToString, strSteuerFeld, strBeBuEintrag)

                            'Status Sub schreiben

                        Next


                        Call DbBhg.WriteBuchung()

                    Else

                        'Buchung nur in Fibu
                        'Prinzip Funktion WriteBuchung() anwenden mit allen Parametern
                        'Beleg-Nummerierung aktivieren
                        'DbBhg.IncrBelNbr = "J"
                        'Belegsnummer abholen
                        intDebBelegsNummer = FBhg.GetNextBelNbr()

                        'Variablen zuweisen
                        strBelegDatum = Format(row("datDebRGDatum"), "yyyyMMdd").ToString
                        strValutaDatum = Format(row("datDebValDatum"), "yyyyMMdd").ToString
                        'strDebiText = row("strDebText")
                        strCurrency = row("strDebCur")
                        If strCurrency <> "CHF" Then 'Muss ergänzt werden => Was ist Leitwährung auf dem Konto
                            dblKurs = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg)
                        Else
                            dblKurs = 1.0#
                        End If

                        selDebiSub = objdtDebitorenSub.Select("strRGNr='" + row("strDebRGNbr") + "'")

                        If selDebiSub.Length = 2 Then

                            For Each SubRow As DataRow In selDebiSub

                                If SubRow("intSollHaben") = 0 Then 'Soll

                                    intSollKonto = SubRow("lngKto")
                                    dblKursSoll = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg, intSollKonto)
                                    'strSteuerInfo = Split(FBhg.GetKontoInfo(intSollKonto.ToString), "{>}")
                                    'Debug.Print("Konto-Info Soll: " + strSteuerInfo(26))
                                    dblSollBetrag = SubRow("dblNetto")
                                    strDebiTextSoll = SubRow("strDebSubText")
                                    If SubRow("dblMwSt") > 0 Then
                                        strSteuerFeldSoll = Main.FcGetSteuerFeld(FBhg, SubRow("lngKto"), strDebiTextSoll, SubRow("dblBrutto") * dblKursSoll, SubRow("strMwStKey"))
                                    End If
                                    If SubRow("lngKST") > 0 Then
                                        strBeBuEintragSoll = SubRow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
                                    End If


                                ElseIf SubRow("intSollHaben") = 1 Then 'Haben

                                    intHabenKonto = SubRow("lngKto")
                                    dblKursHaben = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg, intHabenKonto)
                                    'strSteuerInfo = Split(FBhg.GetKontoInfo(intSollKonto.ToString), "{>}")
                                    'Debug.Print("Konto-Info Haben: " + strSteuerInfo(26))
                                    dblHabenBetrag = SubRow("dblNetto")
                                    'dblHabenBetrag = dblSollBetrag
                                    strDebiTextHaben = SubRow("strDebSubText")
                                    If SubRow("dblMwSt") > 0 Then
                                        strSteuerFeldHaben = Main.FcGetSteuerFeld(FBhg, SubRow("lngKto"), strDebiTextHaben, SubRow("dblBrutto") * dblKursHaben, SubRow("strMwStKey"))
                                    End If
                                    If SubRow("lngKST") > 0 Then
                                        strBeBuEintragHaben = SubRow("lngKST").ToString + "{<}" + strDebiTextHaben + "{<}" + "CALCULATE" + "{>}"
                                    End If

                                Else

                                    MsgBox("Nicht definierter Wert Sub-Buchungs-SollHaben: " + SubRow("intSollHaben").ToString)

                                End If

                            Next

                            'Tieferer Betrag für die Gesamt-Buchung herausfinden
                            If dblSollBetrag <= dblHabenBetrag Then
                                dblNettoBetrag = dblSollBetrag
                            ElseIf dblHabenBetrag < dblSollBetrag Then
                                dblNettoBetrag = dblHabenBetrag
                            End If

                            'Buchung ausführen
                            Call FBhg.WriteBuchung(0, intDebBelegsNummer, strBelegDatum,
                                                   intSollKonto.ToString, strDebiTextSoll, strCurrency, dblKursSoll.ToString, (dblNettoBetrag * dblKursSoll).ToString, strSteuerFeldSoll,
                                                   intHabenKonto.ToString, strDebiTextHaben, strCurrency, dblKursHaben.ToString, (dblNettoBetrag * dblKursHaben).ToString, strSteuerFeldHaben,
                                                   strCurrency, dblKurs.ToString, dblNettoBetrag.ToString, (dblNettoBetrag * dblKurs).ToString, strBeBuEintragSoll, strBeBuEintragHaben, strValutaDatum)

                        Else
                            MsgBox("Nicht 2 Subbuchungen.")
                        End If



                    End If

                    'Status Head schreiben
                    row("strDebBookStatus") = row("strDebStatusBitLog")
                    row("booBooked") = True
                    row("datBooked") = Now()
                    row("lngBelegNr") = intDebBelegsNummer

                    'Status in File RG-Tabelle schreiben
                    intReturnValue = Main.FcWriteToRGTable(cmbBuha.SelectedValue, row("strDebRGNbr"), row("datBooked"), row("lngBelegNr"), objdbAccessConn, objOracleConn, objdbConn)
                    If intReturnValue <> 0 Then
                        'Throw an exception
                    End If

                End If

            Next


        Catch ex As Exception
            MessageBox.Show(ex.Message)

        Finally
            'Neu aufbauen
            butDebitoren_Click(butDebitoren, EventArgs.Empty)

            Me.Cursor = Cursors.Default

        End Try

    End Sub


    Private Sub dgvDebitoren_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvDebitoren.CellValueChanged

        If e.ColumnIndex = 2 And e.RowIndex >= 0 Then

            If dgvDebitoren.Rows(e.RowIndex).Cells("booDebBook").Value Then

                'MsgBox("Geändert zu checked " + dgvDebitoren.Rows(e.RowIndex).Cells("strDebRGNbr").Value + ", " + dgvDebitoren.Rows(e.RowIndex).Cells("booDebBook").Value.ToString + Val(dgvDebitoren.Rows(e.RowIndex).Cells("strDebStatusBitLog").Value).ToString)
                'Zulassen? = keine Fehler
                If Val(dgvDebitoren.Rows(e.RowIndex).Cells("strDebStatusBitLog").Value) <> 0 And Val(dgvDebitoren.Rows(e.RowIndex).Cells("strDebStatusBitLog").Value) <> 10000 Then
                    MsgBox("Rechnung ist nicht buchbar.", vbOKOnly + vbExclamation, "Nicht buchbar")
                    dgvDebitoren.Rows(e.RowIndex).Cells("booDebBook").Value = False
                End If
            End If

        End If

    End Sub


    Private Sub dgvDebitoren_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvDebitoren.CellContentClick

        'Dim dsDebSub() As DataRow
        'Sub-Grid zeigen, vorher mit Datne abfüllen
        'dgvDebSubActual.DataSource = objdtDebitorenSub.Select("strRGNr='" + dgvDebitoren.Rows(e.RowIndex).Cells("strDebRGNbr").Value + "'").CopyToDataTable()
        dgvDebitorenSub.DataSource = objdtDebitorenSub.Select("strRGNr='" + dgvDebitoren.Rows(e.RowIndex).Cells("strDebRGNbr").Value + "'").CopyToDataTable


        '        dsDebSub = objdtDebitorenSub.Select("lngID > 1")

        'dgvDebitorenSub.DataSource = dsDebSub

        dgvDebitorenSub.Update()

        'dgvDebitorenSub.DataSource = objdtDebitorenSub

        'dgvDebitorenSub.Update()

        'dgvDebSubActual.Visible = True


    End Sub
End Class
