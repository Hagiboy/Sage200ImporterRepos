Option Strict Off
Option Explicit On

Imports MySql.Data.MySqlClient
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
    Public objDABuchhaltungen As New MySqlDataAdapter("SELECT * FROM buchhaltungen WHERE NOT Buchh200_Name IS NULL", objdbConn)
    'Public objDACarsGrid As New MySqlDataAdapter("SELECT tblcars.idCar, tblunits.strUnit, tblplates.strPlate, tblcars.strVIN, tblmodelle.strModell FROM tblcars LEFT JOIN tblunits ON tblcars.refUnit = tblunits.idUnit LEFT JOIN tblplates ON tblcars.refPlate = tblplates.idPlate LEFT JOIN tblmodelle ON tblcars.refModell = tblmodelle.idModell", objdbConn)
    'Public objdtDebitor As New DataTable("tbliDebitor")
    Public objdtBuchhaltungen As New DataTable("tbliBuchhaltungen")
    Public objdtDebitorenHead As New DataTable("tbliDebiHead")
    Public objdtDebitorenHeadRead As New DataTable("tbliDebitorenHeadR")



    Public Sub InitVar()
        'UPGRADE_NOTE: Object PIFin may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        PIFin = Nothing
        'UPGRADE_NOTE: Object KrBhg may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        KrBhg = Nothing
        'UPGRADE_NOTE: Object FBhg may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        FBhg = Nothing
        'UPGRADE_NOTE: Object DbBhg may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        DbBhg = Nothing
        'UPGRADE_NOTE: Object BsExt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        BsExt = Nothing
        'UPGRADE_NOTE: Object BeBu may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        BeBu = Nothing
        'UPGRADE_NOTE: Object Adr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Adr = Nothing
        'UPGRADE_NOTE: Object Finanz may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
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

        '        Dim booAccOk As Boolean
        '        Dim strMandant As String
        '        Dim b As Object
        '       Dim s As Object
        '       b = Nothing
        '       On Error GoTo ErrorHandler

        Me.Cursor = Cursors.WaitCursor

        Call InitVar()

        Call Main.fcLoginSage(objdbConn, Finanz, FBhg, DbBhg, cmbBuha.SelectedValue)

        Call Main.fcFillDebit(cmbBuha.SelectedValue, objdtDebitorenHeadRead, objdbConn, objdbAccessConn)

        'Call InitdgvDebitoren()
        Call Main.InsertDataTableColumnName(objdtDebitorenHeadRead, objdtDebitorenHead)

        'Grid neu aufbauen
        dgvDebitoren.Update()
        dgvDebitoren.Refresh()
        'dgvDebitoren.DataSource = objdtDebitorenHead
        'Debug.Print(objdtDebitorenHead.Rows.Count.ToString)
        'Call InitdgvDebitoren()

        Call Main.fcCheckDebit(cmbBuha.SelectedValue, objdtDebitorenHead, Finanz, FBhg, DbBhg)

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
        dgvDebitoren.Columns("strDebRGNbr").Width = 80
        dgvDebitoren.Columns("lngDebNbr").DisplayIndex = 2
        dgvDebitoren.Columns("lngDebNbr").HeaderText = "Debitor"
        dgvDebitoren.Columns("lngDebNbr").Width = 80
        dgvDebitoren.Columns("strDebBez").DisplayIndex = 3
        dgvDebitoren.Columns("strDebBez").HeaderText = "Bezeichnung"
        dgvDebitoren.Columns("strDebBez").Width = 150
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
        dgvDebitoren.Columns("dblDebNetto").Width = 90
        dgvDebitoren.Columns("dblDebNetto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvDebitoren.Columns("dblDebMwSt").DisplayIndex = 8
        dgvDebitoren.Columns("dblDebMwSt").HeaderText = "MwSt"
        dgvDebitoren.Columns("dblDebMwSt").Width = 80
        dgvDebitoren.Columns("dblDebMwSt").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvDebitoren.Columns("dblDebBrutto").DisplayIndex = 9
        dgvDebitoren.Columns("dblDebBrutto").HeaderText = "Brutto"
        dgvDebitoren.Columns("dblDebBrutto").Width = 90
        dgvDebitoren.Columns("dblDebBrutto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvDebitoren.Columns("intSubBookings").DisplayIndex = 10
        dgvDebitoren.Columns("intSubBookings").HeaderText = "Sub"
        dgvDebitoren.Columns("intSubBookings").Width = 50
        dgvDebitoren.Columns("intSubBookings").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvDebitoren.Columns("dblSumSubBookings").DisplayIndex = 11
        dgvDebitoren.Columns("dblSumSubBookings").HeaderText = "Sub-Summe"
        dgvDebitoren.Columns("dblSumSubBookings").Width = 90
        dgvDebitoren.Columns("dblSumSubBookings").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvDebitoren.Columns("lngDebIdentNbr").DisplayIndex = 12
        dgvDebitoren.Columns("lngDebIdentNbr").HeaderText = "Ident"
        dgvDebitoren.Columns("lngDebIdentNbr").Width = 80
        Dim comBoxCol As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn
        comBoxCol.HeaderText = "Buchungsart"
        comBoxCol.Width = 80
        comBoxCol.Name = "cmbBuchungsart"
        comBoxCol.DataSource = objdtDebitorenHead
        'comBoxCol.Items.Add("OP")
        'comBoxCol.Items.Add("KKT")
        'comBoxCol.Items.Add("Sum Up")
        'comBoxCol.Items.Add("Cash T")
        comBoxCol.ValueMember = "intBuchungsart"
        dgvDebitoren.Columns.Add(comBoxCol)
        dgvDebitoren.Columns("cmbBuchungsart").DisplayIndex = 13
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
        dgvDebitoren.Columns("strDebiBank").Width = 80
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

        objdbConn.Open()

        'Comboxen
        objdtBuchhaltungen.Clear()
        objDABuchhaltungen.Fill(objdtBuchhaltungen)
        'cmbMarken.Sorted = True
        cmbBuha.DataSource = objdtBuchhaltungen
        cmbBuha.DisplayMember = "Buchh_Bez"
        cmbBuha.ValueMember = "Buchh_Nr"

        'Tabelle Head erstellen
        objdtDebitorenHead = Main.tblDebitorenHead()

        'DGV
        dgvDebitoren.DataSource = objdtDebitorenHead
        Call InitdgvDebitoren()
        'dgvDebitoren.AllowUserToAddRows = False
        'dgvDebitoren.AllowUserToDeleteRows = False
        'dgvDebitoren.Columns("booDebBook").DisplayIndex = 0
        'dgvDebitoren.Columns("booDebBook").HeaderText = "ok"
        'dgvDebitoren.Columns("booDebBook").Width = 40
        'dgvDebitoren.Columns("booDebBook").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        'dgvDebitoren.Columns("booDebBook").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'dgvDebitoren.Columns("strDebRGNbr").DisplayIndex = 1
        'dgvDebitoren.Columns("strDebRGNbr").HeaderText = "RG-Nr"
        'dgvDebitoren.Columns("strDebRGNbr").Width = 80
        'dgvDebitoren.Columns("lngDebNbr").DisplayIndex = 2
        'dgvDebitoren.Columns("lngDebNbr").HeaderText = "Debitor"
        'dgvDebitoren.Columns("lngDebNbr").Width = 80
        'dgvDebitoren.Columns("strDebBez").DisplayIndex = 3
        'dgvDebitoren.Columns("strDebBez").HeaderText = "Bezeichnung"
        'dgvDebitoren.Columns("strDebBez").Width = 150
        'dgvDebitoren.Columns("lngDebKtoNbr").DisplayIndex = 4
        'dgvDebitoren.Columns("lngDebKtoNbr").HeaderText = "Konto"
        'dgvDebitoren.Columns("lngDebKtoNbr").Width = 50
        'dgvDebitoren.Columns("strDebKtoBez").DisplayIndex = 5
        'dgvDebitoren.Columns("strDebKtoBez").HeaderText = "Bezeichnung"
        'dgvDebitoren.Columns("strDebKtoBez").Width = 150
        'dgvDebitoren.Columns("strDebCur").DisplayIndex = 6
        'dgvDebitoren.Columns("strDebCur").HeaderText = "Währung"
        'dgvDebitoren.Columns("strDebCur").Width = 60
        'dgvDebitoren.Columns("dblDebNetto").DisplayIndex = 7
        'dgvDebitoren.Columns("dblDebNetto").HeaderText = "Netto"
        'dgvDebitoren.Columns("dblDebNetto").Width = 90
        'dgvDebitoren.Columns("dblDebNetto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        'dgvDebitoren.Columns("dblDebMwSt").DisplayIndex = 8
        'dgvDebitoren.Columns("dblDebMwSt").HeaderText = "MwSt"
        'dgvDebitoren.Columns("dblDebMwSt").Width = 80
        'dgvDebitoren.Columns("dblDebMwSt").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        'dgvDebitoren.Columns("dblDebBrutto").DisplayIndex = 9
        'dgvDebitoren.Columns("dblDebBrutto").HeaderText = "Brutto"
        'dgvDebitoren.Columns("dblDebBrutto").Width = 90
        'dgvDebitoren.Columns("dblDebBrutto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        'dgvDebitoren.Columns("intSubBookings").DisplayIndex = 10
        'dgvDebitoren.Columns("intSubBookings").HeaderText = "Sub"
        'dgvDebitoren.Columns("intSubBookings").Width = 50
        'dgvDebitoren.Columns("intSubBookings").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        'dgvDebitoren.Columns("dblSumSubBookings").DisplayIndex = 11
        'dgvDebitoren.Columns("dblSumSubBookings").HeaderText = "Sub-Summe"
        'dgvDebitoren.Columns("dblSumSubBookings").Width = 90
        'dgvDebitoren.Columns("dblSumSubBookings").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        'dgvDebitoren.Columns("lngDebIdentNbr").DisplayIndex = 12
        'dgvDebitoren.Columns("lngDebIdentNbr").HeaderText = "Ident"
        'dgvDebitoren.Columns("lngDebIdentNbr").Width = 80
        'Dim comBoxCol As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn
        'comBoxCol.HeaderText = "Buchungsart"
        'comBoxCol.Width = 80
        'comBoxCol.Name = "cmbBuchungsart"
        'comBoxCol.DataSource = objdtDebitorenHead
        ''comBoxCol.Items.Add("OP")
        ''comBoxCol.Items.Add("KKT")
        ''comBoxCol.Items.Add("Sum Up")
        ''comBoxCol.Items.Add("Cash T")
        'comBoxCol.ValueMember = "intBuchungsart"
        'dgvDebitoren.Columns.Add(comBoxCol)
        'dgvDebitoren.Columns("cmbBuchungsart").DisplayIndex = 13
        'dgvDebitoren.Columns("strOPNr").DisplayIndex = 14
        'dgvDebitoren.Columns("strOPNr").HeaderText = "OP-Nr"
        'dgvDebitoren.Columns("strOPNr").Width = 80
        'dgvDebitoren.Columns("datDebRGDatum").DisplayIndex = 15
        'dgvDebitoren.Columns("datDebRGDatum").HeaderText = "RG Datum"
        'dgvDebitoren.Columns("datDebRGDatum").Width = 70
        'dgvDebitoren.Columns("datDebValDatum").DisplayIndex = 16
        'dgvDebitoren.Columns("datDebValDatum").HeaderText = "Val Datum"
        'dgvDebitoren.Columns("datDebValDatum").Width = 70
        'dgvDebitoren.Columns("strDebiBank").DisplayIndex = 17
        'dgvDebitoren.Columns("strDebiBank").HeaderText = "Bank"
        'dgvDebitoren.Columns("strDebiBank").Width = 80
        'dgvDebitoren.Columns("strDebStatusText").DisplayIndex = 18
        'dgvDebitoren.Columns("strDebStatusText").HeaderText = "Status"
        'dgvDebitoren.Columns("strDebStatusText").Width = 200
        'dgvDebitoren.Columns("intBuchhaltung").Visible = False
        'dgvDebitoren.Columns("intBuchungsart").Visible = False
        'dgvDebitoren.Columns("intRGArt").Visible = False
        'dgvDebitoren.Columns("strRGArt").Visible = False
        'dgvDebitoren.Columns("lngLinkedRG").Visible = False
        'dgvDebitoren.Columns("booLinked").Visible = False
        'dgvDebitoren.Columns("strRGName").Visible = False
        'dgvDebitoren.Columns("strDebIdentnbr2").Visible = False
        'dgvDebitoren.Columns("strDebText").Visible = False
        'dgvDebitoren.Columns("strRGBemerkung").Visible = False
        'dgvDebitoren.Columns("strDebRef").Visible = False
        'dgvDebitoren.Columns("strZahlBed").Visible = False
        'dgvDebitoren.Columns("strDebStatusBitLog").Visible = False
        'dgvDebitoren.Columns("strDebBookStatus").Visible = False
        'dgvDebitoren.Columns("booBooked").Visible = False
        'dgvDebitoren.Columns("datBooked").Visible = False
        'dgvDebitoren.Columns("lngBelegNr").Visible = False


    End Sub
End Class
