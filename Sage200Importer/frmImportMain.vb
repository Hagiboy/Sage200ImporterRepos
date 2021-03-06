﻿Option Strict Off
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
    Public objDABuchhaltungen As New MySqlDataAdapter("SELECT * FROM t_sage_buchhaltungen WHERE NOT Buchh200_Name IS NULL ORDER BY Buchh_Bez", objdbConn)
    'Public objDACarsGrid As New MySqlDataAdapter("SELECT tblcars.idCar, tblunits.strUnit, tblplates.strPlate, tblcars.strVIN, tblmodelle.strModell FROM tblcars LEFT JOIN tblunits ON tblcars.refUnit = tblunits.idUnit LEFT JOIN tblplates ON tblcars.refPlate = tblplates.idPlate LEFT JOIN tblmodelle ON tblcars.refModell = tblmodelle.idModell", objdbConn)
    'Public objdtDebitor As New DataTable("tbliDebitor")
    Public objdtBuchhaltungen As New DataTable("tbliBuchhaltungen")
    Public objdtDebitorenHead As New DataTable("tbliDebiHead")
    Public objdtDebitorenHeadRead As New DataTable("tbliDebitorenHeadR")
    Public objdtDebitorenSub As New DataTable("tbliDebiSub")
    Public objdtKreditorenHead As New DataTable("tbliKrediHead")
    Public objdtKreditorenHeadRead As New DataTable("tbliKreditorenHeadR")
    Public objdtKreditorenSub As New DataTable("tbliKrediSub")
    Public objdtInfo As New DataTable("tbliInfo")
    Public objOracleConn As New OracleConnection("Data Source=(DESCRIPTION=" _
                    + "(ADDRESS=(PROTOCOL=TCP)(HOST=10.0.0.29)(PORT=1521))" _
                    + "(CONNECT_DATA=(SERVICE_NAME=CISNEW)));" _
                    + "User Id=cis;Password=sugus;")
    Public objOracleCmd As New OracleCommand()
    Public intMode As Int16
    'Public boodgvSet As Boolean = False

    Public Sub InitVar()

        PIFin = Nothing
        KrBhg = Nothing
        FBhg = Nothing
        DbBhg = Nothing
        BsExt = Nothing
        BeBu = Nothing
        Adr = Nothing
        Finanz = Nothing

        'objdbcommand = Nothing
        'objdbcommandZHDB02 = Nothing
        'objdtDebitorenHeadRead.Clear()
        'objdtDebitorenHead.Clear()
        'objdtDebitorenSub.Clear()
        objdtKreditorenHeadRead.Clear()
        objdtKreditorenHead.Clear()
        objdtKreditorenSub.Clear()
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

        ''Compute - Text
        'Dim tblCompute As New DataTable()
        'Dim booResult As Boolean
        'booResult = Convert.ToBoolean(tblCompute.Compute("#" + DateTime.Now.ToString("yyyy-MM-dd") + "#" + ">=#2020-11-19#", Nothing))
        'Debug.Print("Result " + "#" + DateTime.Now.ToString("yyyy-MM-dd") + "#" + ">=#2020-11-19#" + booResult.ToString)
        'Stop

        Me.Cursor = Cursors.WaitCursor

        intMode = 0

        objdtDebitorenHead = Nothing
        objdtDebitorenHeadRead = Nothing
        objdtDebitorenSub = Nothing

        'Tabelle Debi/ Kredi Head erstellen
        objdtDebitorenHead = Main.tblDebitorenHead()
        objdtDebitorenHeadRead = Main.tblDebitorenHead()
        'objdtKreditorenHead = Main.tblKreditorenHead()

        'Tabelle Debi/ Kredi Sub erstellen
        objdtDebitorenSub = Main.tblDebitorenSub()
        'objdtKreditorenSub = Main.tblKreditorenSub()


        'objdtDebitorenHead.Clear()
        'objdtDebitorenHeadRead.Clear()
        'objdtDebitorenSub.Clear()
        objdtInfo.Clear()

        'dgvBookings.Rows.Clear()
        If dgvBookings.Columns.Contains("intBuchungsart") Then
            dgvBookings.Columns.Remove("intBuchungsart")
        End If

        'DGV Debitoren
        dgvBookings.DataSource = objdtDebitorenHead
        dgvBookingSub.DataSource = objdtDebitorenSub
        objdbConn.Open()
        Call InitdgvDebitoren()
        Call InitdgvDebitorenSub()
        objdbConn.Close()

        Call InitVar()

        Call Main.FcLoginSage(objdbConn,
                              objdbMSSQLConn,
                              objdbSQLcommand,
                              Finanz,
                              FBhg,
                              DbBhg,
                              PIFin,
                              KrBhg,
                              cmbBuha.SelectedValue,
                              objdtInfo,
                              cmbPerioden.SelectedItem)

        'Transitorische Buchungen?
        Call Main.fcCheckTransitorischeDebit(cmbBuha.SelectedValue,
                                             objdbConn,
                                             objdbAccessConn)

        'Gibt es eine Query auszuführen bevor dem Buchen?
        Call MainDebitor.FcExecuteBeforeDebit(cmbBuha.SelectedValue, objdbConn)

        Call MainDebitor.FcFillDebit(cmbBuha.SelectedValue,
                                     objdtDebitorenHeadRead,
                                     objdtDebitorenSub,
                                     objdbConn,
                                     objdbAccessConn,
                                     objOracleConn,
                                     objOracleCmd)

        Call Main.InsertDataTableColumnName(objdtDebitorenHeadRead, objdtDebitorenHead)

        'Grid neu aufbauen
        dgvBookingSub.Update()
        dgvBookings.Update()
        dgvBookings.Refresh()

        Call Main.FcCheckDebit(cmbBuha.SelectedValue,
                               objdtDebitorenHead,
                               objdtDebitorenSub,
                               Finanz,
                               FBhg,
                               DbBhg,
                               PIFin,
                               objdbConn,
                               objdbConnZHDB02,
                               objdbcommand,
                               objdbcommandZHDB02,
                               objOracleConn,
                               objOracleCmd,
                               objdbAccessConn,
                               objdtInfo,
                               cmbBuha.Text)

        'Anzahl schreiben
        txtNumber.Text = objdtDebitorenHead.Rows.Count.ToString

        ''Ipmort Kredit hiden
        Me.butImportK.Enabled = False
        Me.butImport.Enabled = True

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

        dgvBookings.ShowCellToolTips = False
        dgvBookings.AllowUserToAddRows = False
        dgvBookings.AllowUserToDeleteRows = False
        dgvBookings.Columns("booDebBook").DisplayIndex = 0
        dgvBookings.Columns("booDebBook").HeaderText = "ok"
        dgvBookings.Columns("booDebBook").Width = 40
        dgvBookings.Columns("booDebBook").ValueType = System.Type.[GetType]("System.Boolean")
        dgvBookings.Columns("booDebBook").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvBookings.Columns("booDebBook").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvBookings.Columns("strDebRGNbr").DisplayIndex = 1
        dgvBookings.Columns("strDebRGNbr").HeaderText = "RG-Nr"
        dgvBookings.Columns("strDebRGNbr").Width = 60
        dgvBookings.Columns("strDebRGNbr").ReadOnly = True
        dgvBookings.Columns("lngDebNbr").DisplayIndex = 2
        dgvBookings.Columns("lngDebNbr").HeaderText = "Debitor"
        dgvBookings.Columns("lngDebNbr").Width = 60
        dgvBookings.Columns("strDebBez").DisplayIndex = 3
        dgvBookings.Columns("strDebBez").HeaderText = "Bezeichnung"
        dgvBookings.Columns("strDebBez").Width = 140
        dgvBookings.Columns("lngDebKtoNbr").DisplayIndex = 4
        dgvBookings.Columns("lngDebKtoNbr").HeaderText = "Konto"
        dgvBookings.Columns("lngDebKtoNbr").Width = 50
        dgvBookings.Columns("strDebKtoBez").DisplayIndex = 5
        dgvBookings.Columns("strDebKtoBez").HeaderText = "Bezeichnung"
        dgvBookings.Columns("strDebKtoBez").Width = 150
        dgvBookings.Columns("strDebCur").DisplayIndex = 6
        dgvBookings.Columns("strDebCur").HeaderText = "Währung"
        dgvBookings.Columns("strDebCur").Width = 60
        dgvBookings.Columns("dblDebNetto").DisplayIndex = 7
        dgvBookings.Columns("dblDebNetto").HeaderText = "Netto"
        dgvBookings.Columns("dblDebNetto").Width = 80
        dgvBookings.Columns("dblDebNetto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvBookings.Columns("dblDebNetto").DefaultCellStyle.Format = "N2"
        dgvBookings.Columns("dblDebNetto").ReadOnly = True
        dgvBookings.Columns("dblDebMwSt").DisplayIndex = 8
        dgvBookings.Columns("dblDebMwSt").HeaderText = "MwSt"
        dgvBookings.Columns("dblDebMwSt").Width = 70
        dgvBookings.Columns("dblDebMwSt").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvBookings.Columns("dblDebMwSt").DefaultCellStyle.Format = "N2"
        dgvBookings.Columns("dblDebBrutto").DisplayIndex = 9
        dgvBookings.Columns("dblDebBrutto").HeaderText = "Brutto"
        dgvBookings.Columns("dblDebBrutto").Width = 80
        dgvBookings.Columns("dblDebBrutto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvBookings.Columns("dblDebBrutto").DefaultCellStyle.Format = "N2"
        dgvBookings.Columns("intSubBookings").DisplayIndex = 10
        dgvBookings.Columns("intSubBookings").HeaderText = "Sub"
        dgvBookings.Columns("intSubBookings").Width = 50
        dgvBookings.Columns("intSubBookings").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvBookings.Columns("dblSumSubBookings").DisplayIndex = 11
        dgvBookings.Columns("dblSumSubBookings").HeaderText = "Sub-Summe"
        dgvBookings.Columns("dblSumSubBookings").Width = 80
        dgvBookings.Columns("dblSumSubBookings").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvBookings.Columns("lngDebIdentNbr").DisplayIndex = 12
        dgvBookings.Columns("lngDebIdentNbr").HeaderText = "Ident"
        dgvBookings.Columns("lngDebIdentNbr").Width = 80
        'If Not boodgvSet Then
        Dim cmbBuchungsart As New DataGridViewComboBoxColumn()
        Dim objdtBA As New DataTable("objidtBA")
        Dim objlocMySQLcmd As New MySqlCommand
        objlocMySQLcmd.CommandText = "SELECT * FROM t_sage_tblbuchungsarten"
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
        dgvBookings.Columns.Add(cmbBuchungsart)
        'boodgvSet = True
        'End If
        'dgvDebitoren.Columns("intBuchungsart").DisplayIndex = 13
        'dgvDebitoren.Columns("intBuchungsart").DisplayIndex = 13
        'dgvDebitoren.Columns("intBuchungsart").HeaderText = "BA"
        'dgvDebitoren.Columns("intBuchungsart").Width = 40
        dgvBookings.Columns("strOPNr").DisplayIndex = 14
        dgvBookings.Columns("strOPNr").HeaderText = "OP-Nr"
        dgvBookings.Columns("strOPNr").Width = 80
        dgvBookings.Columns("datDebRGDatum").DisplayIndex = 15
        dgvBookings.Columns("datDebRGDatum").HeaderText = "RG Datum"
        dgvBookings.Columns("datDebRGDatum").Width = 70
        dgvBookings.Columns("datDebValDatum").DisplayIndex = 16
        dgvBookings.Columns("datDebValDatum").HeaderText = "Val Datum"
        dgvBookings.Columns("datDebValDatum").Width = 70
        dgvBookings.Columns("strDebiBank").DisplayIndex = 17
        dgvBookings.Columns("strDebiBank").HeaderText = "Bank"
        dgvBookings.Columns("strDebiBank").Width = 60
        dgvBookings.Columns("strDebStatusText").DisplayIndex = 18
        dgvBookings.Columns("strDebStatusText").HeaderText = "Status"
        dgvBookings.Columns("strDebStatusText").Width = 200
        dgvBookings.Columns("intBuchhaltung").Visible = False
        dgvBookings.Columns("intBuchungsart").Visible = False
        'dgvBookings.Columns("intRGArt").Visible = False
        dgvBookings.Columns("strRGArt").Visible = False
        'dgvBookings.Columns("lngLinkedRG").Visible = False
        'dgvBookings.Columns("booLinked").Visible = False
        dgvBookings.Columns("strRGName").Visible = False
        dgvBookings.Columns("strDebIdentnbr2").Visible = False
        'dgvBookings.Columns("strDebText").Visible = False
        dgvBookings.Columns("strRGBemerkung").Visible = False
        'dgvBookings.Columns("strDebRef").Visible = False
        dgvBookings.Columns("strZahlBed").Visible = False
        dgvBookings.Columns("strDebStatusBitLog").Visible = False
        dgvBookings.Columns("strDebBookStatus").Visible = False
        dgvBookings.Columns("booBooked").Visible = False
        dgvBookings.Columns("datBooked").Visible = False
        dgvBookings.Columns("lngBelegNr").Visible = False


    End Sub

    Private Sub InitdgvDebitorenSub()

        dgvBookingSub.ShowCellToolTips = False
        dgvBookingSub.AllowUserToAddRows = False
        dgvBookingSub.AllowUserToDeleteRows = False
        dgvBookingSub.Columns("strRGNr").DisplayIndex = 0
        dgvBookingSub.Columns("strRGNr").Width = 55
        dgvBookingSub.Columns("strRGNr").HeaderText = "RG-Nr"
        dgvBookingSub.Columns("intSollHaben").Width = 20
        dgvBookingSub.Columns("intSollHaben").HeaderText = "S/H"
        dgvBookingSub.Columns("lngKto").Width = 50
        dgvBookingSub.Columns("lngKto").HeaderText = "Konto"
        dgvBookingSub.Columns("strKtoBez").HeaderText = "Bezeichnung"
        dgvBookingSub.Columns("lngKST").Width = 40
        dgvBookingSub.Columns("lngKST").HeaderText = "KST"
        dgvBookingSub.Columns("strKSTBez").Width = 60
        dgvBookingSub.Columns("strKSTBez").HeaderText = "Bezeichnung"
        dgvBookingSub.Columns("dblNetto").Width = 60
        dgvBookingSub.Columns("dblNetto").HeaderText = "Netto"
        dgvBookingSub.Columns("dblNetto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvBookingSub.Columns("dblNetto").DefaultCellStyle.Format = "N2"
        dgvBookingSub.Columns("dblMwSt").Width = 50
        dgvBookingSub.Columns("dblMwSt").HeaderText = "MwSt"
        dgvBookingSub.Columns("dblMwSt").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvBookingSub.Columns("dblMwSt").DefaultCellStyle.Format = "N2"
        dgvBookingSub.Columns("dblBrutto").Width = 60
        dgvBookingSub.Columns("dblBrutto").HeaderText = "Brutto"
        dgvBookingSub.Columns("dblBrutto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvBookingSub.Columns("dblBrutto").DefaultCellStyle.Format = "N2"
        dgvBookingSub.Columns("dblMwStSatz").Width = 40
        dgvBookingSub.Columns("dblMwStSatz").HeaderText = "MwStS"
        dgvBookingSub.Columns("strMwStKey").Width = 40
        dgvBookingSub.Columns("strMwStKey").HeaderText = "MwStK"
        dgvBookingSub.Columns("strStatusUBText").HeaderText = "Status"
        dgvBookingSub.Columns("strStatusUBText").Width = 135

        dgvBookingSub.Columns("lngID").Visible = False
        dgvBookingSub.Columns("strArtikel").Visible = False
        dgvBookingSub.Columns("strStatusUBBitLog").Visible = False
        dgvBookingSub.Columns("strDebSubText").Visible = False
        dgvBookingSub.Columns("strDebBookStatus").Visible = False

    End Sub

    Private Sub InitdgvKreditoren()

        dgvBookings.ShowCellToolTips = False
        dgvBookings.AllowUserToAddRows = False
        dgvBookings.AllowUserToDeleteRows = False
        dgvBookings.Columns("booKredBook").DisplayIndex = 0
        dgvBookings.Columns("booKredBook").HeaderText = "ok"
        dgvBookings.Columns("booKredBook").Width = 40
        dgvBookings.Columns("booKredBook").ValueType = System.Type.[GetType]("System.Boolean")
        dgvBookings.Columns("booKredBook").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvBookings.Columns("booKredBook").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'dgvBookings.Columns("strKredRGNbr").DisplayIndex = 1
        'dgvBookings.Columns("strKredRGNbr").HeaderText = "KRG-Nr"
        'dgvBookings.Columns("strKredRGNbr").Width = 60
        'dgvBookings.Columns("strKredRGNbr").ReadOnly = True
        dgvBookings.Columns("lngKredID").DisplayIndex = 1
        dgvBookings.Columns("lngKredID").HeaderText = "Kred-ID"
        dgvBookings.Columns("lngKredID").Width = 60
        dgvBookings.Columns("lngKredID").ReadOnly = True
        dgvBookings.Columns("lngKredNbr").DisplayIndex = 2
        dgvBookings.Columns("lngKredNbr").HeaderText = "Kreditor"
        dgvBookings.Columns("lngKredNbr").Width = 60
        dgvBookings.Columns("strKredBez").DisplayIndex = 3
        dgvBookings.Columns("strKredBez").HeaderText = "Bezeichnung"
        dgvBookings.Columns("strKredBez").Width = 140
        dgvBookings.Columns("lngKredKtoNbr").DisplayIndex = 4
        dgvBookings.Columns("lngKredKtoNbr").HeaderText = "Konto"
        dgvBookings.Columns("lngKredKtoNbr").Width = 50
        dgvBookings.Columns("strKredKtoBez").DisplayIndex = 5
        dgvBookings.Columns("strKredKtoBez").HeaderText = "Bezeichnung"
        dgvBookings.Columns("strKredKtoBez").Width = 150
        dgvBookings.Columns("strKredCur").DisplayIndex = 6
        dgvBookings.Columns("strKredCur").HeaderText = "Währung"
        dgvBookings.Columns("strKredCur").Width = 60
        dgvBookings.Columns("dblKredNetto").DisplayIndex = 7
        dgvBookings.Columns("dblKredNetto").HeaderText = "Netto"
        dgvBookings.Columns("dblKredNetto").Width = 80
        dgvBookings.Columns("dblKredNetto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        'dgvBookings.Columns("dblKredNetto").DefaultCellStyle.Format = "N2"
        dgvBookings.Columns("dblKredNetto").ReadOnly = True
        dgvBookings.Columns("dblKredMwSt").DisplayIndex = 8
        dgvBookings.Columns("dblKredMwSt").HeaderText = "MwSt"
        dgvBookings.Columns("dblKredMwSt").Width = 70
        dgvBookings.Columns("dblKredMwSt").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        'dgvBookings.Columns("dblKredMwSt").DefaultCellStyle.Format = "N2"
        dgvBookings.Columns("dblKredBrutto").DisplayIndex = 9
        dgvBookings.Columns("dblKredBrutto").HeaderText = "Brutto"
        dgvBookings.Columns("dblKredBrutto").Width = 80
        dgvBookings.Columns("dblKredBrutto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        'dgvBookings.Columns("dblKredBrutto").DefaultCellStyle.Format = "N2"
        dgvBookings.Columns("intSubBookings").DisplayIndex = 10
        dgvBookings.Columns("intSubBookings").HeaderText = "Sub"
        dgvBookings.Columns("intSubBookings").Width = 50
        dgvBookings.Columns("intSubBookings").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvBookings.Columns("dblSumSubBookings").DisplayIndex = 11
        dgvBookings.Columns("dblSumSubBookings").DefaultCellStyle.Format = "N2"
        dgvBookings.Columns("dblSumSubBookings").HeaderText = "Sub-Summe"
        dgvBookings.Columns("dblSumSubBookings").Width = 80
        dgvBookings.Columns("dblSumSubBookings").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvBookings.Columns("lngKredIdentNbr").DisplayIndex = 12
        dgvBookings.Columns("lngKredIdentNbr").HeaderText = "Ident"
        dgvBookings.Columns("lngKredIdentNbr").Width = 80
        Dim cmbBuchungsart As New DataGridViewComboBoxColumn()
        Dim objdtBA As New DataTable("objidtBA")
        Dim objlocMySQLcmd As New MySqlCommand
        'objdbConn.Open()
        objlocMySQLcmd.CommandText = "SELECT * FROM t_sage_tblbuchungsarten"
        objlocMySQLcmd.Connection = objdbConn
        objdtBA.Load(objlocMySQLcmd.ExecuteReader)
        cmbBuchungsart.DataSource = objdtBA
        'objdbConn.Close()
        cmbBuchungsart.DisplayMember = "strBuchungsart"
        cmbBuchungsart.ValueMember = "idBuchungsart"
        cmbBuchungsart.HeaderText = "BA"
        cmbBuchungsart.Name = "intBuchungsart"
        cmbBuchungsart.DataPropertyName = "intBuchungsart"
        cmbBuchungsart.DisplayIndex = 13
        cmbBuchungsart.Width = 70
        dgvBookings.Columns.Add(cmbBuchungsart)
        ''dgvDebitoren.Columns("intBuchungsart").DisplayIndex = 13
        ''dgvDebitoren.Columns("intBuchungsart").DisplayIndex = 13
        ''dgvDebitoren.Columns("intBuchungsart").HeaderText = "BA"
        ''dgvDebitoren.Columns("intBuchungsart").Width = 40
        dgvBookings.Columns("strKredRGNbr").DisplayIndex = 14
        dgvBookings.Columns("strKredRGNbr").HeaderText = "RG-Nr"
        dgvBookings.Columns("strKredRGNbr").Width = 100
        dgvBookings.Columns("datKredRGDatum").DisplayIndex = 15
        dgvBookings.Columns("datKredRGDatum").HeaderText = "RG Datum"
        dgvBookings.Columns("datKredRGDatum").Width = 70
        dgvBookings.Columns("datKredValDatum").DisplayIndex = 16
        dgvBookings.Columns("datKredValDatum").HeaderText = "Val Datum"
        dgvBookings.Columns("datKredValDatum").Width = 70
        dgvBookings.Columns("strKrediBank").DisplayIndex = 17
        dgvBookings.Columns("strKrediBank").HeaderText = "KBank"
        dgvBookings.Columns("strKrediBank").Width = 60
        dgvBookings.Columns("strKredStatusText").DisplayIndex = 18
        dgvBookings.Columns("strKredStatusText").HeaderText = "Status"
        dgvBookings.Columns("strKredStatusText").Width = 200
        'dgvBookings.Columns("intBuchhaltung").Visible = False
        'dgvBookings.Columns("intBuchungsart").Visible = False
        'dgvBookings.Columns("intRGArt").Visible = False
        'dgvBookings.Columns("strRGArt").Visible = False
        'dgvBookings.Columns("lngLinkedRG").Visible = False
        'dgvBookings.Columns("booLinked").Visible = False
        'dgvBookings.Columns("strRGName").Visible = False
        'dgvBookings.Columns("strDebIdentnbr2").Visible = False
        'dgvBookings.Columns("strDebText").Visible = False
        'dgvBookings.Columns("strRGBemerkung").Visible = False
        'dgvBookings.Columns("strDebRef").Visible = False
        'dgvBookings.Columns("strZahlBed").Visible = False
        'dgvBookings.Columns("strDebStatusBitLog").Visible = False
        'dgvBookings.Columns("strDebBookStatus").Visible = False
        'dgvBookings.Columns("booBooked").Visible = False
        'dgvBookings.Columns("datBooked").Visible = False
        'dgvBookings.Columns("lngBelegNr").Visible = False


    End Sub


    Private Sub InitdgvKreditorenSub()

        dgvBookingSub.ShowCellToolTips = False
        dgvBookingSub.AllowUserToAddRows = False
        dgvBookingSub.AllowUserToDeleteRows = False
        'dgvBookingSub.Columns("strRGNr").DisplayIndex = 0
        'dgvBookingSub.Columns("strRGNr").Width = 60
        'dgvBookingSub.Columns("strRGNr").HeaderText = "RG-Nr"
        dgvBookingSub.Columns("lngKredID").DisplayIndex = 0
        dgvBookingSub.Columns("lngKredID").Width = 60
        dgvBookingSub.Columns("lngKredID").HeaderText = "Kred-ID"
        dgvBookingSub.Columns("intSollHaben").Width = 30
        dgvBookingSub.Columns("intSollHaben").HeaderText = "S/H"
        dgvBookingSub.Columns("lngKto").Width = 50
        dgvBookingSub.Columns("lngKto").HeaderText = "Konto"
        dgvBookingSub.Columns("strKtoBez").HeaderText = "Bezeichnung"
        dgvBookingSub.Columns("lngKST").Width = 50
        dgvBookingSub.Columns("lngKST").HeaderText = "KST"
        dgvBookingSub.Columns("strKSTBez").Width = 60
        dgvBookingSub.Columns("strKSTBez").HeaderText = "Bezeichnung"
        dgvBookingSub.Columns("dblNetto").Width = 60
        dgvBookingSub.Columns("dblNetto").HeaderText = "Netto"
        dgvBookingSub.Columns("dblNetto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        'dgvBookingSub.Columns("dblNetto").DefaultCellStyle.Format = "N2"
        dgvBookingSub.Columns("dblMwSt").Width = 50
        dgvBookingSub.Columns("dblMwSt").HeaderText = "MwSt"
        dgvBookingSub.Columns("dblMwSt").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        'dgvBookingSub.Columns("dblMwSt").DefaultCellStyle.Format = "N2"
        dgvBookingSub.Columns("dblBrutto").Width = 60
        dgvBookingSub.Columns("dblBrutto").HeaderText = "Brutto"
        dgvBookingSub.Columns("dblBrutto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        'dgvBookingSub.Columns("dblBrutto").DefaultCellStyle.Format = "N2"
        dgvBookingSub.Columns("dblMwStSatz").Width = 40
        dgvBookingSub.Columns("dblMwStSatz").HeaderText = "MwStS"
        dgvBookingSub.Columns("dblMwStSatz").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvBookingSub.Columns("dblMwStSatz").DefaultCellStyle.Format = "N1"
        dgvBookingSub.Columns("strMwStKey").Width = 40
        dgvBookingSub.Columns("strMwStKey").HeaderText = "MwStK"
        dgvBookingSub.Columns("strStatusUBText").HeaderText = "Status"

        dgvBookingSub.Columns("lngID").Visible = False
        'dgvBookingSub.Columns("strArtikel").Visible = False
        'dgvBookingSub.Columns("strStatusUBBitLog").Visible = False
        'dgvBookingSub.Columns("strDebSubText").Visible = False
        'dgvBookingSub.Columns("strDebBookStatus").Visible = False

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

        'Tabelle Debi/ Kredi Head erstellen
        'objdtDebitorenHead = Main.tblDebitorenHead()
        objdtKreditorenHead = Main.tblKreditorenHead()

        'Tabelle Debi/ Kredi Sub erstellen
        'objdtDebitorenSub = Main.tblDebitorenSub()
        objdtKreditorenSub = Main.tblKreditorenSub()

        'Info - Tabelle erstellen
        objdtInfo = Main.tblInfo()

        'Subbuchungen ausblenden, kann für Testzwecke aktiviert werden
        'dgvBookingSub.DataSource = objdtDebitorenSub
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

        'Version
        'Debug.Print("1 " + System.Reflection.Assembly.GetExecutingAssembly().GetName.Version.ToString)
        lblVersion.Text = "V " + System.Reflection.Assembly.GetExecutingAssembly().GetName.Version.ToString
        'Debug.Print("2 " + My.Application.Info.Version.ToString)
        'Debug.Print("3 " + Application.ProductVersion.ToString)

        'Call InitdgvDebitoren()

        ''DGV Debitoren
        'dgvBookings.DataSource = objdtDebitorenHead
        'objdbConn.Open()
        'Call InitdgvDebitoren()
        'Call InitdgvDebitorenSub()
        'objdbConn.Close()

    End Sub

    Private Sub butImport_Click(sender As Object, e As EventArgs) Handles butImport.Click


        Dim intReturnValue As Int32
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
        Dim strBeBuEintrag As String = ""
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
        Dim intEigeneBank As Int16

        Dim selDebiSub() As DataRow
        Dim strSteuerInfo() As String
        Dim strDebitor() As String
        Dim strDebiLine As String

        'Sammelbeleg
        Dim intCommonKonto As Int32
        Dim strDebiCurrency As String
        Dim strKrediCurrency As String
        Dim dblBuchBetrag As Double
        Dim dblBasisBetrag As Double
        Dim strErfassungsDatum As String
        Dim strRGNbr As String
        Dim booBooingok As Boolean

        Dim intLaufNbr As Int32
        Dim strBeleg As String
        Dim strBelegArr() As String
        Dim dblSplitPayed As Double


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

                        'Verdopplung interne BelegsNummer verhindern
                        DbBhg.CheckDoubleIntBelNbr = "J"

                        If row("dblDebBrutto") < 0 Then
                            'Gutschrift
                            'Zuerst Beleg-Nummerieungung aktivieren
                            DbBhg.IncrBelNbr = "J"
                            'Belegsnummer abholen
                            intDebBelegsNummer = DbBhg.GetNextBelNbr("G")
                            strExtBelegNbr = row("strOPNr")
                            'Beträge drehen
                            row("dblDebBrutto") = row("dblDebBrutto") * -1
                            row("dblDebMwSt") = row("dblDebMwSt") * -1
                            row("dblDebNetto") = row("dblDebNetto") * -1

                            strBuchType = "G"

                        Else

                            If IIf(IsDBNull(row("strOPNr")), "", row("strOPNr")) = "" Then
                                'strExtBelegNbr = row("strOPNr")

                                'Zuerst Beleg-Nummerieungung aktivieren
                                DbBhg.IncrBelNbr = "J"
                                'Belegsnummer abholen
                                intDebBelegsNummer = DbBhg.GetNextBelNbr("R")
                            Else
                                If Strings.Len(row("strOPNr")) > 9 Then
                                    'Zahl zu gross
                                    DbBhg.IncrBelNbr = "J"
                                    'Belegsnummer abholen
                                    intDebBelegsNummer = DbBhg.GetNextBelNbr("R")
                                    strExtBelegNbr = row("strOPNr")
                                Else
                                    'Beleg-Nummerierung abschalten
                                    DbBhg.IncrBelNbr = "N"
                                    intDebBelegsNummer = Main.FcCleanRGNrStrict(row("strOPNr"))
                                    strExtBelegNbr = row("strOPNr")
                                End If

                            End If

                            strBuchType = "R"

                        End If

                        'Variablen zuweisen
                        'Sachbearbeiter aus Debitor auslesen
                        strDebiLine = DbBhg.ReadDebitor3(row("lngDebNbr") * -1, "")
                        strDebitor = Split(strDebiLine, "{>}")
                        strSachBID = strDebitor(30)
                        'strExtBelegNbr = row("strDebRGNbr")
                        intDebitorNbr = row("lngDebNbr")
                        strValutaDatum = Format(row("datDebValDatum"), "yyyyMMdd").ToString
                        strBelegDatum = Format(row("datDebRGDatum"), "yyyyMMdd").ToString
                        If IsDBNull(row("datDebDue")) Then
                            strVerfallDatum = ""
                        Else
                            strVerfallDatum = Format(row("datDebDue"), "yyyyMMdd").ToString
                        End If
                        strReferenz = row("strDebReferenz")
                        strMahnerlaubnis = "" 'Format(row("datDebRGDatum"), "yyyyMMdd").ToString
                        dblBetrag = row("dblDebBrutto")
                        'Bei SplittBill 2ter Rechnung Text anfügen
                        If row("booLinked") Then
                            strDebiText = row("strDebText") + ", Splitt-Bill"
                        Else
                            strDebiText = row("strDebText")
                        End If
                        'strDebiText = row("strDebText")
                        strCurrency = row("strDebCur")
                        If strCurrency <> "CHF" Then 'Muss ergänzt werden => Was ist Leitwährung auf dem Konto
                            dblKurs = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg)
                        Else
                            dblKurs = 1.0#
                        End If
                        intEigeneBank = row("strDebiBank")

                        Try
                            booBooingok = True
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
                                                     strCurrency,
                                                     "",
                                                     intEigeneBank.ToString)

                        Catch ex As Exception
                            MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Belegkopf " + intDebBelegsNummer.ToString + ", RG " + strRGNbr + ", Debitor " + intDebitorNbr.ToString)
                            If (Err.Number And 65535) < 10000 Then
                                booBooingok = False
                            Else
                                booBooingok = True
                            End If

                        End Try


                        selDebiSub = objdtDebitorenSub.Select("strRGNr='" + row("strDebRGNbr") + "'")
                        strRGNbr = row("strDebRGNbr")

                        For Each SubRow As DataRow In selDebiSub

                            'Bei zweiter Splitt-Bill 2ter Rechung hier eingreifen
                            'Gegenkonto auf 1092, MwStKey auf 'null' setzen, KST = 0
                            If row("booLinked") Then
                                intGegenKonto = 1092
                                SubRow("dblNetto") = SubRow("dblBrutto")
                                SubRow("strMwStKey") = "null"
                                SubRow("lngKST") = 0
                            Else
                                intGegenKonto = SubRow("lngKto")
                            End If
                            strFibuText = SubRow("strDebSubText")
                            If intGegenKonto <> 6906 Then
                                If strBuchType = "R" Then
                                    dblNettoBetrag = SubRow("dblNetto") * -1
                                Else
                                    dblNettoBetrag = SubRow("dblNetto")
                                End If
                            Else 'Rundungsdifferenzen
                                If strBuchType = "R" Then
                                    dblNettoBetrag = SubRow("dblBrutto") * -1
                                Else
                                    dblNettoBetrag = SubRow("dblBrutto")
                                End If
                            End If
                            'dblBebuBetrag = 1000.0#
                            If SubRow("lngKST") > 0 Then
                                strBeBuEintrag = SubRow("lngKST").ToString + "{<}" + SubRow("strDebSubText") + "{<}" + "CALCULATE" + "{>}"    '"PROD{<}BebuText{<}" + dblBebuBetrag.ToString + "{>}"
                            Else
                                'strBeBuEintrag = "999999" + "{<}" + SubRow("strDebSubText") + "{<}" + "CALCULATE" + "{>}"
                                strBeBuEintrag = ""
                            End If
                            If Not IsDBNull(SubRow("strMwStKey")) And SubRow("strMwStKey") <> "null" Then 'And SubRow("strMwStKey") <> "25" Then
                                If strBuchType = "R" Then
                                    strSteuerFeld = Main.FcGetSteuerFeld(FBhg, SubRow("lngKto"), SubRow("strDebSubText"), SubRow("dblBrutto") * -1, SubRow("strMwStKey"), SubRow("dblMwSt") * -1)     '"25{<}DEBI D Bhg Export MwSt{<}0{>}"
                                Else
                                    strSteuerFeld = Main.FcGetSteuerFeld(FBhg, SubRow("lngKto"), SubRow("strDebSubText"), SubRow("dblBrutto"), SubRow("strMwStKey"), SubRow("dblMwSt"))     '"25{<}DEBI D Bhg Export MwSt{<}0{>}"
                                End If
                            Else
                                strSteuerFeld = "STEUERFREI"
                            End If
                            'strSteuerInfo = Split(FBhg.GetKontoInfo(intGegenKonto.ToString), "{>}")
                            'Debug.Print("Konto-Info: " + strSteuerInfo(26))

                            Try

                                booBooingok = True
                                Call DbBhg.SetVerteilung(intGegenKonto.ToString, strFibuText, dblNettoBetrag.ToString, strSteuerFeld, strBeBuEintrag)

                            Catch ex As Exception
                                MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Verteilung " + intDebBelegsNummer.ToString + ", RG " + strRGNbr + ", Konto " + SubRow("lngKto").ToString)
                                If (Err.Number And 65535) < 10000 Then
                                    booBooingok = False
                                Else
                                    booBooingok = True
                                End If

                            End Try

                            strSteuerFeld = ""
                            strBeBuEintrag = ""

                            'Status Sub schreiben
                            Application.DoEvents()

                        Next

                        Try

                            booBooingok = True
                            Call DbBhg.WriteBuchung()

                            'Bei SplittBill 2ter Rechnung TZahlung auf LinkedRG machen
                            'Prinzip: Beleg einlesen anhand und Betrag ausrechnen => Summe Beleg - diesen Beleg
                            If row("booLinked") Then
                                'Betrag von Beleg 1 holen
                                intLaufNbr = DbBhg.doesBelegExist2(row("lngLinkedDeb").ToString,
                                                                   row("strDebCur"),
                                                                   row("lngLinkedRG").ToString,
                                                                   "NOT_SET",
                                                                   "R",
                                                                   "NOT_SET",
                                                                   "NOT_SET",
                                                                   "NOT_SET")

                                If intLaufNbr > 0 Then
                                    strBeleg = DbBhg.GetBeleg(row("lngLinkedDeb").ToString,
                                                              intLaufNbr.ToString)

                                    strBelegArr = Split(strBeleg, "{>}")
                                    If strBelegArr(4) = "B" Then 'schon bezahlt
                                        'Ausbuchen?
                                    Else

                                        'Welcher Betrag wurde schon bezahlt?
                                        dblSplitPayed = dblBetrag

                                        'Teilzahlung buchen
                                        Call DbBhg.SetZahlung(344,
                                                          strBelegDatum,
                                                          strValutaDatum,
                                                          row("strDebCur"),
                                                          dblKurs,
                                                          "",
                                                          "",
                                                          row("lngLinkedDeb"),
                                                          dblSplitPayed.ToString,
                                                          row("strDebCur"),
                                                          ,
                                                          strDebiText + ", TZ auf RG1")

                                        Call DbBhg.WriteTeilzahlung4(intLaufNbr.ToString,
                                                                 strDebiText + ", TZ auf RG1",
                                                                 "NOT_SET",
                                                                 ,
                                                                 "NOT_SET",
                                                                 "NOT_SET",
                                                                 "DEFAULT",
                                                                 "DEFAULT")

                                    End If

                                End If

                            End If

                        Catch ex As Exception
                            'MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Belegerstellung " + intDebBelegsNummer.ToString + ", RG " + strRGNbr)
                            If (Err.Number And 65535) < 10000 Then
                                MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Belegerstellung nicht möglich" + intDebBelegsNummer.ToString + ", RG " + strRGNbr)
                                'booBooingok = False
                            Else
                                MessageBox.Show(ex.Message, "Warnung " + (Err.Number And 65535).ToString + " Belegerstellung " + intDebBelegsNummer.ToString + ", RG " + strRGNbr)
                                booBooingok = True
                            End If

                        End Try


                    Else

                        'Buchung nur in Fibu
                        'Prinzip Funktion WriteBuchung() anwenden mit allen Parametern

                        'Verdopplung interne BelegsNummer verhindern
                        FBhg.CheckDoubleIntBelNbr = "J"

                        If IIf(IsDBNull(row("strOPNr")), "", row("strOPNr")) <> "" And IIf(IsDBNull(row("lngDebIdentNbr")), 0, row("lngDebIdentNbr")) <> 0 Then
                            'Belegsnummer abholen fall keine Beleg-Nummer angegeben
                            intDebBelegsNummer = FBhg.GetNextBelNbr()
                            'Prüfen ob wirklich frei
                            intReturnValue = 10
                            Do Until intReturnValue = 0
                                intReturnValue = FBhg.doesBelegExist(intDebBelegsNummer,
                                                                     "NOT_SET",
                                                                     "NOT_SET",
                                                                     Microsoft.VisualBasic.Left(cmbPerioden.SelectedItem, 4) + "0101",
                                                                     Microsoft.VisualBasic.Left(cmbPerioden.SelectedItem, 4) + "1231")
                                If intReturnValue <> 0 Then
                                    intDebBelegsNummer += 1
                                End If
                            Loop
                            Debug.Print("Belegnummer taken: " + intDebBelegsNummer.ToString)
                        Else
                            If IIf(IsDBNull(row("strOPNr")), "", row("strOPNr")) <> "" Then
                                intDebBelegsNummer = Convert.ToInt32(row("strOPNr"))
                            Else
                                intDebBelegsNummer = row("lngDebIdentNbr")
                            End If
                        End If
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
                        strRGNbr = row("strDebRGNbr")

                        If selDebiSub.Length = 2 Then

                            'Initialisieren
                            dblNettoBetrag = 0
                            dblSollBetrag = 0
                            dblHabenBetrag = 0
                            strBeBuEintrag = ""
                            strBeBuEintragSoll = ""
                            strBeBuEintragHaben = ""
                            strSteuerFeld = ""
                            strSteuerFeldHaben = ""
                            strSteuerFeldSoll = ""

                            For Each SubRow As DataRow In selDebiSub

                                If SubRow("intSollHaben") = 0 Then 'Soll

                                    intSollKonto = SubRow("lngKto")
                                    dblKursSoll = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg, intSollKonto)
                                    'strSteuerInfo = Split(FBhg.GetKontoInfo(intSollKonto.ToString), "{>}")
                                    'Debug.Print("Konto-Info Soll: " + strSteuerInfo(26))
                                    dblSollBetrag = SubRow("dblNetto")
                                    strDebiTextSoll = SubRow("strDebSubText")
                                    If SubRow("dblMwSt") > 0 Then
                                        strSteuerFeldSoll = Main.FcGetSteuerFeld(FBhg,
                                                                                 SubRow("lngKto"),
                                                                                 strDebiTextSoll,
                                                                                 SubRow("dblBrutto") * dblKursSoll,
                                                                                 SubRow("strMwStKey"),
                                                                                 SubRow("dblMwSt"))
                                    Else
                                        strSteuerFeldSoll = "STEUERFREI"
                                    End If
                                    If SubRow("lngKST") > 0 Then
                                        strBeBuEintragSoll = SubRow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
                                    End If


                                ElseIf SubRow("intSollHaben") = 1 Then 'Haben

                                    intHabenKonto = SubRow("lngKto")
                                    dblKursHaben = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg, intHabenKonto)
                                    'strSteuerInfo = Split(FBhg.GetKontoInfo(intSollKonto.ToString), "{>}")
                                    'Debug.Print("Konto-Info Haben: " + strSteuerInfo(26))
                                    dblHabenBetrag = SubRow("dblNetto") * -1
                                    'dblHabenBetrag = dblSollBetrag
                                    strDebiTextHaben = SubRow("strDebSubText")
                                    If SubRow("dblMwSt") * -1 > 0 Then
                                        strSteuerFeldHaben = Main.FcGetSteuerFeld(FBhg,
                                                                                  SubRow("lngKto"),
                                                                                  strDebiTextHaben,
                                                                                  SubRow("dblBrutto") * dblKursHaben * -1,
                                                                                  SubRow("strMwStKey"),
                                                                                  SubRow("dblMwSt") * -1)
                                    Else
                                        strSteuerFeldHaben = "STEUERFREI"
                                    End If
                                    If SubRow("lngKST") > 0 Then
                                        strBeBuEintragHaben = SubRow("lngKST").ToString + "{<}" + strDebiTextHaben + "{<}" + "CALCULATE" + "{>}"
                                    End If

                                Else

                                    MsgBox("Nicht definierter Wert Sub-Buchungs-SollHaben: " + SubRow("intSollHaben").ToString)

                                End If
                                Application.DoEvents()

                            Next

                            'Tieferer Betrag für die Gesamt-Buchung herausfinden
                            If dblSollBetrag <= dblHabenBetrag Then
                                dblNettoBetrag = dblSollBetrag
                            ElseIf dblHabenBetrag < dblSollBetrag Then
                                dblNettoBetrag = dblHabenBetrag
                            End If

                            Try

                                booBooingok = True
                                'Buchung ausführen
                                Call FBhg.WriteBuchung(0,
                                                   intDebBelegsNummer,
                                                   strBelegDatum,
                                                   intSollKonto.ToString,
                                                   strDebiTextSoll,
                                                   strCurrency,
                                                   dblKursSoll.ToString,
                                                   (dblNettoBetrag * dblKursSoll).ToString,
                                                   strSteuerFeldSoll,
                                                   intHabenKonto.ToString,
                                                   strDebiTextHaben,
                                                   strCurrency,
                                                   dblKursHaben.ToString,
                                                   (dblNettoBetrag * dblKursHaben).ToString,
                                                   strSteuerFeldHaben,
                                                   strCurrency,
                                                   dblKurs.ToString,
                                                   (dblNettoBetrag * dblKurs).ToString,
                                                   dblNettoBetrag.ToString,
                                                   strBeBuEintragSoll,
                                                   strBeBuEintragHaben,
                                                   strValutaDatum)

                            Catch ex As Exception
                                MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Belegerstellung " + intDebBelegsNummer.ToString + ", RG " + strRGNbr)
                                If (Err.Number And 65535) < 10000 Then
                                    booBooingok = False
                                Else
                                    booBooingok = True
                                End If

                            End Try

                            'Initialisieren
                            'dblNettoBetrag = 0
                            'dblSollBetrag = 0
                            'dblHabenBetrag = 0
                            'strBeBuEintrag = ""
                            'strBeBuEintragSoll = ""
                            'strBeBuEintragHaben = ""
                            'strSteuerFeld = ""
                            'strSteuerFeldHaben = ""
                            'strSteuerFeldSoll = ""


                            'Vergebene Nummer checken
                            'intDebBelegsNummer = FBhg.GetBuchLaufnr()

                        Else
                            'Sammelbeleg
                            'MsgBox("Nicht 2 Subbuchungen.")
                            'Variablen initiieren
                            strDebiText = row("strDebText")
                            intCommonKonto = row("lngDebKtoNbr") 'Sammelkonto

                            'Beleg-Kopf
                            Call FBhg.SetSammelBhgCommonT2(strValutaDatum,
                                                           intDebBelegsNummer.ToString,
                                                           intCommonKonto.ToString,
                                                           strDebiText,
                                                           strBelegDatum)

                            'Buchungen
                            For Each SubRow As DataRow In selDebiSub

                                'Common - Konto ausblenden da sonst Doppelbuchung
                                If SubRow("lngKto") <> intCommonKonto Then

                                    intSollKonto = 0
                                    strDebiTextSoll = ""
                                    strDebiCurrency = ""
                                    dblKursSoll = 0
                                    dblSollBetrag = 0
                                    strSteuerFeldSoll = ""
                                    intHabenKonto = 0
                                    strDebiTextHaben = ""
                                    strKrediCurrency = ""
                                    dblKursHaben = 0
                                    dblHabenBetrag = 0
                                    strSteuerFeldHaben = ""
                                    dblBuchBetrag = 0
                                    dblBasisBetrag = 0
                                    strBeBuEintragSoll = ""
                                    strBeBuEintragHaben = ""
                                    strErfassungsDatum = Format(Date.Today(), "yyyyMMdd").ToString

                                    If SubRow("intSollHaben") = 0 And SubRow("lngKto") <> intCommonKonto Then 'Soll

                                        intSollKonto = SubRow("lngKto")
                                        strDebiTextSoll = SubRow("strDebSubText")
                                        strDebiCurrency = strCurrency
                                        dblKursSoll = 1 / Main.FcGetKurs(strCurrency, strValutaDatum, FBhg, intSollKonto)
                                        dblSollBetrag = SubRow("dblNetto")
                                        If SubRow("dblMwSt") > 0 Then
                                            strSteuerFeldSoll = Main.FcGetSteuerFeld(FBhg, SubRow("lngKto"), strDebiTextSoll, SubRow("dblBrutto") * dblKursSoll, SubRow("strMwStKey"), SubRow("dblMwSt"))
                                        Else
                                            strSteuerFeldSoll = "STEUERFREI"
                                        End If
                                        If SubRow("lngKST") > 0 Then
                                            strBeBuEintragSoll = SubRow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
                                        End If

                                        'Haben Seite Common-Konto
                                        intHabenKonto = intCommonKonto
                                        strDebiTextHaben = SubRow("strDebSubText")
                                        strKrediCurrency = strCurrency
                                        dblKursHaben = dblKursSoll
                                        dblHabenBetrag = SubRow("dblNetto")
                                        'If SubRow("dblMwSt") > 0 Then
                                        ' strSteuerFeldHaben = Main.FcGetSteuerFeld(FBhg, SubRow("lngKto"), strDebiTextSoll, SubRow("dblBrutto") * dblKursSoll, SubRow("strMwStKey"), SubRow("dblMwSt"))
                                        'End If
                                        If SubRow("lngKST") > 0 Then
                                            strBeBuEintragHaben = SubRow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
                                        End If

                                        'Betrag
                                        dblBuchBetrag = SubRow("dblBrutto")
                                        dblBasisBetrag = SubRow("dblBrutto") 'Bei nicht CHF umrechnen

                                    ElseIf SubRow("intSollHaben") = 1 Then 'Haben

                                        intHabenKonto = SubRow("lngKto")
                                        strDebiTextHaben = SubRow("strDebSubText")
                                        strKrediCurrency = strCurrency
                                        dblKursHaben = 1 / Main.FcGetKurs(strCurrency, strValutaDatum, FBhg, intHabenKonto)
                                        dblHabenBetrag = SubRow("dblNetto") * -1
                                        If (SubRow("dblMwSt") * -1) > 0 Then
                                            strSteuerFeldHaben = Main.FcGetSteuerFeld(FBhg, SubRow("lngKto"), strDebiTextHaben, SubRow("dblBrutto") * dblKursHaben * -1, SubRow("strMwStKey"), SubRow("dblMwSt") * -1)
                                        Else
                                            strSteuerFeldHaben = "STEUERFREI"
                                        End If
                                        If SubRow("lngKST") > 0 Then
                                            strBeBuEintragHaben = SubRow("lngKST").ToString + "{<}" + strDebiTextHaben + "{<}" + "CALCULATE" + "{>}"
                                        End If

                                        'Soll - Seite Common - Konto 
                                        intSollKonto = intCommonKonto
                                        strDebiTextSoll = SubRow("strDebSubText")
                                        strDebiCurrency = strCurrency
                                        dblKursSoll = dblKursHaben
                                        dblSollBetrag = SubRow("dblNetto") * -1

                                        'If SubRow("dblMwSt") * -1 > 0 Then
                                        ' strSteuerFeldSoll = Main.FcGetSteuerFeld(FBhg, SubRow("lngKto"), strDebiTextSoll, SubRow("dblBrutto") * dblKursSoll * -1, SubRow("strMwStKey"), SubRow("dblMwSt") * -1)
                                        'End If
                                        If SubRow("lngKST") > 0 Then
                                            strBeBuEintragSoll = SubRow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
                                        End If

                                        dblBuchBetrag = SubRow("dblBrutto") * -1
                                        dblBasisBetrag = SubRow("dblBrutto") * -1 'Bei nicht CHF umrechnen

                                    End If
                                    'Buchungsbetrag von Kopfbuchung nehmen
                                    'dblBuchBetrag = row("dblDebBrutto")
                                    'dblBasisBetrag = row("dblDebBrutto")

                                    Call FBhg.SetSammelBhgT(intSollKonto.ToString,
                                                            strDebiTextSoll,
                                                            strDebiCurrency,
                                                            dblKursSoll.ToString,
                                                            dblSollBetrag.ToString,
                                                            strSteuerFeldSoll,
                                                            intHabenKonto.ToString,
                                                            strDebiTextHaben,
                                                            strKrediCurrency,
                                                            dblKursHaben.ToString,
                                                            dblHabenBetrag.ToString,
                                                            strSteuerFeldHaben,
                                                            strCurrency,
                                                            dblKurs.ToString,
                                                            dblBuchBetrag.ToString,
                                                            dblBasisBetrag.ToString,
                                                            strBeBuEintragSoll,
                                                            strBeBuEintragHaben,
                                                            strErfassungsDatum)
                                End If

                            Next

                            'Sammelbeleg schreiben
                            Try

                                booBooingok = True
                                Call FBhg.WriteSammelBhgT()

                            Catch ex As Exception
                                MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Belegerstellung " + intDebBelegsNummer.ToString + ", RG " + strRGNbr)
                                If (Err.Number And 65535) < 10000 Then
                                    booBooingok = False
                                Else
                                    booBooingok = True
                                End If
                            End Try


                        End If

                    End If

                    If booBooingok Then
                        'Status Head schreiben
                        row("strDebBookStatus") = row("strDebStatusBitLog")
                        row("booBooked") = True
                        row("datBooked") = Now()
                        row("lngBelegNr") = intDebBelegsNummer
                        Application.DoEvents()

                        'Status in File RG-Tabelle schreiben
                        intReturnValue = MainDebitor.FcWriteToRGTable(cmbBuha.SelectedValue, row("strDebRGNbr"), row("datBooked"), row("lngBelegNr"), objdbAccessConn, objOracleConn, objdbConn)
                        If intReturnValue <> 0 Then
                            'Throw an exception
                        End If

                        'Evtl. Query nach Buchung ausführen
                        Call MainDebitor.FcExecuteAfterDebit(cmbBuha.SelectedValue, objdbConn)
                    End If

                End If

            Next


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Belegerstellung " + intDebBelegsNummer.ToString + ", RG " + strRGNbr)

        Finally
            'Neu aufbauen
            butDebitoren_Click(butDebitoren, EventArgs.Empty)

            Me.Cursor = Cursors.Default

        End Try

    End Sub


    Private Sub dgvDebitoren_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvBookings.CellValueChanged

        Dim intDecidiveCell As Integer

        If intMode = 0 Then
            intDecidiveCell = 2
        Else
            intDecidiveCell = 3
        End If

        If e.ColumnIndex = intDecidiveCell And e.RowIndex >= 0 Then


            If intMode = 0 Then
                If dgvBookings.Rows(e.RowIndex).Cells("booDebBook").Value Then

                    'MsgBox("Geändert zu checked " + dgvDebitoren.Rows(e.RowIndex).Cells("strDebRGNbr").Value + ", " + dgvDebitoren.Rows(e.RowIndex).Cells("booDebBook").Value.ToString + Val(dgvDebitoren.Rows(e.RowIndex).Cells("strDebStatusBitLog").Value).ToString)
                    'Zulassen? = keine Fehler
                    If Val(dgvBookings.Rows(e.RowIndex).Cells("strDebStatusBitLog").Value) <> 0 And Val(dgvBookings.Rows(e.RowIndex).Cells("strDebStatusBitLog").Value) <> 10000 Then
                        MsgBox("Rechnung ist nicht buchbar.", vbOKOnly + vbExclamation, "Nicht buchbar")
                        dgvBookings.Rows(e.RowIndex).Cells("booDebBook").Value = False
                    End If

                End If

            Else

                If dgvBookings.Rows(e.RowIndex).Cells("booKredBook").Value Then

                    'MsgBox("Geändert zu checked " + dgvDebitoren.Rows(e.RowIndex).Cells("strDebRGNbr").Value + ", " + dgvDebitoren.Rows(e.RowIndex).Cells("booDebBook").Value.ToString + Val(dgvDebitoren.Rows(e.RowIndex).Cells("strDebStatusBitLog").Value).ToString)
                    'Zulassen? = keine Fehler
                    If Val(dgvBookings.Rows(e.RowIndex).Cells("strKredStatusBitLog").Value) <> 0 Then
                        MsgBox("Rechnung ist nicht buchbar.", vbOKOnly + vbExclamation, "Nicht buchbar")
                        dgvBookings.Rows(e.RowIndex).Cells("booKredBook").Value = False
                    End If

                End If

            End If

        End If

    End Sub


    Private Sub dgvDebitoren_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvBookings.CellContentClick

        Try

            If e.RowIndex >= 0 Then

                If intMode = 0 Then 'Debitoren
                    dgvBookingSub.DataSource = objdtDebitorenSub.Select("strRGNr='" + dgvBookings.Rows(e.RowIndex).Cells("strDebRGNbr").Value + "'").CopyToDataTable
                Else
                    dgvBookingSub.DataSource = objdtKreditorenSub.Select("lngKredID=" + dgvBookings.Rows(e.RowIndex).Cells("lngKredID").Value.ToString).CopyToDataTable
                End If

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)

        Finally
            dgvBookingSub.Update()

        End Try



    End Sub

    Private Sub butKreditoren_Click(sender As Object, e As EventArgs) Handles butKreditoren.Click

        Dim intReturnValue As Int16

        Me.Cursor = Cursors.WaitCursor

        intMode = 1

        objdtKreditorenHead = Nothing
        objdtKreditorenHeadRead = Nothing
        objdtKreditorenSub = Nothing

        'Tabelle Kredi - Head erstellen
        objdtKreditorenHead = Main.tblKreditorenHead()
        objdtKreditorenHeadRead = Main.tblKreditorenHead()

        'Tabelle Kreditoren - Sub erstellen
        objdtKreditorenSub = Main.tblKreditorenSub()

        'objdtKreditorenHead.Clear()
        'objdtKreditorenSub.Clear()
        'objdtKreditorenHeadRead.Clear()
        objdtInfo.Clear()

        If dgvBookings.Columns.Contains("intBuchungsart") Then
            dgvBookings.Columns.Remove("intBuchungsart")
        End If

        'DGV Kreditoren
        dgvBookings.DataSource = objdtKreditorenHead
        dgvBookingSub.DataSource = objdtKreditorenSub
        objdbConn.Open()
        Call InitdgvKreditoren()
        Call InitdgvKreditorenSub()
        objdbConn.Close()

        Call InitVar()

        Call Main.FcLoginSage(objdbConn, objdbMSSQLConn, objdbSQLcommand, Finanz, FBhg, DbBhg, PIFin, KrBhg, cmbBuha.SelectedValue, objdtInfo, cmbPerioden.SelectedValue)

        intReturnValue = MainKreditor.FcFillKredit(cmbBuha.SelectedValue, objdtKreditorenHeadRead, objdtKreditorenSub, objdbConn, objdbAccessConn)
        If intReturnValue = 1 Then
            MessageBox.Show("Keine Kreditoren-Defintion hinterlegt.", "Keine Definition")
        End If

        Call Main.InsertDataTableColumnName(objdtKreditorenHeadRead, objdtKreditorenHead)

        'Grid neu aufbauen
        dgvBookingSub.Update()
        dgvBookings.Update()
        dgvBookings.Refresh()

        Call Main.FcCheckKredit(cmbBuha.SelectedValue,
                                objdtKreditorenHead,
                                objdtKreditorenSub,
                                Finanz,
                                FBhg,
                                KrBhg,
                                PIFin,
                                objdbConn,
                                objdbConnZHDB02,
                                objdbcommand,
                                objdbcommandZHDB02,
                                objOracleConn,
                                objOracleCmd,
                                objdbAccessConn,
                                objdtInfo,
                                cmbBuha.Text)

        'Anzahl schreiben
        txtNumber.Text = objdtKreditorenHead.Rows.Count.ToString

        'Import Debitoren desattivate
        Me.butImport.Enabled = False
        Me.butImportK.Enabled = True

        Me.Cursor = Cursors.Default

    End Sub

    Private Sub butImportK_Click(sender As Object, e As EventArgs) Handles butImportK.Click

        Dim intReturnValue As Int16
        Dim intKredBelegsNummer As UInt32
        Dim strExtKredBelegsNummer As String

        Dim intKreditorNbr As Int32
        Dim strBuchType As String
        Dim strBelegDatum As String
        Dim strValutaDatum As String
        Dim strVerfallDatum As String
        Dim strReferenz As String
        Dim intKondition As Int32
        Dim intKonditionLN As Int16
        Dim strSachBID As String = ""
        Dim strVerkID As String = ""
        Dim strMahnerlaubnis As String
        Dim sngAktuelleMahnstufe As Single
        Dim dblBetrag As Double
        Dim dblKurs As Double
        Dim strExtBelegNbr As String = ""
        Dim strSkonto As String = ""
        Dim strCurrency As String
        Dim strKrediText As String
        Dim intBankNbr As Int16
        Dim strZahlSperren As String = "N"
        Dim strVorausZahlung As String = "N"
        Dim strErfassungsArt As String = "K"
        Dim intTeilnehmer As Int32
        Dim strTeilnehmer As String
        Dim intEigeneBank As Int32

        Dim intGegenKonto As Int32
        Dim strFibuText As String
        Dim dblNettoBetrag As Double
        Dim dblBruttoBetrag As Double
        Dim dblMwStBetrag As Double
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
        Dim strKrediTextSoll As String = ""
        Dim strKrediTextHaben As String = ""
        Dim dblKursSoll As Double = 0
        Dim dblKursHaben As Double = 0

        Dim selKrediSub() As DataRow
        Dim strSteuerInfo() As String
        Dim strDebiLine As String
        Dim strDebitor() As String

        Dim booBookingok As Boolean

        'Sammelbeleg
        Dim intCommonKonto As Int32



        Try

            'Debitor erstellen, minimal - Angaben

            Me.Cursor = Cursors.WaitCursor

            'Kopfbuchung
            For Each row As DataRow In objdtKreditorenHead.Rows

                If IIf(IsDBNull(row("booKredBook")), False, row("booKredBook")) Then

                    'Test ob OP - Buchung
                    If row("intBuchungsart") = 1 Then

                        'Immer zuerst Belegs-Nummerierung aktivieren, falls vorhanden externe Nummer = OP - Nr. Rg
                        'Führt zu Problemen beim Ausbuchen des DTA - Files
                        'Resultat Besprechnung 17.09.20 mit Muhi/ Andy
                        'DbBhg.IncrBelNbr = "J"
                        'Belegsnummer abholen
                        'intDebBelegsNummer = DbBhg.GetNextBelNbr("R")

                        'Auf Provisorisch setzen
                        Call KrBhg.SetBuchMode("P")

                        'Automatische ESR - Zahlungsverbindung
                        KrBhg.EnableAutoESRZlgVerb = "J"

                        'Eindeutigkeit der internen Beleg-Nummer setzen
                        KrBhg.CheckDoubleIntBelNbr = "J"

                        'If IsDBNull(row("strOPNr")) Or row("StrOPNr") = "" Then
                        'strExtBelegNbr = row("strOPNr")

                        'Zuerst Beleg-Nummerieungung aktivieren
                        KrBhg.IncrBelNbr = "J"
                        'Belegsnummer abholen
                        'intKredBelegsNummer = KrBhg.GetNextBelNbr("R")
                        'Else
                        'Beleg-Nummerierung abschalten
                        'KrBhg.IncrBelNbr = "N"
                        'intKredBelegsNummer = row("strOPNr")
                        'strExtBelegNbr = row("strOPNr")
                        'End If
                        strExtKredBelegsNummer = row("strKredRGNbr")

                        'Variablen zuweisen
                        intKreditorNbr = row("lngKredNbr")
                        If row("dblKredBrutto") < 0 Then
                            strBuchType = "G"
                            'strZahlSperren = "J"
                            row("dblKredBrutto") = row("dblKredBrutto") * -1
                            'Belegsnummer abholen
                            intKredBelegsNummer = KrBhg.GetNextBelNbr("G")
                        Else
                            strBuchType = "R"
                            'strZahlSperren = "N"
                            'Belegsnummer abholen
                            intKredBelegsNummer = KrBhg.GetNextBelNbr("R")
                        End If

                        strValutaDatum = Format(row("datKredValDatum"), "yyyyMMdd").ToString
                        strBelegDatum = Format(row("datKredRGDatum"), "yyyyMMdd").ToString
                        strVerfallDatum = ""
                        'strReferenz = IIf(IsDBNull(row("strKredRef")), "", row("strKredRef"))
                        'If IsDBNull(row("strKrediBank")) Then
                        'intTeilnehmer = 0
                        'Else
                        'Teilnehmer nur bei ESR setzen
                        If row("intPayType") <> 9 Then 'nicht IBAN
                            'QR-Referenz
                            strReferenz = IIf(IsDBNull(row("strKredRef")), "", row("strKredRef"))
                            If row("intPayType") = 10 Then
                                strTeilnehmer = row("strKrediBank")
                            Else
                                strTeilnehmer = Val(row("strKrediBank"))
                            End If
                            intBankNbr = 0
                        Else
                            'IBAN
                            strReferenz = IIf(IsDBNull(row("strKredRef")), "", row("strKredRef"))
                            intBankNbr = IIf(IsDBNull(row("intEBank")), 0, row("intEBank"))
                            strTeilnehmer = ""
                        End If
                        'End If
                        strMahnerlaubnis = "" 'Format(row("datDebRGDatum"), "yyyyMMdd").ToString
                        'Sachbearbeiter aus Debitor auslesen
                        strDebiLine = KrBhg.ReadKreditor3(row("lngKredNbr") * -1, "")
                        strDebitor = Split(strDebiLine, "{>}")
                        strSachBID = strDebitor(29)

                        dblBetrag = row("dblKredBrutto")
                        strKrediText = IIf(IsDBNull(row("strKredText")), "", row("strKredText"))
                        strCurrency = row("strKredCur")
                        'intBankNbr = 0
                        intKondition = IIf(IsDBNull(row("intZKond")), 1, row("intZKond"))
                        'LN 0=automatsich ersterfasste Kondition, -1=Schlechteste Kondition, -2=Beste Kondition
                        intKonditionLN = 0
                        intEigeneBank = row("intintBank")

                        If strCurrency <> "CHF" Then 'Muss ergänzt werden => Was ist Leitwährung auf dem Konto
                            dblKurs = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg)
                        Else
                            dblKurs = 1.0#
                        End If

                        Try
                            booBookingok = True
                            Call KrBhg.SetBelegKopf2(intKredBelegsNummer,
                                                     strValutaDatum,
                                                     intKreditorNbr,
                                                     strExtKredBelegsNummer,
                                                     strBelegDatum,
                                                     strVerfallDatum,
                                                     strKrediText,
                                                     intBankNbr,
                                                     strBuchType,
                                                     strZahlSperren,
                                                     strVorausZahlung,
                                                     intKondition,
                                                     intKonditionLN,
                                                     strSachBID,
                                                     strReferenz,
                                                     strSkonto,
                                                     dblBetrag.ToString,
                                                     strErfassungsArt,
                                                     dblKurs.ToString,
                                                     strCurrency,
                                                     "",
                                                     strTeilnehmer,
                                                     intEigeneBank.ToString)

                        Catch ex As Exception
                            MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Belegkopf ")
                            If (Err.Number And 65535) < 10000 Then
                                booBookingok = False
                            Else
                                booBookingok = True
                            End If

                        End Try

                        selKrediSub = objdtKreditorenSub.Select("lngKredID=" + row("lngKredID").ToString)

                        For Each SubRow As DataRow In selKrediSub

                            intGegenKonto = SubRow("lngKto")
                            strFibuText = SubRow("strKredSubText")
                            'Soll auf Minus setzen
                            'If SubRow("intSollHaben") = 1 Then
                            'dblNettoBetrag = SubRow("dblNetto") * -1
                            'dblMwStBetrag = SubRow("dblMwSt") * -1
                            'dblBruttoBetrag = SubRow("dblBrutto") * -1
                            'Else
                            If intGegenKonto <> 6906 Then
                                If strBuchType = "R" Then
                                    dblNettoBetrag = SubRow("dblNetto")
                                    dblMwStBetrag = SubRow("dblMwSt")
                                    dblBruttoBetrag = SubRow("dblBrutto")
                                Else
                                    dblNettoBetrag = SubRow("dblNetto") * -1
                                    dblMwStBetrag = SubRow("dblMwSt") * -1
                                    dblBruttoBetrag = SubRow("dblBrutto") * -1
                                End If
                            Else 'Rundungsdifferenzen
                                If strBuchType = "R" Then
                                    dblNettoBetrag = SubRow("dblBrutto")
                                    dblMwStBetrag = SubRow("dblMwSt")
                                    dblBruttoBetrag = SubRow("dblBrutto")
                                Else
                                    dblNettoBetrag = SubRow("dblBrutto") * -1
                                    dblMwStBetrag = SubRow("dblMwSt") * -1
                                    dblBruttoBetrag = SubRow("dblBrutto") * -1
                                End If

                            End If

                            'If strBuchType = "R" Then
                            '    dblNettoBetrag = SubRow("dblNetto")
                            '    dblMwStBetrag = SubRow("dblMwSt")
                            '    dblBruttoBetrag = SubRow("dblBrutto")
                            'Else
                            '    dblNettoBetrag = SubRow("dblNetto") * -1
                            '    dblMwStBetrag = SubRow("dblMwSt") * -1
                            '    dblBruttoBetrag = SubRow("dblBrutto") * -1
                            'End If
                            'End If
                            'dblBebuBetrag = 1000.0#
                            If SubRow("lngKST") > 0 Then
                                strBeBuEintrag = SubRow("lngKST").ToString + "{<}" + SubRow("strKredSubText") + "{<}" + "CALCULATE" + "{>}"    '"PROD{<}BebuText{<}" + dblBebuBetrag.ToString + "{>}"
                            Else
                                'strBeBuEintrag = "00" + "{<}" + SubRow("strKredSubText") + "{<}" + "0" + "{>}"
                            End If
                            If Not IsDBNull(SubRow("strMwStKey")) And SubRow("strMwStKey") <> "null" Then ' And SubRow("strMwStKey") <> "25" Then
                                strSteuerFeld = Main.FcGetSteuerFeld(FBhg, SubRow("lngKto"), SubRow("strKredSubText"), dblBruttoBetrag, SubRow("strMwStKey"), dblMwStBetrag)     '"25{<}DEBI D Bhg Export MwSt{<}0{>}"
                            Else
                                strSteuerFeld = "STEUERFREI"
                            End If

                            'strSteuerInfo = Split(FBhg.GetKontoInfo(intGegenKonto.ToString), "{>}")
                            'Debug.Print("Konto-Info: " + strSteuerInfo(26))

                            Try
                                booBookingok = True
                                Call KrBhg.SetVerteilung(intGegenKonto.ToString, strFibuText, dblNettoBetrag.ToString, strSteuerFeld, strBeBuEintrag)

                            Catch ex As Exception
                                MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Verteilung " + SubRow("strKredSubText") + ", Konto " + SubRow("lngKto").ToString)
                                If (Err.Number And 65535) < 10000 Then
                                    booBookingok = False
                                Else
                                    booBookingok = True
                                End If

                            End Try


                            strSteuerFeld = ""
                            strBeBuEintrag = ""

                            'Status Sub schreiben
                            Application.DoEvents()

                        Next


                        Try
                            booBookingok = True
                            Call KrBhg.WriteBuchung()

                        Catch ex As Exception
                            If (Err.Number And 65535) < 10000 Then
                                MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Belegerstellung nicht möglich")
                                booBookingok = False
                            Else
                                MessageBox.Show(ex.Message, "Warnung " + (Err.Number And 65535).ToString + " Belegerstellung")
                                booBookingok = True
                            End If

                        End Try

                        Application.DoEvents()

                        strBeBuEintrag = ""
                        strSteuerFeld = ""
                        dblNettoBetrag = 0
                        dblMwStBetrag = 0
                        dblBruttoBetrag = 0

                    Else

                        'Buchung nur in Fibu
                        'Prinzip Funktion WriteBuchung() anwenden mit allen Parametern
                        'Beleg-Nummerierung aktivieren
                        'DbBhg.IncrBelNbr = "J"
                        'Belegsnummer abholen
                        intKredBelegsNummer = FBhg.GetNextBelNbr()

                        booBookingok = True

                        'Variablen zuweisen
                        strBelegDatum = Format(row("datKredRGDatum"), "yyyyMMdd").ToString
                        strValutaDatum = Format(row("datKredValDatum"), "yyyyMMdd").ToString
                        'strDebiText = row("strDebText")
                        strCurrency = row("strKredCur")
                        If strCurrency <> "CHF" Then 'Muss ergänzt werden => Was ist Leitwährung auf dem Konto
                            dblKurs = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg)
                        Else
                            dblKurs = 1.0#
                        End If

                        selKrediSub = objdtDebitorenSub.Select("lngKredID=" + row("lngKredID").ToString)

                        If selKrediSub.Length = 2 Then

                            For Each SubRow As DataRow In selKrediSub

                                If SubRow("intSollHaben") = 0 Then 'Soll

                                    intSollKonto = SubRow("lngKto")
                                    dblKursSoll = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg, intSollKonto)
                                    'strSteuerInfo = Split(FBhg.GetKontoInfo(intSollKonto.ToString), "{>}")
                                    'Debug.Print("Konto-Info Soll: " + strSteuerInfo(26))
                                    dblSollBetrag = SubRow("dblNetto")
                                    strKrediTextSoll = SubRow("strDebSubText")
                                    If SubRow("dblMwSt") > 0 Then
                                        strSteuerFeldSoll = Main.FcGetSteuerFeld(FBhg, SubRow("lngKto"), strKrediTextSoll, SubRow("dblBrutto") * dblKursSoll, SubRow("strMwStKey"), SubRow("dblMwSt"))
                                    End If
                                    If SubRow("lngKST") > 0 Then
                                        strBeBuEintragSoll = SubRow("lngKST").ToString + "{<}" + strKrediTextSoll + "{<}" + "CALCULATE" + "{>}"
                                    End If


                                ElseIf SubRow("intSollHaben") = 1 Then 'Haben

                                    intHabenKonto = SubRow("lngKto")
                                    dblKursHaben = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg, intHabenKonto)
                                    'strSteuerInfo = Split(FBhg.GetKontoInfo(intSollKonto.ToString), "{>}")
                                    'Debug.Print("Konto-Info Haben: " + strSteuerInfo(26))
                                    dblHabenBetrag = SubRow("dblNetto")
                                    'dblHabenBetrag = dblSollBetrag
                                    strKrediTextHaben = SubRow("strKredSubText")
                                    If SubRow("dblMwSt") > 0 Then
                                        strSteuerFeldHaben = Main.FcGetSteuerFeld(FBhg, SubRow("lngKto"), strKrediTextHaben, SubRow("dblBrutto") * dblKursHaben, SubRow("strMwStKey"), SubRow("dblMwSt"))
                                    End If
                                    If SubRow("lngKST") > 0 Then
                                        strBeBuEintragHaben = SubRow("lngKST").ToString + "{<}" + strKrediTextHaben + "{<}" + "CALCULATE" + "{>}"
                                    End If

                                Else

                                    'Sammelbeleg
                                    MsgBox("Nicht definierter Wert Sub-Buchungs-SollHaben: " + SubRow("intSollHaben").ToString)
                                    'strKrediText = IIf(IsDBNull(row("strKredText")), "", row("strKredText"))

                                End If

                            Next

                            'Tieferer Betrag für die Gesamt-Buchung herausfinden
                            If dblSollBetrag <= dblHabenBetrag Then
                                dblNettoBetrag = dblSollBetrag
                            ElseIf dblHabenBetrag < dblSollBetrag Then
                                dblNettoBetrag = dblHabenBetrag
                            End If

                            'Buchung ausführen
                            Call FBhg.WriteBuchung(0,
                                                   intKredBelegsNummer,
                                                   strBelegDatum,
                                                   intSollKonto.ToString,
                                                   strKrediTextSoll,
                                                   strCurrency,
                                                   dblKursSoll.ToString,
                                                   (dblNettoBetrag * dblKursSoll).ToString,
                                                   strSteuerFeldSoll,
                                                   intHabenKonto.ToString,
                                                   strKrediTextHaben,
                                                   strCurrency,
                                                   dblKursHaben.ToString,
                                                   (dblNettoBetrag * dblKursHaben).ToString,
                                                   strSteuerFeldHaben,
                                                   strCurrency,
                                                   dblKurs.ToString,
                                                   dblNettoBetrag.ToString,
                                                   (dblNettoBetrag * dblKurs).ToString,
                                                   strBeBuEintragSoll,
                                                   strBeBuEintragHaben,
                                                   strValutaDatum)

                        Else
                            MsgBox("Nicht 2 Subbuchungen.")
                        End If

                    End If

                    If booBookingok Then
                        'Status Head schreiben
                        row("strKredBookStatus") = row("strKredStatusBitLog")
                        row("booBooked") = True
                        row("datBooked") = Now()
                        row("lngBelegNr") = intKredBelegsNummer

                        'Status in File RG-Tabelle schreiben
                        intReturnValue = MainKreditor.FcWriteToKrediRGTable(cmbBuha.SelectedValue,
                                                                        row("lngKredID"),
                                                                        row("datBooked"),
                                                                        row("lngBelegNr"),
                                                                        objdbAccessConn,
                                                                        objOracleConn,
                                                                        objdbConn)
                        If intReturnValue <> 0 Then
                            'Throw an exception
                        End If

                    End If

                End If

            Next


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " bei Buchung-Beleg " + strExtKredBelegsNummer)

        Finally
            'Neu aufbauen
            butKreditoren_Click(butDebitoren, EventArgs.Empty)

            Me.Cursor = Cursors.Default

        End Try


    End Sub

    Private Sub cmbBuha_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cmbBuha.SelectionChangeCommitted

        Call Main.FcReadPeriodsFromMandant(objdbConn,
                                           Finanz,
                                           cmbBuha.SelectedValue,
                                           cmbPerioden)


    End Sub

    Private Sub butMail_Click(sender As Object, e As EventArgs) Handles butMail.Click

        Dim strMailText As String
        Dim intColCounter As Int16
        Dim intRowCounter As Int32

        'String zusammensetzen für Mailtext

        strMailText = "<table border=""1px solid black"">" + vbCrLf
        strMailText += "    <thead>" + vbCrLf
        strMailText += "    <caption>Debitoren</caption>"
        strMailText += "    <tr>" + vbCrLf

        intColCounter = 0

        'Zuerst Titel zusammen setzen
        For intColCounter = 0 To dgvBookings.Columns.Count - 1
            strMailText += "        <th>" + dgvBookings.Columns(intColCounter).HeaderText + "</th>" + vbCrLf
        Next
        strMailText += "    </tr>" + vbCrLf
        strMailText += "    </thead>" + vbCrLf

        'Footer mit Legende
        strMailText += "    <tfoot>" + vbCrLf
        strMailText += "    <tr>" + vbCrLf
        strMailText += "    <td colspan=""6"">ValD = Valuta-Datum nicht möglich</td>"
        strMailText += "    </tr>" + vbCrLf
        strMailText += "    <tr>" + vbCrLf
        strMailText += "    <td colspan=""6"">RgD = Rechnungs-Datum nicht möglich</td>"
        strMailText += "    </tr>" + vbCrLf
        strMailText += "    </tfoot>" + vbCrLf

        'Durch die Tabelle steppen
        strMailText += "    <tbody>" + vbCrLf
        For intRowCounter = 0 To dgvBookings.Rows.Count - 1
            strMailText += "    <tr>" + vbCrLf
            For intColCounter = 0 To dgvBookings.Columns.Count - 1
                strMailText += "        <td>" + dgvBookings.Rows(intRowCounter).Cells(intColCounter).Value.ToString + "</td>" + vbCrLf
            Next
            strMailText += "    </tr>" + vbCrLf
        Next
        strMailText += "    </tbody>" + vbCrLf


        strMailText += "</table>"
        Debug.Print(strMailText)

    End Sub

End Class
