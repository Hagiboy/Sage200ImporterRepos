Option Strict Off
Option Explicit On

Imports MySql.Data.MySqlClient
Imports System.Data.OracleClient
Imports System.Data.SqlClient
Imports CLClassSage200.WFSage200Import
'Imports System.Data.OleDb

Friend Class frmImportMain

    'Wird gebraucht um die GJ zu lesen
    Public Finanz As SBSXASLib.AXFinanz
    'Public FBhg As SBSXASLib.AXiFBhg
    'Public DbBhg As SBSXASLib.AXiDbBhg
    'Public KrBhg As SBSXASLib.AXiKrBhg
    'Public BsExt As SBSXASLib.AXiBSExt
    'Public Adr As SBSXASLib.AXiAdr
    'Public BeBu As SBSXASLib.AXiBeBu
    'Public PIFin As SBSXASLib.AXiPlFin

    'Public Methode As String
    'Public DidOpenmandant As Boolean

    'Public FELD_SEP As String
    'Public REC_SEP As String
    'Public KSTKTR_SEP As String
    'Public FELD_SEP_OUT As String
    'Public REC_SEP_OUT As String
    'Public nID As String

    Dim objdbConn As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString"))
    'Public objdbConnZHDB02 As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))
    'Public objdbMSSQLConn As New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("SQLConnectionString"))
    'Public objdbMySQLMail As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringMail"))
    'Public objdbAccessConn As New OleDb.OleDbConnection
    'Public objdbcommand As New MySqlCommand
    'Public objdbcommandZHDB02 As New MySqlCommand
    'Public objdbSQLcommand As New SqlCommand
    'Public objDABuchhaltungen As New MySqlDataAdapter("SELECT * FROM t_sage_buchhaltungen WHERE NOT Buchh200_Name IS NULL AND NOT Buchh_TableDeb IS NULL ORDER BY Buchh_Bez", objdbConn)
    'Public objdtBuchhaltungen As New DataTable("tbliBuchhaltungen")
    'Public objDACarsGrid As New MySqlDataAdapter("SELECT tblcars.idCar, tblunits.strUnit, tblplates.strPlate, tblcars.strVIN, tblmodelle.strModell FROM tblcars LEFT JOIN tblunits ON tblcars.refUnit = tblunits.idUnit LEFT JOIN tblplates ON tblcars.refPlate = tblplates.idPlate LEFT JOIN tblmodelle ON tblcars.refModell = tblmodelle.idModell", objdbConn)
    'Public objdtDebitor As New DataTable("tbliDebitor")
    'Public objdtDebitorenHead As New DataTable("tbliDebiHead")
    'Public objdtDebitorenHeadRead As New DataTable("tbliDebitorenHeadR")
    'Public objdtDebitorenSub As New DataTable("tbliDebiSub")
    'Public objdtKreditorenHead As New DataTable("tbliKrediHead")
    'Public objdtKreditorenHeadRead As New DataTable("tbliKreditorenHeadR")
    'Public objdtKreditorenSub As New DataTable("tbliKrediSub")
    'Public objdtInfo As New DataTable("tbliInfo")
    'Public objOracleConn As New OracleConnection("Data Source=(DESCRIPTION=" _
    '                + "(ADDRESS=(PROTOCOL=TCP)(HOST=10.0.0.29)(PORT=1521))" _
    '                + "(CONNECT_DATA=(SERVICE_NAME=CISNEW)));" _
    '                + "User Id=cis;Password=sugus;")
    'Public objOracleCmd As New OracleCommand()
    'Public intMode As Int16
    'Public strYear As String
    'Public intTeqNbr As Int16
    'Public intTeqNbrLY As Int16
    'Public intTeqNbrPLY As Int16
    'Public datPeriodFrom As Date
    'Public datPeriodTo As Date
    'Public strPeriodStatus As String
    'Public boodgvSet As Boolean = False

    Public Sub InitVar()


        'PIFin = Nothing
        'KrBhg = Nothing
        'FBhg = Nothing
        'DbBhg = Nothing
        'BsExt = Nothing
        'BeBu = Nothing
        'Adr = Nothing
        'Finanz = Nothing

        ''objdbcommand = Nothing
        ''objdbcommandZHDB02 = Nothing
        ''objdtDebitorenHeadRead.Clear()
        ''objdtDebitorenHead.Clear()
        ''objdtDebitorenSub.Clear()
        'objdtKreditorenHeadRead.Clear()
        'objdtKreditorenHead.Clear()
        'objdtKreditorenSub.Clear()
        ''Call Check_CheckStateChanged(Check, New System.EventArgs())

        'FELD_SEP = "{<}"
        'REC_SEP = "{>}"
        'KSTKTR_SEP = "{-}"

        'FELD_SEP_OUT = "{>}"
        'REC_SEP_OUT = "{<}"

        'AXFinanzForm.rec1.Text = REC_SEP
        'AXFinanzForm.feld1.Text = FELD_SEP
        'AXFinanzForm.kst1.Text = KSTKTR_SEP

        'AXFinanzForm.rec2.Text = REC_SEP_OUT
        'AXFinanzForm.feld2.Text = FELD_SEP_OUT

        'lblVersion.Text = "SBSxas V-" & Version
    End Sub

    Public Sub InitDB()

        Dim strIdentityName As String

        Try

            strIdentityName = System.Security.Principal.WindowsIdentity.GetCurrent().Name
            strIdentityName = Strings.Replace(strIdentityName, "\", "/")

            'Dim daDebitorenHead As New MySqlDataAdapter()
            mysqlconn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString")
            'Read cmd DebiHead
            mysqlcmdDebRead.CommandText = "SELECT * FROM tbldebitorenjhead WHERE IdentityName='" + strIdentityName + "' AND ProcessID= " + Process.GetCurrentProcess().Id.ToString

            'Del cmd DebiHead
            mysqlcmdDebDel.CommandText = "DELETE FROM tbldebitorenjhead WHERE IdentityName='" + strIdentityName + "' AND ProcessID= " + Process.GetCurrentProcess().Id.ToString


            'Debitoren Sub
            'Read
            mysqlcmdDebSubRead.CommandText = "Select * FROM tbldebitorensub WHERE IdentityName='" + strIdentityName + "' AND ProcessID= " + Process.GetCurrentProcess().Id.ToString

            'Del cmd Debi Sub
            mysqlcmdDebSubDel.CommandText = "DELETE FROM tbldebitorensub WHERE IdentityName='" + strIdentityName + "' AND ProcessID= " + Process.GetCurrentProcess().Id.ToString

        Catch ex As Exception


        End Try



    End Sub

    Public Sub butDebitoren_Click(sender As Object, e As EventArgs) Handles butDebitoren.Click


        'Dim strIncrBelNbr As String = String.Empty
        Dim intNbrDebiForms As Int16

        Try

            System.GC.Collect()

            'Zuerst prüfen ob schon eine Insstanz für Mandant existiert
            For Each frm As Form In Application.OpenForms()
                If frm.Text = "Debitor " + lstBoxMandant.Text Then
                    'Es existiert schon eine Instanz
                    intNbrDebiForms += 1
                End If
            Next
            If intNbrDebiForms > 0 Then
                MessageBox.Show("Es ist schone ein Formular für " + "Debitor " + lstBoxMandant.Text + " geöffnet.", "Pro Mandant nur 1 Formular möglich", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
                intNbrDebiForms = 0
                Dim frmDebDisp As New frmDebDisp
                frmDebDisp.Cursor = Cursors.WaitCursor
                frmDebDisp.Text = "Debitor " + lstBoxMandant.Text
                frmDebDisp.MdiParent = Me
                frmDebDisp.Show()
                frmDebDisp.Top = 100
                'Für korrekte Plazierung
                For Each frm As Form In Application.OpenForms()
                    If frm.Name = "frmDebDisp" Then
                        'Es existiert schon eine Instanz
                        intNbrDebiForms += 1
                    End If
                Next
                'Je nach Anzahl von Links her verschieben
                frmDebDisp.Left = intNbrDebiForms * 15
                frmDebDisp.FcDebiDisplay(lstBoxMandant.SelectedValue,
                                     lstBoxPerioden)

            End If

            frmDebDisp.Cursor = Cursors.Default
            'Application.DoEvents()


            'Form unloaden
            'frmDebDisp.Hide()
            'frmDebDisp = Nothing


            'Me.Cursor = Cursors.WaitCursor

            'intMode = 0

            ''DGV Debitoren
            'dgvBookings.DataSource = Nothing
            'dgvBookingSub.DataSource = Nothing

            'dsDebitoren.Reset()
            ''dsDebitoren.Clear()

            ''Zuerst evtl. vorhandene DS löschen in Tabelle
            'MySQLdaDebitoren.DeleteCommand.Connection.Open()
            'MySQLdaDebitoren.DeleteCommand.ExecuteNonQuery()
            'MySQLdaDebitoren.DeleteCommand.Connection.Close()

            'MySQLdaDebitorenSub.DeleteCommand.Connection.Open()
            'MySQLdaDebitorenSub.DeleteCommand.ExecuteNonQuery()
            'MySQLdaDebitorenSub.DeleteCommand.Connection.Close()

            ''Info neu erstellen
            'dsDebitoren.Tables.Add("tblDebitorenInfo")
            'Dim col1 As DataColumn = New DataColumn("strInfoT")
            'col1.DataType = System.Type.GetType("System.String")
            'col1.MaxLength = 50
            'col1.Caption = "Info-Titel"
            'dsDebitoren.Tables("tblDebitorenInfo").Columns.Add(col1)
            'Dim col2 As DataColumn = New DataColumn("strInfoV")
            'col2.DataType = System.Type.GetType("System.String")
            'col2.MaxLength = 50
            'col2.Caption = "Info-Wert"
            'dsDebitoren.Tables("tblDebitorenInfo").Columns.Add(col2)

            'dgvInfo.DataSource = dsDebitoren.Tables("tblDebitorenInfo")

            ''dsDebitoren.Tables("tblDebitorenInfo").Constraints.Clear()
            ''dsDebitoren.Tables("tblDebitorenInfo").Rows.Clear()

            ''dgvInfo.Refresh()

            'Call InitVar()

            'Call Main.FcLoginSage2(objdbConn,
            '                      objdbMSSQLConn,
            '                      objdbSQLcommand,
            '                      Finanz,
            '                      FBhg,
            '                      DbBhg,
            '                      PIFin,
            '                      BeBu,
            '                      KrBhg,
            '                      cmbBuha.SelectedValue,
            '                      dsDebitoren.Tables("tblDebitorenInfo"),
            '                      cmbPerioden.SelectedItem,
            '                      strYear,
            '                      intTeqNbr,
            '                      intTeqNbrLY,
            '                      intTeqNbrPLY,
            '                      datPeriodFrom,
            '                      datPeriodTo,
            '                      strPeriodStatus)

            ''Transitorische Buchungen?
            ''Call Main.fcCheckTransitorischeDebit(cmbBuha.SelectedValue,
            ''                                     objdbConn,
            ''                                     objdbAccessConn)

            ''Gibt es eine Query auszuführen bevor dem Buchen?
            ''Call MainDebitor.FcExecuteBeforeDebit(cmbBuha.SelectedValue, objdbConn)


            ''In DS löschen
            ''If dsDebitoren.Tables("tblDebiHeadsFromUser") IsNot Nothing Then
            ''    dsDebitoren.Tables("tblDebiHeadsFromUser").Constraints.Clear()
            ''    dsDebitoren.Tables("tblDebiHeadsFromUser").Rows.Clear()
            ''    dsDebitoren.Tables("tblDebiHeadsFromUser").Columns.Clear()
            ''End If
            ''MySQLdaDebitoren.Fill(dsDebitoren, "tblDebiHeadsFromUser")

            ''In DS löschen
            ''If dsDebitoren.Tables("tblDebiSubsFromUser") IsNot Nothing Then
            ''    dsDebitoren.Tables("tblDebiSubsFromUser").Constraints.Clear()
            ''    dsDebitoren.Tables("tblDebiSubsFromUser").Rows.Clear()
            ''    dsDebitoren.Tables("tblDebiSubsFromUser").Columns.Clear()
            ''End If
            ''MySQLdaDebitorenSub.Fill(dsDebitoren, "tblDebiSubsFromUser")


            'Dim clImp As New ClassImport
            'clImp.FcDebitFill(cmbBuha.SelectedValue)
            'clImp = Nothing

            ''Call MainDebitor.FcFillDebit(cmbBuha.SelectedValue,
            ''                             objdbAccessConn,
            ''                             objdbConn)

            ''Grid neu aufbauen
            'MySQLdaDebitoren.Fill(dsDebitoren, "tblDebiHeadsFromUser")
            'MySQLdaDebitorenSub.Fill(dsDebitoren, "tblDebiSubsFromUser")

            'dgvBookings.DataSource = dsDebitoren.Tables("tblDebiHeadsFromUser")
            'dgvBookingSub.DataSource = dsDebitoren.Tables("tblDebiSubsFromUser")

            'Call InitdgvInfo()
            'Call InitdgvDebitoren()
            'Call InitdgvDebitorenSub()

            'Me.Refresh()

            'Dim clCheck As New ClassCheck
            'clCheck.FcClCheckDebit(cmbBuha.SelectedValue,
            '                       dsDebitoren,
            '                       Finanz,
            '                       FBhg,
            '                       DbBhg,
            '                       PIFin,
            '                       BeBu,
            '                       dsDebitoren.Tables("tblDebitorenInfo"),
            '                       cmbBuha.Text,
            '                       intTeqNbr,
            '                       intTeqNbrLY,
            '                       intTeqNbrPLY,
            '                       strYear,
            '                       cmbPerioden.SelectedItem,
            '                       datPeriodFrom,
            '                       datPeriodTo,
            '                       strPeriodStatus,
            '                       chkValutaCorrect.Checked,
            '                       dtpValutaCorrect.Value)
            'clCheck = Nothing

            ''Call Main.FcCheckDebit(cmbBuha.SelectedValue,
            ''                       dsDebitoren,
            ''                       Finanz,
            ''                       FBhg,
            ''                       DbBhg,
            ''                       PIFin,
            ''                       BeBu,
            ''                       dsDebitoren.Tables("tblDebitorenInfo"),
            ''                       cmbBuha.Text,
            ''                       intTeqNbr,
            ''                       intTeqNbrLY,
            ''                       intTeqNbrPLY,
            ''                       strYear,
            ''                       cmbPerioden.SelectedItem,
            ''                       datPeriodFrom,
            ''                       datPeriodTo,
            ''                       strPeriodStatus,
            ''                       chkValutaCorrect.Checked,
            ''                       dtpValutaCorrect.Value)



            'Application.DoEvents()


            ''Anzahl schreiben
            'txtNumber.Text = dsDebitoren.Tables("tblDebiHeadsFromUser").Rows.Count.ToString
            'Me.Cursor = Cursors.Default

            '''Ipmort Kredit hiden
            'Me.butImportK.Enabled = False
            'Me.butImport.Enabled = True

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem Debitorenauflistung", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        Finally

            Me.Cursor = Cursors.Default

        End Try


    End Sub

    Public Sub InitdgvInfo()

        ''DGV - Info
        ''dgvInfo.DataSource = objdtInfo
        'dgvInfo.AllowUserToAddRows = False
        'dgvInfo.AllowUserToDeleteRows = False
        ''dgvInfo.Enabled = False
        'dgvInfo.RowHeadersVisible = False
        'dgvInfo.Columns("strInfoT").HeaderText = "Info"
        'dgvInfo.Columns("strInfoT").Width = 100
        'dgvInfo.Columns("strInfoV").HeaderText = "Wert"
        'dgvInfo.Columns("strInfoV").Width = 250

    End Sub


    Private Sub InitdgvDebitoren()

        Try

            'dgvBookingsO.ShowCellToolTips = False
            'dgvBookingsO.AllowUserToAddRows = False
            'dgvBookingsO.AllowUserToDeleteRows = False
            'Dim okcol As New DataGridViewCheckBoxColumn
            'okcol.DataPropertyName = "booDebBook"
            'okcol.HeaderText = "ok"
            'okcol.DisplayIndex = 0
            'okcol.Width = 40
            'dgvBookingsO.Columns.Add(okcol)
            'dgvBookingsO.Columns("booDebBook").Visible = False
            ''dgvBookings.Columns("booDebBook").DisplayIndex = 0
            ''dgvBookings.Columns("booDebBook").HeaderText = "ok"
            ''dgvBookings.Columns("booDebBook").Width = 40
            ''dgvBookings.Columns("booDebBook").ValueType = System.Type.[GetType]("System.Boolean")
            ''dgvBookings.Columns("booDebBook").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            ''dgvBookings.Columns("booDebBook").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            'dgvBookingsO.Columns("strDebRGNbr").DisplayIndex = 1
            'dgvBookingsO.Columns("strDebRGNbr").HeaderText = "RG-Nr"
            'dgvBookingsO.Columns("strDebRGNbr").Width = 60
            'dgvBookingsO.Columns("strDebRGNbr").ReadOnly = True
            'dgvBookingsO.Columns("lngDebNbr").DisplayIndex = 2
            'dgvBookingsO.Columns("lngDebNbr").HeaderText = "Debitor"
            'dgvBookingsO.Columns("lngDebNbr").Width = 60
            'dgvBookingsO.Columns("strDebBez").DisplayIndex = 3
            'dgvBookingsO.Columns("strDebBez").HeaderText = "Bezeichnung"
            'dgvBookingsO.Columns("strDebBez").Width = 140
            'dgvBookingsO.Columns("lngDebKtoNbr").DisplayIndex = 4
            'dgvBookingsO.Columns("lngDebKtoNbr").HeaderText = "Konto"
            'dgvBookingsO.Columns("lngDebKtoNbr").Width = 50
            'dgvBookingsO.Columns("strDebKtoBez").DisplayIndex = 5
            'dgvBookingsO.Columns("strDebKtoBez").HeaderText = "Bezeichnung"
            'dgvBookingsO.Columns("strDebKtoBez").Width = 150
            'dgvBookingsO.Columns("strDebCur").DisplayIndex = 6
            'dgvBookingsO.Columns("strDebCur").HeaderText = "Währung"
            'dgvBookingsO.Columns("strDebCur").Width = 60
            'dgvBookingsO.Columns("dblDebNetto").DisplayIndex = 7
            'dgvBookingsO.Columns("dblDebNetto").HeaderText = "Netto"
            'dgvBookingsO.Columns("dblDebNetto").Width = 80
            'dgvBookingsO.Columns("dblDebNetto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            'dgvBookingsO.Columns("dblDebNetto").DefaultCellStyle.Format = "N4"
            'dgvBookingsO.Columns("dblDebNetto").ReadOnly = True
            'dgvBookingsO.Columns("dblDebMwSt").DisplayIndex = 8
            'dgvBookingsO.Columns("dblDebMwSt").HeaderText = "MwSt"
            'dgvBookingsO.Columns("dblDebMwSt").Width = 70
            'dgvBookingsO.Columns("dblDebMwSt").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            'dgvBookingsO.Columns("dblDebMwSt").DefaultCellStyle.Format = "N4"
            'dgvBookingsO.Columns("dblDebBrutto").DisplayIndex = 9
            'dgvBookingsO.Columns("dblDebBrutto").HeaderText = "Brutto"
            'dgvBookingsO.Columns("dblDebBrutto").Width = 80
            'dgvBookingsO.Columns("dblDebBrutto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            'dgvBookingsO.Columns("dblDebBrutto").DefaultCellStyle.Format = "N4"
            'dgvBookingsO.Columns("intSubBookings").DisplayIndex = 10
            'dgvBookingsO.Columns("intSubBookings").HeaderText = "Sub"
            'dgvBookingsO.Columns("intSubBookings").Width = 50
            'dgvBookingsO.Columns("intSubBookings").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            'dgvBookingsO.Columns("dblSumSubBookings").DisplayIndex = 11
            'dgvBookingsO.Columns("dblSumSubBookings").HeaderText = "Sub-Summe"
            'dgvBookingsO.Columns("dblSumSubBookings").Width = 80
            'dgvBookingsO.Columns("dblSumSubBookings").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            'dgvBookingsO.Columns("dblSumSubBookings").DefaultCellStyle.Format = "N4"
            'dgvBookingsO.Columns("lngDebIdentNbr").DisplayIndex = 12
            'dgvBookingsO.Columns("lngDebIdentNbr").HeaderText = "Ident"
            'dgvBookingsO.Columns("lngDebIdentNbr").Width = 80
            'dgvBookingsO.Columns("intBuchungsart").DisplayIndex = 13
            'dgvBookingsO.Columns("intBuchungsart").DisplayIndex = 13
            'dgvBookingsO.Columns("intBuchungsart").HeaderText = "BA"
            'dgvBookingsO.Columns("intBuchungsart").Width = 40
            'dgvBookingsO.Columns("strOPNr").DisplayIndex = 14
            'dgvBookingsO.Columns("strOPNr").HeaderText = "OP-Nr"
            'dgvBookingsO.Columns("strOPNr").Width = 80
            'dgvBookingsO.Columns("datDebRGDatum").DisplayIndex = 15
            'dgvBookingsO.Columns("datDebRGDatum").HeaderText = "RG Datum"
            'dgvBookingsO.Columns("datDebRGDatum").Width = 70
            'dgvBookingsO.Columns("datDebValDatum").DisplayIndex = 16
            'dgvBookingsO.Columns("datDebValDatum").HeaderText = "Val Datum"
            'dgvBookingsO.Columns("datDebValDatum").Width = 70
            'dgvBookingsO.Columns("strDebiBank").DisplayIndex = 17
            'dgvBookingsO.Columns("strDebiBank").HeaderText = "Bank"
            'dgvBookingsO.Columns("strDebiBank").Width = 60
            'dgvBookingsO.Columns("strDebStatusText").DisplayIndex = 18
            'dgvBookingsO.Columns("strDebStatusText").HeaderText = "Status"
            'dgvBookingsO.Columns("strDebStatusText").Width = 200
            'dgvBookingsO.Columns("intBuchhaltung").Visible = False
            ''dgvBookings.Columns("intBuchungsart").Visible = False
            ''dgvBookings.Columns("intRGArt").Visible = False
            ''dgvBookings.Columns("strRGArt").Visible = False
            ''dgvBookings.Columns("lngLinkedRG").Visible = False
            ''dgvBookings.Columns("booLinked").Visible = False
            ''dgvBookings.Columns("strRGName").Visible = False
            ''dgvBookings.Columns("strDebIdentnbr2").Visible = False
            ''dgvBookings.Columns("strDebText").Visible = False
            ''dgvBookings.Columns("strRGBemerkung").Visible = False
            ''dgvBookings.Columns("strDebRef").Visible = False
            ''dgvBookings.Columns("strZahlBed").Visible = False
            ''dgvBookings.Columns("strDebStatusBitLog").Visible = False
            ''dgvBookings.Columns("strDebBookStatus").Visible = False
            ''dgvBookings.Columns("booBooked").Visible = False
            ''dgvBookings.Columns("datBooked").Visible = False
            ''dgvBookings.Columns("lngBelegNr").Visible = False

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error " + Err.Number.ToString)

        End Try



    End Sub

    Private Sub InitdgvDebitorenSub()

        Try

            'dgvBookingSubO.RowHeadersWidth = 24

            'dgvBookingSubO.ShowCellToolTips = False
            'dgvBookingSubO.AllowUserToAddRows = False
            'dgvBookingSubO.AllowUserToDeleteRows = False
            'dgvBookingSubO.Columns("strRGNr").DisplayIndex = 0
            'dgvBookingSubO.Columns("strRGNr").Width = 50
            'dgvBookingSubO.Columns("strRGNr").HeaderText = "RG-Nr"
            'dgvBookingSubO.Columns("intSollHaben").DisplayIndex = 1
            'dgvBookingSubO.Columns("intSollHaben").Width = 20
            'dgvBookingSubO.Columns("intSollHaben").HeaderText = "S/H"
            'dgvBookingSubO.Columns("lngKto").DisplayIndex = 2
            'dgvBookingSubO.Columns("lngKto").Width = 45
            'dgvBookingSubO.Columns("lngKto").HeaderText = "Konto"
            'dgvBookingSubO.Columns("strKtoBez").DisplayIndex = 3
            'dgvBookingSubO.Columns("strKtoBez").HeaderText = "Bezeichnung"
            'dgvBookingSubO.Columns("lngKST").DisplayIndex = 4
            'dgvBookingSubO.Columns("lngKST").Width = 30
            'dgvBookingSubO.Columns("lngKST").HeaderText = "KST"
            'dgvBookingSubO.Columns("strKSTBez").DisplayIndex = 5
            'dgvBookingSubO.Columns("strKSTBez").Width = 60
            'dgvBookingSubO.Columns("strKSTBez").HeaderText = "Bezeichnung"
            'dgvBookingSubO.Columns("lngProj").DisplayIndex = 6
            'dgvBookingSubO.Columns("lngProj").Width = 30
            'dgvBookingSubO.Columns("lngProj").HeaderText = "Proj"
            'dgvBookingSubO.Columns("strProjBez").DisplayIndex = 7
            'dgvBookingSubO.Columns("strProjBez").HeaderText = "Pr.-Bez."
            'dgvBookingSubO.Columns("strProjBez").Width = 55
            'dgvBookingSubO.Columns("dblNetto").DisplayIndex = 8
            'dgvBookingSubO.Columns("dblNetto").Width = 65
            'dgvBookingSubO.Columns("dblNetto").HeaderText = "Netto"
            'dgvBookingSubO.Columns("dblNetto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            'dgvBookingSubO.Columns("dblNetto").DefaultCellStyle.Format = "N4"
            'dgvBookingSubO.Columns("dblMwSt").DisplayIndex = 9
            'dgvBookingSubO.Columns("dblMwSt").Width = 60
            'dgvBookingSubO.Columns("dblMwSt").HeaderText = "MwSt"
            'dgvBookingSubO.Columns("dblMwSt").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            'dgvBookingSubO.Columns("dblMwSt").DefaultCellStyle.Format = "N4"
            'dgvBookingSubO.Columns("dblBrutto").DisplayIndex = 10
            'dgvBookingSubO.Columns("dblBrutto").Width = 65
            'dgvBookingSubO.Columns("dblBrutto").HeaderText = "Brutto"
            'dgvBookingSubO.Columns("dblBrutto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            'dgvBookingSubO.Columns("dblBrutto").DefaultCellStyle.Format = "N4"
            'dgvBookingSubO.Columns("dblMwStSatz").DisplayIndex = 11
            'dgvBookingSubO.Columns("dblMwStSatz").Width = 30
            'dgvBookingSubO.Columns("dblMwStSatz").HeaderText = "MwStS"
            'dgvBookingSubO.Columns("strMwStKey").DisplayIndex = 12
            'dgvBookingSubO.Columns("strMwStKey").Width = 30
            'dgvBookingSubO.Columns("strMwStKey").HeaderText = "MwStK"
            'dgvBookingSubO.Columns("strStatusUBText").DisplayIndex = 13
            'dgvBookingSubO.Columns("strStatusUBText").HeaderText = "Status"
            'dgvBookingSubO.Columns("strStatusUBText").Width = 135

            ''dgvBookingSub.Columns("lngID").Visible = False
            ''dgvBookingSub.Columns("strArtikel").Visible = False
            ''dgvBookingSub.Columns("strStatusUBBitLog").Visible = False
            ''dgvBookingSub.Columns("strDebSubText").Visible = False
            ''dgvBookingSub.Columns("strDebBookStatus").Visible = False

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error " + Err.Number.ToString)

        End Try


    End Sub

    Private Sub InitdgvKreditoren()

        Try

            'dgvBookingsO.ShowCellToolTips = False
            'dgvBookingsO.AllowUserToAddRows = False
            'dgvBookingsO.AllowUserToDeleteRows = False
            'dgvBookingsO.Columns("booKredBook").DisplayIndex = 0
            'dgvBookingsO.Columns("booKredBook").HeaderText = "ok"
            'dgvBookingsO.Columns("booKredBook").Width = 40
            'dgvBookingsO.Columns("booKredBook").ValueType = System.Type.[GetType]("System.Boolean")
            'dgvBookingsO.Columns("booKredBook").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            'dgvBookingsO.Columns("booKredBook").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            ''dgvBookings.Columns("strKredRGNbr").DisplayIndex = 1
            ''dgvBookings.Columns("strKredRGNbr").HeaderText = "KRG-Nr"
            ''dgvBookings.Columns("strKredRGNbr").Width = 60
            ''dgvBookings.Columns("strKredRGNbr").ReadOnly = True
            'dgvBookingsO.Columns("lngKredID").DisplayIndex = 1
            'dgvBookingsO.Columns("lngKredID").HeaderText = "Kred-ID"
            'dgvBookingsO.Columns("lngKredID").Width = 60
            'dgvBookingsO.Columns("lngKredID").ReadOnly = True
            'dgvBookingsO.Columns("lngKredNbr").DisplayIndex = 2
            'dgvBookingsO.Columns("lngKredNbr").HeaderText = "Kreditor"
            'dgvBookingsO.Columns("lngKredNbr").Width = 60
            'dgvBookingsO.Columns("strKredBez").DisplayIndex = 3
            'dgvBookingsO.Columns("strKredBez").HeaderText = "Bezeichnung"
            'dgvBookingsO.Columns("strKredBez").Width = 140
            'dgvBookingsO.Columns("lngKredKtoNbr").DisplayIndex = 4
            'dgvBookingsO.Columns("lngKredKtoNbr").HeaderText = "Konto"
            'dgvBookingsO.Columns("lngKredKtoNbr").Width = 50
            'dgvBookingsO.Columns("strKredKtoBez").DisplayIndex = 5
            'dgvBookingsO.Columns("strKredKtoBez").HeaderText = "Bezeichnung"
            'dgvBookingsO.Columns("strKredKtoBez").Width = 150
            'dgvBookingsO.Columns("strKredCur").DisplayIndex = 6
            'dgvBookingsO.Columns("strKredCur").HeaderText = "Währung"
            'dgvBookingsO.Columns("strKredCur").Width = 60
            'dgvBookingsO.Columns("dblKredNetto").DisplayIndex = 7
            'dgvBookingsO.Columns("dblKredNetto").HeaderText = "Netto"
            'dgvBookingsO.Columns("dblKredNetto").Width = 80
            'dgvBookingsO.Columns("dblKredNetto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            ''dgvBookings.Columns("dblKredNetto").DefaultCellStyle.Format = "N2"
            'dgvBookingsO.Columns("dblKredNetto").ReadOnly = True
            'dgvBookingsO.Columns("dblKredMwSt").DisplayIndex = 8
            'dgvBookingsO.Columns("dblKredMwSt").HeaderText = "MwSt"
            'dgvBookingsO.Columns("dblKredMwSt").Width = 70
            'dgvBookingsO.Columns("dblKredMwSt").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            ''dgvBookings.Columns("dblKredMwSt").DefaultCellStyle.Format = "N2"
            'dgvBookingsO.Columns("dblKredBrutto").DisplayIndex = 9
            'dgvBookingsO.Columns("dblKredBrutto").HeaderText = "Brutto"
            'dgvBookingsO.Columns("dblKredBrutto").Width = 80
            'dgvBookingsO.Columns("dblKredBrutto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            ''dgvBookings.Columns("dblKredBrutto").DefaultCellStyle.Format = "N2"
            'dgvBookingsO.Columns("intSubBookings").DisplayIndex = 10
            'dgvBookingsO.Columns("intSubBookings").HeaderText = "Sub"
            'dgvBookingsO.Columns("intSubBookings").Width = 50
            'dgvBookingsO.Columns("intSubBookings").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            'dgvBookingsO.Columns("dblSumSubBookings").DisplayIndex = 11
            'dgvBookingsO.Columns("dblSumSubBookings").DefaultCellStyle.Format = "N2"
            'dgvBookingsO.Columns("dblSumSubBookings").HeaderText = "Sub-Summe"
            'dgvBookingsO.Columns("dblSumSubBookings").Width = 80
            'dgvBookingsO.Columns("dblSumSubBookings").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            'dgvBookingsO.Columns("lngKredIdentNbr").DisplayIndex = 12
            'dgvBookingsO.Columns("lngKredIdentNbr").HeaderText = "Ident"
            'dgvBookingsO.Columns("lngKredIdentNbr").Width = 80
            ''Dim cmbBuchungsart As New DataGridViewComboBoxColumn()
            ''Dim objdtBA As New DataTable("objidtBA")
            ''Dim objlocMySQLcmd As New MySqlCommand
            '''objdbConn.Open()
            ''objlocMySQLcmd.CommandText = "Select * FROM t_sage_tblbuchungsarten"
            ''objlocMySQLcmd.Connection = objdbConn
            ''objdtBA.Load(objlocMySQLcmd.ExecuteReader)
            ''cmbBuchungsart.DataSource = objdtBA
            '''objdbConn.Close()
            ''cmbBuchungsart.DisplayMember = "strBuchungsart"
            ''cmbBuchungsart.ValueMember = "idBuchungsart"
            ''cmbBuchungsart.HeaderText = "BA"
            ''cmbBuchungsart.Name = "intBuchungsart"
            ''cmbBuchungsart.DataPropertyName = "intBuchungsart"
            ''cmbBuchungsart.DisplayIndex = 13
            ''cmbBuchungsart.Width = 70
            ''dgvBookings.Columns.Add(cmbBuchungsart)
            'dgvBookingsO.Columns("intBuchungsart").DisplayIndex = 13
            'dgvBookingsO.Columns("intBuchungsart").DisplayIndex = 13
            'dgvBookingsO.Columns("intBuchungsart").HeaderText = "BA"
            'dgvBookingsO.Columns("intBuchungsart").Width = 40
            'dgvBookingsO.Columns("strKredRGNbr").DisplayIndex = 14
            'dgvBookingsO.Columns("strKredRGNbr").HeaderText = "RG-Nr"
            'dgvBookingsO.Columns("strKredRGNbr").Width = 100
            'dgvBookingsO.Columns("datKredRGDatum").DisplayIndex = 15
            'dgvBookingsO.Columns("datKredRGDatum").HeaderText = "RG Datum"
            'dgvBookingsO.Columns("datKredRGDatum").Width = 70
            'dgvBookingsO.Columns("datKredValDatum").DisplayIndex = 16
            'dgvBookingsO.Columns("datKredValDatum").HeaderText = "Val Datum"
            'dgvBookingsO.Columns("datKredValDatum").Width = 70
            'dgvBookingsO.Columns("strKrediBank").DisplayIndex = 17
            'dgvBookingsO.Columns("strKrediBank").HeaderText = "KBank"
            'dgvBookingsO.Columns("strKrediBank").Width = 60
            'dgvBookingsO.Columns("strKredStatusText").DisplayIndex = 18
            'dgvBookingsO.Columns("strKredStatusText").HeaderText = "Status"
            'dgvBookingsO.Columns("strKredStatusText").Width = 200
            ''dgvBookings.Columns("intBuchhaltung").Visible = False
            ''dgvBookings.Columns("intBuchungsart").Visible = False
            ''dgvBookings.Columns("intRGArt").Visible = False
            ''dgvBookings.Columns("strRGArt").Visible = False
            ''dgvBookings.Columns("lngLinkedRG").Visible = False
            ''dgvBookings.Columns("booLinked").Visible = False
            ''dgvBookings.Columns("strRGName").Visible = False
            ''dgvBookings.Columns("strDebIdentnbr2").Visible = False
            ''dgvBookings.Columns("strDebText").Visible = False
            ''dgvBookings.Columns("strRGBemerkung").Visible = False
            ''dgvBookings.Columns("strDebRef").Visible = False
            ''dgvBookings.Columns("strZahlBed").Visible = False
            ''dgvBookings.Columns("strDebStatusBitLog").Visible = False
            ''dgvBookings.Columns("strDebBookStatus").Visible = False
            ''dgvBookings.Columns("booBooked").Visible = False
            ''dgvBookings.Columns("datBooked").Visible = False
            ''dgvBookings.Columns("lngBelegNr").Visible = False

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error " + Err.Number.ToString)

        End Try



    End Sub


    Private Sub InitdgvKreditorenSub()

        Try

            'dgvBookingSub.ShowCellToolTips = False
            'dgvBookingSub.AllowUserToAddRows = False
            'dgvBookingSub.AllowUserToDeleteRows = False
            ''dgvBookingSub.Columns("strRGNr").DisplayIndex = 0
            ''dgvBookingSub.Columns("strRGNr").Width = 60
            ''dgvBookingSub.Columns("strRGNr").HeaderText = "RG-Nr"
            'dgvBookingSub.Columns("lngKredID").DisplayIndex = 0
            'dgvBookingSub.Columns("lngKredID").Width = 60
            'dgvBookingSub.Columns("lngKredID").HeaderText = "Kred-ID"
            'dgvBookingSub.Columns("intSollHaben").Width = 30
            'dgvBookingSub.Columns("intSollHaben").HeaderText = "S/H"
            'dgvBookingSub.Columns("lngKto").Width = 50
            'dgvBookingSub.Columns("lngKto").HeaderText = "Konto"
            'dgvBookingSub.Columns("strKtoBez").HeaderText = "Bezeichnung"
            'dgvBookingSub.Columns("lngKST").Width = 50
            'dgvBookingSub.Columns("lngKST").HeaderText = "KST"
            'dgvBookingSub.Columns("strKSTBez").Width = 60
            'dgvBookingSub.Columns("strKSTBez").HeaderText = "Bezeichnung"
            'dgvBookingSub.Columns("dblNetto").Width = 60
            'dgvBookingSub.Columns("dblNetto").HeaderText = "Netto"
            'dgvBookingSub.Columns("dblNetto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            ''dgvBookingSub.Columns("dblNetto").DefaultCellStyle.Format = "N2"
            'dgvBookingSub.Columns("dblMwSt").Width = 50
            'dgvBookingSub.Columns("dblMwSt").HeaderText = "MwSt"
            'dgvBookingSub.Columns("dblMwSt").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            ''dgvBookingSub.Columns("dblMwSt").DefaultCellStyle.Format = "N2"
            'dgvBookingSub.Columns("dblBrutto").Width = 60
            'dgvBookingSub.Columns("dblBrutto").HeaderText = "Brutto"
            'dgvBookingSub.Columns("dblBrutto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            ''dgvBookingSub.Columns("dblBrutto").DefaultCellStyle.Format = "N2"
            'dgvBookingSub.Columns("dblMwStSatz").Width = 40
            'dgvBookingSub.Columns("dblMwStSatz").HeaderText = "MwStS"
            'dgvBookingSub.Columns("dblMwStSatz").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            'dgvBookingSub.Columns("dblMwStSatz").DefaultCellStyle.Format = "N1"
            'dgvBookingSub.Columns("strMwStKey").Width = 40
            'dgvBookingSub.Columns("strMwStKey").HeaderText = "MwStK"
            'dgvBookingSub.Columns("strStatusUBText").HeaderText = "Status"

            'dgvBookingSub.Columns("lngID").Visible = False
            ''dgvBookingSub.Columns("strArtikel").Visible = False
            ''dgvBookingSub.Columns("strStatusUBBitLog").Visible = False
            ''dgvBookingSub.Columns("strDebSubText").Visible = False
            ''dgvBookingSub.Columns("strDebBookStatus").Visible = False

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error " + Err.Number.ToString)

        End Try


    End Sub


    Private Sub frmImportMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim strDKDef As String
        Dim objDABuchhaltungen As New MySqlDataAdapter("SELECT * FROM t_sage_buchhaltungen WHERE NOT Buchh200_Name IS NULL AND NOT Buchh_TableDeb IS NULL ORDER BY Buchh_Bez", objdbConn)
        Dim objdtBuchhaltungen As New DataTable("tbliBuchhaltungen")
        'Dim intFunctionReturns As Int16
        Dim objOracleConn As New OracleConnection("Data Source=(DESCRIPTION=" _
                        + "(ADDRESS=(PROTOCOL=TCP)(HOST=10.0.0.29)(PORT=1521))" _
                        + "(CONNECT_DATA=(SERVICE_NAME=CISNEW)));" _
                        + "User Id=cis;Password=sugus;")
        Dim objOracleCmd As New OracleCommand()


        Try

            'MySQL - Connection öffnen
            'objdbConn.Open()
            'Oracle - Connection öffnen
            'objOracleConn.ConnectionString = strOraDB
            'objOracleConn.Open()
            objOracleCmd.Connection = objOracleConn

            'Comboxen
            objdtBuchhaltungen.Clear()
            objDABuchhaltungen.Fill(objdtBuchhaltungen)

            For Each drBuha In objdtBuchhaltungen.Rows
                strDKDef = Main.FcGetDKDef(drBuha("Buchh_Nr"))
                drBuha("Buchh_Bez") += strDKDef
            Next
            'cmbMarken.Sorted = True
            'cmbBuha.DataSource = objdtBuchhaltungen
            'cmbBuha.DisplayMember = "Buchh_Bez"
            'cmbBuha.ValueMember = "Buchh_Nr"

            'cmbBuha.DropDownWidth = 330

            lstBoxMandant.DataSource = objdtBuchhaltungen
            lstBoxMandant.DisplayMember = "Buchh_Bez"
            lstBoxMandant.ValueMember = "Buchh_Nr"

            'Tabelle Debi/ Kredi Head erstellen
            'objdtDebitorenHead = Main.tblDebitorenHead()
            'objdtKreditorenHead = Main.tblKreditorenHead()

            'Tabelle Debi/ Kredi Sub erstellen
            'objdtDebitorenSub = Main.tblDebitorenSub()
            'objdtKreditorenSub = Main.tblKreditorenSub()

            'Info - Tabelle erstellen
            'objdtInfo = Main.tblInfo()
            'dsDebitoren.Tables.Add("tblDebitorenInfo")
            'Dim col1 As DataColumn = New DataColumn("strInfoT")
            'col1.DataType = System.Type.GetType("System.String")
            'col1.MaxLength = 50
            'col1.Caption = "Info-Titel"
            'dsDebitoren.Tables("tblDebitorenInfo").Columns.Add(col1)
            'Dim col2 As DataColumn = New DataColumn("strInfoV")
            'col2.DataType = System.Type.GetType("System.String")
            'col2.MaxLength = 50
            'col2.Caption = "Info-Wert"
            'dsDebitoren.Tables("tblDebitorenInfo").Columns.Add(col2)

            'dgvInfo.DataSource = dsDebitoren.Tables("tblDebitorenInfo")

            'Me.mysqlcongen.ConnectionString = "server=ZHDB02.sdlc.mssag.ch;uid=workbench;database=AHZ;Pwd=sesam;Convert Zero Datetime=True"
            'Me.mysqlcongen.Open()
            'Me.mysqlcongen.Close()

            'Main.FcInitAccessConnecation(objdbAccessConn, "NetAuction_data_Server.mdb")
            'objdbAccessConn.Open()
            'objdbAccessConn.Close()
            'Main.FcInitAccessConnecation(objdbAccessConn, "Fleetservice_data_Server.mdb")
            'objdbAccessConn.Open()
            'objdbAccessConn.Close()

            'dgvBookings.DataSource = dsDebitoren.Tables("tblDebiHeadsFromUser")


            'Subbuchungen ausblenden, kann für Testzwecke aktiviert werden
            'dgvBookingSub.DataSource = objdtDebitorenSub
            'dgvDebitorenSub.Visible = False

            'DGV - Info
            'dgvInfo.DataSource = objdtInfo
            'dgvInfo.AllowUserToAddRows = False
            'dgvInfo.AllowUserToDeleteRows = False
            'dgvInfo.Enabled = False
            'dgvInfo.RowHeadersVisible = False
            'dgvInfo.Columns("strInfoT").HeaderText = "Info"
            'dgvInfo.Columns("strInfoT").Width = 100
            'dgvInfo.Columns("strInfoV").HeaderText = "Wert"
            'dgvInfo.Columns("strInfoV").Width = 250

            'Version
            'Debug.Print("1 " + System.Reflection.Assembly.GetExecutingAssembly().GetName.Version.ToString)
            LblVersion.Text = "V " + System.Reflection.Assembly.GetExecutingAssembly().GetName.Version.ToString
            'Debug.Print("2 " + My.Application.Info.Version.ToString)
            'Debug.Print("3 " + Application.ProductVersion.ToString)

            'Call InitdgvDebitoren()

            ''DGV Debitoren
            'dgvBookings.DataSource = dsDebitoren.Tables("tblDebiSubsFromUser")
            'dgvBookingsSub.DataSource = dsDebitoren.Tables("tblDebiSubsFromUser")
            'objdbConn.Open()
            'Call InitdgvDebitoren()
            'Call InitdgvDebitorenSub()
            'objdbConn.Close()

            Call InitDB()


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error In Load" + Err.Number.ToString)

        End Try



    End Sub

    Private Sub butImport_Click(sender As Object, e As EventArgs)

        ''Dim objdtDebitorenHead As New DataTable("tbliDebiHead")
        ''Dim objdtDebitorenSub As New DataTable("tbliDebiSub")

        'Dim intReturnValue As Int32
        'Dim intDebBelegsNummer As Int32

        'Dim intDebitorNbr As Int32
        'Dim strBuchType As String
        'Dim strBelegDatum As String
        'Dim strValutaDatum As String
        'Dim strVerfallDatum As String
        'Dim strReferenz As String
        'Dim intKondition As Int32
        'Dim strSachBID As String = String.Empty
        'Dim strVerkID As String = String.Empty
        'Dim strMahnerlaubnis As String
        'Dim sngAktuelleMahnstufe As Single
        'Dim dblBetrag As Double
        'Dim dblKurs As Double
        'Dim strExtBelegNbr As String = String.Empty
        'Dim strSkonto As String = String.Empty
        'Dim strCurrency As String
        'Dim strDebiText As String

        'Dim intGegenKonto As Int32
        'Dim strFibuText As String
        'Dim dblNettoBetrag As Double
        'Dim dblBebuBetrag As Double
        'Dim strBeBuEintrag As String = String.Empty
        'Dim strSteuerFeld As String

        'Dim intSollKonto As Int32
        'Dim intHabenKonto As Int32
        'Dim dblSollBetrag As Double
        'Dim dblHabenBetrag As Double
        'Dim strSteuerFeldSoll As String = String.Empty
        'Dim strSteuerFeldHaben As String = String.Empty
        'Dim strBeBuEintragSoll As String = String.Empty
        'Dim strBeBuEintragHaben As String = String.Empty
        'Dim strDebiTextSoll As String = String.Empty
        'Dim strDebiTextHaben As String = String.Empty
        'Dim dblKursSoll As Double = 0
        'Dim dblKursHaben As Double = 0
        'Dim intEigeneBank As Int16

        'Dim selDebiSub() As DataRow
        'Dim strSteuerInfo() As String
        'Dim strDebitor() As String
        'Dim strDebiLine As String

        ''Sammelbeleg
        'Dim intCommonKonto As Int32
        'Dim strDebiCurrency As String
        'Dim strKrediCurrency As String
        'Dim dblBuchBetrag As Double
        'Dim dblBasisBetrag As Double
        'Dim strErfassungsDatum As String
        'Dim strRGNbr As String
        'Dim booBooingok As Boolean
        'Dim booErfOPExt As Boolean

        'Dim intLaufNbr As Int32
        'Dim strBeleg As String
        'Dim strBelegArr() As String
        'Dim dblSplitPayed As Double
        'Dim strErrMessage As String

        ''Debitoren - Head
        ''Dim daDebitorenHead As New MySqlDataAdapter()
        ''Dim daselDebitorenHead As New MySqlCommand("SELECT * FROM tbldebitorenjhead WHERE IdentityName='" + System.Security.Principal.WindowsIdentity.GetCurrent().Name + "' AND ProcessID= " + Process.GetCurrentProcess().Id.ToString, objdbConn)
        ''Dim cmbDebitorenHead As New MySqlCommandBuilder(daDebitorenHead)
        ''Dim dadelDebitorenHead As New MySqlCommand("DELETE FROM tbldebitorenjhead WHERE IdentityName='" + System.Security.Principal.WindowsIdentity.GetCurrent().Name + "' AND ProcessID= " + Process.GetCurrentProcess().Id.ToString, objdbConn)
        ''Dim dsDebitorenHead As New DataSet()
        ''daDebitorenHead.SelectCommand = daselDebitorenHead
        ''daDebitorenHead.DeleteCommand = dadelDebitorenHead

        ''daDebitorenHead.Fill(dsDebitorenHead, "tblDebiHeadsFromUser")

        ''Debitoren - Sub
        ''Dim daDebitorenSub As New MySqlDataAdapter()
        ''Dim daselDebitorenSub As New MySqlCommand("SELECT * FROM tbldebitorensub WHERE IdentityName='" + System.Security.Principal.WindowsIdentity.GetCurrent().Name + "' AND ProcessID= " + Process.GetCurrentProcess().Id.ToString, objdbConn)
        ''Dim cmdDebitorenSub As New MySqlCommandBuilder(daDebitorenSub)
        ''Dim dadelDebitorenSub As New MySqlCommand("DELETE FROM tbldebitorensub WHERE IdentityName='" + System.Security.Principal.WindowsIdentity.GetCurrent().Name + "' AND ProcessID= " + Process.GetCurrentProcess().Id.ToString, objdbConn)
        ''Dim dsDebitorenSub As New DataSet()
        ''daDebitorenSub.SelectCommand = daselDebitorenSub
        ''daDebitorenSub.DeleteCommand = dadelDebitorenSub

        ''daDebitorenSub.Fill(dsDebitorenSub, "tblDebiSubsFromUser")

        'Try


        '    Me.Cursor = Cursors.WaitCursor
        '    'Butteon desaktivieren
        '    'Me.butImport.Enabled = False

        '    'Start in Sync schreiben
        '    intReturnValue = WFDBClass.FcWriteStartToSync(objdbConn,
        '                                                  cmbBuha.SelectedValue,
        '                                                  1,
        '                                                  dsDebitoren.Tables("tblDebiHeadsFromUser").Rows.Count)

        '    'Setting soll erfasste OP als externe Beleg-Nr. genommen werden und lngDebIdentNbr als Beleg-Nr.
        '    objdbConn.Open()
        '    booErfOPExt = Convert.ToBoolean(Convert.ToInt16(Main.FcReadFromSettings(objdbConn, "Buchh_ErfOPExt", cmbBuha.SelectedValue)))
        '    objdbConn.Close()

        '    'Kopfbuchung
        '    For Each row In dsDebitoren.Tables("tblDebiHeadsFromUser").Rows

        '        If IIf(IsDBNull(row("booDebBook")), False, row("booDebBook")) Then

        '            'Test ob OP - Buchung
        '            If row("intBuchungsart") = 1 Then

        '                'Immer zuerst Belegs-Nummerierung aktivieren, falls vorhanden externe Nummer = OP - Nr. Rg
        '                'Führt zu Problemen beim Ausbuchen des DTA - Files
        '                'Resultat Besprechnung 17.09.20 mit Muhi/ Andy
        '                'DbBhg.IncrBelNbr = "J"
        '                'Belegsnummer abholen
        '                'intDebBelegsNummer = DbBhg.GetNextBelNbr("R")

        '                'Verdopplung interne BelegsNummer verhindern
        '                DbBhg.CheckDoubleIntBelNbr = "J"

        '                If row("dblDebBrutto") < 0 Then
        '                    'Gutschrift
        '                    'Falls booGSToInv (Gutschrift zu Rechnung) dann OP-Nummer vorgeben, sonst hochzählen lassen
        '                    If row("booCrToInv") Then
        '                        'Beleg-Nummerierung desaktivieren
        '                        DbBhg.IncrBelNbr = "N"
        '                        'Eingelesene OP-Nummer (=Verknüpfte OP-Nr.) = interne Beleg-Nummer
        '                        intDebBelegsNummer = Main.FcCleanRGNrStrict(row("strOPNr"))
        '                        strExtBelegNbr = row("strDebRGNbr")
        '                    Else
        '                        'Zuerst Beleg-Nummerieungung aktivieren
        '                        DbBhg.IncrBelNbr = "J"
        '                        'Belegsnummer abholen
        '                        intDebBelegsNummer = DbBhg.GetNextBelNbr("G")
        '                        'Prüfen ob wirklich frei und falls nicht hochzählen
        '                        intReturnValue = MainDebitor.FcCheckDebiExistance(objdbMSSQLConn,
        '                                                                          objdbSQLcommand,
        '                                                                          intDebBelegsNummer,
        '                                                                          "G",
        '                                                                          intTeqNbr,
        '                                                                          intTeqNbrLY,
        '                                                                          intTeqNbrPLY)


        '                        'intReturnValue = 10
        '                        'Do Until intReturnValue = 0

        '                        '    intReturnValue = DbBhg.doesBelegExist(row("lngDebNbr").ToString,
        '                        '                                      row("strDebCur"),
        '                        '                                      intDebBelegsNummer.ToString,
        '                        '                                      "NOT_SET",
        '                        '                                      "G",
        '                        '                                      "NOT_SET")
        '                        '    If intReturnValue <> 0 Then
        '                        '        intDebBelegsNummer += 1
        '                        '    End If
        '                        'Loop
        '                        strExtBelegNbr = row("strOPNr")
        '                    End If

        '                    'Beträge drehen
        '                    row("dblDebBrutto") = row("dblDebBrutto") * -1
        '                    row("dblDebMwSt") = row("dblDebMwSt") * -1
        '                    row("dblDebNetto") = row("dblDebNetto") * -1

        '                    strBuchType = "G"

        '                Else

        '                    If String.IsNullOrEmpty(row("strOPNr")) Then
        '                        'strExtBelegNbr = row("strOPNr")

        '                        'Zuerst Beleg-Nummerieungung aktivieren
        '                        DbBhg.IncrBelNbr = "J"
        '                        'Belegsnummer abholen
        '                        intDebBelegsNummer = DbBhg.GetNextBelNbr("R")
        '                        intReturnValue = MainDebitor.FcCheckDebiExistance(objdbMSSQLConn,
        '                                                                          objdbSQLcommand,
        '                                                                          intDebBelegsNummer,
        '                                                                          "R",
        '                                                                          intTeqNbr,
        '                                                                          intTeqNbrLY,
        '                                                                          intTeqNbrPLY)
        '                    Else
        '                        If Strings.Len(Main.FcCleanRGNrStrict(row("strOPNr"))) > 9 Then
        '                            'Zahl zu gross
        '                            DbBhg.IncrBelNbr = "J"
        '                            'Belegsnummer abholen
        '                            intDebBelegsNummer = DbBhg.GetNextBelNbr("R")
        '                            intReturnValue = MainDebitor.FcCheckDebiExistance(objdbMSSQLConn,
        '                                                                              objdbSQLcommand,
        '                                                                              intDebBelegsNummer,
        '                                                                              "R",
        '                                                                              intTeqNbr,
        '                                                                              intTeqNbrLY,
        '                                                                              intTeqNbrPLY)
        '                            strExtBelegNbr = row("strOPNr")
        '                        Else
        '                            'Beleg-Nummerierung abschalten
        '                            DbBhg.IncrBelNbr = "N"
        '                            'Gemäss Setting Erfasste OP-Nr. Nummern vergeben
        '                            If Not booErfOPExt Then
        '                                intDebBelegsNummer = Main.FcCleanRGNrStrict(row("strOPNr"))
        '                                strExtBelegNbr = row("strOPNr")
        '                            Else
        '                                'bei t_debi: IdentNbr wird genommen da dort die RG-Nr. drin ist. RgNr = ID
        '                                intDebBelegsNummer = row("lngDebIdentNbr")
        '                                strExtBelegNbr = row("strOPNr")
        '                            End If

        '                        End If

        '                    End If

        '                    strBuchType = "R"

        '                End If

        '                'Variablen zuweisen
        '                'Sachbearbeiter aus Debitor auslesen
        '                strDebiLine = DbBhg.ReadDebitor3(row("lngDebNbr") * -1, "")
        '                strDebitor = Split(strDebiLine, "{>}")
        '                strSachBID = strDebitor(30)
        '                'strExtBelegNbr = row("strDebRGNbr")
        '                intDebitorNbr = row("lngDebNbr")
        '                strValutaDatum = Format(row("datDebValDatum"), "yyyyMMdd").ToString
        '                strBelegDatum = Format(row("datDebRGDatum"), "yyyyMMdd").ToString
        '                If IsDBNull(row("datDebDue")) Then
        '                    strVerfallDatum = String.Empty
        '                Else
        '                    strVerfallDatum = Format(row("datDebDue"), "yyyyMMdd").ToString
        '                End If
        '                strReferenz = row("strDebReferenz")
        '                strMahnerlaubnis = String.Empty 'Format(row("datDebRGDatum"), "yyyyMMdd").ToString
        '                dblBetrag = row("dblDebBrutto")
        '                'Bei SplittBill 2ter Rechnung Text anfügen
        '                If row("booLinked") Then
        '                    strDebiText = row("strDebText") + ", FRG "
        '                Else
        '                    strDebiText = row("strDebText")
        '                End If
        '                'strDebiText = row("strDebText")
        '                strCurrency = row("strDebCur")
        '                If strCurrency <> "CHF" Then 'Muss ergänzt werden => Was ist Leitwährung auf dem Konto
        '                    dblKurs = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg)
        '                Else
        '                    dblKurs = 1.0#
        '                End If
        '                intEigeneBank = row("strDebiBank")
        '                'Zahl-Kondition
        '                intKondition = IIf(IsDBNull(row("intZKond")), 1, row("intZKond"))

        '                Try
        '                    booBooingok = True
        '                    Call DbBhg.SetBelegKopf2(intDebBelegsNummer,
        '                                             strValutaDatum,
        '                                             intDebitorNbr,
        '                                             strBuchType,
        '                                             strBelegDatum,
        '                                             strVerfallDatum,
        '                                             strDebiText,
        '                                             strReferenz,
        '                                             intKondition,
        '                                             strSachBID,
        '                                             strVerkID,
        '                                             strMahnerlaubnis,
        '                                             sngAktuelleMahnstufe,
        '                                             dblBetrag.ToString,
        '                                             dblKurs.ToString,
        '                                             strExtBelegNbr,
        '                                             strSkonto,
        '                                             strCurrency,
        '                                             "",
        '                                             intEigeneBank.ToString)

        '                    'Application.DoEvents()

        '                Catch ex As Exception
        '                    strErrMessage = "Problem " + (Err.Number And 65535).ToString + " Belegkopf zu" + intDebBelegsNummer.ToString + vbCrLf
        '                    strErrMessage += "RG " + strRGNbr + vbCrLf
        '                    strErrMessage += "Debitor " + intDebitorNbr.ToString

        '                    MessageBox.Show(ex.Message + vbCrLf + strErrMessage, "Problem " + (Err.Number And 65535).ToString, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        '                    If (Err.Number And 65535) < 10000 Then
        '                        booBooingok = False
        '                    Else
        '                        booBooingok = True
        '                    End If

        '                End Try

        '                selDebiSub = dsDebitoren.Tables("tblDebiSubsFromUser").Select("strRGNr='" + row("strDebRGNbr") + "'")
        '                strRGNbr = row("strDebRGNbr")

        '                For Each SubRow As DataRow In selDebiSub

        '                    'Bei zweiter Splitt-Bill Rechung hier eingreifen
        '                    'Gegenkonto auf 1092, MwStKey auf 'null' setzen, KST = 0
        '                    'If row("booLinked") Then
        '                    '    If row("booLinkedPayed") Then
        '                    '        intGegenKonto = 2331
        '                    '    Else
        '                    '        intGegenKonto = 1092
        '                    '    End If
        '                    '    SubRow("dblNetto") = SubRow("dblBrutto")
        '                    '    SubRow("strMwStKey") = "null"
        '                    '    SubRow("lngKST") = 0
        '                    'Else
        '                    intGegenKonto = SubRow("lngKto")
        '                    'End If
        '                    strFibuText = SubRow("strDebSubText")
        '                    If intGegenKonto <> 6906 Then
        '                        If strBuchType = "R" Then
        '                            dblNettoBetrag = SubRow("dblNetto") * -1
        '                        Else
        '                            dblNettoBetrag = SubRow("dblNetto")
        '                        End If
        '                    Else 'Rundungsdifferenzen
        '                        If strBuchType = "R" Then
        '                            dblNettoBetrag = SubRow("dblBrutto") * -1
        '                        Else
        '                            dblNettoBetrag = SubRow("dblBrutto")
        '                        End If
        '                    End If
        '                    'dblBebuBetrag = 1000.0#
        '                    If SubRow("lngKST") > 0 Then
        '                        strBeBuEintrag = SubRow("lngKST").ToString + "{<}" + SubRow("strDebSubText") + "{<}" + "CALCULATE" + "{>}"    '"PROD{<}BebuText{<}" + dblBebuBetrag.ToString + "{>}"
        '                    Else
        '                        'strBeBuEintrag = "999999" + "{<}" + SubRow("strDebSubText") + "{<}" + "CALCULATE" + "{>}"
        '                        strBeBuEintrag = Nothing
        '                    End If
        '                    If Not IsDBNull(SubRow("strMwStKey")) And
        '                            SubRow("strMwStKey") <> "null" And
        '                            SubRow("lngKto") <> 6906 Then 'And SubRow("strMwStKey") <> "25" Then
        '                        If strBuchType = "R" Then
        '                            strSteuerFeld = Main.FcGetSteuerFeld(FBhg,
        '                                                                 SubRow("lngKto"),
        '                                                                 SubRow("strDebSubText"),
        '                                                                 SubRow("dblBrutto") * -1,
        '                                                                 SubRow("strMwStKey"),
        '                                                                 SubRow("dblMwSt") * -1)     '"25{<}DEBI D Bhg Export MwSt{<}0{>}"
        '                        Else
        '                            strSteuerFeld = Main.FcGetSteuerFeld(FBhg,
        '                                                                 SubRow("lngKto"),
        '                                                                 SubRow("strDebSubText"),
        '                                                                 SubRow("dblBrutto"),
        '                                                                 SubRow("strMwStKey"),
        '                                                                 SubRow("dblMwSt"))     '"25{<}DEBI D Bhg Export MwSt{<}0{>}"
        '                        End If
        '                    Else
        '                        strSteuerFeld = "STEUERFREI"
        '                    End If
        '                    'strSteuerInfo = Split(FBhg.GetKontoInfo(intGegenKonto.ToString), "{>}")
        '                    'Debug.Print("Konto-Info: " + strSteuerInfo(26))

        '                    Try

        '                        booBooingok = True
        '                        Call DbBhg.SetVerteilung(intGegenKonto.ToString,
        '                                                 strFibuText,
        '                                                 dblNettoBetrag.ToString,
        '                                                 strSteuerFeld,
        '                                                 strBeBuEintrag)

        '                        'Application.DoEvents()

        '                    Catch ex As Exception
        '                        strErrMessage = "Problem " + (Err.Number And 65535).ToString + " Verteilung " + intDebBelegsNummer.ToString + vbCrLf
        '                        strErrMessage += "RG " + strRGNbr + vbCrLf
        '                        strErrMessage += "Konto " + SubRow("lngKto").ToString + vbCrLf
        '                        strErrMessage += "Gegenkonto " + intGegenKonto.ToString + vbCrLf
        '                        strErrMessage += "Betrag " + dblNettoBetrag.ToString + vbCrLf
        '                        strErrMessage += "Steuer " + strSteuerFeld + vbCrLf
        '                        strErrMessage += "Bebu " + strBeBuEintrag

        '                        MessageBox.Show(ex.Message + vbCrLf + strErrMessage, "Problem " + (Err.Number And 65535).ToString, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        '                        If (Err.Number And 65535) < 10000 Then
        '                            booBooingok = False
        '                        Else
        '                            booBooingok = True
        '                        End If

        '                    End Try

        '                    strSteuerFeld = Nothing
        '                    strBeBuEintrag = Nothing

        '                    'Status Sub schreiben
        '                    'Application.DoEvents()

        '                Next

        '                Try

        '                    booBooingok = True
        '                    Call DbBhg.WriteBuchung()

        '                    'Bei SplittBill 2ter Rechnung TZahlung auf LinkedRG machen
        '                    'Prinzip: Beleg einlesen anhand und Betrag ausrechnen => Summe Beleg - diesen Beleg
        '                    If row("booLinked") And Mid(row("strDebStatusBitLog"), 13, 1) = "0" Then 'Nur wenn Beleg in gleicher Buha
        '                        'Betrag von Beleg 1 holen
        '                        intLaufNbr = DbBhg.doesBelegExist2(row("lngLinkedDeb").ToString,
        '                                                           row("strDebCur"),
        '                                                           row("lngLinkedRG").ToString,
        '                                                           "NOT_SET",
        '                                                           "R",
        '                                                           "NOT_SET",
        '                                                           "NOT_SET",
        '                                                           "NOT_SET")

        '                        If intLaufNbr > 0 Then
        '                            strBeleg = DbBhg.GetBeleg(row("lngLinkedDeb").ToString,
        '                                                      intLaufNbr.ToString)

        '                            strBelegArr = Split(strBeleg, "{>}")
        '                            If strBelegArr(4) = "B" Then 'schon bezahlt
        '                                'Ausbuchen?, wohin mit dem Betrag?
        '                            Else

        '                                'Welcher Betrag wurde schon bezahlt?
        '                                dblSplitPayed = dblBetrag

        '                                'Teilzahlung buchen
        '                                Call DbBhg.SetZahlung(1944,
        '                                                  strBelegDatum,
        '                                                  strValutaDatum,
        '                                                  row("strDebCur"),
        '                                                  dblKurs,
        '                                                  "",
        '                                                  "",
        '                                                  row("lngLinkedDeb"),
        '                                                  dblSplitPayed.ToString,
        '                                                  row("strDebCur"),
        '                                                  ,
        '                                                  row("lngDebIdentNbr").ToString + ", TZ " + row("strDebRGNbr").ToString)
        '                                'Application.DoEvents()

        '                                Call DbBhg.WriteTeilzahlung4(intLaufNbr.ToString,
        '                                                         row("lngDebIdentNbr").ToString + ", TZ " + row("strDebRGNbr").ToString,
        '                                                         "NOT_SET",
        '                                                         ,
        '                                                         "NOT_SET",
        '                                                         "NOT_SET",
        '                                                         "Default",
        '                                                         "Default")
        '                                'Application.DoEvents()

        '                            End If

        '                        End If

        '                    End If

        '                Catch ex As Exception
        '                    'MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Belegerstellung " + intDebBelegsNummer.ToString + ", RG " + strRGNbr)
        '                    If (Err.Number And 65535) < 10000 Then
        '                        strErrMessage = "Belegerstellung RG " + strRGNbr + " Beleg " + intDebBelegsNummer.ToString + " NICHT möglich!"
        '                        MessageBox.Show(ex.Message + vbCrLf + strErrMessage, "Problem " + (Err.Number And 65535).ToString, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        '                        booBooingok = False
        '                    Else
        '                        strErrMessage = "Belegerstellung RG " + strRGNbr + " Beleg " + intDebBelegsNummer.ToString + " möglich mit Warnung"
        '                        MessageBox.Show(ex.Message, "Warnung " + (Err.Number And 65535).ToString, MessageBoxButtons.OK, MessageBoxIcon.Warning)
        '                        booBooingok = True
        '                    End If

        '                End Try


        '            Else

        '                'Buchung nur in Fibu
        '                'Prinzip Funktion WriteBuchung() anwenden mit allen Parametern

        '                'Verdopplung interne BelegsNummer verhindern
        '                FBhg.CheckDoubleIntBelNbr = "J"

        '                If IIf(IsDBNull(row("strOPNr")), "", row("strOPNr")) <> "" And IIf(IsDBNull(row("lngDebIdentNbr")), 0, row("lngDebIdentNbr")) <> 0 Then
        '                    'Belegsnummer abholen fall keine Beleg-Nummer angegeben
        '                    intDebBelegsNummer = FBhg.GetNextBelNbr()
        '                    'Prüfen ob wirklich frei
        '                    intReturnValue = 10
        '                    Do Until intReturnValue = 0
        '                        intReturnValue = FBhg.doesBelegExist(intDebBelegsNummer,
        '                                                             "NOT_SET",
        '                                                             "NOT_SET",
        '                                                             String.Concat(Microsoft.VisualBasic.Left(cmbPerioden.SelectedItem, 4) - 1, "0101"),
        '                                                             String.Concat(Microsoft.VisualBasic.Left(cmbPerioden.SelectedItem, 4), "1231"))
        '                        If intReturnValue <> 0 Then
        '                            intDebBelegsNummer += 1
        '                        End If
        '                    Loop
        '                    'Debug.Print("Belegnummer taken:  " + intDebBelegsNummer.ToString)
        '                Else
        '                    If IIf(IsDBNull(row("strOPNr")), "", row("strOPNr")) <> "" Then
        '                        intDebBelegsNummer = Convert.ToInt32(row("strOPNr"))
        '                    Else
        '                        intDebBelegsNummer = row("lngDebIdentNbr")
        '                    End If
        '                End If
        '                'Variablen zuweisen
        '                strBelegDatum = Format(row("datDebRGDatum"), "yyyyMMdd").ToString
        '                strValutaDatum = Format(row("datDebValDatum"), "yyyyMMdd").ToString
        '                'strDebiText = row("strDebText")
        '                strCurrency = row("strDebCur")
        '                If strCurrency <> "CHF" Then 'Muss ergänzt werden => Was ist Leitwährung auf dem Konto
        '                    dblKurs = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg)
        '                Else
        '                    dblKurs = 1.0#
        '                End If

        '                selDebiSub = dsDebitoren.Tables("tblDebiSubsFromUser").Select("strRGNr='" + row("strDebRGNbr") + "'")
        '                strRGNbr = row("strDebRGNbr")

        '                If selDebiSub.Length = 2 Then

        '                    'Initialisieren
        '                    dblNettoBetrag = 0
        '                    dblSollBetrag = 0
        '                    dblHabenBetrag = 0
        '                    strBeBuEintrag = String.Empty
        '                    strBeBuEintragSoll = String.Empty
        '                    strBeBuEintragHaben = String.Empty
        '                    strSteuerFeld = String.Empty
        '                    strSteuerFeldHaben = String.Empty
        '                    strSteuerFeldSoll = String.Empty

        '                    For Each SubRow As DataRow In selDebiSub

        '                        If SubRow("intSollHaben") = 0 Then 'Soll

        '                            intSollKonto = SubRow("lngKto")
        '                            dblKursSoll = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg, intSollKonto)
        '                            'strSteuerInfo = Split(FBhg.GetKontoInfo(intSollKonto.ToString), "{>}")
        '                            'Debug.Print("Konto-Info Soll: " + strSteuerInfo(26))
        '                            dblSollBetrag = SubRow("dblNetto")
        '                            strDebiTextSoll = SubRow("strDebSubText")
        '                            If SubRow("dblMwSt") > 0 Then
        '                                strSteuerFeldSoll = Main.FcGetSteuerFeld(FBhg,
        '                                                                         SubRow("lngKto"),
        '                                                                         strDebiTextSoll,
        '                                                                         SubRow("dblBrutto") * dblKursSoll,
        '                                                                         SubRow("strMwStKey"),
        '                                                                         SubRow("dblMwSt"))
        '                            Else
        '                                strSteuerFeldSoll = "STEUERFREI"
        '                            End If
        '                            If SubRow("lngKST") > 0 Then
        '                                strBeBuEintragSoll = SubRow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
        '                            End If


        '                        ElseIf SubRow("intSollHaben") = 1 Then 'Haben

        '                            intHabenKonto = SubRow("lngKto")
        '                            dblKursHaben = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg, intHabenKonto)
        '                            'strSteuerInfo = Split(FBhg.GetKontoInfo(intSollKonto.ToString), "{>}")
        '                            'Debug.Print("Konto-Info Haben: " + strSteuerInfo(26))
        '                            dblHabenBetrag = SubRow("dblNetto") * -1
        '                            'dblHabenBetrag = dblSollBetrag
        '                            strDebiTextHaben = SubRow("strDebSubText")
        '                            If SubRow("dblMwSt") * -1 > 0 Then
        '                                strSteuerFeldHaben = Main.FcGetSteuerFeld(FBhg,
        '                                                                          SubRow("lngKto"),
        '                                                                          strDebiTextHaben,
        '                                                                          SubRow("dblBrutto") * dblKursHaben * -1,
        '                                                                          SubRow("strMwStKey"),
        '                                                                          SubRow("dblMwSt") * -1)
        '                            Else
        '                                strSteuerFeldHaben = "STEUERFREI"
        '                            End If
        '                            If SubRow("lngKST") > 0 Then
        '                                strBeBuEintragHaben = SubRow("lngKST").ToString + "{<}" + strDebiTextHaben + "{<}" + "CALCULATE" + "{>}"
        '                            End If

        '                        Else

        '                            MsgBox("Nicht definierter Wert Sub-Buchungs-SollHaben: " + SubRow("intSollHaben").ToString)

        '                        End If
        '                        'Application.DoEvents()

        '                    Next

        '                    'Tieferer Betrag für die Gesamt-Buchung herausfinden
        '                    If dblSollBetrag <= dblHabenBetrag Then
        '                        dblNettoBetrag = dblSollBetrag
        '                    ElseIf dblHabenBetrag < dblSollBetrag Then
        '                        dblNettoBetrag = dblHabenBetrag
        '                    End If

        '                    Try

        '                        booBooingok = True
        '                        'Buchung ausführen
        '                        Call FBhg.WriteBuchung(0,
        '                                           intDebBelegsNummer,
        '                                           strBelegDatum,
        '                                           intSollKonto.ToString,
        '                                           strDebiTextSoll,
        '                                           strCurrency,
        '                                           dblKursSoll.ToString,
        '                                           (dblNettoBetrag * dblKursSoll).ToString,
        '                                           strSteuerFeldSoll,
        '                                           intHabenKonto.ToString,
        '                                           strDebiTextHaben,
        '                                           strCurrency,
        '                                           dblKursHaben.ToString,
        '                                           (dblNettoBetrag * dblKursHaben).ToString,
        '                                           strSteuerFeldHaben,
        '                                           strCurrency,
        '                                           dblKurs.ToString,
        '                                           (dblNettoBetrag * dblKurs).ToString,
        '                                           dblNettoBetrag.ToString,
        '                                           strBeBuEintragSoll,
        '                                           strBeBuEintragHaben,
        '                                           strValutaDatum)

        '                        'Application.DoEvents()

        '                    Catch ex As Exception
        '                        MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Belegerstellung " + intDebBelegsNummer.ToString + ", RG " + strRGNbr, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        '                        If (Err.Number And 65535) < 10000 Then
        '                            booBooingok = False
        '                        Else
        '                            booBooingok = True
        '                        End If

        '                    End Try

        '                    'Initialisieren
        '                    'dblNettoBetrag = 0
        '                    'dblSollBetrag = 0
        '                    'dblHabenBetrag = 0
        '                    'strBeBuEintrag = ""
        '                    'strBeBuEintragSoll = ""
        '                    'strBeBuEintragHaben = ""
        '                    'strSteuerFeld = ""
        '                    'strSteuerFeldHaben = ""
        '                    'strSteuerFeldSoll = ""


        '                    'Vergebene Nummer checken
        '                    'intDebBelegsNummer = FBhg.GetBuchLaufnr()

        '                Else
        '                    'Sammelbeleg
        '                    'MsgBox("Nicht 2 Subbuchungen.")
        '                    'Variablen initiieren
        '                    strDebiText = row("strDebText")
        '                    intCommonKonto = row("lngDebKtoNbr") 'Sammelkonto

        '                    'Beleg-Kopf
        '                    Call FBhg.SetSammelBhgCommonT2(strValutaDatum,
        '                                                   intDebBelegsNummer.ToString,
        '                                                   intCommonKonto.ToString,
        '                                                   strDebiText,
        '                                                   strBelegDatum)

        '                    'Buchungen
        '                    For Each SubRow As DataRow In selDebiSub

        '                        'Common - Konto ausblenden da sonst Doppelbuchung
        '                        If SubRow("lngKto") <> intCommonKonto Then

        '                            intSollKonto = 0
        '                            strDebiTextSoll = String.Empty
        '                            strDebiCurrency = String.Empty
        '                            dblKursSoll = 0
        '                            dblSollBetrag = 0
        '                            strSteuerFeldSoll = String.Empty
        '                            intHabenKonto = 0
        '                            strDebiTextHaben = String.Empty
        '                            strKrediCurrency = String.Empty
        '                            dblKursHaben = 0
        '                            dblHabenBetrag = 0
        '                            strSteuerFeldHaben = String.Empty
        '                            dblBuchBetrag = 0
        '                            dblBasisBetrag = 0
        '                            strBeBuEintragSoll = String.Empty
        '                            strBeBuEintragHaben = String.Empty
        '                            strErfassungsDatum = Format(Date.Today(), "yyyyMMdd").ToString

        '                            If SubRow("intSollHaben") = 0 And SubRow("lngKto") <> intCommonKonto Then 'Soll

        '                                intSollKonto = SubRow("lngKto")
        '                                strDebiTextSoll = SubRow("strDebSubText")
        '                                strDebiCurrency = strCurrency
        '                                dblKursSoll = 1 / Main.FcGetKurs(strCurrency, strValutaDatum, FBhg, intSollKonto)
        '                                dblSollBetrag = SubRow("dblNetto")
        '                                If SubRow("dblMwSt") > 0 Then
        '                                    strSteuerFeldSoll = Main.FcGetSteuerFeld(FBhg, SubRow("lngKto"), strDebiTextSoll, SubRow("dblBrutto") * dblKursSoll, SubRow("strMwStKey"), SubRow("dblMwSt"))
        '                                Else
        '                                    strSteuerFeldSoll = "STEUERFREI"
        '                                End If
        '                                If SubRow("lngKST") > 0 Then
        '                                    strBeBuEintragSoll = SubRow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
        '                                End If

        '                                'Haben Seite Common-Konto
        '                                intHabenKonto = intCommonKonto
        '                                strDebiTextHaben = SubRow("strDebSubText")
        '                                strKrediCurrency = strCurrency
        '                                dblKursHaben = dblKursSoll
        '                                dblHabenBetrag = SubRow("dblNetto")
        '                                'If SubRow("dblMwSt") > 0 Then
        '                                ' strSteuerFeldHaben = Main.FcGetSteuerFeld(FBhg, SubRow("lngKto"), strDebiTextSoll, SubRow("dblBrutto") * dblKursSoll, SubRow("strMwStKey"), SubRow("dblMwSt"))
        '                                'End If
        '                                If SubRow("lngKST") > 0 Then
        '                                    strBeBuEintragHaben = SubRow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
        '                                End If

        '                                'Betrag
        '                                dblBuchBetrag = SubRow("dblBrutto")
        '                                dblBasisBetrag = SubRow("dblBrutto") 'Bei nicht CHF umrechnen

        '                            ElseIf SubRow("intSollHaben") = 1 Then 'Haben

        '                                intHabenKonto = SubRow("lngKto")
        '                                strDebiTextHaben = SubRow("strDebSubText")
        '                                strKrediCurrency = strCurrency
        '                                dblKursHaben = 1 / Main.FcGetKurs(strCurrency, strValutaDatum, FBhg, intHabenKonto)
        '                                dblHabenBetrag = SubRow("dblNetto") * -1
        '                                If (SubRow("dblMwSt") * -1) > 0 Then
        '                                    strSteuerFeldHaben = Main.FcGetSteuerFeld(FBhg, SubRow("lngKto"), strDebiTextHaben, SubRow("dblBrutto") * dblKursHaben * -1, SubRow("strMwStKey"), SubRow("dblMwSt") * -1)
        '                                Else
        '                                    strSteuerFeldHaben = "STEUERFREI"
        '                                End If
        '                                If SubRow("lngKST") > 0 Then
        '                                    strBeBuEintragHaben = SubRow("lngKST").ToString + "{<}" + strDebiTextHaben + "{<}" + "CALCULATE" + "{>}"
        '                                End If

        '                                'Soll - Seite Common - Konto 
        '                                intSollKonto = intCommonKonto
        '                                strDebiTextSoll = SubRow("strDebSubText")
        '                                strDebiCurrency = strCurrency
        '                                dblKursSoll = dblKursHaben
        '                                dblSollBetrag = SubRow("dblNetto") * -1

        '                                'If SubRow("dblMwSt") * -1 > 0 Then
        '                                ' strSteuerFeldSoll = Main.FcGetSteuerFeld(FBhg, SubRow("lngKto"), strDebiTextSoll, SubRow("dblBrutto") * dblKursSoll * -1, SubRow("strMwStKey"), SubRow("dblMwSt") * -1)
        '                                'End If
        '                                If SubRow("lngKST") > 0 Then
        '                                    strBeBuEintragSoll = SubRow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
        '                                End If

        '                                dblBuchBetrag = SubRow("dblBrutto") * -1
        '                                dblBasisBetrag = SubRow("dblBrutto") * -1 'Bei nicht CHF umrechnen

        '                            End If
        '                            'Buchungsbetrag von Kopfbuchung nehmen
        '                            'dblBuchBetrag = row("dblDebBrutto")
        '                            'dblBasisBetrag = row("dblDebBrutto")

        '                            Call FBhg.SetSammelBhgT(intSollKonto.ToString,
        '                                                    strDebiTextSoll,
        '                                                    strDebiCurrency,
        '                                                    dblKursSoll.ToString,
        '                                                    dblSollBetrag.ToString,
        '                                                    strSteuerFeldSoll,
        '                                                    intHabenKonto.ToString,
        '                                                    strDebiTextHaben,
        '                                                    strKrediCurrency,
        '                                                    dblKursHaben.ToString,
        '                                                    dblHabenBetrag.ToString,
        '                                                    strSteuerFeldHaben,
        '                                                    strCurrency,
        '                                                    dblKurs.ToString,
        '                                                    dblBuchBetrag.ToString,
        '                                                    dblBasisBetrag.ToString,
        '                                                    strBeBuEintragSoll,
        '                                                    strBeBuEintragHaben,
        '                                                    strErfassungsDatum)

        '                            'Application.DoEvents()

        '                        End If

        '                    Next

        '                    'Sammelbeleg schreiben
        '                    Try

        '                        booBooingok = True
        '                        Call FBhg.WriteSammelBhgT()

        '                    Catch ex As Exception

        '                        MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Belegerstellung " + intDebBelegsNummer.ToString + ", RG " + strRGNbr, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        '                        If (Err.Number And 65535) < 10000 Then
        '                            booBooingok = False
        '                        Else
        '                            booBooingok = True
        '                        End If
        '                    End Try


        '                End If

        '            End If

        '            If booBooingok Then
        '                If row("booPGV") Then
        '                    'Bei PGV Buchungen
        '                    If IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")) = "" Or
        '                        (IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")) = "RV" And row("intPGVMthsAY") + row("intPGVMthsNY") > 1) Then

        '                        intReturnValue = MainDebitor.FcPGVDTreatment(FBhg,
        '                                                           Finanz,
        '                                                           DbBhg,
        '                                                           PIFin,
        '                                                           BeBu,
        '                                                           KrBhg,
        '                                                           dsDebitoren.Tables("tblDebiSubsFromUser"),
        '                                                           row("strDebRGNbr"),
        '                                                           intDebBelegsNummer,
        '                                                           row("strDebCur"),
        '                                                           row("datDebValDatum"),
        '                                                           "M",
        '                                                           row("datPGVFrom"),
        '                                                           row("datPGVTo"),
        '                                                           row("intPGVMthsAY") + row("intPGVMthsNY"),
        '                                                           row("intPGVMthsAY"),
        '                                                           row("intPGVMthsNY"),
        '                                                           1311,
        '                                                           1312,
        '                                                           cmbPerioden.SelectedItem,
        '                                                           objdbConn,
        '                                                           objdbMSSQLConn,
        '                                                           objdbSQLcommand,
        '                                                           cmbBuha.SelectedValue,
        '                                                           objdtInfo,
        '                                                           strYear,
        '                                                           intTeqNbr,
        '                                                           intTeqNbrLY,
        '                                                           intTeqNbrPLY,
        '                                                           IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")),
        '                                                           datPeriodFrom,
        '                                                           datPeriodTo,
        '                                                           strPeriodStatus)


        '                    Else
        '                        intReturnValue = MainDebitor.FcPGVDTreatmentYC(FBhg,
        '                                                           Finanz,
        '                                                           DbBhg,
        '                                                           PIFin,
        '                                                           BeBu,
        '                                                           KrBhg,
        '                                                           dsDebitoren.Tables("tblDebiSubsFromUser"),
        '                                                           row("strDebRGNbr"),
        '                                                           intDebBelegsNummer,
        '                                                           row("strDebCur"),
        '                                                           row("datDebValDatum"),
        '                                                           "M",
        '                                                           row("datPGVFrom"),
        '                                                           row("datPGVTo"),
        '                                                           row("intPGVMthsAY") + row("intPGVMthsNY"),
        '                                                           row("intPGVMthsAY"),
        '                                                           row("intPGVMthsNY"),
        '                                                           1311,
        '                                                           1312,
        '                                                           cmbPerioden.SelectedItem,
        '                                                           objdbConn,
        '                                                           objdbMSSQLConn,
        '                                                           objdbSQLcommand,
        '                                                           cmbBuha.SelectedValue,
        '                                                           objdtInfo,
        '                                                           strYear,
        '                                                           intTeqNbr,
        '                                                           intTeqNbrLY,
        '                                                           intTeqNbrPLY,
        '                                                           IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")),
        '                                                           datPeriodFrom,
        '                                                           datPeriodTo,
        '                                                           strPeriodStatus)
        '                    End If


        '                End If

        '                'Status Head schreiben
        '                row("strDebBookStatus") = row("strDebStatusBitLog")
        '                row("booBooked") = True
        '                row("datBooked") = Now()
        '                row("lngBelegNr") = intDebBelegsNummer
        '                'Application.DoEvents()

        '                'Status in File RG-Tabelle schreiben
        '                intReturnValue = MainDebitor.FcWriteToRGTable(cmbBuha.SelectedValue,
        '                                                              row("strDebRGNbr"),
        '                                                              row("datBooked"),
        '                                                              row("lngBelegNr"),
        '                                                              objdbAccessConn,
        '                                                              objOracleConn,
        '                                                              objdbConn,
        '                                                              row("booDatChanged"),
        '                                                              row("datDebRGDatum"),
        '                                                              row("datDebValDatum"))
        '                If intReturnValue <> 0 Then
        '                    'Throw an exception
        '                End If

        '                'Evtl. Query nach Buchung ausführen
        '                Call MainDebitor.FcExecuteAfterDebit(cmbBuha.SelectedValue, objdbConn)
        '            End If

        '        End If

        '    Next
        '    'Status in t_sage_syncstatus schreiben
        '    'intReturnValue = MainDebitor.FcWriteEndToSync(objdbConn,
        '    '                                              cmbBuha.SelectedValue,
        '    '                                              1,
        '    '                                              Date.Now,
        '    '                                              0,
        '    '                                              IIf(booBooingok, "ok", "Probleme"))

        '    intReturnValue = WFDBClass.FcWriteEndToSync(objdbConn,
        '                                                cmbBuha.SelectedValue,
        '                                                1,
        '                                                0,
        '                                                IIf(booBooingok, "ok", "Probleme"))




        'Catch ex As Exception
        '    MessageBox.Show(ex.Message, "Generelles Problem " + (Err.Number And 65535).ToString + " Belegerstellung " + intDebBelegsNummer.ToString + ", RG " + strRGNbr, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        'Finally

        '    'Application.DoEvents()
        '    'Grid neu aufbauen, Daten von Mandant einlesen
        '    Call butDebitoren.PerformClick()

        '    Me.Cursor = Cursors.Default
        '    'Me.butImport.Enabled = True
        '    'Application.Restart()

        'End Try

    End Sub


    Private Sub dgvBookings_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs)

        Dim intDecidiveCell As Integer

        Try

            'If intMode = 0 Then
            '    intDecidiveCell = 2
            'Else
            '    intDecidiveCell = 3
            'End If

            'If e.ColumnIndex = intDecidiveCell And e.RowIndex >= 0 Then


            '    If intMode = 0 Then
            '        If dgvBookingsO.Rows(e.RowIndex).Cells("booDebBook").Value Then

            '            'MsgBox("Geändert zu checked " + dgvDebitoren.Rows(e.RowIndex).Cells("strDebRGNbr").Value + ", " + dgvDebitoren.Rows(e.RowIndex).Cells("booDebBook").Value.ToString + Val(dgvDebitoren.Rows(e.RowIndex).Cells("strDebStatusBitLog").Value).ToString)
            '            'Zulassen? = keine Fehler
            '            If Val(dgvBookingsO.Rows(e.RowIndex).Cells("strDebStatusBitLog").Value) <> 0 And Val(dgvBookingsO.Rows(e.RowIndex).Cells("strDebStatusBitLog").Value) <> 10000 Then
            '                If Val(Strings.Mid(dgvBookingsO.Rows(e.RowIndex).Cells("strDebStatusBitLog").Value, 13, 1)) = 1 Then
            '                    MsgBox("Erst - RG Splitt-Bill ist nicht auffindbar. Wird trotzdem gebucht, wird nur die Zweit-RG gebucht.", vbOKOnly + vbExclamation, "Splitt-Bill No RG 1")
            '                Else
            '                    MsgBox("Debi-Rechnung ist nicht buchbar.", vbOKOnly + vbExclamation, "Nicht buchbar")
            '                    dgvBookingsO.Rows(e.RowIndex).Cells("booDebBook").Value = False
            '                End If
            '            End If

            '        End If

            '    Else

            '        If dgvBookingsO.Rows(e.RowIndex).Cells("booKredBook").Value Then

            '            'MsgBox("Geändert zu checked " + dgvDebitoren.Rows(e.RowIndex).Cells("strDebRGNbr").Value + ", " + dgvDebitoren.Rows(e.RowIndex).Cells("booDebBook").Value.ToString + Val(dgvDebitoren.Rows(e.RowIndex).Cells("strDebStatusBitLog").Value).ToString)
            '            'Zulassen? = keine Fehler
            '            If Val(dgvBookingsO.Rows(e.RowIndex).Cells("strKredStatusBitLog").Value) <> 0 Then
            '                MsgBox("Kredi-Rechnung ist nicht buchbar.", vbOKOnly + vbExclamation, "Nicht buchbar")
            '                dgvBookingsO.Rows(e.RowIndex).Cells("booKredBook").Value = False
            '            End If

            '        End If

            '    End If

            'End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Generelles Problem " + Err.Number.ToString)

        End Try


    End Sub


    Private Sub dgvBookings_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

        Try

            'If e.RowIndex >= 0 Then

            '    If intMode = 0 Then 'Debitoren
            '        'dgvBookingSubO.DataSource = dsDebitoren.Tables("tblDebiSubsFromUser").Select("strRGNr='" + dgvBookingsO.Rows(e.RowIndex).Cells("strDebRGNbr").Value + "'").CopyToDataTable
            '    Else
            '        'dgvBookingSubO.DataSource = objdtKreditorenSub.Select("lngKredID=" + dgvBookingsO.Rows(e.RowIndex).Cells("lngKredID").Value.ToString).CopyToDataTable
            '    End If

            'End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)

        Finally
            'dgvBookingSub.Update()

        End Try



    End Sub

    Private Sub butKreditoren_Click(sender As Object, e As EventArgs) Handles butKreditoren.Click

        'Dim intReturnValue As Int16
        Dim intNbrKrediForms As Int16

        Try

            Me.Cursor = Cursors.WaitCursor
            System.GC.Collect()

            'Zuerst überprüfen, ob schon eine Instanz für mandant existiert
            For Each frm As Form In Application.OpenForms()
                If frm.Text = "Kreditor " + lstBoxMandant.Text Then
                    intNbrKrediForms += 1
                End If
            Next
            If intNbrKrediForms > 0 Then
                MessageBox.Show("Es ist schone ein Formular für " + "Kredtor " + lstBoxMandant.Text + " geöffnet.", "Pro Mandant nur 1 Formular möglich", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
                intNbrKrediForms = 0
                Dim frmKredDisp As New frmKredDisp
                frmKredDisp.Cursor = Cursors.WaitCursor
                frmKredDisp.Text = "Kreditor " + lstBoxMandant.Text
                frmKredDisp.MdiParent = Me
                frmKredDisp.Show()
                frmKredDisp.Top = 105
                'Für korrekte Plazierung
                For Each frm As Form In Application.OpenForms()
                    If frm.Name = "frmKredDisp" Then
                        intNbrKrediForms += 1
                    End If
                Next
                'Je nach Anzahl von Links her verschieben
                frmKredDisp.Left = intNbrKrediForms * 20
                frmKredDisp.FcKrediDisplay(lstBoxMandant.SelectedValue,
                                     lstBoxPerioden)

            End If

            frmKredDisp.Cursor = Cursors.Default
            'Application.DoEvents()

            'intMode = 1

            'objdtKreditorenHead = Nothing
            'objdtKreditorenHeadRead = Nothing
            'objdtKreditorenSub = Nothing

            ''Tabelle Kredi - Head erstellen
            'objdtKreditorenHead = Main.tblKreditorenHead()
            'objdtKreditorenHeadRead = Main.tblKreditorenHead()

            ''Tabelle Kreditoren - Sub erstellen
            'objdtKreditorenSub = Main.tblKreditorenSub()

            ''objdtKreditorenHead.Clear()
            ''objdtKreditorenSub.Clear()
            ''objdtKreditorenHeadRead.Clear()
            'objdtInfo.Clear()

            ''If dgvBookings.Columns.Contains("intBuchungsart") Then
            ''    dgvBookings.Columns.Remove("intBuchungsart")
            ''End If

            ''DGV Kreditoren
            'dgvBookings.DataSource = Nothing
            'dgvBookings.Rows.Clear()
            'dgvBookings.Columns.Clear()
            ''dgvBookings.DataSource = objdtKreditorenHead
            ''dgvBookingSub.DataSource = Nothing
            ''dgvBookingSub.Rows.Clear()
            ''dgvBookingSub.Columns.Clear()
            ''dgvBookingSub.DataSource = objdtKreditorenSub

            ''objdbConn.Open()
            ''Call InitdgvKreditoren()
            ''Call InitdgvKreditorenSub()
            ''objdbConn.Close()

            'Call InitVar()

            'Call Main.FcLoginSage2(objdbConn,
            '                  objdbMSSQLConn,
            '                  objdbSQLcommand,
            '                  Finanz,
            '                  FBhg,
            '                  DbBhg,
            '                  PIFin,
            '                  BeBu,
            '                  KrBhg,
            '                  cmbBuha.SelectedValue,
            '                  objdtInfo,
            '                  cmbPerioden.SelectedItem,
            '                  strYear,
            '                  intTeqNbr,
            '                  intTeqNbrLY,
            '                  intTeqNbrPLY,
            '                  datPeriodFrom,
            '                  datPeriodTo,
            '                  strPeriodStatus)

            ''Transitorische Buchungen?
            ''Call Main.fcCheckTransitorischeKredit(cmbBuha.SelectedValue,
            ''                                 objdbConn,
            ''                                 objdbAccessConn)


            'intReturnValue = MainKreditor.FcFillKredit(cmbBuha.SelectedValue,
            '                                       objdtKreditorenHeadRead,
            '                                       objdtKreditorenSub,
            '                                       objdbConn,
            '                                       objdbAccessConn)
            'If intReturnValue = 1 Then
            '    MessageBox.Show("Keine Kreditoren-Defintion hinterlegt.", "Keine Definition")
            'End If

            'Call Main.InsertDataTableColumnName(objdtKreditorenHeadRead, objdtKreditorenHead)

            ''Grid neu aufbauen
            ''dgvBookingSub.Update()
            ''dgvBookings.Update()
            ''dgvBookings.Refresh()
            'Application.DoEvents()

            'Call Main.FcCheckKredit(cmbBuha.SelectedValue,
            '                    objdtKreditorenHead,
            '                    objdtKreditorenSub,
            '                    Finanz,
            '                    FBhg,
            '                    KrBhg,
            '                    PIFin,
            '                    objdbConn,
            '                    objdbConnZHDB02,
            '                    objdbcommand,
            '                    objdbcommandZHDB02,
            '                    objOracleConn,
            '                    objOracleCmd,
            '                    objdbAccessConn,
            '                    objdtInfo,
            '                    cmbBuha.Text,
            '                    strYear,
            '                    cmbPerioden.SelectedItem,
            '                    datPeriodFrom,
            '                    datPeriodTo,
            '                    strPeriodStatus,
            '                    chkValutaCorrect.Checked,
            '                    dtpValutaCorrect.Value)

            'dgvBookings.DataSource = objdtKreditorenHead
            ''dgvBookingSub.DataSource = objdtKreditorenSub
            'objdbConn.Open()
            'Call InitdgvKreditoren()
            'Call InitdgvKreditorenSub()
            'objdbConn.Close()


            ''Anzahl schreiben
            'txtNumber.Text = objdtKreditorenHead.Rows.Count.ToString

            ''Import Debitoren desattivate
            'Me.butImport.Enabled = False
            'Me.butImportK.Enabled = True

            'Me.Cursor = Cursors.Default

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem Kreditorenauflistung", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        Finally

            Me.Cursor = Cursors.Default

        End Try


    End Sub

    Private Sub butImportK_Click(sender As Object, e As EventArgs)

        ''Dim objdtDebitorenHead As New DataTable("tbliDebiHead")

        'Dim intReturnValue As Int32
        'Dim intKredBelegsNummer As UInt32
        'Dim strExtKredBelegsNummer As String

        'Dim intKreditorNbr As Int32
        'Dim strBuchType As String
        'Dim strBelegDatum As String
        'Dim strValutaDatum As String
        'Dim strVerfallDatum As String
        'Dim strReferenz As String
        'Dim intKondition As Int32
        'Dim intKonditionLN As Int16
        'Dim strSachBID As String = String.Empty
        'Dim strVerkID As String = String.Empty
        'Dim strMahnerlaubnis As String
        'Dim sngAktuelleMahnstufe As Single
        'Dim dblBetrag As Double
        'Dim dblKurs As Double
        'Dim strExtBelegNbr As String = String.Empty
        'Dim strSkonto As String = String.Empty
        'Dim strCurrency As String
        'Dim strKrediText As String
        'Dim intBankNbr As Int16
        'Dim strZahlSperren As String = "N"
        'Dim strVorausZahlung As String = "N"
        'Dim strErfassungsArt As String = "K"
        'Dim intTeilnehmer As Int32
        'Dim strTeilnehmer As String
        'Dim intEigeneBank As Int32

        'Dim intGegenKonto As Int32
        'Dim strFibuText As String
        'Dim dblNettoBetrag As Double
        'Dim dblBruttoBetrag As Double
        'Dim dblMwStBetrag As Double
        'Dim dblBebuBetrag As Double
        'Dim strBeBuEintrag As String
        'Dim strSteuerFeld As String

        'Dim intSollKonto As Int32
        'Dim intHabenKonto As Int32
        'Dim dblSollBetrag As Double
        'Dim dblHabenBetrag As Double
        'Dim strSteuerFeldSoll As String = String.Empty
        'Dim strSteuerFeldHaben As String = String.Empty
        'Dim strBeBuEintragSoll As String = String.Empty
        'Dim strBeBuEintragHaben As String = String.Empty
        'Dim strKrediTextSoll As String = String.Empty
        'Dim strKrediTextHaben As String = String.Empty
        'Dim dblKursSoll As Double = 0
        'Dim dblKursHaben As Double = 0

        'Dim selKrediSub() As DataRow
        'Dim strSteuerInfo() As String
        'Dim strDebiLine As String
        'Dim strDebitor() As String

        'Dim booBookingok As Boolean

        ''Sammelbeleg
        'Dim intCommonKonto As Int32
        'Dim strKRGReferTo As String

        ''Dim intTeqNbr As Int32

        'Try


        '    Me.Cursor = Cursors.WaitCursor
        '    'Application.DoEvents()
        '    'Button disablen damit er nicht noch einmal geklickt werden kann.
        '    'Me.butImportK.Enabled = False

        '    'Start in Sync schreiben
        '    intReturnValue = WFDBClass.FcWriteStartToSync(objdbConn,
        '                                                  cmbBuha.SelectedValue,
        '                                                  2,
        '                                                  objdtKreditorenHeadRead.Rows.Count)

        '    'intTeqNbr = Conversion.Val(Strings.Right(objdtInfo.Rows(1).Item(1), 3))

        '    'Kopfbuchung
        '    For Each row As DataRow In objdtKreditorenHead.Rows

        '        If IIf(IsDBNull(row("booKredBook")), False, row("booKredBook")) Then

        '            'Test ob OP - Buchung
        '            If row("intBuchungsart") = 1 Then

        '                'Immer zuerst Belegs-Nummerierung aktivieren, falls vorhanden externe Nummer = OP - Nr. Rg
        '                'Führt zu Problemen beim Ausbuchen des DTA - Files
        '                'Resultat Besprechnung 17.09.20 mit Muhi/ Andy
        '                'DbBhg.IncrBelNbr = "J"
        '                'Belegsnummer abholen
        '                'intDebBelegsNummer = DbBhg.GetNextBelNbr("R")

        '                'Auf Provisorisch setzen
        '                Call KrBhg.SetBuchMode("P")

        '                'Automatische ESR - Zahlungsverbindung
        '                KrBhg.EnableAutoESRZlgVerb = "J"

        '                'Eindeutigkeit der internen Beleg-Nummer setzen
        '                KrBhg.CheckDoubleIntBelNbr = "N"

        '                'Eindeutigkeit externer Beleg-Nummer setzen
        '                KrBhg.CheckDoubleExtBelNbr = "J"

        '                'If IsDBNull(row("strOPNr")) Or row("StrOPNr") = "" Then
        '                'strExtBelegNbr = row("strOPNr")

        '                'Zuerst Beleg-Nummerieungung aktivieren
        '                'KrBhg.IncrBelNbr = "N"
        '                'Belegsnummer abholen
        '                'intKredBelegsNummer = KrBhg.GetNextBelNbr("R")
        '                'Else
        '                'Beleg-Nummerierung abschalten
        '                'KrBhg.IncrBelNbr = "N"
        '                'intKredBelegsNummer = row("strOPNr")
        '                'strExtBelegNbr = row("strOPNr")
        '                'End If
        '                strExtKredBelegsNummer = row("strKredRGNbr")

        '                'Variablen zuweisen
        '                intKreditorNbr = row("lngKredNbr")
        '                If row("dblKredBrutto") < 0 Then
        '                    strBuchType = "G"
        '                    'strZahlSperren = "J"
        '                    row("dblKredBrutto") = row("dblKredBrutto") * -1
        '                    'Belegsnummer abholen
        '                    KrBhg.IncrBelNbr = "J"
        '                    intKredBelegsNummer = KrBhg.GetNextBelNbr("G")
        '                    KrBhg.IncrBelNbr = "N"

        '                    intReturnValue = MainKreditor.FcCheckKrediExistance(objdbMSSQLConn,
        '                                                                        objdbSQLcommand,
        '                                                                        intKredBelegsNummer,
        '                                                                        "G",
        '                                                                        intTeqNbr,
        '                                                                        intTeqNbrLY,
        '                                                                        intTeqNbrPLY,
        '                                                                        KrBhg)

        '                Else
        '                    strBuchType = "R"
        '                    'strZahlSperren = "N"
        '                    'Belegsnummer abholen
        '                    KrBhg.IncrBelNbr = "J"
        '                    intKredBelegsNummer = KrBhg.GetNextBelNbr("R")
        '                    'Muss auf Nicht hochzählen gesetzt werden da Sage 200 nicht merkt, dass Beleg-Nr. schon vergeben worden sind. => In den Einstellungen muss von Zeit zu Zeit der Zähler geändert werden
        '                    KrBhg.IncrBelNbr = "N"

        '                    intReturnValue = MainKreditor.FcCheckKrediExistance(objdbMSSQLConn,
        '                                                                        objdbSQLcommand,
        '                                                                        intKredBelegsNummer,
        '                                                                        "R",
        '                                                                        intTeqNbr,
        '                                                                        intTeqNbrLY,
        '                                                                        intTeqNbrPLY,
        '                                                                        KrBhg)

        '                End If

        '                strValutaDatum = Format(row("datKredValDatum"), "yyyyMMdd").ToString
        '                strBelegDatum = Format(row("datKredRGDatum"), "yyyyMMdd").ToString
        '                strVerfallDatum = String.Empty
        '                'strReferenz = IIf(IsDBNull(row("strKredRef")), "", row("strKredRef"))
        '                'If IsDBNull(row("strKrediBank")) Then
        '                'intTeilnehmer = 0
        '                'Else
        '                'Teilnehmer nur bei ESR setzen
        '                If row("intPayType") <> 9 Then 'nicht IBAN
        '                    'QR-Referenz
        '                    strReferenz = IIf(IsDBNull(row("strKredRef")), "", row("strKredRef"))
        '                    If row("intPayType") = 10 Then
        '                        strTeilnehmer = row("strKrediBank")
        '                    Else
        '                        strTeilnehmer = Val(row("strKrediBank"))
        '                    End If
        '                    intBankNbr = 0
        '                Else
        '                    'IBAN
        '                    strReferenz = IIf(IsDBNull(row("strKredRef")), "", row("strKredRef"))
        '                    intBankNbr = IIf(IsDBNull(row("intEBank")), 0, row("intEBank"))
        '                    strTeilnehmer = String.Empty
        '                End If
        '                'End If
        '                strMahnerlaubnis = String.Empty 'Format(row("datDebRGDatum"), "yyyyMMdd").ToString
        '                'Sachbearbeiter aus Debitor auslesen
        '                strDebiLine = KrBhg.ReadKreditor3(row("lngKredNbr") * -1, "")
        '                strDebitor = Split(strDebiLine, "{>}")
        '                strSachBID = strDebitor(29)

        '                dblBetrag = row("dblKredBrutto")
        '                strKrediText = IIf(IsDBNull(row("strKredText")), "", row("strKredText"))
        '                strCurrency = row("strKredCur")
        '                'intBankNbr = 0
        '                intKondition = IIf(IsDBNull(row("intZKond")), 1, row("intZKond"))
        '                'LN 0=automatsich ersterfasste Kondition, -1=Schlechteste Kondition, -2=Beste Kondition
        '                intKonditionLN = 0
        '                intEigeneBank = row("intintBank")

        '                If strCurrency <> "CHF" Then 'Muss ergänzt werden => Was ist Leitwährung auf dem Konto
        '                    dblKurs = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg)
        '                Else
        '                    dblKurs = 1.0#
        '                End If

        '                Try
        '                    booBookingok = True
        '                    Call KrBhg.SetBelegKopf2(intKredBelegsNummer,
        '                                             strValutaDatum,
        '                                             intKreditorNbr,
        '                                             strExtKredBelegsNummer,
        '                                             strBelegDatum,
        '                                             strVerfallDatum,
        '                                             strKrediText,
        '                                             intBankNbr,
        '                                             strBuchType,
        '                                             strZahlSperren,
        '                                             strVorausZahlung,
        '                                             intKondition,
        '                                             intKonditionLN,
        '                                             strSachBID,
        '                                             strReferenz,
        '                                             strSkonto,
        '                                             dblBetrag.ToString,
        '                                             strErfassungsArt,
        '                                             dblKurs.ToString,
        '                                             strCurrency,
        '                                             "",
        '                                             strTeilnehmer,
        '                                             intEigeneBank.ToString)

        '                Catch ex As Exception
        '                    MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Belegkopf ")
        '                    If (Err.Number And 65535) < 10000 Then
        '                        booBookingok = False
        '                    Else
        '                        booBookingok = True
        '                    End If

        '                End Try

        '                selKrediSub = objdtKreditorenSub.Select("lngKredID=" + row("lngKredID").ToString)

        '                For Each SubRow As DataRow In selKrediSub

        '                    intGegenKonto = SubRow("lngKto")
        '                    strFibuText = SubRow("strKredSubText")
        '                    'Soll auf Minus setzen
        '                    'If SubRow("intSollHaben") = 1 Then
        '                    'dblNettoBetrag = SubRow("dblNetto") * -1
        '                    'dblMwStBetrag = SubRow("dblMwSt") * -1
        '                    'dblBruttoBetrag = SubRow("dblBrutto") * -1
        '                    'Else
        '                    If intGegenKonto <> 6906 Then
        '                        If strBuchType = "R" Then
        '                            dblNettoBetrag = SubRow("dblNetto")
        '                            dblMwStBetrag = SubRow("dblMwSt")
        '                            dblBruttoBetrag = SubRow("dblBrutto")
        '                        Else
        '                            dblNettoBetrag = SubRow("dblNetto") * -1
        '                            dblMwStBetrag = SubRow("dblMwSt") * -1
        '                            dblBruttoBetrag = SubRow("dblBrutto") * -1
        '                        End If
        '                    Else 'Rundungsdifferenzen
        '                        If strBuchType = "R" Then
        '                            dblNettoBetrag = SubRow("dblBrutto")
        '                            dblMwStBetrag = SubRow("dblMwSt")
        '                            dblBruttoBetrag = SubRow("dblBrutto")
        '                        Else
        '                            dblNettoBetrag = SubRow("dblBrutto") * -1
        '                            dblMwStBetrag = SubRow("dblMwSt") * -1
        '                            dblBruttoBetrag = SubRow("dblBrutto") * -1
        '                        End If

        '                    End If

        '                    'If strBuchType = "R" Then
        '                    '    dblNettoBetrag = SubRow("dblNetto")
        '                    '    dblMwStBetrag = SubRow("dblMwSt")
        '                    '    dblBruttoBetrag = SubRow("dblBrutto")
        '                    'Else
        '                    '    dblNettoBetrag = SubRow("dblNetto") * -1
        '                    '    dblMwStBetrag = SubRow("dblMwSt") * -1
        '                    '    dblBruttoBetrag = SubRow("dblBrutto") * -1
        '                    'End If
        '                    'End If
        '                    'dblBebuBetrag = 1000.0#
        '                    If SubRow("lngKST") > 0 Then
        '                        strBeBuEintrag = SubRow("lngKST").ToString + "{<}" + SubRow("strKredSubText") + "{<}" + "CALCULATE" + "{>}"    '"PROD{<}BebuText{<}" + dblBebuBetrag.ToString + "{>}"
        '                    Else
        '                        'strBeBuEintrag = "00" + "{<}" + SubRow("strKredSubText") + "{<}" + "0" + "{>}"
        '                    End If
        '                    If Not IsDBNull(SubRow("strMwStKey")) And SubRow("strMwStKey") <> "null" Then ' And SubRow("strMwStKey") <> "25" Then
        '                        strSteuerFeld = Main.FcGetSteuerFeld(FBhg, SubRow("lngKto"), SubRow("strKredSubText"), dblBruttoBetrag, SubRow("strMwStKey"), dblMwStBetrag)     '"25{<}DEBI D Bhg Export MwSt{<}0{>}"
        '                    Else
        '                        strSteuerFeld = "STEUERFREI"
        '                    End If

        '                    'strSteuerInfo = Split(FBhg.GetKontoInfo(intGegenKonto.ToString), "{>}")
        '                    'Debug.Print("Konto-Info: " + strSteuerInfo(26))

        '                    Try
        '                        booBookingok = True
        '                        Call KrBhg.SetVerteilung(intGegenKonto.ToString, strFibuText, dblNettoBetrag.ToString, strSteuerFeld, strBeBuEintrag)

        '                    Catch ex As Exception
        '                        MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Verteilung " + SubRow("strKredSubText") + ", Konto " + SubRow("lngKto").ToString)
        '                        If (Err.Number And 65535) < 10000 Then
        '                            booBookingok = False
        '                        Else
        '                            booBookingok = True
        '                        End If

        '                    End Try


        '                    strSteuerFeld = String.Empty
        '                    strBeBuEintrag = String.Empty

        '                    'Status Sub schreiben
        '                    'Application.DoEvents()

        '                Next


        '                Try
        '                    booBookingok = True
        '                    Call KrBhg.WriteBuchung()

        '                Catch ex As Exception
        '                    If (Err.Number And 65535) < 10000 Then
        '                        MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Belegerstellung nicht möglich")
        '                        booBookingok = False
        '                    Else
        '                        MessageBox.Show(ex.Message, "Warnung " + (Err.Number And 65535).ToString + " Belegerstellung")
        '                        booBookingok = True
        '                    End If

        '                End Try

        '                'Application.DoEvents()

        '                strBeBuEintrag = String.Empty
        '                strSteuerFeld = String.Empty
        '                dblNettoBetrag = 0
        '                dblMwStBetrag = 0
        '                dblBruttoBetrag = 0

        '            Else

        '                'Buchung nur in Fibu
        '                'Prinzip Funktion WriteBuchung() anwenden mit allen Parametern
        '                'Beleg-Nummerierung aktivieren
        '                'DbBhg.IncrBelNbr = "J"
        '                'Belegsnummer abholen
        '                intKredBelegsNummer = FBhg.GetNextBelNbr()

        '                'Prüfen, ob wirklich frei
        '                intReturnValue = 10
        '                Do Until intReturnValue = 0
        '                    intReturnValue = FBhg.doesBelegExist(intKredBelegsNummer,
        '                                                         "NOT_SET",
        '                                                         "NOT_SET",
        '                                                         Strings.Left(cmbPerioden.SelectedItem, 4) + "0101",
        '                                                         Strings.Left(cmbPerioden.SelectedItem, 4) + "1231")
        '                    If intReturnValue <> 0 Then
        '                        intKredBelegsNummer += 1
        '                    End If
        '                Loop

        '                booBookingok = True

        '                'Variablen zuweisen
        '                strBelegDatum = Format(row("datKredRGDatum"), "yyyyMMdd").ToString
        '                strValutaDatum = Format(row("datKredValDatum"), "yyyyMMdd").ToString
        '                'strDebiText = row("strDebText")
        '                strCurrency = row("strKredCur")
        '                If strCurrency <> "CHF" Then 'Muss ergänzt werden => Was ist Leitwährung auf dem Konto
        '                    dblKurs = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg)
        '                Else
        '                    dblKurs = 1.0#
        '                End If

        '                selKrediSub = objdtKreditorenSub.Select("lngKredID=" + row("lngKredID").ToString)

        '                If selKrediSub.Length = 2 Then

        '                    For Each SubRow As DataRow In selKrediSub

        '                        If SubRow("intSollHaben") = 0 Then 'Soll

        '                            intSollKonto = SubRow("lngKto")
        '                            dblKursSoll = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg, intSollKonto)
        '                            'strSteuerInfo = Split(FBhg.GetKontoInfo(intSollKonto.ToString), "{>}")
        '                            'Debug.Print("Konto-Info Soll: " + strSteuerInfo(26))
        '                            dblSollBetrag = SubRow("dblNetto")
        '                            strKrediTextSoll = SubRow("strKredSubText")
        '                            If SubRow("dblMwSt") > 0 Then
        '                                strSteuerFeldSoll = Main.FcGetSteuerFeld(FBhg,
        '                                                                         SubRow("lngKto"),
        '                                                                         strKrediTextSoll,
        '                                                                         SubRow("dblBrutto") * dblKursSoll,
        '                                                                         SubRow("strMwStKey"),
        '                                                                         SubRow("dblMwSt"))
        '                            End If
        '                            If SubRow("lngKST") > 0 Then
        '                                strBeBuEintragSoll = SubRow("lngKST").ToString + "{<}" + strKrediTextSoll + "{<}" + "CALCULATE" + "{>}"
        '                            End If


        '                        ElseIf SubRow("intSollHaben") = 1 Then 'Haben

        '                            intHabenKonto = SubRow("lngKto")
        '                            dblKursHaben = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg, intHabenKonto)
        '                            'strSteuerInfo = Split(FBhg.GetKontoInfo(intSollKonto.ToString), "{>}")
        '                            'Debug.Print("Konto-Info Haben: " + strSteuerInfo(26))
        '                            dblHabenBetrag = SubRow("dblNetto") * -1
        '                            'dblHabenBetrag = dblSollBetrag
        '                            strKrediTextHaben = SubRow("strKredSubText")
        '                            If SubRow("dblMwSt") > 0 Then
        '                                strSteuerFeldHaben = Main.FcGetSteuerFeld(FBhg,
        '                                                                          SubRow("lngKto"),
        '                                                                          strKrediTextHaben,
        '                                                                          SubRow("dblBrutto") * dblKursHaben * -1,
        '                                                                          SubRow("strMwStKey"),
        '                                                                          SubRow("dblMwSt") * -1)
        '                            End If
        '                            If SubRow("lngKST") > 0 Then
        '                                strBeBuEintragHaben = SubRow("lngKST").ToString + "{<}" + strKrediTextHaben + "{<}" + "CALCULATE" + "{>}"
        '                            End If

        '                        Else

        '                            'Sammelbeleg
        '                            MsgBox("Nicht definierter Wert Sub-Buchungs-SollHaben: " + SubRow("intSollHaben").ToString)
        '                            'strKrediText = IIf(IsDBNull(row("strKredText")), "", row("strKredText"))

        '                        End If

        '                    Next

        '                    'Tieferer Betrag für die Gesamt-Buchung herausfinden
        '                    If dblSollBetrag <= dblHabenBetrag Then
        '                        dblNettoBetrag = dblSollBetrag
        '                    ElseIf dblHabenBetrag < dblSollBetrag Then
        '                        dblNettoBetrag = dblHabenBetrag
        '                    End If

        '                    'Buchung ausführen
        '                    Call FBhg.WriteBuchung(0,
        '                                           intKredBelegsNummer,
        '                                           strBelegDatum,
        '                                           intSollKonto.ToString,
        '                                           strKrediTextSoll,
        '                                           strCurrency,
        '                                           dblKursSoll.ToString,
        '                                           (dblNettoBetrag * dblKursSoll).ToString,
        '                                           strSteuerFeldSoll,
        '                                           intHabenKonto.ToString,
        '                                           strKrediTextHaben,
        '                                           strCurrency,
        '                                           dblKursHaben.ToString,
        '                                           (dblNettoBetrag * dblKursHaben).ToString,
        '                                           strSteuerFeldHaben,
        '                                           strCurrency,
        '                                           dblKurs.ToString,
        '                                           dblNettoBetrag.ToString,
        '                                           (dblNettoBetrag * dblKurs).ToString,
        '                                           strBeBuEintragSoll,
        '                                           strBeBuEintragHaben,
        '                                           strValutaDatum)

        '                Else
        '                    MsgBox("Nicht 2 Subbuchungen.")
        '                End If

        '            End If

        '            If booBookingok Then
        '                If row("booPGV") Then
        '                    'Bei PGV Buchungen
        '                    If IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")) = "" Or
        '                       (IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")) = "RV" And row("intPGVMthsAY") + row("intPGVMthsNY") > 1) Then

        '                        intReturnValue = MainKreditor.FcPGVKTreatment(FBhg,
        '                                                               Finanz,
        '                                                               DbBhg,
        '                                                               PIFin,
        '                                                               BeBu,
        '                                                               KrBhg,
        '                                                               objdtKreditorenSub,
        '                                                               row("lngKredID"),
        '                                                               intKredBelegsNummer,
        '                                                               row("strKredCur"),
        '                                                               row("datKredValDatum"),
        '                                                               "M",
        '                                                               row("datPGVFrom"),
        '                                                               row("datPGVTo"),
        '                                                               row("intPGVMthsAY") + row("intPGVMthsNY"),
        '                                                               row("intPGVMthsAY"),
        '                                                               row("intPGVMthsNY"),
        '                                                               2311,
        '                                                               2312,
        '                                                               cmbPerioden.SelectedItem,
        '                                                               objdbConn,
        '                                                               objdbMSSQLConn,
        '                                                               objdbSQLcommand,
        '                                                               cmbBuha.SelectedValue,
        '                                                               objdtInfo,
        '                                                               strYear,
        '                                                               intTeqNbr,
        '                                                               intTeqNbrLY,
        '                                                               intTeqNbrPLY,
        '                                                               IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")),
        '                                                               datPeriodFrom,
        '                                                               datPeriodTo,
        '                                                               strPeriodStatus)

        '                    Else

        '                        'TP
        '                        intReturnValue = MainKreditor.FcPGVKTreatmentYC(FBhg,
        '                                                               Finanz,
        '                                                               DbBhg,
        '                                                               PIFin,
        '                                                               BeBu,
        '                                                               KrBhg,
        '                                                               objdtKreditorenSub,
        '                                                               row("lngKredID"),
        '                                                               intKredBelegsNummer,
        '                                                               row("strKredCur"),
        '                                                               row("datKredValDatum"),
        '                                                               "M",
        '                                                               row("datPGVFrom"),
        '                                                               row("datPGVTo"),
        '                                                               row("intPGVMthsAY") + row("intPGVMthsNY"),
        '                                                               row("intPGVMthsAY"),
        '                                                               row("intPGVMthsNY"),
        '                                                               2311,
        '                                                               2312,
        '                                                               cmbPerioden.SelectedItem,
        '                                                               objdbConn,
        '                                                               objdbMSSQLConn,
        '                                                               objdbSQLcommand,
        '                                                               cmbBuha.SelectedValue,
        '                                                               objdtInfo,
        '                                                               strYear,
        '                                                               intTeqNbr,
        '                                                               intTeqNbrLY,
        '                                                               intTeqNbrPLY,
        '                                                               IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")),
        '                                                               datPeriodFrom,
        '                                                               datPeriodTo,
        '                                                               strPeriodStatus)


        '                    End If


        '                End If

        '                'Status Head schreiben
        '                row("strKredBookStatus") = row("strKredStatusBitLog")
        '                row("booBooked") = True
        '                row("datBooked") = Now()
        '                row("lngBelegNr") = intKredBelegsNummer
        '                'strKRGReferTo = "lngKredID"
        '                'strKRGReferTo = "strKredRGNbr"
        '                If objdbConn.State = ConnectionState.Closed Then
        '                    objdbConn.Open()
        '                End If
        '                strKRGReferTo = Main.FcReadFromSettings(objdbConn, "Buchh_TableKRGReferTo", cmbBuha.SelectedValue)
        '                If objdbConn.State = ConnectionState.Open Then
        '                    objdbConn.Close()
        '                End If
        '                'Status in File RG-Tabelle schreiben
        '                intReturnValue = MainKreditor.FcWriteToKrediRGTable(cmbBuha.SelectedValue,
        '                                                                row(strKRGReferTo),
        '                                                                row("datBooked"),
        '                                                                row("lngBelegNr"),
        '                                                                objdbAccessConn,
        '                                                                objOracleConn,
        '                                                                objdbConn)
        '                If intReturnValue <> 0 Then
        '                    'Throw an exception
        '                    MessageBox.Show("Achtung, Beleg-Nummer: " + row("lngBelegNr").ToString + " konnte nicht In die RG-Tabelle geschrieben werden auf RG-ID: " + row("lngKredID").ToString + ".", "RG-Table Update nicht möglich", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        '                End If

        '            End If

        '        End If

        '    Next

        '    'In sync-Tabelle schreiben
        '    intReturnValue = WFDBClass.FcWriteEndToSync(objdbConn,
        '                                                cmbBuha.SelectedValue,
        '                                                2,
        '                                                0,
        '                                                IIf(booBookingok, "ok", "Probleme"))


        'Catch ex As Exception
        '    MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " bei Buchung-Beleg " + strExtKredBelegsNummer)

        'Finally
        '    'Neu aufbauen
        '    'butKreditoren_Click(butDebitoren, EventArgs.Empty)

        '    Me.Cursor = Cursors.Default
        '    'Me.butImportK.Enabled = True
        '    Application.Restart()

        'End Try


    End Sub


    Private Sub butMail_Click(sender As Object, e As EventArgs)

        Dim strMailText As String
        Dim intColCounter As Int16
        Dim intRowCounter As Int32
        Dim intAffected As Int16
        Dim strMailTo As String

        Try

            'Me.Cursor = Cursors.WaitCursor

            ''String zusammensetzen für Mailtext
            'strMailText = "<html>" + vbCrLf
            'strMailText += "<head>" + vbCrLf
            'strMailText += "<style>" + vbCrLf
            'strMailText += "   table, th, td {border:   1px solid; border-collapse: collapse;}" + vbCrLf
            'strMailText += "</style>" + vbCrLf
            'strMailText += "</head>" + vbCrLf
            'strMailText += "<body>" + vbCrLf
            'strMailText += "<table border=""1px solid black"">" + vbCrLf
            'strMailText += "    <thead>" + vbCrLf
            'strMailText += "    <caption><h1> " + cmbBuha.Text + " Debitoren</h1></caption>"
            'strMailText += "    <tr>" + vbCrLf

            'intColCounter = 0

            ''Zuerst Titel zusammen setzen
            'For intColCounter = 0 To dgvBookingsO.Columns.Count - 1
            '    'Ausblenden intBuchhaltung, ok
            '    'Debug.Print("Spalte " + intColCounter.ToString + ", " + dgvBookings.Columns(intColCounter).HeaderText)
            '    If intColCounter = 0 Or intColCounter = 11 Or
            '    intColCounter = 12 Or intColCounter = 10 Or
            '    intColCounter = 15 Or intColCounter = 16 Or
            '    intColCounter = 17 Or intColCounter = 18 Or
            '    intColCounter = 20 Or
            '    intColCounter = 21 Or intColCounter = 23 Or
            '    intColCounter = 26 Or intColCounter = 27 Or
            '    intColCounter = 39 Then
            '        strMailText += "        <th>" + dgvBookingsO.Columns(intColCounter).HeaderText + "</th>" + vbCrLf
            '    End If
            'Next
            'strMailText += "        <th style=""background-color:#FF0000"">" + dgvBookingsO.Columns(34).HeaderText + "</th>" + vbCrLf
            'strMailText += "    </tr>" + vbCrLf
            'strMailText += "    </thead>" + vbCrLf


            ''Durch die Tabelle steppen
            'strMailText += "    <tbody>" + vbCrLf
            'For intRowCounter = 0 To dgvBookingsO.Rows.Count - 1
            '    strMailText += "    <tr>" + vbCrLf
            '    For intColCounter = 0 To dgvBookingsO.Columns.Count - 1
            '        If intColCounter = 0 Or intColCounter = 11 Or
            '    intColCounter = 12 Or intColCounter = 10 Or
            '    intColCounter = 15 Or intColCounter = 16 Or
            '    intColCounter = 17 Or intColCounter = 18 Or
            '    intColCounter = 20 Or
            '    intColCounter = 21 Or intColCounter = 23 Or
            '    intColCounter = 26 Or intColCounter = 27 Or
            '    intColCounter = 39 Then
            '            strMailText += "        <td>" + dgvBookingsO.Rows(intRowCounter).Cells(intColCounter).Value.ToString + "</td>" + vbCrLf
            '        End If
            '    Next
            '    strMailText += "        <td style=""background-color:#FF0000"">" + dgvBookingsO.Rows(intRowCounter).Cells(34).Value.ToString + "</td>" + vbCrLf
            '    strMailText += "    </tr>" + vbCrLf
            'Next
            'strMailText += "    </tbody>" + vbCrLf

            ''Footer mit Legende
            'strMailText += "    <tfoot>" + vbCrLf
            'strMailText += "    <tr>" + vbCrLf
            'strMailText += "    <td colspan=""6"" style=""background-color:#FFA07A""><b>FEHLERMELDUNG</b></td>"
            'strMailText += "    <td colspan=""6"" style=""background-color:#8FBC8F""><b>LOESUNG</b></td>"
            'strMailText += "    </tr>" + vbCrLf
            'strMailText += "    <tr>" + vbCrLf
            'strMailText += "    <td colspan=""6"" style=""background-color:#FFA07A"">ValD = Valuta-Datum nicht möglich</td>"
            'strMailText += "    <td colspan=""6"" style=""background-color:#8FBC8F"">Valuta Datum auf aktuelle Periode anpassen</td>"
            'strMailText += "    </tr>" + vbCrLf
            'strMailText += "    <tr>" + vbCrLf
            'strMailText += "    <td colspan=""6"" style=""background-color:#FFA07A"">RgD = Rechnungs-Datum nicht möglich</td>"
            'strMailText += "    </tr>" + vbCrLf
            'strMailText += "    <tr>" + vbCrLf
            'strMailText += "    <td colspan=""6"" style=""background-color:#FFA07A"">Sub = Problem in den Buchungen</td>"
            'strMailText += "    </tr>" + vbCrLf
            'strMailText += "    <tr>" + vbCrLf
            'strMailText += "    <td colspan=""6"" style=""background-color:#FFA07A"">Rnd>1 = Rundungsbetrag zu gross</td>"
            'strMailText += "    </tr>" + vbCrLf
            'strMailText += "    <tr>" + vbCrLf
            'strMailText += "    <td colspan=""6"" style=""background-color:#FFA07A"">SplB = Verlinkte RG nicht gefunden</td>"
            'strMailText += "    </tr>" + vbCrLf
            'strMailText += "    <tr>" + vbCrLf
            'strMailText += "    <td colspan=""6"" style=""background-color:#FFA07A"">OPDbl = OP existiert schon</td>"
            'strMailText += "    </tr>" + vbCrLf
            'strMailText += "    <tr>" + vbCrLf
            'strMailText += "    <td colspan=""6"" style=""background-color:#FFA07A"">Deb keine Ref = Der Debitor ist nicht spezifiziert</td>"
            'strMailText += "    </tr>" + vbCrLf
            'strMailText += "    </tfoot>" + vbCrLf

            'strMailText += "</table>" + vbCrLf
            'strMailText += "</body>" + vbCrLf
            'strMailText += "</html>"

            'Debug.Print(strMailText)

            ''In Mail Tabelle schreiben
            'If objdbConn.State <> ConnectionState.Open Then
            '    objdbConn.Open()
            'End If
            'strMailTo = Main.FcReadFromSettings(objdbConn, "Buchh_ErrDMailTo", cmbBuha.SelectedValue)
            ''objdbConn.Close()
            'objdbMySQLMail.Open()
            'objdbcommand.Connection = objdbMySQLMail
            'objdbcommand.CommandText = "INSERT INTO t_mailout (MailBoxID, Totext, SubjectText, " +
            '                                               "BodyFormat, BodyText, Importance, " +
            '                                               "InterfaceID, SentStatusID, CreatedBy, " +
            '                                               "CreatedByTool) VALUES (" +
            '                                                "7, '" + strMailTo + "', 'Import Sage 200 - Debitoren', " +
            '                                                "1, '" + strMailText + "', 1, " +
            '                                                "3, 1, 'Sage200I', " +
            '                                                "'Sage200 - Importer')"
            'intAffected = objdbcommand.ExecuteNonQuery()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem " + (Err.Number.ToString))

        Finally
            'objdbConn.Close()
            'objdbMySQLMail.Close()
            'Me.Cursor = Cursors.Default

        End Try


    End Sub

    Private Sub butDblDebis_Click(sender As Object, e As EventArgs) Handles butDblDebis.Click

        Dim intCheckDblDebis As Int16
        Dim tblDebiBelege As New DataTable
        Dim tblDebiSearch As New DataTable
        Dim intteqnbr As Int32

        Application.Restart()

        'Überprüfung ob doppelte Debitoren existieren
        'Gleiche exterene Beleg-Nr., gleicher Betrag, gleiches Belegdatum

        'Try

        '    intCheckDblDebis = MessageBox.Show("Soll wirklich eine Überprüfung auf doppelte Debitoren - Belege erfolgen?", "Doppelte Debi-Belege", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        '    If intCheckDblDebis = vbYes Then

        '        Me.Cursor = Cursors.WaitCursor

        '        'durch die Debitoren - Belege steppen
        '        objdbMSSQLConn.Open()
        '        intteqnbr = Conversion.Val(Strings.Right(objdtInfo.Rows(1).Item(1), 3))

        '        'Zuerst nach Rechnungen suchen
        '        objdbSQLcommand.CommandText = "SELECT * FROM debibuchung WHERE teqnbr=" + intteqnbr.ToString + " AND typ='R' AND NOT belnr IS NULL ORDER BY belnr"
        '        objdbSQLcommand.Connection = objdbMSSQLConn
        '        tblDebiBelege.Load(objdbSQLcommand.ExecuteReader)
        '        For Each drdebibelege In tblDebiBelege.Rows
        '            'Gibt es mehr als einen Beleg mit gleichem Betrag und gleicher externer Beleg-Nr.?
        '            tblDebiSearch.Rows.Clear()
        '            objdbSQLcommand.CommandText = "SELECT COUNT(belnr) FROM debibuchung WHERE teqnbr=" + intteqnbr.ToString + " AND typ='R' AND belnr='" + drdebibelege.item("belnr") + "' AND skontobetrag=" + drdebibelege.item("skontobetrag").ToString
        '            tblDebiSearch.Load(objdbSQLcommand.ExecuteReader)
        '            If tblDebiSearch.Rows(0).Item(0) > 1 Then
        '                MessageBox.Show("Mögliche Verdopplung gefunden 'R' belnr " + drdebibelege.item("belnr") + vbCrLf + "Betrag " + drdebibelege.item("skontobetrag").ToString, "Verdopplung gefunden", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        '            End If
        '        Next

        '        'Dann nach Gutschriften
        '        tblDebiBelege.Dispose()
        '        objdbSQLcommand.CommandText = "SELECT * FROM debibuchung WHERE teqnbr=" + intteqnbr.ToString + " AND typ='G' AND NOT belnr IS NULL ORDER BY belnr"
        '        objdbSQLcommand.Connection = objdbMSSQLConn
        '        tblDebiBelege.Load(objdbSQLcommand.ExecuteReader)
        '        For Each drdebibelege In tblDebiBelege.Rows
        '            'Gibt es mehr als einen Beleg mit gleichem Betrag und gleicher externer Beleg-Nr.?
        '            tblDebiSearch.Rows.Clear()
        '            objdbSQLcommand.CommandText = "SELECT COUNT(belnr) FROM debibuchung WHERE teqnbr=" + intteqnbr.ToString + " AND typ='G' AND belnr='" + drdebibelege.item("belnr") + "' AND skontobetrag=" + drdebibelege.item("skontobetrag").ToString
        '            tblDebiSearch.Load(objdbSQLcommand.ExecuteReader)
        '            If tblDebiSearch.Rows(0).Item(0) > 1 Then
        '                MessageBox.Show("Mögliche Verdopplung gefunden 'G' belnr " + drdebibelege.item("belnr") + vbCrLf + "Betrag " + drdebibelege.item("skontobetrag").ToString, "Verdopplung gefunden", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        '            End If
        '        Next

        '        MessageBox.Show("Suche beendet.", "Suche beendet", MessageBoxButtons.OK, MessageBoxIcon.Information)

        '    End If

        'Catch ex As Exception
        '    MessageBox.Show(ex.Message, "Problem " + (Err.Number.ToString))

        'Finally
        '    objdbMSSQLConn.Close()
        '    tblDebiBelege.Dispose()
        '    tblDebiSearch.Dispose()
        '    Me.Cursor = Cursors.Default

        'End Try

    End Sub

    Private Sub butDblKredis_Click(sender As Object, e As EventArgs) Handles butDblKredis.Click

        Dim intCheckDblKredis As Int16
        Dim tblKrediBelege As New DataTable
        Dim tblKrediSearch As New DataTable
        'Dim intteqnbr As Int32

        'Überprüfung ob doppelte Kreditoren existieren
        'Gleiche exterene Beleg-Nr.

        Try

            '    intCheckDblKredis = MessageBox.Show("Soll wirklich eine Überprüfung auf doppelte Kreditoren - Belege erfolgen?", "Doppelte Kredi-Belege", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            '    If intCheckDblKredis = vbYes Then

            '        Me.Cursor = Cursors.WaitCursor

            '        'durch die Debitoren - Belege steppen
            '        objdbMSSQLConn.Open()
            '        'intteqnbr = Conversion.Val(Strings.Right(objdtInfo.Rows(1).Item(1), 3))

            '        'Zuerst nach Rechnungen suchen Sortiert nach Beleg-Nr.
            '        'objdbSQLcommand.CommandText = "SELECT * FROM kredibuchung WHERE teqnbr IN (" + intTeqNbr.ToString + ", " + intTeqNbrLY.ToString + ", " + intTeqNbrPLY.ToString + ") AND typ='R' AND NOT belnr IS NULL ORDER BY belnr"
            '        'objdbSQLcommand.Connection = objdbMSSQLConn
            '        'tblKrediBelege.Load(objdbSQLcommand.ExecuteReader)
            '        'For Each drKredibelege In tblKrediBelege.Rows
            '        '    'Gibt es mehr als einen Beleg mit gleichem Betrag und gleicher externer Beleg-Nr.?
            '        '    tblKrediSearch.Rows.Clear()
            '        '    'Debug.Print("Suche " + Replace(drKredibelege.item("belnr"), "'", "''"))
            '        '    objdbSQLcommand.CommandText = "SELECT COUNT(belnr) FROM kredibuchung WHERE teqnbr IN (" + intTeqNbr.ToString + ", " + intTeqNbrLY.ToString + ", " + intTeqNbrPLY.ToString + ") AND typ='R' AND belnr='" + Replace(drKredibelege.item("belnr"), "'", "''") + "'" ' AND skontobetrag=" + drdebibelege.item("skontobetrag").ToString
            '        '    tblKrediSearch.Load(objdbSQLcommand.ExecuteReader)
            '        '    If tblKrediSearch.Rows(0).Item(0) > 1 Then
            '        '        MessageBox.Show("Mögliche Verdopplung gefunden 'R' belnr " + drKredibelege.item("belnr"), "Verdopplung gefunden", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            '        '        Debug.Print("Verdopplung " + Replace(drKredibelege.item("belnr"), "'", "''"))
            '        '    End If
            '        'Next

            '        'Nach Betrag suchen
            '        'Zuerst Tabelle mit den Summen bilden
            '        objdbSQLcommand.CommandText = "SELECT lfnbrk, " +
            '                                            "MAX(belnbrint) AS belnbrint, " +
            '                                            "MAX(belnr) AS belnbrext, " +
            '                                            "MAX(beltext) As beltext, " +
            '                                            "SUM(buchung.betrag) AS sumBetrag, " +
            '                                            "COUNT(buchung.betrag) AS nbrSubs " +
            '                                      "FROM kredibuchung " +
            '                                      "INNER JOIN buchung ON kredibuchung.lfnbrk = buchung.dklfnbr " +
            '                                      "WHERE kredibuchung.teqnbr IN (" + intTeqNbr.ToString + ", " + intTeqNbrLY.ToString + ", " + intTeqNbrPLY.ToString + ") " +
            '                                            "And kredibuchung.typ = 'R' " +
            '                                            "AND buchung.ktotyp = 'HK' " +
            '                                      "GROUP BY kredibuchung.lfnbrk"

            '        objdbSQLcommand.Connection = objdbMSSQLConn
            '        tblKrediBelege.Load(objdbSQLcommand.ExecuteReader)

            '        'Spalte hinzufügen für Kombi-Feld Text / Betrag
            '        tblKrediBelege.Columns.Add("combinedTA", GetType(String))
            '        'Durch die Tabelle gehen und das Feld abfüllen
            '        For Each drKBelege As DataRow In tblKrediBelege.Rows
            '            drKBelege.Item("combinedTA") = drKBelege.Item("beltext") + drKBelege.Item("sumBetrag").ToString
            '            'Debug.Print("combined: " + drKBelege.Item("combinedTA"))
            '        Next

            '        'Gruppiren nach diesem Feld
            '        'Zuerst Tabelle in DB löschen
            '        objdbcommandZHDB02.Connection = objdbConnZHDB02
            '        objdbcommandZHDB02.CommandText = "DELETE FROM tblKrediBelegeDbl"
            '        objdbConnZHDB02.Open()
            '        objdbcommandZHDB02.ExecuteNonQuery()

            '        'Tabelle erstellen und befüllen
            '        'objdbcommandZHDB02.CommandText = "CREATE TABLE tblKrediBelegeDbl(lfnbrk int, " +
            '        '                                                        "beltext text, " +
            '        '                                                        "sumBetrag double, " +
            '        '                                                        "nbrSubs int)"
            '        objdbcommandZHDB02.ExecuteNonQuery()
            '        For Each drkBelege In tblKrediBelege.Rows
            '            objdbcommandZHDB02.CommandText = "INSERT INTO tblKrediBelegeDbl VALUES(" + drkBelege.item("lfnbrk").ToString + ", " +
            '                                                                                       drkBelege.item("belnbrint").ToString + ", '" +
            '                                                                                       drkBelege.item("belnbrext") + "', '" +
            '                                                                                       drkBelege.item("combinedTA") + "', " +
            '                                                                                       drkBelege.item("sumBetrag").ToString + ", " +
            '                                                                                       drkBelege.item("nbrSubs").ToString + ")"
            '            objdbcommandZHDB02.ExecuteNonQuery()
            '        Next
            '        objdbConnZHDB02.Close()




            '        For Each drKredibelege In tblKrediBelege.Rows
            '            'Gibt es mehr als einen Beleg mit gleichem Betrag und gleicher externer Beleg-Nr.?
            '            tblKrediSearch.Constraints.Clear()
            '            tblKrediSearch.Rows.Clear()
            '            'Debug.Print("Suche " + Replace(drKredibelege.item("belnr"), "'", "''"))
            '            objdbSQLcommand.CommandText = "SELECT COUNT(betrag) FROM kredibuchung INNER JOIN buchung ON lfnbrk=dklfnbr " +
            '                                            " WHERE kredibuchung.teqnbr IN (" + intTeqNbr.ToString + ", " + intTeqNbrLY.ToString + ", " + intTeqNbrPLY.ToString + ") AND typ='R' AND NOT belnr IS NULL AND betrag=" + drKredibelege.item("betrag").ToString  ' AND skontobetrag=" + drdebibelege.item("skontobetrag").ToString
            '            tblKrediSearch.Load(objdbSQLcommand.ExecuteReader)
            '            If tblKrediSearch.Rows(0).Item(0) > 1 Then
            '                MessageBox.Show("Mögliche Verdopplung gefunden 'R' betrag " + drKredibelege.Item("betrag").ToString + " auf Beleg E: " + drKredibelege.item("belnr") + ", I: " + drKredibelege.item("belnbrint").ToString + ", K: " + drKredibelege.item("kredinbr").ToString, "Verdopplung gefunden", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            '                Debug.Print("Verdopplung " + Replace(drKredibelege.item("belnr"), "'", "''"))
            '            End If
            '        Next

            '        'Dann nach Gutschriften
            '        tblKrediBelege.Dispose()
            '        objdbSQLcommand.CommandText = "SELECT * FROM kredibuchung WHERE teqnbr IN (" + intTeqNbr.ToString + ", " + intTeqNbrLY.ToString + ", " + intTeqNbrPLY.ToString + ") AND typ='G' AND NOT belnr IS NULL ORDER BY belnr"
            '        objdbSQLcommand.Connection = objdbMSSQLConn
            '        tblKrediBelege.Load(objdbSQLcommand.ExecuteReader)
            '        For Each drKredibelege In tblKrediBelege.Rows
            '            'Gibt es mehr als einen Beleg mit gleichem Betrag und gleicher externer Beleg-Nr.?
            '            tblKrediSearch.Rows.Clear()
            '            objdbSQLcommand.CommandText = "SELECT COUNT(belnr) FROM kredibuchung WHERE teqnbr IN (" + intTeqNbr.ToString + ", " + intTeqNbrLY.ToString + ", " + intTeqNbrPLY.ToString + ") AND typ='G' AND belnr='" + Replace(drKredibelege.item("belnr"), "'", "''") + "'" ' AND skontobetrag=" + drdebibelege.item("skontobetrag").ToString
            '            tblKrediSearch.Load(objdbSQLcommand.ExecuteReader)
            '            If tblKrediSearch.Rows(0).Item(0) > 1 Then
            '                MessageBox.Show("Mögliche Verdopplung gefunden 'G' belnr " + drKredibelege.item("belnr"), "Verdopplung gefunden", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            '                Debug.Print("Verdopplung " + Replace(drKredibelege.item("belnr"), "'", "''"))
            '            End If
            '        Next

            '        MessageBox.Show("Suche beendet.", "Suche beendet", MessageBoxButtons.OK, MessageBoxIcon.Information)

            '    End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem " + (Err.Number.ToString))

        Finally
            '    objdbMSSQLConn.Close()
            '    tblKrediBelege.Dispose()
            '    tblKrediSearch.Dispose()
            '    Me.Cursor = Cursors.Default

        End Try

    End Sub

    Private Sub cmbPerioden_SelectionChangeCommitted(sender As Object, e As EventArgs)

        Try

            'dsDebitoren.Tables("tblDebitorenInfo").Rows.Clear()
            'Me.Refresh()

            'Call Main.FcLoginSage2(objdbConn,
            '                  objdbMSSQLConn,
            '                  objdbSQLcommand,
            '                  Finanz,
            '                  FBhg,
            '                  DbBhg,
            '                  PIFin,
            '                  BeBu,
            '                  KrBhg,
            '                  cmbBuha.SelectedValue,
            '                  dsDebitoren.Tables("tblDebitorenInfo"),
            '                  cmbPerioden.SelectedItem,
            '                  strYear,
            '                  intTeqNbr,
            '                  intTeqNbrLY,
            '                  intTeqNbrPLY,
            '                  datPeriodFrom,
            '                  datPeriodTo,
            '                  strPeriodStatus)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem " + (Err.Number.ToString))

        Finally

        End Try


    End Sub

    Private Sub chkValutaCorrect_CheckedChanged(sender As Object, e As EventArgs) Handles chkValutaCorrect.CheckedChanged

        If chkValutaCorrect.Checked Then
            dtpValutaCorrect.Enabled = True
        Else
            dtpValutaCorrect.Enabled = False
        End If

    End Sub

    Private Sub frmImportMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        'DS in tabelle löschen
        MySQLdaDebitoren.DeleteCommand.Connection.Open()
        MySQLdaDebitoren.DeleteCommand.ExecuteNonQuery()
        MySQLdaDebitoren.DeleteCommand.Connection.Close()

        MySQLdaDebitorenSub.DeleteCommand.Connection.Open()
        MySQLdaDebitorenSub.DeleteCommand.ExecuteNonQuery()
        MySQLdaDebitorenSub.DeleteCommand.Connection.Close()


    End Sub


    Private Sub lstBoxMandant_MouseClick(sender As Object, e As MouseEventArgs) Handles lstBoxMandant.MouseClick

        Try


            Call Main.FcReadPeriodsFromMandantLst(objdbConn,
                                           Finanz,
                                           lstBoxMandant.SelectedValue,
                                           lstBoxPerioden)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem " + Err.Number.ToString)

        End Try

    End Sub
End Class
