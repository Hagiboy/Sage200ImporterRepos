Option Strict Off
Option Explicit On

Imports MySql.Data.MySqlClient
Imports System.Data.OracleClient
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic.ApplicationServices
'Imports CLClassSage200.WFSage200Import
Imports System.IO

Public Class frmDebDisp

    'Dim Finanz As SBSXASLib.AXFinanz
    'Dim FBhg As SBSXASLib.AXiFBhg
    'Dim DbBhg As SBSXASLib.AXiDbBhg
    'Dim KrBhg As SBSXASLib.AXiKrBhg
    'Dim BsExt As SBSXASLib.AXiBSExt
    'Dim Adr As SBSXASLib.AXiAdr
    'Dim BeBu As SBSXASLib.AXiBeBu
    'Dim PIFin As SBSXASLib.AXiPlFin

    Dim FELD_SEP As String
    Dim REC_SEP As String
    Dim KSTKTR_SEP As String
    Dim FELD_SEP_OUT As String
    Dim REC_SEP_OUT As String
    Dim nID As String

    Dim intTeqNbr As Int32
    Dim intTeqNbrLY As Int32
    Dim intTeqNbrPLY As Int32
    Dim strYear As String
    Dim datPeriodFrom As Date
    Dim datPeriodTo As Date
    Dim strPeriodStatus As String

    'Dim objdbConn As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString"))
    'Dim objdbMSSQLConn As New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("SQLConnectionString"))
    'Dim objdbSQLcommand As New SqlCommand
    'Dim objdbAccessConn As New OleDb.OleDbConnection
    'Dim objOracleConn As New OracleConnection("Data Source=(DESCRIPTION=" _
    '                + "(ADDRESS=(PROTOCOL=TCP)(HOST=10.0.0.29)(PORT=1521))" _
    '                + "(CONNECT_DATA=(SERVICE_NAME=CISNEW)));" _
    '                + "User Id=cis;Password=sugus;")

    Public Class BgWCheckDebitArgs
        Public intMandant As Int16
        Public strMandant As String
        Public intTeqNbr As Int16
        Public intTeqNbrLY As Int16
        Public intTeqNbrPLY As Int16
        Public strYear As String
        Public strPeriode As String
        Public booValutaCor As Boolean
        Public datValutaCor As Date
    End Class


    Public Sub InitDB()

        Dim strIdentityName As String

        Try

            strIdentityName = System.Security.Principal.WindowsIdentity.GetCurrent().Name
            strIdentityName = Strings.Replace(strIdentityName, "\", "/")

            frmImportMain.LblIdentity.Text = strIdentityName
            frmImportMain.LblTaskID.Text = Process.GetCurrentProcess().Id.ToString

            mysqlconn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString")

            'Zuerst alle evtl. vorhandene DS des Users löschen
            mysqlcmdDebDel.CommandText = "DELETE FROM tbldebitorenjhead WHERE IdentityName='" + strIdentityName + "'"
            mysqlcmdDebDel.Connection.Open()
            mysqlcmdDebDel.ExecuteNonQuery()
            mysqlcmdDebDel.Connection.Close()

            mysqlcmdDebSubDel.CommandText = "DELETE FROM tbldebitorensub WHERE IdentityName='" + strIdentityName + "'"
            mysqlcmdDebSubDel.Connection.Open()
            mysqlcmdDebSubDel.ExecuteNonQuery()
            mysqlcmdDebSubDel.Connection.Close()

            'setzen für weiteren Gebraucht mit Process ID
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

    Private Sub frmDebDisp_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        FELD_SEP = "{<}"
        REC_SEP = "{>}"
        KSTKTR_SEP = "{-}"

        FELD_SEP_OUT = "{>}"
        REC_SEP_OUT = "{<}"

        Call InitDB()

    End Sub

    Friend Function FcDebiDisplay(intMandant As Int32,
                                  LstMandnat As ListBox,
                                  LstBPerioden As ListBox) As Int16

        Dim objdbConn As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString"))
        Dim objdbtaskcmd As New MySqlCommand
        Dim objdbMSSQLConn As New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("SQLConnectionString"))
        Dim objdbSQLcommand As New SqlCommand

        Dim intFcReturns As Int16
        Dim strPeriode As String
        Dim strYearCh As String
        Dim BgWCheckDebitLocArgs As New BgWCheckDebitArgs
        Dim objdbtasks As New DataTable

        'Dim intTeqNbr As Int32
        'Dim intTeqNbrLY As Int32
        'Dim intTeqNbrPLY As Int32
        'Dim strYear As String

        Dim objFinanz As New SBSXASLib.AXFinanz
        Dim objfiBuha As New SBSXASLib.AXiFBhg
        Dim objdbBuha As New SBSXASLib.AXiDbBhg
        Dim objdbPIFb As New SBSXASLib.AXiPlFin
        Dim objFiBebu As New SBSXASLib.AXiBeBu
        Dim objKrBuha As New SBSXASLib.AXiKrBhg


        Try

            Me.Cursor = Cursors.WaitCursor
            'Zuerst in tblImportTasks setzen
            objdbtaskcmd.Connection = objdbConn
            objdbtaskcmd.Connection.Open()
            objdbtaskcmd.CommandText = "SELECT * FROM tblimporttasks WHERE IdentityName='" + frmImportMain.LblIdentity.Text + "' AND Type='D'"
            objdbtasks.Load(objdbtaskcmd.ExecuteReader())
            If objdbtasks.Rows.Count > 0 Then
                'update
                objdbtaskcmd.CommandText = "UPDATE tblimporttasks SET Mandant=" + Convert.ToString(LstMandnat.SelectedIndex) + ", Periode=" + Convert.ToString(LstBPerioden.SelectedIndex) + " WHERE IdentityName='" + frmImportMain.LblIdentity.Text + "' AND Type='D'"
            Else
                'insert
                objdbtaskcmd.CommandText = "INSERT INTO tblimporttasks (IdentityName, Type, Mandant, Periode) VALUES ('" + frmImportMain.LblIdentity.Text + "', 'D', " + Convert.ToString(LstMandnat.SelectedIndex) + ", " + Convert.ToString(LstBPerioden.SelectedIndex) + ")"
            End If
            objdbtaskcmd.ExecuteNonQuery()
            objdbtaskcmd.Connection.Close()

            'intMode = 0

            Me.butImport.Enabled = False

            'DGV Debitoren
            dgvBookings.DataSource = Nothing
            dgvBookingSub.DataSource = Nothing

            'dsDebitoren.Reset()
            'dsDebitoren.Clear()

            'Zuerst evtl. vorhandene DS löschen in Tabelle
            MySQLdaDebitoren.DeleteCommand.Connection.Open()
            MySQLdaDebitoren.DeleteCommand.ExecuteNonQuery()
            MySQLdaDebitoren.DeleteCommand.Connection.Close()

            MySQLdaDebitorenSub.DeleteCommand.Connection.Open()
            MySQLdaDebitorenSub.DeleteCommand.ExecuteNonQuery()
            MySQLdaDebitorenSub.DeleteCommand.Connection.Close()

            'Info neu erstellen
            dsDebitoren.Tables.Add("tblDebitorenInfo")
            Dim col1 As DataColumn = New DataColumn("strInfoT")
            col1.DataType = System.Type.GetType("System.String")
            col1.MaxLength = 50
            col1.Caption = "Info-Titel"
            dsDebitoren.Tables("tblDebitorenInfo").Columns.Add(col1)
            Dim col2 As DataColumn = New DataColumn("strInfoV")
            col2.DataType = System.Type.GetType("System.String")
            col2.MaxLength = 50
            col2.Caption = "Info-Wert"
            dsDebitoren.Tables("tblDebitorenInfo").Columns.Add(col2)

            dgvInfo.DataSource = dsDebitoren.Tables("tblDebitorenInfo")

            'Datums-Tabelle erstellen
            dsDebitoren.Tables.Add("tblDebitorenDates")
            Dim col7 As DataColumn = New DataColumn("intYear")
            col7.DataType = System.Type.GetType("System.Int16")
            col7.Caption = "Year"
            dsDebitoren.Tables("tblDebitorenDates").Columns.Add(col7)
            Dim col3 As DataColumn = New DataColumn("strDatType")
            col3.DataType = System.Type.GetType("System.String")
            col3.MaxLength = 50
            col3.Caption = "Datum-Typ"
            dsDebitoren.Tables("tblDebitorenDates").Columns.Add(col3)
            Dim col4 As DataColumn = New DataColumn("datFrom")
            col4.DataType = System.Type.GetType("System.DateTime")
            col4.Caption = "Von"
            dsDebitoren.Tables("tblDebitorenDates").Columns.Add(col4)
            Dim col5 As DataColumn = New DataColumn("datTo")
            col5.DataType = System.Type.GetType("System.DateTime")
            col5.Caption = "Bis"
            dsDebitoren.Tables("tblDebitorenDates").Columns.Add(col5)
            Dim col6 As DataColumn = New DataColumn("strStatus")
            col6.DataType = System.Type.GetType("System.String")
            col6.Caption = "S"
            dsDebitoren.Tables("tblDebitorenDates").Columns.Add(col6)
            dgvDates.DataSource = dsDebitoren.Tables("tblDebitorenDates")

            strPeriode = LstBPerioden.GetItemText(LstBPerioden.SelectedItem)

            Call Main.FcLoginSage3(objdbConn,
                                  objdbMSSQLConn,
                                  objdbSQLcommand,
                                  objFinanz,
                                  objfiBuha,
                                  objdbBuha,
                                  objdbPIFb,
                                  objFiBebu,
                                  objKrBuha,
                                  intMandant,
                                  dsDebitoren.Tables("tblDebitorenInfo"),
                                  dsDebitoren.Tables("tblDebitorenDates"),
                                  strPeriode,
                                  strYear,
                                  intTeqNbr,
                                  intTeqNbrLY,
                                  intTeqNbrPLY,
                                  datPeriodFrom,
                                  datPeriodTo,
                                  strPeriodStatus)

            'Gibt es mehr als ein Jahr?
            If LstBPerioden.Items.Count > 1 Then

                'Gibt es ein Vorjahr?
                If LstBPerioden.SelectedIndex + 1 > 1 Then
                    strPeriode = LstBPerioden.Items(LstBPerioden.SelectedIndex - 1)
                    'Peeriodendef holen
                    Call Main.FcLoginSage4(intMandant,
                                       dsDebitoren.Tables("tblDebitorenDates"),
                                       strPeriode)
                Else
                    'Periode ezreugen und auf N stellen
                    strYearCh = Convert.ToString(Val(strYear) - 1)
                    dsDebitoren.Tables("tblDebitorenDates").Rows.Add(strYearCh, "GJ n/a", DateSerial(Convert.ToUInt16(strYearCh), 1, 1), DateSerial(Convert.ToUInt16(strYearCh), 12, 31), "N")
                End If

                'Gibt es ein Folgehahr?
                If LstBPerioden.SelectedIndex + 1 < LstBPerioden.Items.Count Then
                    strPeriode = LstBPerioden.Items(LstBPerioden.SelectedIndex + 1)
                    'Peeriodendef holen
                    Call Main.FcLoginSage4(intMandant,
                                       dsDebitoren.Tables("tblDebitorenDates"),
                                       strPeriode)
                Else
                    'Periode ezreugen und auf N stellen
                    strYearCh = Convert.ToString(Val(strYear) + 1)
                    dsDebitoren.Tables("tblDebitorenDates").Rows.Add(strYearCh, "GJ n/a", DateSerial(Convert.ToUInt16(strYearCh), 1, 1), DateSerial(Convert.ToUInt16(strYearCh), 12, 31), "N")
                End If
            ElseIf LstBPerioden.Items.Count = 1 Then 'es gibt genau 1 Jahr
                'gewähltes Jahr checken
                Call Main.FcLoginSage4(intMandant,
                                       dsDebitoren.Tables("tblDebitorenDates"),
                                       strPeriode)
                'VJ erzeugen
                strYearCh = Convert.ToString(Val(strYear) - 1)
                dsDebitoren.Tables("tblDebitorenDates").Rows.Add(strYearCh, "GJ n/a", DateSerial(Convert.ToUInt16(strYearCh), 1, 1), DateSerial(Convert.ToUInt16(strYearCh), 12, 31), "N")

                'FJ erzeugen
                strYearCh = Convert.ToString(Val(strYear) + 1)
                dsDebitoren.Tables("tblDebitorenDates").Rows.Add(strYearCh, "GJ n/a", DateSerial(Convert.ToUInt16(strYearCh), 1, 1), DateSerial(Convert.ToUInt16(strYearCh), 12, 31), "N")

            End If


            'Dim clImp As New ClassImport
            'clImp.FcDebitFill(intMandant)
            'clImp = Nothing

            BgWLoadDebi.RunWorkerAsync(intMandant)

            Do While BgWLoadDebi.IsBusy
                Application.DoEvents()
            Loop

            'Tabellentyp darstellen
            Me.lblDB.Text = Main.FcReadFromSettingsII("Buchh_RGTableType", intMandant)


            MySQLdaDebitoren.Fill(dsDebitoren, "tblDebiHeadsFromUser")
            MySQLdaDebitorenSub.Fill(dsDebitoren, "tblDebiSubsFromUser")

            'Application.DoEvents()

            'Dim clCheck As New ClassCheck
            'clCheck.FcClCheckDebit(intMandant,
            '                       dsDebitoren,
            '                       Finanz,
            '                       FBhg,
            '                       DbBhg,
            '                       PIFin,
            '                       BeBu,
            '                       dsDebitoren.Tables("tblDebitorenInfo"),
            '                       dsDebitoren.Tables("tblDebitorenDates"),
            '                       frmImportMain.lstBoxMandant.Text,
            '                       intTeqNbr,
            '                       intTeqNbrLY,
            '                       intTeqNbrPLY,
            '                       strYear,
            '                       frmImportMain.chkValutaCorrect.Checked,
            '                       frmImportMain.dtpValutaCorrect.Value)
            'clCheck = Nothing

            BgWCheckDebitLocArgs.intMandant = intMandant
            BgWCheckDebitLocArgs.strMandant = frmImportMain.lstBoxMandant.GetItemText(frmImportMain.lstBoxMandant.SelectedItem)
            BgWCheckDebitLocArgs.intTeqNbr = intTeqNbr
            BgWCheckDebitLocArgs.intTeqNbrLY = intTeqNbrLY
            BgWCheckDebitLocArgs.intTeqNbrPLY = intTeqNbrPLY
            BgWCheckDebitLocArgs.strYear = strYear
            BgWCheckDebitLocArgs.strPeriode = LstBPerioden.GetItemText(LstBPerioden.SelectedItem)
            BgWCheckDebitLocArgs.booValutaCor = frmImportMain.chkValutaCorrect.Checked
            BgWCheckDebitLocArgs.datValutaCor = frmImportMain.dtpValutaCorrect.Value

            BgWCheckDebi.RunWorkerAsync(BgWCheckDebitLocArgs)

            Do While BgWCheckDebi.IsBusy
                Application.DoEvents()
            Loop

            System.GC.Collect()

            'Grid neu aufbauen
            dgvBookings.DataSource = dsDebitoren.Tables("tblDebiHeadsFromUser")
            dgvBookingSub.DataSource = dsDebitoren.Tables("tblDebiSubsFromUser")

            intFcReturns = FcInitdgvInfo(dgvInfo)
            intFcReturns = FcInitdgvBookings(dgvBookings)
            intFcReturns = FcInitdgvDebiSub(dgvBookingSub)
            intFcReturns = FcInitdgvDate(dgvDates)

            'Anzahl schreiben
            txtNumber.Text = Me.dsDebitoren.Tables("tblDebiHeadsFromUser").Rows.Count.ToString
            Me.Cursor = Cursors.Default

            Me.butImport.Enabled = True
            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Generelles Problem " + Convert.ToString(Err.Number) + "Check Debitoren", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Err.Clear()
            Return 1

        Finally
            objFinanz = Nothing
            objfiBuha = Nothing
            objdbBuha = Nothing
            objdbPIFb = Nothing
            objFiBebu = Nothing
            objKrBuha = Nothing

            objdbConn = Nothing
            objdbMSSQLConn = Nothing
            objdbSQLcommand = Nothing
            objdbtaskcmd = Nothing
            objdbtasks = Nothing

            'System.GC.Collect()

        End Try



    End Function

    Friend Function FcInitdgvDate(ByRef dgvDate As DataGridView) As Int16

        'DGV - Info
        'dgvInfo.DataSource = objdtInfo
        dgvDate.AllowUserToAddRows = False
        dgvDate.AllowUserToDeleteRows = False
        'dgvInfo.Enabled = False
        dgvDate.RowHeadersVisible = False
        dgvDate.Columns("intYear").HeaderText = "Jahr"
        dgvDate.Columns("intYear").Width = 35
        dgvDate.Columns("strDatType").HeaderText = "Type"
        dgvDate.Columns("strDatType").Width = 80
        dgvDate.Columns("datFrom").HeaderText = "Von"
        dgvDate.Columns("datFrom").Width = 70
        dgvDate.Columns("datto").HeaderText = "Bis"
        dgvDate.Columns("datTo").Width = 70
        dgvDate.Columns("strStatus").HeaderText = "S"
        dgvDate.Columns("strStatus").Width = 30
        Return 0

    End Function


    Friend Function FcInitdgvInfo(ByRef dgvInfo As DataGridView) As Int16

        'DGV - Info
        'dgvInfo.DataSource = objdtInfo
        dgvInfo.AllowUserToAddRows = False
        dgvInfo.AllowUserToDeleteRows = False
        'dgvInfo.Enabled = False
        dgvInfo.RowHeadersVisible = False
        dgvInfo.Columns("strInfoT").HeaderText = "Info"
        dgvInfo.Columns("strInfoT").Width = 100
        dgvInfo.Columns("strInfoV").HeaderText = "Wert"
        dgvInfo.Columns("strInfoV").Width = 250
        Return 0

    End Function

    Friend Function FcInitdgvBookings(ByRef dgvBookings As DataGridView) As Int16

        dgvBookings.ShowCellToolTips = False
        dgvBookings.AllowUserToAddRows = False
        dgvBookings.AllowUserToDeleteRows = False
        Dim okcol As New DataGridViewCheckBoxColumn
        okcol.DataPropertyName = "booDebBook"
        okcol.HeaderText = "ok"
        okcol.DisplayIndex = 0
        okcol.Width = 40
        dgvBookings.Columns.Add(okcol)
        dgvBookings.Columns("booDebBook").Visible = False
        'dgvBookings.Columns("booDebBook").DisplayIndex = 0
        'dgvBookings.Columns("booDebBook").HeaderText = "ok"
        'dgvBookings.Columns("booDebBook").Width = 40
        'dgvBookings.Columns("booDebBook").ValueType = System.Type.[GetType]("System.Boolean")
        'dgvBookings.Columns("booDebBook").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        'dgvBookings.Columns("booDebBook").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
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
        dgvBookings.Columns("dblDebNetto").DefaultCellStyle.Format = "N4"
        dgvBookings.Columns("dblDebNetto").ReadOnly = True
        dgvBookings.Columns("dblDebMwSt").DisplayIndex = 8
        dgvBookings.Columns("dblDebMwSt").HeaderText = "MwSt"
        dgvBookings.Columns("dblDebMwSt").Width = 70
        dgvBookings.Columns("dblDebMwSt").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvBookings.Columns("dblDebMwSt").DefaultCellStyle.Format = "N4"
        dgvBookings.Columns("dblDebBrutto").DisplayIndex = 9
        dgvBookings.Columns("dblDebBrutto").HeaderText = "Brutto"
        dgvBookings.Columns("dblDebBrutto").Width = 80
        dgvBookings.Columns("dblDebBrutto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvBookings.Columns("dblDebBrutto").DefaultCellStyle.Format = "N4"
        dgvBookings.Columns("intSubBookings").DisplayIndex = 10
        dgvBookings.Columns("intSubBookings").HeaderText = "Sub"
        dgvBookings.Columns("intSubBookings").Width = 50
        dgvBookings.Columns("intSubBookings").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvBookings.Columns("dblSumSubBookings").DisplayIndex = 11
        dgvBookings.Columns("dblSumSubBookings").HeaderText = "Sub-Summe"
        dgvBookings.Columns("dblSumSubBookings").Width = 80
        dgvBookings.Columns("dblSumSubBookings").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvBookings.Columns("dblSumSubBookings").DefaultCellStyle.Format = "N4"
        dgvBookings.Columns("lngDebIdentNbr").DisplayIndex = 12
        dgvBookings.Columns("lngDebIdentNbr").HeaderText = "Ident"
        dgvBookings.Columns("lngDebIdentNbr").Width = 80
        dgvBookings.Columns("intBuchungsart").DisplayIndex = 13
        dgvBookings.Columns("intBuchungsart").DisplayIndex = 13
        dgvBookings.Columns("intBuchungsart").HeaderText = "BA"
        dgvBookings.Columns("intBuchungsart").Width = 40
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
        Return 0

    End Function

    Friend Function FcInitdgvDebiSub(ByRef dgvBookingSub As DataGridView) As Int16

        dgvBookingSub.RowHeadersWidth = 24

        dgvBookingSub.ShowCellToolTips = False
        dgvBookingSub.AllowUserToAddRows = False
        dgvBookingSub.AllowUserToDeleteRows = False
        dgvBookingSub.Columns("strRGNr").DisplayIndex = 0
        dgvBookingSub.Columns("strRGNr").Width = 50
        dgvBookingSub.Columns("strRGNr").HeaderText = "RG-Nr"
        dgvBookingSub.Columns("intSollHaben").DisplayIndex = 1
        dgvBookingSub.Columns("intSollHaben").Width = 20
        dgvBookingSub.Columns("intSollHaben").HeaderText = "S/H"
        dgvBookingSub.Columns("lngKto").DisplayIndex = 2
        dgvBookingSub.Columns("lngKto").Width = 45
        dgvBookingSub.Columns("lngKto").HeaderText = "Konto"
        dgvBookingSub.Columns("strKtoBez").DisplayIndex = 3
        dgvBookingSub.Columns("strKtoBez").HeaderText = "Bezeichnung"
        dgvBookingSub.Columns("lngKST").DisplayIndex = 4
        dgvBookingSub.Columns("lngKST").Width = 30
        dgvBookingSub.Columns("lngKST").HeaderText = "KST"
        dgvBookingSub.Columns("strKSTBez").DisplayIndex = 5
        dgvBookingSub.Columns("strKSTBez").Width = 60
        dgvBookingSub.Columns("strKSTBez").HeaderText = "Bezeichnung"
        dgvBookingSub.Columns("lngProj").DisplayIndex = 6
        dgvBookingSub.Columns("lngProj").Width = 30
        dgvBookingSub.Columns("lngProj").HeaderText = "Proj"
        dgvBookingSub.Columns("strProjBez").DisplayIndex = 7
        dgvBookingSub.Columns("strProjBez").HeaderText = "Pr.-Bez."
        dgvBookingSub.Columns("strProjBez").Width = 55
        dgvBookingSub.Columns("dblNetto").DisplayIndex = 8
        dgvBookingSub.Columns("dblNetto").Width = 65
        dgvBookingSub.Columns("dblNetto").HeaderText = "Netto"
        dgvBookingSub.Columns("dblNetto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvBookingSub.Columns("dblNetto").DefaultCellStyle.Format = "N4"
        dgvBookingSub.Columns("dblMwSt").DisplayIndex = 9
        dgvBookingSub.Columns("dblMwSt").Width = 60
        dgvBookingSub.Columns("dblMwSt").HeaderText = "MwSt"
        dgvBookingSub.Columns("dblMwSt").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvBookingSub.Columns("dblMwSt").DefaultCellStyle.Format = "N4"
        dgvBookingSub.Columns("dblBrutto").DisplayIndex = 10
        dgvBookingSub.Columns("dblBrutto").Width = 65
        dgvBookingSub.Columns("dblBrutto").HeaderText = "Brutto"
        dgvBookingSub.Columns("dblBrutto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvBookingSub.Columns("dblBrutto").DefaultCellStyle.Format = "N4"
        dgvBookingSub.Columns("dblMwStSatz").DisplayIndex = 11
        dgvBookingSub.Columns("dblMwStSatz").Width = 30
        dgvBookingSub.Columns("dblMwStSatz").HeaderText = "MwStS"
        dgvBookingSub.Columns("strMwStKey").DisplayIndex = 12
        dgvBookingSub.Columns("strMwStKey").Width = 30
        dgvBookingSub.Columns("strMwStKey").HeaderText = "MwStK"
        dgvBookingSub.Columns("strStatusUBText").DisplayIndex = 13
        dgvBookingSub.Columns("strStatusUBText").HeaderText = "Status"
        dgvBookingSub.Columns("strStatusUBText").Width = 135

        'dgvBookingSub.Columns("lngID").Visible = False
        'dgvBookingSub.Columns("strArtikel").Visible = False
        'dgvBookingSub.Columns("strStatusUBBitLog").Visible = False
        'dgvBookingSub.Columns("strDebSubText").Visible = False
        'dgvBookingSub.Columns("strDebBookStatus").Visible = False
        Return 0


    End Function

    Private Sub frmDebDisp_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        'DS in tabelle löschen
        MySQLdaDebitoren.DeleteCommand.Connection.Open()
        MySQLdaDebitoren.DeleteCommand.ExecuteNonQuery()
        MySQLdaDebitoren.DeleteCommand.Connection.Close()

        MySQLdaDebitorenSub.DeleteCommand.Connection.Open()
        MySQLdaDebitorenSub.DeleteCommand.ExecuteNonQuery()
        MySQLdaDebitorenSub.DeleteCommand.Connection.Close()

    End Sub

    Private Sub butImport_Click(sender As Object, e As EventArgs) Handles butImport.Click

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
        Dim BgWImportDebiLocArgs As New BgWCheckDebitArgs

        Try

            'Variablen zuweisen
            BgWImportDebiLocArgs.intMandant = frmImportMain.lstBoxMandant.SelectedValue
            BgWImportDebiLocArgs.intTeqNbr = intTeqNbr
            BgWImportDebiLocArgs.intTeqNbrLY = intTeqNbrLY
            BgWImportDebiLocArgs.intTeqNbrPLY = intTeqNbrPLY
            BgWImportDebiLocArgs.strYear = strYear
            BgWImportDebiLocArgs.strPeriode = frmImportMain.lstBoxPerioden.GetItemText(frmImportMain.lstBoxPerioden.SelectedItem)


            Me.Cursor = Cursors.WaitCursor
            Me.butImport.Enabled = False
            BgWImportDebi.RunWorkerAsync(BgWImportDebiLocArgs)

            Do While BgWImportDebi.IsBusy
                Application.DoEvents()
            Loop

            Me.Cursor = Cursors.Default

            'Me.Cursor = Cursors.WaitCursor
            ''Butteon desaktivieren
            'Me.butImport.Enabled = False

            ''Start in Sync schreiben
            'intReturnValue = WFDBClass.FcWriteStartToSync(objdbConn,
            '                                              frmImportMain.lstBoxMandant.SelectedValue,
            '                                              1,
            '                                              dsDebitoren.Tables("tblDebiHeadsFromUser").Rows.Count)

            ''Setting soll erfasste OP als externe Beleg-Nr. genommen werden und lngDebIdentNbr als Beleg-Nr.
            'objdbConn.Open()
            'booErfOPExt = Convert.ToBoolean(Convert.ToInt16(Main.FcReadFromSettings(objdbConn, "Buchh_ErfOPExt", frmImportMain.lstBoxMandant.SelectedValue)))
            'objdbConn.Close()

            ''Kopfbuchung
            'For Each row In Me.dsDebitoren.Tables("tblDebiHeadsFromUser").Rows

            '    If IIf(IsDBNull(row("booDebBook")), False, row("booDebBook")) Then

            '        'Test ob OP - Buchung
            '        If row("intBuchungsart") = 1 Then

            '            'Immer zuerst Belegs-Nummerierung aktivieren, falls vorhanden externe Nummer = OP - Nr. Rg
            '            'Führt zu Problemen beim Ausbuchen des DTA - Files
            '            'Resultat Besprechnung 17.09.20 mit Muhi/ Andy
            '            'DbBhg.IncrBelNbr = "J"
            '            'Belegsnummer abholen
            '            'intDebBelegsNummer = DbBhg.GetNextBelNbr("R")

            '            'Verdopplung interne BelegsNummer verhindern
            '            DbBhg.CheckDoubleIntBelNbr = "J"

            '            If row("dblDebBrutto") < 0 Then
            '                'Gutschrift
            '                'Falls booGSToInv (Gutschrift zu Rechnung) dann OP-Nummer vorgeben, sonst hochzählen lassen
            '                If row("booCrToInv") Then
            '                    'Beleg-Nummerierung desaktivieren
            '                    DbBhg.IncrBelNbr = "N"
            '                    'Eingelesene OP-Nummer (=Verknüpfte OP-Nr.) = interne Beleg-Nummer
            '                    intDebBelegsNummer = Main.FcCleanRGNrStrict(row("strOPNr"))
            '                    strExtBelegNbr = row("strDebRGNbr")
            '                Else
            '                    'Zuerst Beleg-Nummerieungung aktivieren
            '                    DbBhg.IncrBelNbr = "J"
            '                    'Belegsnummer abholen
            '                    intDebBelegsNummer = DbBhg.GetNextBelNbr("G")
            '                'Prüfen ob wirklich frei und falls nicht hochzählen
            '                intReturnValue = MainDebitor.FcCheckDebiExistance(objdbMSSQLConn,
            '                                                                      objdbSQLcommand,
            '                                                                      intDebBelegsNummer,
            '                                                                      "G",
            '                                                                      intTeqNbr)


            '                'intReturnValue = 10
            '                'Do Until intReturnValue = 0

            '                '    intReturnValue = DbBhg.doesBelegExist(row("lngDebNbr").ToString,
            '                '                                      row("strDebCur"),
            '                '                                      intDebBelegsNummer.ToString,
            '                '                                      "NOT_SET",
            '                '                                      "G",
            '                '                                      "NOT_SET")
            '                '    If intReturnValue <> 0 Then
            '                '        intDebBelegsNummer += 1
            '                '    End If
            '                'Loop
            '                strExtBelegNbr = row("strOPNr")
            '                End If

            '                'Beträge drehen
            '                row("dblDebBrutto") = row("dblDebBrutto") * -1
            '                row("dblDebMwSt") = row("dblDebMwSt") * -1
            '                row("dblDebNetto") = row("dblDebNetto") * -1

            '                strBuchType = "G"

            '            Else

            '                If String.IsNullOrEmpty(row("strOPNr")) Then
            '                    'strExtBelegNbr = row("strOPNr")

            '                    'Zuerst Beleg-Nummerieungung aktivieren
            '                    DbBhg.IncrBelNbr = "J"
            '                    'Belegsnummer abholen
            '                    intDebBelegsNummer = DbBhg.GetNextBelNbr("R")
            '                intReturnValue = MainDebitor.FcCheckDebiExistance(objdbMSSQLConn,
            '                                                                      objdbSQLcommand,
            '                                                                      intDebBelegsNummer,
            '                                                                      "R",
            '                                                                      intTeqNbr)
            '            Else
            '                    If Strings.Len(Main.FcCleanRGNrStrict(row("strOPNr"))) > 9 Then
            '                        'Zahl zu gross
            '                        DbBhg.IncrBelNbr = "J"
            '                        'Belegsnummer abholen
            '                        intDebBelegsNummer = DbBhg.GetNextBelNbr("R")
            '                    intReturnValue = MainDebitor.FcCheckDebiExistance(objdbMSSQLConn,
            '                                                                          objdbSQLcommand,
            '                                                                          intDebBelegsNummer,
            '                                                                          "R",
            '                                                                          intTeqNbr)
            '                    strExtBelegNbr = row("strOPNr")
            '                    Else
            '                        'Beleg-Nummerierung abschalten
            '                        DbBhg.IncrBelNbr = "N"
            '                        'Gemäss Setting Erfasste OP-Nr. Nummern vergeben
            '                        If Not booErfOPExt Then
            '                            intDebBelegsNummer = Main.FcCleanRGNrStrict(row("strOPNr"))
            '                            strExtBelegNbr = row("strOPNr")
            '                        Else
            '                            'bei t_debi: IdentNbr wird genommen da dort die RG-Nr. drin ist. RgNr = ID
            '                            intDebBelegsNummer = row("lngDebIdentNbr")
            '                            strExtBelegNbr = row("strOPNr")
            '                        End If

            '                    End If

            '                End If

            '                strBuchType = "R"

            '            End If

            '            'Variablen zuweisen
            '            'Sachbearbeiter aus Debitor auslesen
            '            strDebiLine = DbBhg.ReadDebitor3(row("lngDebNbr") * -1, "")
            '            strDebitor = Split(strDebiLine, "{>}")
            '            strSachBID = strDebitor(30)
            '            'strExtBelegNbr = row("strDebRGNbr")
            '            intDebitorNbr = row("lngDebNbr")
            '            strValutaDatum = Format(row("datDebValDatum"), "yyyyMMdd").ToString
            '            strBelegDatum = Format(row("datDebRGDatum"), "yyyyMMdd").ToString
            '            If IsDBNull(row("datDebDue")) Then
            '                strVerfallDatum = String.Empty
            '            Else
            '                strVerfallDatum = Format(row("datDebDue"), "yyyyMMdd").ToString
            '            End If
            '            strReferenz = row("strDebReferenz")
            '            strMahnerlaubnis = String.Empty 'Format(row("datDebRGDatum"), "yyyyMMdd").ToString
            '            dblBetrag = row("dblDebBrutto")
            '            'Bei SplittBill 2ter Rechnung Text anfügen
            '            If row("booLinked") Then
            '                strDebiText = row("strDebText") + ", FRG "
            '            Else
            '                strDebiText = row("strDebText")
            '            End If
            '            'strDebiText = row("strDebText")
            '            strCurrency = row("strDebCur")
            '            If strCurrency <> "CHF" Then 'Muss ergänzt werden => Was ist Leitwährung auf dem Konto
            '                dblKurs = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg)
            '            Else
            '                dblKurs = 1.0#
            '            End If
            '            intEigeneBank = row("strDebiBank")
            '            'Zahl-Kondition
            '            intKondition = IIf(IsDBNull(row("intZKond")), 1, row("intZKond"))

            '            Try
            '                booBooingok = True
            '                Call DbBhg.SetBelegKopf2(intDebBelegsNummer,
            '                                         strValutaDatum,
            '                                         intDebitorNbr,
            '                                         strBuchType,
            '                                         strBelegDatum,
            '                                         strVerfallDatum,
            '                                         strDebiText,
            '                                         strReferenz,
            '                                         intKondition,
            '                                         strSachBID,
            '                                         strVerkID,
            '                                         strMahnerlaubnis,
            '                                         sngAktuelleMahnstufe,
            '                                         dblBetrag.ToString,
            '                                         dblKurs.ToString,
            '                                         strExtBelegNbr,
            '                                         strSkonto,
            '                                         strCurrency,
            '                                         "",
            '                                         intEigeneBank.ToString)

            '                'Application.DoEvents()

            '            Catch ex As Exception
            '                strErrMessage = "Problem " + (Err.Number And 65535).ToString + " Belegkopf zu" + intDebBelegsNummer.ToString + vbCrLf
            '                strErrMessage += "RG " + strRGNbr + vbCrLf
            '                strErrMessage += "Debitor " + intDebitorNbr.ToString

            '                MessageBox.Show(ex.Message + vbCrLf + strErrMessage, "Problem " + (Err.Number And 65535).ToString, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            '                If (Err.Number And 65535) < 10000 Then
            '                    booBooingok = False
            '                Else
            '                    booBooingok = True
            '                End If

            '            End Try

            '            selDebiSub = dsDebitoren.Tables("tblDebiSubsFromUser").Select("strRGNr='" + row("strDebRGNbr") + "'")
            '            strRGNbr = row("strDebRGNbr")

            '            For Each SubRow As DataRow In selDebiSub

            '                'Bei zweiter Splitt-Bill Rechung hier eingreifen
            '                'Gegenkonto auf 1092, MwStKey auf 'null' setzen, KST = 0
            '                'If row("booLinked") Then
            '                '    If row("booLinkedPayed") Then
            '                '        intGegenKonto = 2331
            '                '    Else
            '                '        intGegenKonto = 1092
            '                '    End If
            '                '    SubRow("dblNetto") = SubRow("dblBrutto")
            '                '    SubRow("strMwStKey") = "null"
            '                '    SubRow("lngKST") = 0
            '                'Else
            '                intGegenKonto = SubRow("lngKto")
            '                'End If
            '                strFibuText = SubRow("strDebSubText")
            '                If intGegenKonto <> 6906 Then
            '                    If strBuchType = "R" Then
            '                        dblNettoBetrag = SubRow("dblNetto") * -1
            '                    Else
            '                        dblNettoBetrag = SubRow("dblNetto")
            '                    End If
            '                Else 'Rundungsdifferenzen
            '                    If strBuchType = "R" Then
            '                        dblNettoBetrag = SubRow("dblBrutto") * -1
            '                    Else
            '                        dblNettoBetrag = SubRow("dblBrutto")
            '                    End If
            '                End If
            '                'dblBebuBetrag = 1000.0#
            '                If SubRow("lngKST") > 0 Then
            '                    strBeBuEintrag = SubRow("lngKST").ToString + "{<}" + SubRow("strDebSubText") + "{<}" + "CALCULATE" + "{>}"    '"PROD{<}BebuText{<}" + dblBebuBetrag.ToString + "{>}"
            '                Else
            '                    'strBeBuEintrag = "999999" + "{<}" + SubRow("strDebSubText") + "{<}" + "CALCULATE" + "{>}"
            '                    strBeBuEintrag = Nothing
            '                End If
            '                If Not IsDBNull(SubRow("strMwStKey")) And
            '                        SubRow("strMwStKey") <> "null" And
            '                        SubRow("lngKto") <> 6906 Then 'And SubRow("strMwStKey") <> "25" Then
            '                    If strBuchType = "R" Then
            '                        strSteuerFeld = Main.FcGetSteuerFeld(FBhg,
            '                                                             SubRow("lngKto"),
            '                                                             SubRow("strDebSubText"),
            '                                                             SubRow("dblBrutto") * -1,
            '                                                             SubRow("strMwStKey"),
            '                                                             SubRow("dblMwSt") * -1)     '"25{<}DEBI D Bhg Export MwSt{<}0{>}"
            '                    Else
            '                        strSteuerFeld = Main.FcGetSteuerFeld(FBhg,
            '                                                             SubRow("lngKto"),
            '                                                             SubRow("strDebSubText"),
            '                                                             SubRow("dblBrutto"),
            '                                                             SubRow("strMwStKey"),
            '                                                             SubRow("dblMwSt"))     '"25{<}DEBI D Bhg Export MwSt{<}0{>}"
            '                    End If
            '                Else
            '                    strSteuerFeld = "STEUERFREI"
            '                End If
            '                'strSteuerInfo = Split(FBhg.GetKontoInfo(intGegenKonto.ToString), "{>}")
            '                'Debug.Print("Konto-Info: " + strSteuerInfo(26))

            '                Try

            '                    booBooingok = True
            '                    Call DbBhg.SetVerteilung(intGegenKonto.ToString,
            '                                             strFibuText,
            '                                             dblNettoBetrag.ToString,
            '                                             strSteuerFeld,
            '                                             strBeBuEintrag)

            '                    'Application.DoEvents()

            '                Catch ex As Exception
            '                    strErrMessage = "Problem " + (Err.Number And 65535).ToString + " Verteilung " + intDebBelegsNummer.ToString + vbCrLf
            '                    strErrMessage += "RG " + strRGNbr + vbCrLf
            '                    strErrMessage += "Konto " + SubRow("lngKto").ToString + vbCrLf
            '                    strErrMessage += "Gegenkonto " + intGegenKonto.ToString + vbCrLf
            '                    strErrMessage += "Betrag " + dblNettoBetrag.ToString + vbCrLf
            '                    strErrMessage += "Steuer " + strSteuerFeld + vbCrLf
            '                    strErrMessage += "Bebu " + strBeBuEintrag

            '                    MessageBox.Show(ex.Message + vbCrLf + strErrMessage, "Problem " + (Err.Number And 65535).ToString, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            '                    If (Err.Number And 65535) < 10000 Then
            '                        booBooingok = False
            '                    Else
            '                        booBooingok = True
            '                    End If

            '                End Try

            '                strSteuerFeld = Nothing
            '                strBeBuEintrag = Nothing

            '                'Status Sub schreiben
            '                'Application.DoEvents()

            '            Next

            '            Try

            '                booBooingok = True
            '                Call DbBhg.WriteBuchung()

            '                'Bei SplittBill 2ter Rechnung TZahlung auf LinkedRG machen
            '                'Prinzip: Beleg einlesen anhand und Betrag ausrechnen => Summe Beleg - diesen Beleg
            '                If row("booLinked") And Mid(row("strDebStatusBitLog"), 13, 1) = "0" Then 'Nur wenn Beleg in gleicher Buha
            '                    'Betrag von Beleg 1 holen
            '                    intLaufNbr = DbBhg.doesBelegExist2(row("lngLinkedDeb").ToString,
            '                                                       row("strDebCur"),
            '                                                       row("lngLinkedRG").ToString,
            '                                                       "NOT_SET",
            '                                                       "R",
            '                                                       "NOT_SET",
            '                                                       "NOT_SET",
            '                                                       "NOT_SET")

            '                    If intLaufNbr > 0 Then
            '                        strBeleg = DbBhg.GetBeleg(row("lngLinkedDeb").ToString,
            '                                                  intLaufNbr.ToString)

            '                        strBelegArr = Split(strBeleg, "{>}")
            '                        If strBelegArr(4) = "B" Then 'schon bezahlt
            '                            'Ausbuchen?, wohin mit dem Betrag?
            '                        Else

            '                            'Betrag von RG 10 auf RG1 als TZ buchen
            '                            dblSplitPayed = dblBetrag

            '                            'Teilzahlung buchen
            '                            Call DbBhg.SetZahlung(1944,
            '                                              strBelegDatum,
            '                                              strValutaDatum,
            '                                              row("strDebCur"),
            '                                              dblKurs,
            '                                              "",
            '                                              "",
            '                                              row("lngLinkedDeb"),
            '                                              dblSplitPayed.ToString,
            '                                              row("strDebCur"),
            '                                              ,
            '                                              row("lngDebIdentNbr").ToString + ", TZ " + row("strDebRGNbr").ToString)
            '                            'Application.DoEvents()

            '                            Call DbBhg.WriteTeilzahlung4(intLaufNbr.ToString,
            '                                                     row("lngDebIdentNbr").ToString + ", TZ " + row("strDebRGNbr").ToString,
            '                                                     "NOT_SET",
            '                                                     ,
            '                                                     "NOT_SET",
            '                                                     "NOT_SET",
            '                                                     "Default",
            '                                                     "Default")
            '                            'Application.DoEvents()

            '                        End If

            '                    End If

            '                End If

            '            Catch ex As Exception
            '                'MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Belegerstellung " + intDebBelegsNummer.ToString + ", RG " + strRGNbr)
            '                If (Err.Number And 65535) < 10000 Then
            '                    strErrMessage = "Belegerstellung RG " + strRGNbr + " Beleg " + intDebBelegsNummer.ToString + " NICHT möglich!"
            '                    MessageBox.Show(ex.Message + vbCrLf + strErrMessage, "Problem " + (Err.Number And 65535).ToString, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            '                    booBooingok = False
            '                Else
            '                    strErrMessage = "Belegerstellung RG " + strRGNbr + " Beleg " + intDebBelegsNummer.ToString + " möglich mit Warnung"
            '                    MessageBox.Show(ex.Message, "Warnung " + (Err.Number And 65535).ToString, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            '                    booBooingok = True
            '                End If

            '            End Try


            '        Else

            '            'Buchung nur in Fibu
            '            'Prinzip Funktion WriteBuchung() anwenden mit allen Parametern

            '            'Verdopplung interne BelegsNummer verhindern
            '            FBhg.CheckDoubleIntBelNbr = "J"

            '            If IIf(IsDBNull(row("strOPNr")), "", row("strOPNr")) <> "" And IIf(IsDBNull(row("lngDebIdentNbr")), 0, row("lngDebIdentNbr")) <> 0 Then
            '                'Belegsnummer abholen fall keine Beleg-Nummer angegeben
            '                intDebBelegsNummer = FBhg.GetNextBelNbr()
            '                'Prüfen ob wirklich frei
            '                intReturnValue = 10
            '                Do Until intReturnValue = 0
            '                    intReturnValue = FBhg.doesBelegExist(intDebBelegsNummer,
            '                                                         "NOT_SET",
            '                                                         "NOT_SET",
            '                                                         String.Concat(Microsoft.VisualBasic.Left(frmImportMain.lstBoxPerioden.Text, 4) - 1, "0101"),
            '                                                         String.Concat(Microsoft.VisualBasic.Left(frmImportMain.lstBoxPerioden.Text, 4), "1231"))
            '                    If intReturnValue <> 0 Then
            '                        intDebBelegsNummer += 1
            '                    End If
            '                Loop
            '                'Debug.Print("Belegnummer taken:  " + intDebBelegsNummer.ToString)
            '            Else
            '                If IIf(IsDBNull(row("strOPNr")), "", row("strOPNr")) <> "" Then
            '                    intDebBelegsNummer = Convert.ToInt32(row("strOPNr"))
            '                Else
            '                    intDebBelegsNummer = row("lngDebIdentNbr")
            '                End If
            '            End If
            '            'Variablen zuweisen
            '            strBelegDatum = Format(row("datDebRGDatum"), "yyyyMMdd").ToString
            '            strValutaDatum = Format(row("datDebValDatum"), "yyyyMMdd").ToString
            '            'strDebiText = row("strDebText")
            '            strCurrency = row("strDebCur")
            '            If strCurrency <> "CHF" Then 'Muss ergänzt werden => Was ist Leitwährung auf dem Konto
            '                dblKurs = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg)
            '            Else
            '                dblKurs = 1.0#
            '            End If

            '            selDebiSub = dsDebitoren.Tables("tblDebiSubsFromUser").Select("strRGNr='" + row("strDebRGNbr") + "'")
            '            strRGNbr = row("strDebRGNbr")

            '            If selDebiSub.Length = 2 Then

            '                'Initialisieren
            '                dblNettoBetrag = 0
            '                dblSollBetrag = 0
            '                dblHabenBetrag = 0
            '                strBeBuEintrag = String.Empty
            '                strBeBuEintragSoll = String.Empty
            '                strBeBuEintragHaben = String.Empty
            '                strSteuerFeld = String.Empty
            '                strSteuerFeldHaben = String.Empty
            '                strSteuerFeldSoll = String.Empty

            '                For Each SubRow As DataRow In selDebiSub

            '                    If SubRow("intSollHaben") = 0 Then 'Soll

            '                        intSollKonto = SubRow("lngKto")
            '                        dblKursSoll = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg, intSollKonto)
            '                        'strSteuerInfo = Split(FBhg.GetKontoInfo(intSollKonto.ToString), "{>}")
            '                        'Debug.Print("Konto-Info Soll: " + strSteuerInfo(26))
            '                        dblSollBetrag = SubRow("dblNetto")
            '                        strDebiTextSoll = SubRow("strDebSubText")
            '                        If SubRow("dblMwSt") > 0 Then
            '                            strSteuerFeldSoll = Main.FcGetSteuerFeld(FBhg,
            '                                                                     SubRow("lngKto"),
            '                                                                     strDebiTextSoll,
            '                                                                     SubRow("dblBrutto") * dblKursSoll,
            '                                                                     SubRow("strMwStKey"),
            '                                                                     SubRow("dblMwSt"))
            '                        Else
            '                            strSteuerFeldSoll = "STEUERFREI"
            '                        End If
            '                        If SubRow("lngKST") > 0 Then
            '                            strBeBuEintragSoll = SubRow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
            '                        End If


            '                    ElseIf SubRow("intSollHaben") = 1 Then 'Haben

            '                        intHabenKonto = SubRow("lngKto")
            '                        dblKursHaben = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg, intHabenKonto)
            '                        'strSteuerInfo = Split(FBhg.GetKontoInfo(intSollKonto.ToString), "{>}")
            '                        'Debug.Print("Konto-Info Haben: " + strSteuerInfo(26))
            '                        dblHabenBetrag = SubRow("dblNetto") * -1
            '                        'dblHabenBetrag = dblSollBetrag
            '                        strDebiTextHaben = SubRow("strDebSubText")
            '                        If SubRow("dblMwSt") * -1 > 0 Then
            '                            strSteuerFeldHaben = Main.FcGetSteuerFeld(FBhg,
            '                                                                      SubRow("lngKto"),
            '                                                                      strDebiTextHaben,
            '                                                                      SubRow("dblBrutto") * dblKursHaben * -1,
            '                                                                      SubRow("strMwStKey"),
            '                                                                      SubRow("dblMwSt") * -1)
            '                        Else
            '                            strSteuerFeldHaben = "STEUERFREI"
            '                        End If
            '                        If SubRow("lngKST") > 0 Then
            '                            strBeBuEintragHaben = SubRow("lngKST").ToString + "{<}" + strDebiTextHaben + "{<}" + "CALCULATE" + "{>}"
            '                        End If

            '                    Else

            '                        MsgBox("Nicht definierter Wert Sub-Buchungs-SollHaben: " + SubRow("intSollHaben").ToString)

            '                    End If
            '                    'Application.DoEvents()

            '                Next

            '                'Tieferer Betrag für die Gesamt-Buchung herausfinden
            '                If dblSollBetrag <= dblHabenBetrag Then
            '                    dblNettoBetrag = dblSollBetrag
            '                ElseIf dblHabenBetrag < dblSollBetrag Then
            '                    dblNettoBetrag = dblHabenBetrag
            '                End If

            '                Try

            '                    booBooingok = True
            '                    'Buchung ausführen
            '                    Call FBhg.WriteBuchung(0,
            '                                       intDebBelegsNummer,
            '                                       strBelegDatum,
            '                                       intSollKonto.ToString,
            '                                       strDebiTextSoll,
            '                                       strCurrency,
            '                                       dblKursSoll.ToString,
            '                                       (dblNettoBetrag * dblKursSoll).ToString,
            '                                       strSteuerFeldSoll,
            '                                       intHabenKonto.ToString,
            '                                       strDebiTextHaben,
            '                                       strCurrency,
            '                                       dblKursHaben.ToString,
            '                                       (dblNettoBetrag * dblKursHaben).ToString,
            '                                       strSteuerFeldHaben,
            '                                       strCurrency,
            '                                       dblKurs.ToString,
            '                                       (dblNettoBetrag * dblKurs).ToString,
            '                                       dblNettoBetrag.ToString,
            '                                       strBeBuEintragSoll,
            '                                       strBeBuEintragHaben,
            '                                       strValutaDatum)

            '                    'Application.DoEvents()

            '                Catch ex As Exception
            '                    MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Belegerstellung " + intDebBelegsNummer.ToString + ", RG " + strRGNbr, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            '                    If (Err.Number And 65535) < 10000 Then
            '                        booBooingok = False
            '                    Else
            '                        booBooingok = True
            '                    End If

            '                End Try

            '                'Initialisieren
            '                'dblNettoBetrag = 0
            '                'dblSollBetrag = 0
            '                'dblHabenBetrag = 0
            '                'strBeBuEintrag = ""
            '                'strBeBuEintragSoll = ""
            '                'strBeBuEintragHaben = ""
            '                'strSteuerFeld = ""
            '                'strSteuerFeldHaben = ""
            '                'strSteuerFeldSoll = ""


            '                'Vergebene Nummer checken
            '                'intDebBelegsNummer = FBhg.GetBuchLaufnr()

            '            Else
            '                'Sammelbeleg
            '                'MsgBox("Nicht 2 Subbuchungen.")
            '                'Variablen initiieren
            '                strDebiText = row("strDebText")
            '                intCommonKonto = row("lngDebKtoNbr") 'Sammelkonto

            '                'Beleg-Kopf
            '                Call FBhg.SetSammelBhgCommonT2(strValutaDatum,
            '                                               intDebBelegsNummer.ToString,
            '                                               intCommonKonto.ToString,
            '                                               strDebiText,
            '                                               strBelegDatum)

            '                'Buchungen
            '                For Each SubRow As DataRow In selDebiSub

            '                    'Common - Konto ausblenden da sonst Doppelbuchung
            '                    If SubRow("lngKto") <> intCommonKonto Then

            '                        intSollKonto = 0
            '                        strDebiTextSoll = String.Empty
            '                        strDebiCurrency = String.Empty
            '                        dblKursSoll = 0
            '                        dblSollBetrag = 0
            '                        strSteuerFeldSoll = String.Empty
            '                        intHabenKonto = 0
            '                        strDebiTextHaben = String.Empty
            '                        strKrediCurrency = String.Empty
            '                        dblKursHaben = 0
            '                        dblHabenBetrag = 0
            '                        strSteuerFeldHaben = String.Empty
            '                        dblBuchBetrag = 0
            '                        dblBasisBetrag = 0
            '                        strBeBuEintragSoll = String.Empty
            '                        strBeBuEintragHaben = String.Empty
            '                        strErfassungsDatum = Format(Date.Today(), "yyyyMMdd").ToString

            '                        If SubRow("intSollHaben") = 0 And SubRow("lngKto") <> intCommonKonto Then 'Soll

            '                            intSollKonto = SubRow("lngKto")
            '                            strDebiTextSoll = SubRow("strDebSubText")
            '                            strDebiCurrency = strCurrency
            '                            dblKursSoll = 1 / Main.FcGetKurs(strCurrency, strValutaDatum, FBhg, intSollKonto)
            '                            dblSollBetrag = SubRow("dblNetto")
            '                            If SubRow("dblMwSt") > 0 Then
            '                                strSteuerFeldSoll = Main.FcGetSteuerFeld(FBhg, SubRow("lngKto"), strDebiTextSoll, SubRow("dblBrutto") * dblKursSoll, SubRow("strMwStKey"), SubRow("dblMwSt"))
            '                            Else
            '                                strSteuerFeldSoll = "STEUERFREI"
            '                            End If
            '                            If SubRow("lngKST") > 0 Then
            '                                strBeBuEintragSoll = SubRow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
            '                            End If

            '                            'Haben Seite Common-Konto
            '                            intHabenKonto = intCommonKonto
            '                            strDebiTextHaben = SubRow("strDebSubText")
            '                            strKrediCurrency = strCurrency
            '                            dblKursHaben = dblKursSoll
            '                            dblHabenBetrag = SubRow("dblNetto")
            '                            'If SubRow("dblMwSt") > 0 Then
            '                            ' strSteuerFeldHaben = Main.FcGetSteuerFeld(FBhg, SubRow("lngKto"), strDebiTextSoll, SubRow("dblBrutto") * dblKursSoll, SubRow("strMwStKey"), SubRow("dblMwSt"))
            '                            'End If
            '                            If SubRow("lngKST") > 0 Then
            '                                strBeBuEintragHaben = SubRow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
            '                            End If

            '                            'Betrag
            '                            dblBuchBetrag = SubRow("dblBrutto")
            '                            dblBasisBetrag = SubRow("dblBrutto") 'Bei nicht CHF umrechnen

            '                        ElseIf SubRow("intSollHaben") = 1 Then 'Haben

            '                            intHabenKonto = SubRow("lngKto")
            '                            strDebiTextHaben = SubRow("strDebSubText")
            '                            strKrediCurrency = strCurrency
            '                            dblKursHaben = 1 / Main.FcGetKurs(strCurrency, strValutaDatum, FBhg, intHabenKonto)
            '                            dblHabenBetrag = SubRow("dblNetto") * -1
            '                            If (SubRow("dblMwSt") * -1) > 0 Then
            '                                strSteuerFeldHaben = Main.FcGetSteuerFeld(FBhg, SubRow("lngKto"), strDebiTextHaben, SubRow("dblBrutto") * dblKursHaben * -1, SubRow("strMwStKey"), SubRow("dblMwSt") * -1)
            '                            Else
            '                                strSteuerFeldHaben = "STEUERFREI"
            '                            End If
            '                            If SubRow("lngKST") > 0 Then
            '                                strBeBuEintragHaben = SubRow("lngKST").ToString + "{<}" + strDebiTextHaben + "{<}" + "CALCULATE" + "{>}"
            '                            End If

            '                            'Soll - Seite Common - Konto 
            '                            intSollKonto = intCommonKonto
            '                            strDebiTextSoll = SubRow("strDebSubText")
            '                            strDebiCurrency = strCurrency
            '                            dblKursSoll = dblKursHaben
            '                            dblSollBetrag = SubRow("dblNetto") * -1

            '                            'If SubRow("dblMwSt") * -1 > 0 Then
            '                            ' strSteuerFeldSoll = Main.FcGetSteuerFeld(FBhg, SubRow("lngKto"), strDebiTextSoll, SubRow("dblBrutto") * dblKursSoll * -1, SubRow("strMwStKey"), SubRow("dblMwSt") * -1)
            '                            'End If
            '                            If SubRow("lngKST") > 0 Then
            '                                strBeBuEintragSoll = SubRow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
            '                            End If

            '                            dblBuchBetrag = SubRow("dblBrutto") * -1
            '                            dblBasisBetrag = SubRow("dblBrutto") * -1 'Bei nicht CHF umrechnen

            '                        End If
            '                        'Buchungsbetrag von Kopfbuchung nehmen
            '                        'dblBuchBetrag = row("dblDebBrutto")
            '                        'dblBasisBetrag = row("dblDebBrutto")

            '                        Call FBhg.SetSammelBhgT(intSollKonto.ToString,
            '                                                strDebiTextSoll,
            '                                                strDebiCurrency,
            '                                                dblKursSoll.ToString,
            '                                                dblSollBetrag.ToString,
            '                                                strSteuerFeldSoll,
            '                                                intHabenKonto.ToString,
            '                                                strDebiTextHaben,
            '                                                strKrediCurrency,
            '                                                dblKursHaben.ToString,
            '                                                dblHabenBetrag.ToString,
            '                                                strSteuerFeldHaben,
            '                                                strCurrency,
            '                                                dblKurs.ToString,
            '                                                dblBuchBetrag.ToString,
            '                                                dblBasisBetrag.ToString,
            '                                                strBeBuEintragSoll,
            '                                                strBeBuEintragHaben,
            '                                                strErfassungsDatum)

            '                        'Application.DoEvents()

            '                    End If

            '                Next

            '                'Sammelbeleg schreiben
            '                Try

            '                    booBooingok = True
            '                    Call FBhg.WriteSammelBhgT()

            '                Catch ex As Exception

            '                    MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Belegerstellung " + intDebBelegsNummer.ToString + ", RG " + strRGNbr, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            '                    If (Err.Number And 65535) < 10000 Then
            '                        booBooingok = False
            '                    Else
            '                        booBooingok = True
            '                    End If
            '                End Try


            '            End If

            '        End If

            '        If booBooingok Then
            '            If row("booPGV") Then
            '                'Bei PGV Buchungen
            '                If IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")) = "" Or
            '                    (IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")) = "RV" And row("intPGVMthsAY") + row("intPGVMthsNY") > 1) Then

            '                    intReturnValue = MainDebitor.FcPGVDTreatment(FBhg,
            '                                                       Finanz,
            '                                                       DbBhg,
            '                                                       PIFin,
            '                                                       BeBu,
            '                                                       KrBhg,
            '                                                       dsDebitoren.Tables("tblDebiSubsFromUser"),
            '                                                       row("strDebRGNbr"),
            '                                                       intDebBelegsNummer,
            '                                                       row("strDebCur"),
            '                                                       row("datDebValDatum"),
            '                                                       "M",
            '                                                       row("datPGVFrom"),
            '                                                       row("datPGVTo"),
            '                                                       row("intPGVMthsAY") + row("intPGVMthsNY"),
            '                                                       row("intPGVMthsAY"),
            '                                                       row("intPGVMthsNY"),
            '                                                       1311,
            '                                                       1312,
            '                                                       frmImportMain.lstBoxPerioden.Text,
            '                                                       objdbConn,
            '                                                       objdbMSSQLConn,
            '                                                       objdbSQLcommand,
            '                                                       frmImportMain.lstBoxMandant.SelectedValue,
            '                                                       dsDebitoren.Tables("tblDebitorenInfo"),
            '                                                       strYear,
            '                                                       intTeqNbr,
            '                                                       intTeqNbrLY,
            '                                                       intTeqNbrPLY,
            '                                                       IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")),
            '                                                       datPeriodFrom,
            '                                                       datPeriodTo,
            '                                                       strPeriodStatus)


            '                Else
            '                    intReturnValue = MainDebitor.FcPGVDTreatmentYC(FBhg,
            '                                                       Finanz,
            '                                                       DbBhg,
            '                                                       PIFin,
            '                                                       BeBu,
            '                                                       KrBhg,
            '                                                       dsDebitoren.Tables("tblDebiSubsFromUser"),
            '                                                       row("strDebRGNbr"),
            '                                                       intDebBelegsNummer,
            '                                                       row("strDebCur"),
            '                                                       row("datDebValDatum"),
            '                                                       "M",
            '                                                       row("datPGVFrom"),
            '                                                       row("datPGVTo"),
            '                                                       row("intPGVMthsAY") + row("intPGVMthsNY"),
            '                                                       row("intPGVMthsAY"),
            '                                                       row("intPGVMthsNY"),
            '                                                       1311,
            '                                                       1312,
            '                                                       frmImportMain.lstBoxPerioden.Text,
            '                                                       objdbConn,
            '                                                       objdbMSSQLConn,
            '                                                       objdbSQLcommand,
            '                                                       frmImportMain.lstBoxMandant.SelectedValue,
            '                                                       dsDebitoren.Tables("tblDebitorenInfo"),
            '                                                       strYear,
            '                                                       intTeqNbr,
            '                                                       intTeqNbrLY,
            '                                                       intTeqNbrPLY,
            '                                                       IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")),
            '                                                       datPeriodFrom,
            '                                                       datPeriodTo,
            '                                                       strPeriodStatus)
            '                End If


            '            End If

            '            'Status Head schreiben
            '            row("strDebBookStatus") = row("strDebStatusBitLog")
            '            row("booBooked") = True
            '            row("datBooked") = Now()
            '            row("lngBelegNr") = intDebBelegsNummer
            '            'Application.DoEvents()
            '            dsDebitoren.Tables("tblDebiHeadsFromUser").AcceptChanges()

            '            'Status in File RG-Tabelle schreiben
            '            intReturnValue = MainDebitor.FcWriteToRGTable(frmImportMain.lstBoxMandant.SelectedValue,
            '                                                          row("strDebRGNbr"),
            '                                                          row("datBooked"),
            '                                                          row("lngBelegNr"),
            '                                                          objdbAccessConn,
            '                                                          objOracleConn,
            '                                                          objdbConn,
            '                                                          row("booDatChanged"),
            '                                                          row("datDebRGDatum"),
            '                                                          row("datDebValDatum"))
            '            If intReturnValue <> 0 Then
            '                'Throw an exception
            '            End If

            '            'Evtl. Query nach Buchung ausführen
            '            Call MainDebitor.FcExecuteAfterDebit(frmImportMain.lstBoxMandant.SelectedValue, objdbConn)
            '        End If

            '    End If

            'Next
            ''Status in t_sage_syncstatus schreiben
            ''intReturnValue = MainDebitor.FcWriteEndToSync(objdbConn,
            ''                                              cmbBuha.SelectedValue,
            ''                                              1,
            ''                                              Date.Now,
            ''                                              0,
            ''                                              IIf(booBooingok, "ok", "Probleme"))

            'intReturnValue = WFDBClass.FcWriteEndToSync(objdbConn,
            '                                            frmImportMain.lstBoxMandant.SelectedValue,
            '                                            1,
            '                                            0,
            '                                            IIf(booBooingok, "ok", "Probleme"))




        Catch ex As Exception
            MessageBox.Show(ex.Message, "Generelles Problem " + (Err.Number And 65535).ToString + " Belegerstellung ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        Finally

            'Application.DoEvents()
            'Grid neu aufbauen, Daten von Mandant einlesen
            'Call butDebitoren.PerformClick()

            Me.Cursor = Cursors.Default
            'Me.butImport.Enabled = False
            Me.Close()
            'Me.Dispose()
            'System.GC.Collect()
            'Application.Restart()

        End Try


    End Sub

    Private Sub dgvBookings_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvBookings.CellContentClick

        Dim intFctReturns As Int16

        Try

            If e.RowIndex >= 0 Then

                dgvBookingSub.DataSource = dsDebitoren.Tables("tblDebiSubsFromUser").Select("strRGNr='" + dgvBookings.Rows(e.RowIndex).Cells("strDebRGNbr").Value + "'").CopyToDataTable
                intFctReturns = FcInitdgvDebiSub(dgvBookingSub)

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)

        Finally
            'dgvBookingSub.Update()

        End Try


    End Sub

    Private Sub dgvBookings_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvBookings.CellValueChanged

        Dim intDecidiveCell As Integer

        Try


            If dgvBookings.Columns(e.ColumnIndex).HeaderText = "ok" And e.RowIndex >= 0 Then

                If IIf(IsDBNull(dgvBookings.Rows(e.RowIndex).Cells("booDebBook").Value), False, dgvBookings.Rows(e.RowIndex).Cells("booDebBook").Value) Then

                    'MsgBox("Geändert zu checked " + dgvDebitoren.Rows(e.RowIndex).Cells("strDebRGNbr").Value + ", " + dgvDebitoren.Rows(e.RowIndex).Cells("booDebBook").Value.ToString + Val(dgvDebitoren.Rows(e.RowIndex).Cells("strDebStatusBitLog").Value).ToString)
                    'Zulassen? = keine Fehler
                    If Val(dgvBookings.Rows(e.RowIndex).Cells("strDebStatusBitLog").Value) <> 0 And Val(dgvBookings.Rows(e.RowIndex).Cells("strDebStatusBitLog").Value) <> 10000 Then
                        If Val(Strings.Mid(dgvBookings.Rows(e.RowIndex).Cells("strDebStatusBitLog").Value, 13, 1)) = 1 Then
                            MsgBox("Erst - RG Splitt-Bill ist nicht auffindbar. Wird trotzdem gebucht, wird nur die Zweit-RG gebucht.", vbOKOnly + vbExclamation, "Splitt-Bill No RG 1")
                        Else
                            MsgBox("Debi-Rechnung ist nicht buchbar.", vbOKOnly + vbExclamation, "Nicht buchbar")
                            dgvBookings.Rows(e.RowIndex).Cells("booDebBook").Value = False
                        End If
                    End If

                End If

            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Generelles Problem " + Err.Number.ToString)

        End Try


    End Sub

    Private Sub BgWLoadDebi_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BgWLoadDebi.DoWork



        Dim strIdentityName As String
        Dim strMDBName As String
        Dim strSQL As String
        Dim strSQLSub As String
        Dim strRGTableType As String
        Dim objdtLocDebiHead As New DataTable
        Dim objdtlocDebiSub As New DataTable
        Dim objdaolelocDebiSubs As New OleDb.OleDbDataAdapter
        Dim objdaolelocDebiHeads As New OleDb.OleDbDataAdapter
        Dim objdaolesubsselcomd As New OleDb.OleDbCommand
        Dim objdaoleheadselcomd As New OleDb.OleDbCommand
        Dim objdalocDebiSubs As New MySqlDataAdapter
        Dim objdalocDebiHeads As New MySqlDataAdapter
        Dim objdasubselcomd As New MySqlCommand
        Dim objdaheadselcomd As New MySqlCommand
        Dim objdslocdebisub As New DataSet
        Dim objdslocdebihead As New DataSet
        Dim strConnection As String
        Dim objRGMySQLConn As New MySqlConnection
        Dim objlocMySQLcmd As New MySqlCommand
        Dim strSQLToParse As String
        Dim objmysqlcomdwritehead As New MySqlCommand
        Dim intFcReturns As Int16
        Dim strmysqlSaveSub As String
        Dim objmysqlcomdwritesub As New MySqlCommand
        Dim intAccounting As Int16 = CInt(e.Argument)
        Dim objdbConnZHDB02 As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))
        Dim objdbAccessConn As New OleDb.OleDbConnection
        Dim objOLEdbcmdLoc As New OleDb.OleDbCommand


        Try

            Debug.Print("BW Start " + Convert.ToString(intAccounting))
            objmysqlcomdwritehead.Connection = objdbConnZHDB02
            objmysqlcomdwritesub.Connection = objdbConnZHDB02

            'Für den Save der Records
            strIdentityName = System.Security.Principal.WindowsIdentity.GetCurrent().Name
            strIdentityName = Strings.Replace(strIdentityName, "\", "/")

            strMDBName = Main.FcReadFromSettingsII("Buchh_RGTableMDB",
                                                        intAccounting)

            strSQL = Main.FcReadFromSettingsII("Buchh_SQLHead",
                                                 intAccounting)

            strRGTableType = Main.FcReadFromSettingsII("Buchh_RGTableType",
                                                         intAccounting)
            objdslocdebihead.EnforceConstraints = False
            objdslocdebihead.AcceptChanges()

            Debug.Print("BW Before Read Head " + Convert.ToString(intAccounting))
            If strRGTableType = "A" Then

                'Access
                Call Main.FcInitAccessConnecation(objdbAccessConn,
                                                  strMDBName)
                objdaoleheadselcomd.Connection = objdbAccessConn
                objdaoleheadselcomd.CommandText = strSQL
                objdaolelocDebiHeads.SelectCommand = objdaoleheadselcomd
                objdaolelocDebiHeads.SelectCommand.Connection.Open()
                objdaolelocDebiHeads.Fill(objdslocdebihead, "tbldebihead")
                objdaolelocDebiHeads.SelectCommand.Connection.Close()
                'objdbAccessConn.Open()
                'objOLEdbcmdLoc.CommandText = strSQL
                'objOLEdbcmdLoc.Connection = objdbAccessConn
                'objdtLocDebiHead.Load(objOLEdbcmdLoc.ExecuteReader)
                'objdbAccessConn.Close()
            ElseIf strRGTableType = "M" Then

                strConnection = System.Configuration.ConfigurationManager.AppSettings(strMDBName)
                objRGMySQLConn.ConnectionString = strConnection
                objdaheadselcomd.Connection = objRGMySQLConn
                objdaheadselcomd.CommandText = strSQL
                objdalocDebiHeads.SelectCommand = objdaheadselcomd
                objdalocDebiHeads.SelectCommand.Connection.Open()
                objdalocDebiHeads.Fill(objdslocdebihead, "tbldebihead")
                objdalocDebiHeads.SelectCommand.Connection.Close()
                'objlocMySQLcmd.Connection = objRGMySQLConn
                'objlocMySQLcmd.CommandText = strSQL
                'objRGMySQLConn.Open()
                'objdtLocDebiHead.Load(objlocMySQLcmd.ExecuteReader)
                'objRGMySQLConn.Close()

            Else
                MessageBox.Show("Tabletype not A or M")
                Exit Sub
            End If

            objdslocdebisub.EnforceConstraints = False
            objdslocdebisub.AcceptChanges()

            strSQLToParse = Main.FcReadFromSettingsII("Buchh_SQLDetail",
                                                        intAccounting)

            intFcReturns = Main.FcInitInsCmdDHeads(objmysqlcomdwritehead)

            Debug.Print("BW Write Heads " + Convert.ToString(intAccounting))
            For Each row As DataRow In objdslocdebihead.Tables("tbldebihead").Rows

                objmysqlcomdwritehead.Connection.Open()
                objmysqlcomdwritehead.Parameters("@IdentityName").Value = strIdentityName
                objmysqlcomdwritehead.Parameters("@ProcessID").Value = Process.GetCurrentProcess().Id
                objmysqlcomdwritehead.Parameters("@intBuchhaltung").Value = intAccounting
                objmysqlcomdwritehead.Parameters("@strDebRGNbr").Value = row("strDebRGNbr")
                objmysqlcomdwritehead.Parameters("@intBuchungsart").Value = row("intBuchungsart")
                objmysqlcomdwritehead.Parameters("@intRGArt").Value = row("intRGArt")
                If row.Table.Columns.Contains("strRGArt") Then
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
                If row.Table.Columns.Contains("strDebReferenz") Then
                    objmysqlcomdwritehead.Parameters("@strDebreferenz").Value = row("strDebReferenz")
                End If
                objmysqlcomdwritehead.Parameters("@datDebRGDatum").Value = row("datDebRGDatum")
                objmysqlcomdwritehead.Parameters("@datDebValDatum").Value = row("datDebValDatum")
                If row.Table.Columns.Contains("datRGCreate") Then
                    objmysqlcomdwritehead.Parameters("@datRGCreate").Value = row("datRGCreate")
                End If
                If row.Table.Columns.Contains("intPayType") Then
                    objmysqlcomdwritehead.Parameters("@intPayType").Value = row("intPayType")
                End If
                objmysqlcomdwritehead.Parameters("@strDebiBank").Value = row("strDebiBank")
                objmysqlcomdwritehead.Parameters("@lngLinkedRG").Value = row("lngLinkedRG")
                objmysqlcomdwritehead.Parameters("@strRGName").Value = row("strRGName")
                If row.Table.Columns.Contains("strDebIdentNbr2") Then
                    objmysqlcomdwritehead.Parameters("@strDebIdentNbr2").Value = row("strDebIdentNbr2")
                End If
                If row.Table.Columns.Contains("booCrToInv") Then
                    objmysqlcomdwritehead.Parameters("@booCrToInv").Value = row("booCrToInv")
                End If
                If row.Table.Columns.Contains("datPGVFrom") Then
                    objmysqlcomdwritehead.Parameters("@datPGVFrom").Value = row("datPGVFrom")
                End If
                If row.Table.Columns.Contains("datPGVTo") Then
                    objmysqlcomdwritehead.Parameters("@datPGVTo").Value = row("datPGVTo")
                End If
                objmysqlcomdwritehead.ExecuteNonQuery()
                objmysqlcomdwritehead.Connection.Close()
                objdtLocDebiHead.AcceptChanges()

                'Subs einlesen
                'Es muss der Weg über das DS genommen werden wegen den constraint-Verlethzungen
                strSQLSub = Main.FcSQLParse2(strSQLToParse,
                                                       row("strDebRGNbr"),
                                                       objdslocdebihead.Tables("tbldebihead"),
                                                       "D")

                If strRGTableType = "A" Then
                    objdaolesubsselcomd.CommandText = strSQLSub
                    objdaolesubsselcomd.Connection = objdbAccessConn
                    objdaolelocDebiSubs.SelectCommand = objdaolesubsselcomd
                    objdaolelocDebiSubs.SelectCommand.Connection.Open()
                    objdaolelocDebiSubs.Fill(objdslocdebisub, "tbldebisubs")
                    objdaolelocDebiSubs.SelectCommand.Connection.Close()
                    'objdbAccessConn.Open()
                    'objOLEdbcmdLoc.CommandText = strSQLSub
                    'objdtlocDebiSub.Load(objOLEdbcmdLoc.ExecuteReader)
                    'objdbAccessConn.Close()
                ElseIf strRGTableType = "M" Then
                    objdasubselcomd.CommandText = strSQLSub
                    objdasubselcomd.Connection = objRGMySQLConn
                    objdalocDebiSubs.SelectCommand = objdasubselcomd
                    'objdalocDebiSubs.SelectCommand.Connection = objRGMySQLConn
                    'objdalocDebiSubs.SelectCommand.CommandText = strSQLSub
                    objdalocDebiSubs.SelectCommand.Connection.Open()
                    objdalocDebiSubs.Fill(objdslocdebisub, "tbldebisubs")
                    objdalocDebiSubs.SelectCommand.Connection.Close()
                    'objRGMySQLConn.Open()
                    'objdtlocDebiSub.Load(objlocMySQLcmd.ExecuteReader)
                    'objRGMySQLConn.Close()
                End If



            Next
            If Not IsNothing(objdslocdebisub.Tables("tbldebisubs")) Then

                Debug.Print("BW Write Subs")
                'Subs schreiben
                intFcReturns = Main.FcInitInscmdSubs(objmysqlcomdwritesub)
                'For Each drsub As DataRow In objdtlocDebiSub.Rows
                For Each drsub As DataRow In objdslocdebisub.Tables("tbldebisubs").Rows

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

            End If


            'Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Debitoren-Daten-Lesen Problem", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            'Return 1

        Finally
            objdbAccessConn = Nothing
            objRGMySQLConn = Nothing

            objdslocdebihead = Nothing
            objdslocdebisub = Nothing

            objdalocDebiSubs = Nothing
            objdalocDebiHeads = Nothing

            objdaolelocDebiSubs = Nothing
            objdaolelocDebiHeads = Nothing

            objdaoleheadselcomd = Nothing
            objdaheadselcomd = Nothing
            objdaolesubsselcomd = Nothing
            objdasubselcomd = Nothing

            strConnection = Nothing
            'System.GC.Collect()

            Debug.Print("BW finsih " + Convert.ToString(intAccounting))

        End Try

    End Sub

    Private Sub BgWCheckDebi_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BgWCheckDebi.DoWork

        Dim BgWCheckDebiArgsInProc As BgWCheckDebitArgs = e.Argument

        Dim strMandant As String
        Dim booAccOk As Boolean
        Dim objFinanz As New SBSXASLib.AXFinanz
        Dim objfiBuha As New SBSXASLib.AXiFBhg
        Dim objdbBuha As New SBSXASLib.AXiDbBhg
        Dim objdbPIFb As New SBSXASLib.AXiPlFin
        Dim objFiBebu As New SBSXASLib.AXiBeBu

        Dim strBitLog As String = String.Empty
        Dim intReturnValue As Integer
        Dim strStatus As String = String.Empty
        Dim booAutoCorrect As Boolean
        Dim booCpyKSTToSub As Boolean
        Dim booSplittBill As Boolean
        Dim booCashSollCorrect As Boolean
        Dim booGeneratePymentBooking As Boolean
        Dim strRGNbr As String
        Dim intDebitorNew As Int32
        Dim intSubNumber As Int16
        Dim dblSubNetto As Double
        Dim dblSubMwSt As Double
        Dim dblSubBrutto As Double
        Dim dblRDiffNetto As Double
        Dim dblRDiffMwSt As Double
        Dim dblRDiffBrutto As Double
        Dim decDebiDiff As Decimal
        Dim strDebiReferenz As String
        Dim booPKPrivate As Boolean
        Dim booValutaCorrect As Boolean
        Dim datValutaCorrect As Date
        Dim booDateChanged As Boolean
        Dim datValutaSave As Date
        Dim intMonthsAJ As Int16
        Dim intMonthsNJ As Int16
        Dim intPGVMonths As Int16
        Dim intiBankSage200 As Int32
        Dim intLinkedDebitor As Int32
        Dim intSBGegenKonto As Int32
        Dim selSBrows() As DataRow
        Dim intDZKond As Int32
        Dim intDZKondS200 As Int32
        Dim booDiffHeadText As Boolean
        Dim booDiffSubText As Boolean
        Dim strDebiHeadText As String
        Dim strDebiSubText As String
        Dim selsubrow() As DataRow

        Try

            Debug.Print("Start Check " + Convert.ToString(BgWCheckDebiArgsInProc.intMandant))
            'Finanz-Obj init
            'Login
            Call objFinanz.ConnectSBSdb(System.Configuration.ConfigurationManager.AppSettings("OwnSageServer"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageDB"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageID"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSagePsw"), "")

            strMandant = Main.FcReadFromSettingsII("Buchh200_Name",
                                            BgWCheckDebiArgsInProc.intMandant)

            booAccOk = objFinanz.CheckMandant(strMandant)
            'Open Mandantg
            objFinanz.OpenMandant(strMandant, BgWCheckDebiArgsInProc.strPeriode)

            objfiBuha = objFinanz.GetFibuObj()
            objdbBuha = objFinanz.GetDebiObj()
            objdbPIFb = objfiBuha.GetCheckObj()
            objFiBebu = objFinanz.GetBeBuObj()

            'Variablen einlesen
            booAutoCorrect = Convert.ToBoolean(Convert.ToInt16(Main.FcReadFromSettingsII("Buchh_HeadAutoCorrect", BgWCheckDebiArgsInProc.intMandant)))
            booCpyKSTToSub = Convert.ToBoolean(Convert.ToInt16(Main.FcReadFromSettingsII("Buchh_KSTHeadToSub", BgWCheckDebiArgsInProc.intMandant)))
            booSplittBill = Convert.ToBoolean(Convert.ToInt16(Main.FcReadFromSettingsII("Buchh_LinkedBookings", BgWCheckDebiArgsInProc.intMandant)))
            booCashSollCorrect = Convert.ToBoolean(Convert.ToInt16(Main.FcReadFromSettingsII("Buchh_CashSollKontoKorr", BgWCheckDebiArgsInProc.intMandant)))
            booGeneratePymentBooking = Convert.ToBoolean(Convert.ToInt16(Main.FcReadFromSettingsII("Buchh_GeneratePaymentBooking", BgWCheckDebiArgsInProc.intMandant)))
            booDiffHeadText = IIf(Main.FcReadFromSettingsII("Buchh_TextSpecial", BgWCheckDebiArgsInProc.intMandant) = "0", False, True)
            booDiffSubText = IIf(Main.FcReadFromSettingsII("Buchh_SubTextSpecial", BgWCheckDebiArgsInProc.intMandant) = "0", False, True)
            booPKPrivate = IIf(Main.FcReadFromSettingsII("Buchh_PKTable", BgWCheckDebiArgsInProc.intMandant) = "t_customer", True, False)
            booValutaCorrect = BgWCheckDebiArgsInProc.booValutaCor
            datValutaCorrect = BgWCheckDebiArgsInProc.datValutaCor
            dsDebitoren.Tables("tblDebiHeadsFromUser").AcceptChanges()

            For Each row As DataRow In dsDebitoren.Tables("tblDebiHeadsFromUser").Rows

                'If row("strDebRGNbr") = "101261" Then Stop
                strRGNbr = row("strDebRGNbr") 'Für Error-Msg

                'Runden
                row("dblDebNetto") = Decimal.Round(row("dblDebNetto"), 4, MidpointRounding.AwayFromZero)
                row("dblDebMwSt") = Decimal.Round(row("dblDebMwst"), 2, MidpointRounding.AwayFromZero)
                row("dblDebBrutto") = Decimal.Round(row("dblDebBrutto"), 4, MidpointRounding.AwayFromZero)
                'RG-Create - Datum auf RG-Datum setzen falls nicht vorhanden
                If IsDBNull(row("datRGCreate")) Then
                    row("datRGCreate") = row("datDebRGDatum")
                End If
                'Credit To Debit (Gutschrift zu Rechnung) - Option auf false setzen falls nicht vorhanden
                If IsDBNull(row("booCrToInv")) Then
                    row("booCrToInv") = False
                End If
                'CreatePaymentBooking auf 0 setzen falls nicht vorhanden
                If IsDBNull(row("intKtoPayed")) Then
                    row("intKtoPayed") = 0
                End If

                'Status-String erstellen
                'Debitor 01
                intReturnValue = MainDebitor.FcGetRefDebiNr(IIf(IsDBNull(row("lngDebNbr")), 0, row("lngDebNbr")),
                                                BgWCheckDebiArgsInProc.intMandant,
                                                intDebitorNew)
                If intReturnValue = 1 Then 'Neue Debi-Nr wurde angelegt
                    strStatus = "NDeb "
                End If
                If intDebitorNew <> 0 Or intReturnValue = 4 Then
                    intReturnValue = MainDebitor.FcCheckDebitor(intDebitorNew,
                                                                row("intBuchungsart"),
                                                                objdbBuha)
                Else
                    intReturnValue = 2
                End If
                strBitLog = Trim(intReturnValue.ToString)

                'Kto 02
                'intReturnValue = FcCheckKonto(row("lngDebKtoNbr"), objfiBuha, row("dblDebMwSt"), 0)
                intReturnValue = 0
                strBitLog += Trim(intReturnValue.ToString)

                'Currency 03
                intReturnValue = Main.FcCheckCurrency(row("strDebCur"), objfiBuha)
                strBitLog += Trim(intReturnValue.ToString)

                'Sub 04
                If booSplittBill And IIf(IsDBNull(row("intRGArt")), 0, row("intRGArt")) = 10 Then
                    row("booLinked") = True
                Else
                    row("booLinked") = False
                End If

                intReturnValue = Main.FcCheckSubBookings(row("strDebRGNbr"),
                                                    dsDebitoren.Tables("tblDebiSubsFromUser"),
                                                    intSubNumber,
                                                    dblSubBrutto,
                                                    dblSubNetto,
                                                    dblSubMwSt,
                                                    objfiBuha,
                                                    objdbPIFb,
                                                    objFiBebu,
                                                    row("intBuchungsart"),
                                                    booAutoCorrect,
                                                    booCpyKSTToSub,
                                                    row("lngDebiKST"),
                                                    row("lngDebKtoNbr"),
                                                    booCashSollCorrect,
                                                    booSplittBill)

                strBitLog += Trim(intReturnValue.ToString)

                'Gibt es eine Bezahlbuchung zu erstellen? 
                'booGeneratePymentBooking = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettings(objdbconn, "Buchh_GeneratePaymentBooking", intAccounting)))
                If booGeneratePymentBooking And row("intBuchungsart") <> 1 And row("intKtoPayed") > 0 Then
                    dsDebitoren.Tables("tblDebiSubsFromUser").AcceptChanges()
                    'Bedingungen erfüllt
                    Dim drPaymentBuchung As DataRow = dsDebitoren.Tables("tblDebiSubsFromUser").NewRow
                    'Felder zuweisen
                    drPaymentBuchung("strRGNr") = row("strDebRGNbr")
                    drPaymentBuchung("intSollHaben") = 0
                    drPaymentBuchung("lngKto") = row("intKtoPayed")
                    drPaymentBuchung("strKtoBez") = "Bezahlung"
                    drPaymentBuchung("lngKST") = 0
                    drPaymentBuchung("strKstBez") = "keine"
                    drPaymentBuchung("lngProj") = 0
                    drPaymentBuchung("strProjBez") = "null"
                    drPaymentBuchung("dblNetto") = row("dblDebBrutto")
                    drPaymentBuchung("dblMwSt") = 0
                    drPaymentBuchung("dblBrutto") = row("dblDebBrutto")
                    drPaymentBuchung("dblMwStSatz") = 0
                    drPaymentBuchung("strMwStKey") = "null"
                    drPaymentBuchung("strArtikel") = "Bezahlvorgang"
                    drPaymentBuchung("strDebSubText") = "Bezahlvorgang"
                    dsDebitoren.Tables("tblDebiSubsFromUser").Rows.Add(drPaymentBuchung)
                    drPaymentBuchung = Nothing
                    'Summe der Sub-Buchungen anpassen
                    dblSubBrutto = Decimal.Round(dblSubBrutto + row("dblDebBrutto"), 2, MidpointRounding.AwayFromZero)

                    'dsDebitoren.Tables("tblDebiSubsFromUser").AcceptChanges()
                End If

                'Bei SplitBill - erste Rechnung evtl. Rückzahlung im Total nicht beachten
                If booSplittBill And row("intRGArt") = 1 And IIf(IsDBNull(row("lngLinkedRG")), 0, row("lngLinkedRG")) > 0 Then
                    row("dblDebBrutto") = Decimal.Round(dblSubBrutto, 2, MidpointRounding.AwayFromZero) * -1
                    row("dblDebNetto") = Decimal.Round(dblSubNetto, 2, MidpointRounding.AwayFromZero) * -1
                    row("dblDebMwSt") = Decimal.Round(dblSubMwSt, 2, MidpointRounding.AwayFromZero) * -1
                End If

                'Autokorrektur 05
                If booAutoCorrect And row("intBuchungsart") = 1 Then
                    decDebiDiff = 0
                    'Git es etwas zu korrigieren?
                    If IIf(IsDBNull(row("dblDebNetto")), 0, row("dblDebNetto")) <> dblSubNetto * -1 Or
                        IIf(IsDBNull(row("dblDebMwSt")), 0, row("dblDebMwSt")) <> dblSubMwSt * -1 Or
                        IIf(IsDBNull(row("dblDebBrutto")), 0, row("dblDebBrutto")) <> dblSubBrutto * -1 Then
                        If Math.Abs(row("dblDebBrutto") + dblSubBrutto) < 1 Then
                            decDebiDiff = Decimal.Round(Math.Abs(row("dblDebBrutto") + dblSubBrutto), 4, MidpointRounding.AwayFromZero)
                            row("dblDebBrutto") = Decimal.Round(dblSubBrutto, 4, MidpointRounding.AwayFromZero) * -1
                            row("dblDebNetto") = Decimal.Round(dblSubNetto, 4, MidpointRounding.AwayFromZero) * -1
                            row("dblDebMwSt") = Decimal.Round(dblSubMwSt, 4, MidpointRounding.AwayFromZero) * -1
                            strBitLog += "1"
                        Else
                            strBitLog += "3"
                        End If
                        ''In Sub korrigieren
                        'selsubrow = objdtDebitSubs.Select("strRGNr='" + row("strDebRGNbr") + "' AND intSollHaben=2")
                        'If selsubrow.Length = 1 Then
                        '    selsubrow(0).Item("dblBrutto") = dblSubBrutto * -1
                        '    selsubrow(0).Item("dblMwSt") = dblSubMwSt * -1
                        '    selsubrow(0).Item("dblNetto") = dblSubNetto * -1
                        'End If

                    Else
                        strBitLog += "0"
                    End If
                Else
                    If row("intBuchungsart") = 1 Then

                        dblRDiffBrutto = 0
                        If IIf(IsDBNull(row("dblDebMwSt")), 0, row("dblDebMwSt")) <> dblSubMwSt * -1 Then
                            row("dblDebMwSt") = dblSubMwSt * -1
                        End If
                        If IIf(IsDBNull(row("dblDebNetto")), 0, row("dblDebNetto")) <> dblSubNetto * -1 Then
                            row("dblDebNetto") = dblSubNetto * -1
                        End If

                        'Für evtl. Rundungsdifferenzen einen Datensatz in die Sub-Tabelle hinzufügen
                        If IIf(IsDBNull(row("dblDebBrutto")), 0, row("dblDebBrutto")) + dblSubBrutto <> 0 Then '0 _

                            dblRDiffBrutto = Decimal.Round(dblSubBrutto + row("dblDebBrutto"), 4, MidpointRounding.AwayFromZero)
                            dblRDiffMwSt = 0 'row("dblDebMwSt") - Decimal.Round(dblSubMwSt, 2, MidpointRounding.AwayFromZero)
                            dblRDiffNetto = 0 ' row("dblDebNetto") - Decimal.Round(dblSubNetto, 2, MidpointRounding.AwayFromZero)

                            dsDebitoren.Tables("tblDebiSubsFromUser").AcceptChanges()
                            'Zu sub-Table hinzifügen
                            Dim objdrDebiSub As DataRow = dsDebitoren.Tables("tblDebiSubsFromUser").NewRow
                            objdrDebiSub("strRGNr") = row("strDebRGNbr")
                            objdrDebiSub("intSollHaben") = 1
                            objdrDebiSub("lngKto") = 6906
                            objdrDebiSub("strKtoBez") = "Rundungsdifferenzen"
                            objdrDebiSub("lngKST") = 40
                            objdrDebiSub("strKstBez") = "SystemKST"
                            objdrDebiSub("dblNetto") = dblRDiffNetto
                            objdrDebiSub("dblMwSt") = dblRDiffMwSt
                            objdrDebiSub("dblBrutto") = dblRDiffBrutto * -1
                            objdrDebiSub("dblMwStSatz") = 0
                            objdrDebiSub("strMwStKey") = "null"
                            objdrDebiSub("strArtikel") = "Rundungsdifferenz"
                            objdrDebiSub("strDebSubText") = "Eingefügt"
                            objdrDebiSub("strStatusUBBitLog") = "00000000"
                            If Math.Abs(dblRDiffBrutto) > 6 Then
                                objdrDebiSub("strStatusUBText") = "Rund > 6"
                            Else
                                objdrDebiSub("strStatusUBText") = "ok"
                            End If
                            dsDebitoren.Tables("tblDebiSubsFromUser").Rows.Add(objdrDebiSub)
                            objdrDebiSub = Nothing

                            dsDebitoren.Tables("tblDebiSubsFromUser").AcceptChanges()

                            'Summe der Sub-Buchungen anpassen
                            dblSubBrutto = Decimal.Round(dblSubBrutto - dblRDiffBrutto, 2, MidpointRounding.AwayFromZero)
                            If Math.Abs(dblRDiffBrutto) > 6 Then
                                strBitLog += "3"
                            Else
                                strBitLog += "0"
                            End If
                        Else
                            strBitLog += "0"

                        End If

                    Else
                        strBitLog += "0"
                    End If

                End If

                'Diff Kopf - Sub? 06
                If row("intBuchungsart") = 1 Then 'OP
                    If IIf(IsDBNull(row("dblDebBrutto")), 0, row("dblDebBrutto")) + dblSubBrutto <> 0 _
                        Or IIf(IsDBNull(row("dblDebMwSt")), 0, row("dblDebMwSt")) + dblSubMwSt <> 0 _
                        Or IIf(IsDBNull(row("dblDebNetto")), 0, row("dblDebNetto")) + dblSubNetto <> 0 Then
                        strBitLog += "1"
                    Else
                        strBitLog += "0"
                    End If
                Else
                    'Test ob sub 0
                    If dblSubBrutto <> 0 Then
                        strBitLog += "1"
                    Else
                        strBitLog += "0"
                    End If

                End If

                'OP Kopf balanced? 07
                intReturnValue = Main.FcCheckBelegHead(row("intBuchungsart"),
                                                  IIf(IsDBNull(row("dblDebBrutto")), 0, row("dblDebBrutto")),
                                                  IIf(IsDBNull(row("dblDebNetto")), 0, row("dblDebNetto")),
                                                  IIf(IsDBNull(row("dblDebMwSt")), 0, row("dblDebMwSt")),
                                                  dblRDiffBrutto)
                strBitLog += Trim(intReturnValue.ToString)

                'Referenz 08
                If IIf(IsDBNull(row("strDebReferenz")), "", row("strDebReferenz")) = "" And row("intBuchungsart") = 1 Then
                    intReturnValue = Main.FcCreateDebRef(BgWCheckDebiArgsInProc.intMandant,
                                                    row("strDebiBank"),
                                                    row("strDebRGNbr"),
                                                    row("strOPNr"),
                                                    row("intBuchungsart"),
                                                    strDebiReferenz,
                                                    row("intPayType"))
                    If Len(strDebiReferenz) > 0 Then
                        row("strDebReferenz") = strDebiReferenz
                    Else
                        intReturnValue = 1
                    End If
                Else
                    strDebiReferenz = IIf(IsDBNull(row("strDebReferenz")), "", row("strDebReferenz"))
                    intReturnValue = 0
                End If
                strBitLog += Trim(intReturnValue.ToString)

                'Status-String auswerten, vorziehen um neue PK - Nummer auszulesen
                'booPKPrivate = IIf(Main.FcReadFromSettingsII("Buchh_PKTable", BgWCheckDebiArgsInProc.intMandant) = "t_customer", True, False)
                'Debitor
                If Strings.Left(strBitLog, 1) <> "0" Then
                    strStatus += "Deb"
                    If Strings.Left(strBitLog, 1) <> "2" Then
                        If booPKPrivate = True Then
                            intReturnValue = MainDebitor.FcIsPrivateDebitorCreatable(intDebitorNew,
                                                                                     objdbBuha,
                                                                                     BgWCheckDebiArgsInProc.strMandant,
                                                                                     BgWCheckDebiArgsInProc.intMandant)
                        Else
                            intReturnValue = MainDebitor.FcIsDebitorCreatable(intDebitorNew,
                                                                              objdbBuha,
                                                                              BgWCheckDebiArgsInProc.strMandant,
                                                                              BgWCheckDebiArgsInProc.intMandant)
                        End If
                        If intReturnValue = 0 Then
                            strStatus += " erstellt"
                            row("strDebBez") = MainDebitor.FcReadDebitorName(objdbBuha,
                                                                         intDebitorNew,
                                                                         row("strDebCur"))

                        ElseIf intReturnValue = 5 Then
                            strStatus += " not approved "
                            row("strDebBez") = "nap"
                        Else
                            strStatus += " nicht erstellt."
                            row("strDebBez") = "n/a"
                        End If
                        row("lngDebNbr") = intDebitorNew
                    Else
                        strStatus += " keine Ref"
                        row("strDebBez") = "n/a"
                    End If
                Else
                    'If row("intBuchungsart") = 1 Then
                    row("strDebBez") = MainDebitor.FcReadDebitorName(objdbBuha,
                                                                     intDebitorNew,
                                                                     row("strDebCur"))
                    'Falls Währung auf Debitor nicht definiert, dann Currency setzen
                    If row("strDebBez") = "EOF" And row("intBuchungsart") = 1 Then
                        strBitLog = Strings.Left(strBitLog, 2) + "1" + Strings.Right(strBitLog, Len(strBitLog) - 3)
                    End If
                    'Else
                    'row("strDebBez") = "Nicht relevant"
                    'End If
                    row("lngDebNbr") = intDebitorNew
                End If

                'OP - Verdopplung 09
                intReturnValue = Main.FcCheckOPDouble(objdbBuha,
                                                 row("lngDebNbr"),
                                                 row("strOPNr"),
                                                 IIf(row("dblDebBrutto") > 0, "R", "G"),
                                                 row("strDebCur"))
                strBitLog += Trim(intReturnValue.ToString)

                'PGV => Prüfung vor Valuta-Datum da Valuta-Datum verändert wird
                If Not IsDBNull(row("datPGVFrom")) Then
                    row("booPGV") = True
                End If

                'Bei Datum-Korrektur vorgängig Datum ersetzen um PGV-Buchungen zu verhindern
                If booValutaCorrect Then
                    If row("datDebRGDatum") < datValutaCorrect Then
                        row("datDebRGDatum") = datValutaCorrect.ToShortDateString
                        strStatus = "RgDCor"
                    End If
                    If row("datDebValDatum") < datValutaCorrect Then
                        row("datDebValDatum") = datValutaCorrect.ToShortDateString
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "ValDCor"
                    End If
                End If

                booDateChanged = False
                'Jahresübergreifend RG- / Valuta-Datum
                If Year(row("datDebRGDatum")) <> Year(row("datDebValDatum")) And Year(row("datDebValDatum")) >= 2023 Then
                    'Not IsDBNull(row("datPGVFrom")) Then
                    row("booPGV") = True
                    'datValutaPGV = row("datDebValDatum")
                    'Bei Valuta-Datum in einem anderen Jahr Valuta-Datum ändern
                    If Year(row("datDebRGDatum")) < Year(row("datDebValDatum")) Then
                        row("strPGVType") = "RV"
                    Else
                        row("strPGVType") = "VR"
                    End If
                    datValutaSave = row("datDebValDatum")

                    If IsDBNull(row("datPGVFrom")) Then
                        If row("strPGVType") = "VR" Then
                            row("datPGVFrom") = Year(datValutaSave).ToString + "-" + Month(datValutaSave).ToString + "-" + DateAndTime.Day(datValutaSave).ToString
                            row("datPGVTo") = Year(datValutaSave).ToString + "-" + Month(datValutaSave).ToString + "-" + DateAndTime.Day(datValutaSave).ToString
                            row("datDebValDatum") = "2024-01-01"
                            booDateChanged = True
                        ElseIf row("strPGVType") = "RV" Then
                            row("datPGVFrom") = Year(datValutaSave).ToString + "-" + Month(datValutaSave).ToString + "-" + DateAndTime.Day(datValutaSave).ToString
                            row("datPGVTo") = Year(datValutaSave).ToString + "-" + Month(datValutaSave).ToString + "-" + DateAndTime.Day(datValutaSave).ToString
                            row("datDebValDatum") = row("datDebRGDatum")
                        End If
                    Else
                        If row("strPGVType") = "RV" Then
                            row("datDebValDatum") = row("datDebRGDatum")
                            booDateChanged = True
                        Else
                            row("strPGVType") = "XX"
                        End If
                    End If
                End If

                If row("booPGV") Then

                    'Anzahl Monate prüfen
                    intMonthsAJ = 0
                    intMonthsNJ = 0

                    intPGVMonths = (DateAndTime.Year(row("datPGVTo")) * 12 + DateAndTime.Month(row("datPGVTo"))) - (DateAndTime.Year(row("datPGVFrom")) * 12 + DateAndTime.Month(row("datPGVFrom"))) + 1
                    For intMonthCounter = 0 To intPGVMonths - 1
                        If Year(DateAdd(DateInterval.Month, intMonthCounter, row("datPGVFrom"))) > Convert.ToInt32(BgWCheckDebiArgsInProc.strYear) Then
                            intMonthsNJ += 1
                        Else
                            intMonthsAJ += 1
                        End If
                    Next
                    row("intPGVMthsAY") = intMonthsAJ
                    row("intPGVMthsNY") = intMonthsNJ

                End If

                'Valuta - Datum 10
                intReturnValue = Main.FcCheckDate2(IIf(IsDBNull(row("datDebValDatum")), #1789-09-17#, row("datDebValDatum")),
                                              BgWCheckDebiArgsInProc.strYear,
                                              dsDebitoren.Tables("tblDebitorenDates"),
                                              False)

                'Bei PGV checken ob PGV-Startdatum in blockierter Periode
                If row("booPGV") And intReturnValue = 0 Then
                    'Ist TA ?
                    If intMonthsAJ + intMonthsNJ = 1 Then
                        'Ist Differenz Jahre grösser 1?
                        If Math.Abs(Convert.ToInt16(BgWCheckDebiArgsInProc.strYear) - Year(row("datPGVTo"))) > 1 Then
                            intReturnValue = 4
                        Else
                            intReturnValue = Main.FcCheckDate2(row("datPGVTo"),
                                                      BgWCheckDebiArgsInProc.strYear,
                                                      dsDebitoren.Tables("tblDebitorenDates"),
                                                      True)
                        End If
                    Else
                        'mehrere Monate PGV
                        For intMonthCounter = 0 To intPGVMonths - 1
                            'Ist Differenz Jahre grösser 1?
                            If Math.Abs(Convert.ToInt16(BgWCheckDebiArgsInProc.strYear) - Year(row("datPGVFrom"))) > 1 Then
                                intReturnValue = 4
                            Else
                                intReturnValue = Main.FcCheckDate2(DateAndTime.DateAdd(DateInterval.Month, intMonthCounter, row("datPGVFrom")),
                                                          BgWCheckDebiArgsInProc.strYear,
                                                          dsDebitoren.Tables("tblDebitorenDates"),
                                                          True)
                            End If
                            If intReturnValue <> 0 Then
                                Exit For
                            End If
                        Next
                    End If

                End If
                strBitLog += Trim(intReturnValue.ToString)

                'RG - Datum 11
                intReturnValue = Main.FcCheckDate2(IIf(IsDBNull(row("datDebRGDatum")), #1789-09-17#, row("datDebRGDatum")),
                                              BgWCheckDebiArgsInProc.strYear,
                                              dsDebitoren.Tables("tblDebitorenDates"),
                                              False)

                strBitLog += Trim(intReturnValue.ToString)

                'Interne Bank 12
                If IsDBNull(row("intPayType")) Then
                    row("intPayType") = 9
                End If
                intReturnValue = MainDebitor.FcCheckDebiIntBank(BgWCheckDebiArgsInProc.intMandant,
                                                                IIf(IsDBNull(row("strDebiBank")), "", row("strDebiBank")),
                                                                row("intPayType"),
                                                                intiBankSage200)
                strBitLog += Trim(intReturnValue.ToString)

                'Bei SplittBill: Existiert verlinkter Beleg? 13
                If row("booLinked") Then
                    dsDebitoren.Tables("tblDebiHeadsFromUser").AcceptChanges()
                    'Zuerst Debitor von erstem Beleg suchen
                    intDebitorNew = MainDebitor.FcGetDebitorFromLinkedRG(IIf(IsDBNull(row("lngLinkedRG")), 0, row("lngLinkedRG")),
                                                                         BgWCheckDebiArgsInProc.intMandant,
                                                                         intLinkedDebitor,
                                                                         BgWCheckDebiArgsInProc.intTeqNbr,
                                                                         BgWCheckDebiArgsInProc.intTeqNbrLY,
                                                                         BgWCheckDebiArgsInProc.intTeqNbrPLY)
                    row("lngLinkedDeb") = intLinkedDebitor

                    intReturnValue = MainDebitor.FcCheckLinkedRG(objdbBuha,
                                                                 intLinkedDebitor,
                                                                 row("strDebCur"),
                                                                 row("lngLinkedRG"))
                    'Falls erste Rechnung bezahlt, dann Flag setzen
                    If intReturnValue = 2 Then
                        row("booLinkedPayed") = True
                        intSBGegenKonto = 2331
                    Else
                        row("booLinkedPayed") = False
                        intSBGegenKonto = 1092
                    End If

                    'UB - Löschen und Buchung erstellen ohne MwSt und ohne KST da schon in RG 1 beinhaltet
                    selSBrows = dsDebitoren.Tables("tblDebiSubsFromUser").Select("strRGNr='" + row("strDebRGNbr") + "'")

                    dsDebitoren.Tables("tblDebiSubsFromUser").AcceptChanges()

                    For Each SBsubrow As DataRow In selSBrows
                        SBsubrow.Delete()
                    Next

                    Dim drSBBuchung As DataRow = dsDebitoren.Tables("tblDebiSubsFromUser").NewRow
                    'Felder zuweisen
                    drSBBuchung("strRGNr") = row("strDebRGNbr")
                    drSBBuchung("intSollHaben") = 1
                    drSBBuchung("lngKto") = intSBGegenKonto
                    drSBBuchung("strKtoBez") = "SB - Buchung"
                    drSBBuchung("lngKST") = 0
                    drSBBuchung("strKstBez") = "keine"
                    drSBBuchung("lngProj") = 0
                    drSBBuchung("strProjBez") = "null"
                    drSBBuchung("dblNetto") = row("dblDebBrutto") * -1
                    drSBBuchung("dblMwSt") = 0
                    drSBBuchung("dblBrutto") = row("dblDebBrutto") * -1
                    drSBBuchung("dblMwStSatz") = 0
                    drSBBuchung("strMwStKey") = "null"
                    drSBBuchung("strArtikel") = "SB - Buchung"
                    drSBBuchung("strDebSubText") = row("lngDebIdentNbr").ToString + ", FRG, " + row("lngLinkedRG").ToString
                    dsDebitoren.Tables("tblDebiSubsFromUser").Rows.Add(drSBBuchung)
                    drSBBuchung = Nothing

                Else
                    intReturnValue = 0
                End If
                strBitLog += Trim(intReturnValue.ToString)

                'Zahlungs-Kondition 14
                'Falls Zahlungskondition vorhanden von RG holen, sonst von Tab_Repbetrieben
                'intZKondT=1 = von Rep_Betrieben, 0=von t_payment...
                If IsDBNull(row("intZKondT")) Then
                    row("intZKondT") = 1
                End If
                If IsDBNull(row("intZKond")) Then
                    row("intZKond") = 0
                    intDZKond = 0
                Else
                    'ID in effektive Sage 200 umwandeln (=von Tabelle lesen)
                    intReturnValue = MainDebitor.FcGetDZKondSageID(row("intZKond"),
                                                                   intDZKondS200)
                    row("intZKond") = intDZKondS200
                End If
                If row("intZKondT") = 1 And row("intZKond") = 0 Then
                    'Fall kein Privatekunde
                    If booPKPrivate = False Then
                        'Daten aus den Tab_Repbetriebe holen
                        intReturnValue = MainDebitor.FcGetDZkondFromRep(row("lngDebNbr"),
                                                                    intDZKond,
                                                                    BgWCheckDebiArgsInProc.intMandant)
                    Else
                        'Daten aus der t_customer holen
                        intReturnValue = MainDebitor.FcGetDZkondFromCust(row("lngDebNbr"),
                                                                         intDZKond,
                                                                         BgWCheckDebiArgsInProc.intMandant)
                    End If
                    row("intZKond") = intDZKond
                End If
                'Prüfem ob Zahlungs-Kondition - ID existiert in Sage 200 bei Mandant
                'strMandant = Main.FcReadFromSettingsII("Buchh200_Name",
                '                                BgWCheckDebiArgsInProc.intMandant)
                intReturnValue = MainDebitor.FcCheckDZKond(strMandant,
                                                           row("intZKond"))
                strBitLog += Trim(intReturnValue.ToString)

                'Status-String auswerten
                ''Debitor
                'Wird vorher behandelt
                'Konto
                If Mid(strBitLog, 2, 1) <> "0" Then
                    If Mid(strBitLog, 2, 1) <> 2 Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Kto"
                    Else
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Kto MwSt"
                    End If
                    row("strDebKtoBez") = "n/a"
                Else
                    row("strDebKtoBez") = MainDebitor.FcReadDebitorKName(objfiBuha,
                                                                         row("lngDebKtoNbr"))
                End If
                'Währung
                If Mid(strBitLog, 3, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Cur"
                End If
                'Subbuchungen
                'Totale in Head schreiben
                row("intSubBookings") = intSubNumber.ToString
                row("dblSumSubBookings") = dblSubBrutto.ToString
                If Mid(strBitLog, 4, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Sub"
                End If
                'Autokorretkur
                If Mid(strBitLog, 5, 1) <> "0" Then
                    If Mid(strBitLog, 5, 1) = "1" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "AutoC " + decDebiDiff.ToString
                    ElseIf Mid(strBitLog, 5, 1) = "2" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Round"
                        'Wieder auf 1 setzen damit Beleg gebucht werden kann
                        Mid(strBitLog, 5, 1) = "1"
                    ElseIf Mid(strBitLog, 5, 1) = "3" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Rnd>5"
                    End If
                End If
                'Diff zu Subbuchungen
                If Mid(strBitLog, 6, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "DiffS"
                End If
                'OP Kopf
                If Mid(strBitLog, 7, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "BelK"
                End If
                'Referenz
                If Mid(strBitLog, 8, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Ref"
                    'Else
                    '    row("strDebRef") = strDebiReferenz
                End If
                'OP
                If Mid(strBitLog, 9, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "OPDbl"
                    'Else
                    '    row("strDebRef") = strDebiReferenz
                End If
                'Valuta Datum 
                If Mid(strBitLog, 10, 1) <> "0" Then
                    If Mid(strBitLog, 10, 1) = "1" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "ValD"
                    ElseIf Mid(strBitLog, 10, 1) = "2" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "VDBlck"
                    ElseIf Mid(strBitLog, 10, 1) = "3" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "PGVBlck"
                    ElseIf Mid(strBitLog, 10, 1) = "4" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "PGVYear>1"
                    End If
                End If
                'RG Datum 
                If Mid(strBitLog, 11, 1) <> "0" Then
                    If Mid(strBitLog, 11, 1) = "1" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "RgD"
                    ElseIf Mid(strBitLog, 11, 1) = "2" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "RgDBlck"
                    End If
                End If
                'interne Bank
                If Mid(strBitLog, 12, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "iBnk"
                Else
                    row("strDebiBank") = intiBankSage200
                End If
                'Splitt-Bill
                If Mid(strBitLog, 13, 1) <> "0" Then
                    If Mid(strBitLog, 13, 1) = "1" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "SplBNo1"
                    Else
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "SplBBez1"
                    End If

                End If
                'Zahlungs-Kondition
                If Mid(strBitLog, 14, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "ZKond"
                End If

                'PGV keine Ziffer
                If row("booPGV") Then
                    If row("intPGVMthsAY") + row("intPGVMthsNY") = 1 Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "TA " + row("strPGVType")
                    Else
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "PGV " + row("strPGVType")
                    End If
                End If

                'Status schreiben
                '5 Autokorrektur trotzdem ok, 10 Valuta 2 trotzdem ok, 11 RG 2 trotdem ok
                If Val(strBitLog) = 0 Or Val(strBitLog) = 1000002200 Or Val(strBitLog) = 2200 Or Val(strBitLog) = 1000000000 Then
                    row("booDebBook") = True
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "ok"
                End If
                row("strDebStatusText") = strStatus
                row("strDebStatusBitLog") = strBitLog

                'Wird ein anderer Text in der Head-Buchung gewünscht?
                'booDiffHeadText = IIf(Main.FcReadFromSettingsII("Buchh_TextSpecial", intAccounting) = "0", False, True)
                If booDiffHeadText Then
                    dsDebitoren.Tables("tblDebiHeadsFromUser").AcceptChanges()
                    strDebiHeadText = MainDebitor.FcSQLParse(Main.FcReadFromSettingsII("Buchh_TextSpecialText",
                                                                                BgWCheckDebiArgsInProc.intMandant),
                                                             row("strDebRGNbr"),
                                                             dsDebitoren.Tables("tblDebiHeadsFromUser"),
                                                             "D")
                    row("strDebText") = strDebiHeadText
                End If

                'Wird ein anderer Text in den Sub-Buchung gewünscht?
                'booDiffSubText = IIf(Main.FcReadFromSettingsII("Buchh_SubTextSpecial", intAccounting) = "0", False, True)
                If booDiffSubText And Not row("booLinked") Then
                    strDebiSubText = MainDebitor.FcSQLParse(Main.FcReadFromSettingsII("Buchh_SubTextSpecialText",
                                                                               BgWCheckDebiArgsInProc.intMandant),
                                                            row("strDebRGNbr"),
                                                            dsDebitoren.Tables("tblDebiHeadsFromUser"),
                                                            "D")
                Else
                    strDebiSubText = row("strDebText")
                End If
                'Falls nicht SB - Linked dann Text in SB ersetzen
                If Not row("booLinked") Then
                    dsDebitoren.Tables("tblDebiSubsFromUser").AcceptChanges()
                    selsubrow = dsDebitoren.Tables("tblDebiSubsFromUser").Select("strRGNr='" + row("strDebRGNbr") + "'")
                    For Each subrow In selsubrow
                        subrow("strDebSubText") = strDebiSubText
                    Next
                End If

                'Init
                strBitLog = String.Empty
                strStatus = String.Empty
                intSubNumber = 0
                dblSubBrutto = 0
                dblSubNetto = 0
                dblSubMwSt = 0
                intDZKond = 0

            Next

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Debitoren-Daten-Lesen Problem", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        Finally
            objFinanz = Nothing
            objfiBuha = Nothing
            objdbBuha = Nothing
            objdbPIFb = Nothing
            objFiBebu = Nothing
            selSBrows = Nothing
            selsubrow = Nothing

            System.GC.Collect()
            Debug.Print("End Check " + Convert.ToString(BgWCheckDebiArgsInProc.intMandant))

        End Try


    End Sub

    Private Sub BgWImportDebi_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BgWImportDebi.DoWork

        Dim BgWImportDebiArgsInProc As BgWCheckDebitArgs = e.Argument
        Dim intReturnValue As Int32
        Dim objdbConnZHDB02 As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))
        Dim objdbMSSQLConn As New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("SQLConnectionString"))
        Dim objdbSQLcommand As New SqlCommand
        Dim objdbAccessConn As New OleDb.OleDbConnection
        Dim objOracleConn As New OracleConnection("Data Source=(DESCRIPTION=" _
                        + "(ADDRESS=(PROTOCOL=TCP)(HOST=10.0.0.29)(PORT=1521))" _
                        + "(CONNECT_DATA=(SERVICE_NAME=CISNEW)));" _
                        + "User Id=cis;Password=sugus;")

        Dim booErfOPExt As Boolean
        Dim strMandant As String
        Dim booAccOk As Boolean
        Dim strPeriode As String = BgWImportDebiArgsInProc.strPeriode
        Dim intDebBelegsNummer As Int32
        Dim strExtBelegNbr As String = String.Empty
        Dim strBuchType As String
        Dim strDebiLine As String
        Dim strDebitor() As String
        Dim strSachBID As String
        Dim intDebitorNbr As Int32
        Dim strValutaDatum As String
        Dim strBelegDatum As String
        Dim strVerfallDatum As String
        Dim strReferenz As String
        Dim strMahnerlaubnis As String
        Dim dblBetrag As Double
        Dim strDebiText As String
        Dim strCurrency As String
        Dim dblKurs As Double
        Dim intEigeneBank As Int16
        Dim intKondition As Int16
        Dim booBooingok As Boolean
        Dim strVerkID As String = String.Empty
        Dim sngAktuelleMahnstufe As Single
        Dim strSkonto As String = String.Empty
        Dim strErrMessage As String
        Dim strRGNbr As String
        Dim selDebiSub() As DataRow
        Dim intGegenKonto As Int32
        Dim strFibuText As String
        Dim dblNettoBetrag As Double
        Dim strBeBuEintrag As String
        Dim strSteuerFeld As String
        Dim intLaufNbr As Int32
        Dim strBeleg As String
        Dim strBelegArr() As String
        Dim dblSplitPayed As Double
        Dim dblSollBetrag As Double
        Dim dblHabenBetrag As Double
        Dim strBeBuEintragSoll As String
        Dim strBeBuEintragHaben As String
        Dim strSteuerFeldHaben As String
        Dim strSteuerFeldSoll As String
        Dim intSollKonto As Int32
        Dim dblKursSoll As Double
        Dim strDebiTextSoll As String
        Dim intHabenKonto As Int32
        Dim dblKursHaben As Double
        Dim strDebiTextHaben As String
        Dim intCommonKonto As Int32
        Dim strDebiCurrency As String
        Dim strKrediCurrency As String
        Dim dblBuchBetrag As Double
        Dim dblBasisBetrag As Double
        Dim strErfassungsDatum As String


        Dim objFinanz As New SBSXASLib.AXFinanz
        Dim objfiBuha As New SBSXASLib.AXiFBhg
        Dim objdbBuha As New SBSXASLib.AXiDbBhg
        Dim objdbPIFb As New SBSXASLib.AXiPlFin
        Dim objFiBebu As New SBSXASLib.AXiBeBu
        Dim objKrBuha As New SBSXASLib.AXiKrBhg

        Try

            'Me.Cursor = Cursors.WaitCursor
            'Button deaktivieren
            'Me.butImport.Enabled = False

            'Finanz-Obj init
            'Login
            Call objFinanz.ConnectSBSdb(System.Configuration.ConfigurationManager.AppSettings("OwnSageServer"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageDB"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageID"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSagePsw"), "")

            strMandant = Main.FcReadFromSettingsII("Buchh200_Name",
                                            BgWImportDebiArgsInProc.intMandant)

            booAccOk = objFinanz.CheckMandant(strMandant)
            'Open Mandantg
            objFinanz.OpenMandant(strMandant, strPeriode)
            objfiBuha = objFinanz.GetFibuObj()
            objdbBuha = objFinanz.GetDebiObj()
            objdbPIFb = objfiBuha.GetCheckObj()
            objFiBebu = objFinanz.GetBeBuObj()
            objKrBuha = objFinanz.GetKrediObj()


            'Start in Sync schreiben
            'intReturnValue = WFDBClass.FcWriteStartToSync(objdbConnZHDB02,
            '                                              BgWImportDebiArgsInProc.intMandant,
            '                                              1,
            '                                              dsDebitoren.Tables("tblDebiHeadsFromUser").Rows.Count)

            'Setting soll erfasste OP als externe Beleg-Nr. genommen werden und lngDebIdentNbr als Beleg-Nr.
            booErfOPExt = Convert.ToBoolean(Convert.ToInt16(Main.FcReadFromSettingsII("Buchh_ErfOPExt", BgWImportDebiArgsInProc.intMandant)))

            'Kopfbuchung
            For Each row In Me.dsDebitoren.Tables("tblDebiHeadsFromUser").Rows

                If IIf(IsDBNull(row("booDebBook")), False, row("booDebBook")) Then

                    'Für Err-Msg
                    strRGNbr = row("strDebRGNbr")

                    'Test ob OP - Buchung
                    If row("intBuchungsart") = 1 Then

                        'Verdopplung interne BelegsNummer verhindern
                        objdbBuha.CheckDoubleIntBelNbr = "J"

                        If row("dblDebBrutto") < 0 Then
                            'Gutschrift
                            'Falls booGSToInv (Gutschrift zu Rechnung) dann OP-Nummer vorgeben, sonst hochzählen lassen
                            If row("booCrToInv") Then
                                'Beleg-Nummerierung desaktivieren
                                objdbBuha.IncrBelNbr = "N"
                                'Eingelesene OP-Nummer (=Verknüpfte OP-Nr.) = interne Beleg-Nummer
                                intDebBelegsNummer = Main.FcCleanRGNrStrict(row("strOPNr"))
                                strExtBelegNbr = row("strDebRGNbr")
                            Else
                                'Zuerst Beleg-Nummerieungung aktivieren
                                objdbBuha.IncrBelNbr = "J"
                                'Belegsnummer abholen
                                intDebBelegsNummer = objdbBuha.GetNextBelNbr("G")
                                'Prüfen ob wirklich frei und falls nicht hochzählen
                                intReturnValue = MainDebitor.FcCheckDebiExistance(intDebBelegsNummer,
                                                                                  "G",
                                                                                  BgWImportDebiArgsInProc.intTeqNbr)

                                strExtBelegNbr = row("strOPNr")
                            End If

                            'Beträge drehen
                            row("dblDebBrutto") = row("dblDebBrutto") * -1
                            row("dblDebMwSt") = row("dblDebMwSt") * -1
                            row("dblDebNetto") = row("dblDebNetto") * -1

                            strBuchType = "G"

                        Else 'RG - Buchung

                            If String.IsNullOrEmpty(row("strOPNr")) Then
                                'strExtBelegNbr = row("strOPNr")

                                'Zuerst Beleg-Nummerieungung aktivieren
                                objdbBuha.IncrBelNbr = "J"
                                'Belegsnummer abholen
                                intDebBelegsNummer = objdbBuha.GetNextBelNbr("R")
                                intReturnValue = MainDebitor.FcCheckDebiExistance(intDebBelegsNummer,
                                                                                  "R",
                                                                                  BgWImportDebiArgsInProc.intTeqNbr)
                            Else
                                If Strings.Len(Main.FcCleanRGNrStrict(row("strOPNr"))) > 9 Then
                                    'Zahl zu gross
                                    objdbBuha.IncrBelNbr = "J"
                                    'Belegsnummer abholen
                                    intDebBelegsNummer = objdbBuha.GetNextBelNbr("R")
                                    intReturnValue = MainDebitor.FcCheckDebiExistance(intDebBelegsNummer,
                                                                                      "R",
                                                                                      BgWImportDebiArgsInProc.intTeqNbr)
                                    strExtBelegNbr = row("strOPNr")
                                Else
                                    'Beleg-Nummerierung abschalten
                                    objdbBuha.IncrBelNbr = "N"
                                    'Gemäss Setting Erfasste OP-Nr. Nummern vergeben
                                    If Not booErfOPExt Then
                                        intDebBelegsNummer = Main.FcCleanRGNrStrict(row("strOPNr"))
                                        strExtBelegNbr = row("strOPNr")
                                    Else
                                        'bei t_debi: IdentNbr wird genommen da dort die RG-Nr. drin ist. RgNr = ID
                                        intDebBelegsNummer = row("lngDebIdentNbr")
                                        strExtBelegNbr = row("strOPNr")
                                    End If

                                End If

                            End If

                            strBuchType = "R"

                        End If

                        'Variablen zuweisen
                        'Sachbearbeiter aus Debitor auslesen
                        strDebiLine = objdbBuha.ReadDebitor3(row("lngDebNbr") * -1, "")
                        strDebitor = Split(strDebiLine, "{>}")
                        strSachBID = strDebitor(30)
                        intDebitorNbr = row("lngDebNbr")
                        strValutaDatum = Format(row("datDebValDatum"), "yyyyMMdd").ToString
                        strBelegDatum = Format(row("datDebRGDatum"), "yyyyMMdd").ToString
                        If IsDBNull(row("datDebDue")) Then
                            strVerfallDatum = String.Empty
                        Else
                            strVerfallDatum = Format(row("datDebDue"), "yyyyMMdd").ToString
                        End If
                        strReferenz = row("strDebReferenz")
                        strMahnerlaubnis = String.Empty
                        dblBetrag = row("dblDebBrutto")
                        'Bei SplittBill 2ter Rechnung Text anfügen
                        If row("booLinked") Then
                            strDebiText = row("strDebText") + ", FRG "
                        Else
                            strDebiText = row("strDebText")
                        End If
                        strCurrency = row("strDebCur")
                        If strCurrency <> "CHF" Then 'Muss ergänzt werden => Was ist Leitwährung auf dem Konto
                            dblKurs = Main.FcGetKurs(strCurrency,
                                                     strValutaDatum,
                                                     objfiBuha)
                        Else
                            dblKurs = 1.0#
                        End If
                        intEigeneBank = row("strDebiBank")
                        'Zahl-Kondition
                        intKondition = IIf(IsDBNull(row("intZKond")), 1, row("intZKond"))

                        Try
                            booBooingok = True
                            Call objdbBuha.SetBelegKopf2(intDebBelegsNummer,
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

                            'Application.DoEvents()

                        Catch ex As Exception
                            strErrMessage = "Problem " + (Err.Number And 65535).ToString + " Belegkopf zu" + intDebBelegsNummer.ToString + vbCrLf
                            strErrMessage += "RG " + strRGNbr + vbCrLf
                            strErrMessage += "Debitor " + intDebitorNbr.ToString

                            MessageBox.Show(ex.Message + vbCrLf + strErrMessage, "Problem " + (Err.Number And 65535).ToString, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            If (Err.Number And 65535) < 10000 Then
                                booBooingok = False
                            Else
                                booBooingok = True
                            End If

                        End Try

                        'Verteilung
                        selDebiSub = dsDebitoren.Tables("tblDebiSubsFromUser").Select("strRGNr='" + row("strDebRGNbr") + "'")
                        For Each SubRow As DataRow In selDebiSub

                            intGegenKonto = SubRow("lngKto")
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

                            If SubRow("lngKST") > 0 Then
                                strBeBuEintrag = SubRow("lngKST").ToString + "{<}" + SubRow("strDebSubText") + "{<}" + "CALCULATE" + "{>}"    '"PROD{<}BebuText{<}" + dblBebuBetrag.ToString + "{>}"
                            Else
                                strBeBuEintrag = Nothing
                            End If

                            If Not IsDBNull(SubRow("strMwStKey")) And
                                        SubRow("strMwStKey") <> "null" And
                                        SubRow("lngKto") <> 6906 Then
                                If strBuchType = "R" Then
                                    strSteuerFeld = Main.FcGetSteuerFeld(objfiBuha,
                                                                         SubRow("lngKto"),
                                                                         SubRow("strDebSubText"),
                                                                         SubRow("dblBrutto") * -1,
                                                                         SubRow("strMwStKey"),
                                                                         SubRow("dblMwSt") * -1)
                                Else
                                    strSteuerFeld = Main.FcGetSteuerFeld(objfiBuha,
                                                                         SubRow("lngKto"),
                                                                         SubRow("strDebSubText"),
                                                                         SubRow("dblBrutto"),
                                                                         SubRow("strMwStKey"),
                                                                         SubRow("dblMwSt"))
                                End If
                            Else
                                strSteuerFeld = "STEUERFREI"
                            End If

                            Try

                                booBooingok = True
                                Call objdbBuha.SetVerteilung(intGegenKonto.ToString,
                                                             strFibuText,
                                                             dblNettoBetrag.ToString,
                                                             strSteuerFeld,
                                                             strBeBuEintrag)

                                'Application.DoEvents()

                            Catch ex As Exception
                                strErrMessage = "Problem " + (Err.Number And 65535).ToString + " Verteilung " + intDebBelegsNummer.ToString + vbCrLf
                                strErrMessage += "RG " + strRGNbr + vbCrLf
                                strErrMessage += "Konto " + SubRow("lngKto").ToString + vbCrLf
                                strErrMessage += "Gegenkonto " + intGegenKonto.ToString + vbCrLf
                                strErrMessage += "Betrag " + dblNettoBetrag.ToString + vbCrLf
                                strErrMessage += "Steuer " + strSteuerFeld + vbCrLf
                                strErrMessage += "Bebu " + strBeBuEintrag

                                MessageBox.Show(ex.Message + vbCrLf + strErrMessage, "Problem " + (Err.Number And 65535).ToString, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                If (Err.Number And 65535) < 10000 Then
                                    booBooingok = False
                                Else
                                    booBooingok = True
                                End If

                            End Try

                            strSteuerFeld = String.Empty
                            strBeBuEintrag = String.Empty

                        Next

                        'Beleg buchen
                        Try

                            booBooingok = True
                            Call objdbBuha.WriteBuchung()

                            'Bei SplittBill 2ter Rechnung TZahlung auf LinkedRG machen
                            'Prinzip: Beleg einlesen anhand und Betrag ausrechnen => Summe Beleg - diesen Beleg
                            If row("booLinked") And Mid(row("strDebStatusBitLog"), 13, 1) = "0" Then 'Nur wenn Beleg in gleicher Buha
                                'Betrag von Beleg 1 holen
                                intLaufNbr = objdbBuha.doesBelegExist2(row("lngLinkedDeb").ToString,
                                                                       row("strDebCur"),
                                                                       row("lngLinkedRG").ToString,
                                                                       "NOT_SET",
                                                                       "R",
                                                                       "NOT_SET",
                                                                       "NOT_SET",
                                                                       "NOT_SET")

                                If intLaufNbr > 0 Then
                                    strBeleg = objdbBuha.GetBeleg(row("lngLinkedDeb").ToString,
                                                                  intLaufNbr.ToString)

                                    strBelegArr = Split(strBeleg, "{>}")
                                    If strBelegArr(4) = "B" Then 'schon bezahlt
                                        'Ausbuchen?, wohin mit dem Betrag?
                                    Else

                                        'Betrag von RG 10 auf RG1 als TZ buchen
                                        dblSplitPayed = dblBetrag

                                        'Teilzahlung buchen
                                        Call objdbBuha.SetZahlung(1944,
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
                                                              row("lngDebIdentNbr").ToString + ", TZ " + row("strDebRGNbr").ToString)

                                        Call objdbBuha.WriteTeilzahlung4(intLaufNbr.ToString,
                                                                     row("lngDebIdentNbr").ToString + ", TZ " + row("strDebRGNbr").ToString,
                                                                     "NOT_SET",
                                                                     ,
                                                                     "NOT_SET",
                                                                     "NOT_SET",
                                                                     "Default",
                                                                     "Default")

                                    End If

                                End If

                            End If

                        Catch ex As Exception
                            If (Err.Number And 65535) < 10000 Then
                                strErrMessage = "Belegerstellung RG " + strRGNbr + " Beleg " + intDebBelegsNummer.ToString + " NICHT möglich!"
                                MessageBox.Show(ex.Message + vbCrLf + strErrMessage, "Problem " + (Err.Number And 65535).ToString, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                booBooingok = False
                            Else
                                strErrMessage = "Belegerstellung RG " + strRGNbr + " Beleg " + intDebBelegsNummer.ToString + " möglich mit Warnung"
                                MessageBox.Show(ex.Message, "Warnung " + (Err.Number And 65535).ToString, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                booBooingok = True
                            End If

                        End Try

                    Else 'keine OP - Buchung

                        'Buchung nur in Fibu
                        'Prinzip Funktion WriteBuchung() anwenden mit allen Parametern

                        'Verdopplung interne BelegsNummer verhindern
                        objfiBuha.CheckDoubleIntBelNbr = "J"

                        If IIf(IsDBNull(row("strOPNr")), "", row("strOPNr")) <> "" And IIf(IsDBNull(row("lngDebIdentNbr")), 0, row("lngDebIdentNbr")) <> 0 Then
                            'Belegsnummer abholen fall keine Beleg-Nummer angegeben
                            intDebBelegsNummer = objfiBuha.GetNextBelNbr()
                            'Prüfen ob wirklich frei
                            intReturnValue = 10
                            Do Until intReturnValue = 0
                                intReturnValue = objfiBuha.doesBelegExist(intDebBelegsNummer,
                                                                         "NOT_SET",
                                                                         "NOT_SET",
                                                                         String.Concat(Strings.Left(BgWImportDebiArgsInProc.strPeriode, 4) - 1, "0101"),
                                                                         String.Concat(Strings.Left(BgWImportDebiArgsInProc.strPeriode, 4), "1231"))
                                If intReturnValue <> 0 Then
                                    intDebBelegsNummer += 1
                                End If
                            Loop

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
                            dblKurs = Main.FcGetKurs(strCurrency, strValutaDatum, objfiBuha)
                        Else
                            dblKurs = 1.0#
                        End If

                        selDebiSub = dsDebitoren.Tables("tblDebiSubsFromUser").Select("strRGNr='" + row("strDebRGNbr") + "'")
                        strRGNbr = row("strDebRGNbr")

                        If selDebiSub.Length = 2 Then

                            'Initialisieren
                            dblNettoBetrag = 0
                            dblSollBetrag = 0
                            dblHabenBetrag = 0
                            strBeBuEintrag = String.Empty
                            strBeBuEintragSoll = String.Empty
                            strBeBuEintragHaben = String.Empty
                            strSteuerFeld = String.Empty
                            strSteuerFeldHaben = String.Empty
                            strSteuerFeldSoll = String.Empty

                            For Each SubRow As DataRow In selDebiSub

                                If SubRow("intSollHaben") = 0 Then 'Soll

                                    intSollKonto = SubRow("lngKto")
                                    dblKursSoll = Main.FcGetKurs(strCurrency, strValutaDatum, objfiBuha, intSollKonto)
                                    'strSteuerInfo = Split(FBhg.GetKontoInfo(intSollKonto.ToString), "{>}")
                                    'Debug.Print("Konto-Info Soll: " + strSteuerInfo(26))
                                    dblSollBetrag = SubRow("dblNetto")
                                    strDebiTextSoll = SubRow("strDebSubText")
                                    If SubRow("dblMwSt") > 0 Then
                                        strSteuerFeldSoll = Main.FcGetSteuerFeld(objfiBuha,
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
                                    dblKursHaben = Main.FcGetKurs(strCurrency, strValutaDatum, objfiBuha, intHabenKonto)
                                    'strSteuerInfo = Split(FBhg.GetKontoInfo(intSollKonto.ToString), "{>}")
                                    'Debug.Print("Konto-Info Haben: " + strSteuerInfo(26))
                                    dblHabenBetrag = SubRow("dblNetto") * -1
                                    'dblHabenBetrag = dblSollBetrag
                                    strDebiTextHaben = SubRow("strDebSubText")
                                    If SubRow("dblMwSt") * -1 > 0 Then
                                        strSteuerFeldHaben = Main.FcGetSteuerFeld(objfiBuha,
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
                                Call objfiBuha.WriteBuchung(0,
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
                                MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Belegerstellung " + intDebBelegsNummer.ToString + ", RG " + strRGNbr, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                If (Err.Number And 65535) < 10000 Then
                                    booBooingok = False
                                Else
                                    booBooingok = True
                                End If

                            End Try

                        Else
                            'Sammelbeleg
                            'Variablen initiieren
                            strDebiText = row("strDebText")
                            intCommonKonto = row("lngDebKtoNbr") 'Sammelkonto

                            'Beleg-Kopf
                            Call objfiBuha.SetSammelBhgCommonT2(strValutaDatum,
                                                               intDebBelegsNummer.ToString,
                                                               intCommonKonto.ToString,
                                                               strDebiText,
                                                               strBelegDatum)

                            'Buchungen
                            For Each SubRow As DataRow In selDebiSub

                                'Common - Konto ausblenden da sonst Doppelbuchung
                                If SubRow("lngKto") <> intCommonKonto Then

                                    intSollKonto = 0
                                    strDebiTextSoll = String.Empty
                                    strDebiCurrency = String.Empty
                                    dblKursSoll = 0
                                    dblSollBetrag = 0
                                    strSteuerFeldSoll = String.Empty
                                    intHabenKonto = 0
                                    strDebiTextHaben = String.Empty
                                    strKrediCurrency = String.Empty
                                    dblKursHaben = 0
                                    dblHabenBetrag = 0
                                    strSteuerFeldHaben = String.Empty
                                    dblBuchBetrag = 0
                                    dblBasisBetrag = 0
                                    strBeBuEintragSoll = String.Empty
                                    strBeBuEintragHaben = String.Empty
                                    strErfassungsDatum = Format(Date.Today(), "yyyyMMdd").ToString

                                    If SubRow("intSollHaben") = 0 And SubRow("lngKto") <> intCommonKonto Then 'Soll

                                        intSollKonto = SubRow("lngKto")
                                        strDebiTextSoll = SubRow("strDebSubText")
                                        strDebiCurrency = strCurrency
                                        dblKursSoll = 1 / Main.FcGetKurs(strCurrency, strValutaDatum, objfiBuha, intSollKonto)
                                        dblSollBetrag = SubRow("dblNetto")
                                        If SubRow("dblMwSt") > 0 Then
                                            strSteuerFeldSoll = Main.FcGetSteuerFeld(objfiBuha, SubRow("lngKto"), strDebiTextSoll, SubRow("dblBrutto") * dblKursSoll, SubRow("strMwStKey"), SubRow("dblMwSt"))
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
                                        dblKursHaben = 1 / Main.FcGetKurs(strCurrency, strValutaDatum, objfiBuha, intHabenKonto)
                                        dblHabenBetrag = SubRow("dblNetto") * -1
                                        If (SubRow("dblMwSt") * -1) > 0 Then
                                            strSteuerFeldHaben = Main.FcGetSteuerFeld(objfiBuha, SubRow("lngKto"), strDebiTextHaben, SubRow("dblBrutto") * dblKursHaben * -1, SubRow("strMwStKey"), SubRow("dblMwSt") * -1)
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

                                    Call objfiBuha.SetSammelBhgT(intSollKonto.ToString,
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
                                Call objfiBuha.WriteSammelBhgT()

                            Catch ex As Exception

                                MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Belegerstellung " + intDebBelegsNummer.ToString + ", RG " + strRGNbr, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                If (Err.Number And 65535) < 10000 Then
                                    booBooingok = False
                                Else
                                    booBooingok = True
                                End If
                            End Try


                        End If

                    End If

                    If booBooingok Then
                        If row("booPGV") Then
                            'Bei PGV Buchungen
                            If IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")) = "" Or
                                    (IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")) = "RV" And row("intPGVMthsAY") + row("intPGVMthsNY") > 1) Then

                                intReturnValue = MainDebitor.FcPGVDTreatment(objfiBuha,
                                                                       objFinanz,
                                                                       objdbBuha,
                                                                       objdbPIFb,
                                                                       objFiBebu,
                                                                       objKrBuha,
                                                                       dsDebitoren.Tables("tblDebiSubsFromUser"),
                                                                       row("strDebRGNbr"),
                                                                       intDebBelegsNummer,
                                                                       row("strDebCur"),
                                                                       row("datDebValDatum"),
                                                                       "M",
                                                                       row("datPGVFrom"),
                                                                       row("datPGVTo"),
                                                                       row("intPGVMthsAY") + row("intPGVMthsNY"),
                                                                       row("intPGVMthsAY"),
                                                                       row("intPGVMthsNY"),
                                                                       1311,
                                                                       1312,
                                                                       BgWImportDebiArgsInProc.strPeriode,  'frmImportMain.lstBoxPerioden.Text,
                                                                       objdbConnZHDB02,
                                                                       objdbMSSQLConn,
                                                                       objdbSQLcommand,
                                                                       BgWImportDebiArgsInProc.intMandant,  'frmImportMain.lstBoxMandant.SelectedValue,
                                                                       dsDebitoren.Tables("tblDebitorenInfo"),
                                                                       BgWImportDebiArgsInProc.strYear,
                                                                       BgWImportDebiArgsInProc.intTeqNbr,
                                                                       BgWImportDebiArgsInProc.intTeqNbrLY,
                                                                       BgWImportDebiArgsInProc.intTeqNbrPLY,
                                                                       IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")),
                                                                       datPeriodFrom,
                                                                       datPeriodTo,
                                                                       strPeriodStatus)


                            Else
                                intReturnValue = MainDebitor.FcPGVDTreatmentYC(objfiBuha,
                                                                       objFinanz,
                                                                       objdbBuha,
                                                                       objdbPIFb,
                                                                       objFiBebu,
                                                                       objKrBuha,
                                                                       dsDebitoren.Tables("tblDebiSubsFromUser"),
                                                                       row("strDebRGNbr"),
                                                                       intDebBelegsNummer,
                                                                       row("strDebCur"),
                                                                       row("datDebValDatum"),
                                                                       "M",
                                                                       row("datPGVFrom"),
                                                                       row("datPGVTo"),
                                                                       row("intPGVMthsAY") + row("intPGVMthsNY"),
                                                                       row("intPGVMthsAY"),
                                                                       row("intPGVMthsNY"),
                                                                       1311,
                                                                       1312,
                                                                       BgWImportDebiArgsInProc.strPeriode,
                                                                       objdbConnZHDB02,
                                                                       objdbMSSQLConn,
                                                                       objdbSQLcommand,
                                                                       BgWImportDebiArgsInProc.intMandant,
                                                                       dsDebitoren.Tables("tblDebitorenInfo"),
                                                                       BgWImportDebiArgsInProc.strYear,
                                                                       BgWImportDebiArgsInProc.intTeqNbr,
                                                                       BgWImportDebiArgsInProc.intTeqNbrLY,
                                                                       BgWImportDebiArgsInProc.intTeqNbrPLY,
                                                                       IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")),
                                                                       datPeriodFrom,
                                                                       datPeriodTo,
                                                                       strPeriodStatus)
                            End If


                        End If

                        'Status Head schreiben
                        row("strDebBookStatus") = row("strDebStatusBitLog")
                        row("booBooked") = True
                        row("datBooked") = Now()
                        row("lngBelegNr") = intDebBelegsNummer
                        dsDebitoren.Tables("tblDebiHeadsFromUser").AcceptChanges()
                        'Application.DoEvents()

                        'Status in File RG-Tabelle schreiben
                        intReturnValue = MainDebitor.FcWriteToRGTable(BgWImportDebiArgsInProc.intMandant,
                                                                          row("strDebRGNbr"),
                                                                          row("datBooked"),
                                                                          row("lngBelegNr"),
                                                                          objdbAccessConn,
                                                                          objOracleConn,
                                                                          objdbConnZHDB02,
                                                                          row("booDatChanged"),
                                                                          row("datDebRGDatum"),
                                                                          row("datDebValDatum"))
                        If intReturnValue <> 0 Then
                            'Throw an exception
                        End If

                        'Evtl. Query nach Buchung ausführen
                        'Call MainDebitor.FcExecuteAfterDebit(BgWImportDebiArgsInProc.intMandant, objdbConn)
                    End If


                End If

            Next

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Generelles Problem " + (Err.Number And 65535).ToString + " Belegerstellung " + intDebBelegsNummer.ToString + ", RG " + strRGNbr, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Err.Clear()

        Finally
            'Buhas freigeben
            objKrBuha = Nothing
            objFiBebu = Nothing
            objdbPIFb = Nothing
            objdbBuha = Nothing
            objfiBuha = Nothing
            objFinanz = Nothing

            objdbConnZHDB02 = Nothing
            objdbMSSQLConn = Nothing
            objdbSQLcommand = Nothing
            objdbAccessConn = Nothing
            objOracleConn = Nothing

            'Me.dsDebitoren = Nothing
            'System.GC.Collect()

        End Try

    End Sub

    Private Sub frmDebDisp_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed

        'System.GC.Collect()
        Me.mysqlcmdDebDel = Nothing
        Me.mysqlcmdDebRead = Nothing
        Me.mysqlcmdDebSubDel = Nothing
        Me.mysqlcmdDebSubRead = Nothing
        Me.mysqlconn = Nothing
        Me.MySQLdaDebitoren = Nothing
        Me.MySQLdaDebitorenSub = Nothing

        Me.dsDebitoren = Nothing
        Me.Dispose()
        System.GC.Collect()
        'System.Diagnostics.Process.Start(Application.ExecutablePath)
        'Environment.Exit(0)
        'System.GC.Collect()
        Application.Restart()

    End Sub

    Private Sub butDeSeöect_Click(sender As Object, e As EventArgs) Handles butDeSeöect.Click

        'Alle selektierten Records werden deselektiert

        For Each row As DataRow In dsDebitoren.Tables("tblDebiHeadsFromUser").Rows
            If row("booDebBook") Then
                row("booDebBook") = False
            End If
        Next
        dsDebitoren.Tables("tblDebiHeadsFromUser").AcceptChanges()
        'Me.Refresh()

    End Sub
End Class