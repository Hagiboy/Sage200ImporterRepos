Option Strict Off
Option Explicit On

Imports MySql.Data.MySqlClient
Imports System.Data.OracleClient
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic.ApplicationServices
'Imports CLClassSage200.WFSage200Import
Imports System.IO
Imports System.Net
Imports System.Xml

Public Class frmDebDisp

    'Dim Finanz As SBSXASLib.AXFinanz
    'Dim FBhg As SBSXASLib.AXiFBhg
    'Dim DbBhg As SBSXASLib.AXiDbBhg
    'Dim KrBhg As SBSXASLib.AXiKrBhg
    'Dim BsExt As SBSXASLib.AXiBSExt
    'Dim Adr As SBSXASLib.AXiAdr
    'Dim BeBu As SBSXASLib.AXiBeBu
    'Dim PIFin As SBSXASLib.AXiPlFin

    Dim objFinanz As New SBSXASLib.AXFinanz
    Dim objfiBuha As New SBSXASLib.AXiFBhg
    Dim objdbBuha As New SBSXASLib.AXiDbBhg
    Dim objdbPIFb As New SBSXASLib.AXiPlFin
    Dim objFiBebu As New SBSXASLib.AXiBeBu
    Dim objKrBuha As New SBSXASLib.AXiKrBhg

    Dim FELD_SEP As String
    Dim REC_SEP As String
    Dim KSTKTR_SEP As String
    Dim FELD_SEP_OUT As String
    Dim REC_SEP_OUT As String
    Dim nID As String

    Dim strPeriodenInfo As String
    Dim intMandant As Int32
    Dim intTeqNbr As Int32
    Dim intTeqNbrLY As Int32
    Dim intTeqNbrPLY As Int32
    Dim strYear As String
    Dim datPeriodFrom As Date
    Dim datPeriodTo As Date
    Dim strPeriodStatus As String

    Dim objdbConn As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString"))
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
        Public booValutaEndCor As Boolean
        Public datValutaEndCor As Date
    End Class


    Public Sub InitDB()

        Dim strIdentityName As String
        'Dim mysqllocda As New MySqlDataAdapter
        'Dim mysqllocdasel As MySqlCommand
        'Dim mysqllocdacon As New MySqlConnection
        'Dim mysqllocdadel As MySqlCommand
        Dim objdbtaskcmd As New MySqlCommand
        Dim objdbtasks As New DataTable

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
            'mysqllocdasel = New MySqlCommand("SELECT * FROM tbldebitorenjhead WHERE IdentityName='" + strIdentityName + "' AND ProcessID= " + Process.GetCurrentProcess().Id.ToString)

            'Del cmd DebiHead
            mysqlcmdDebDel.CommandText = "DELETE FROM tbldebitorenjhead WHERE IdentityName='" + strIdentityName + "' AND ProcessID= " + Process.GetCurrentProcess().Id.ToString
            'mysqllocdadel = New MySqlCommand("DELETE FROM tbldebitorenjhead WHERE IdentityName='" + strIdentityName + "' AND ProcessID= " + Process.GetCurrentProcess().Id.ToString)

            'Debitoren Sub
            'Read
            mysqlcmdDebSubRead.CommandText = "Select * FROM tbldebitorensub WHERE IdentityName='" + strIdentityName + "' AND ProcessID= " + Process.GetCurrentProcess().Id.ToString

            'Del cmd Debi Sub
            mysqlcmdDebSubDel.CommandText = "DELETE FROM tbldebitorensub WHERE IdentityName='" + strIdentityName + "' AND ProcessID= " + Process.GetCurrentProcess().Id.ToString

            'mysqllocdacon.ConnectionString = "server=ZHDB02.sdlc.mssag.ch;uid=workbench;database=DBZ;Pwd=sesam"
            'mysqllocdasel.Connection.ConnectionString = "Server=ZHDB02.sdlc.mssag.ch;User ID=workbench;Database=DBZ;Connection Timeout=30;Convert Zero DateTime=True"
            'mysqllocdasel.CommandText = mysqlcmdDebRead.CommandText
            'mysqllocdasel.Connection = mysqllocdacon
            'mysqllocda.SelectCommand = mysqllocdasel

            'm'ysqllocdadel.Connection = mysqllocdacon
            'mysqllocda.DeleteCommand = mysqllocdadel
            'Dim mysqldacmdbld As New MySqlCommandBuilder(mysqllocda)
            'mysqldacmdbld.GetUpdateCommand()
            'mysqldacmdbld.GetInsertCommand()
            'MySQLdaDebitoren.UpdateCommand.CommandText = mysqldacmdbld.GetUpdateCommand().CommandText
            'MySQLdaDebitoren.InsertCommand.CommandText = mysqldacmdbld.GetInsertCommand().CommandText

            'Mandant holen
            objdbtaskcmd.Connection = objdbConn
            objdbtaskcmd.Connection.Open()
            objdbtaskcmd.CommandText = "SELECT * FROM tblimporttasks WHERE IdentityName='" + strIdentityName + "' AND Type='D'"
            objdbtasks.Load(objdbtaskcmd.ExecuteReader())
            If objdbtasks.Rows.Count > 0 Then
                intMandant = objdbtasks.Rows(0).Item("Mandant")
            Else
                intMandant = 1
                MessageBox.Show("Mandant konnte nicht gelesen werden. => Setzen auf AHZ")
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Generelles Problem " + Convert.ToString(Err.Number) + "Init Debitoren", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Err.Clear()

        Finally
            'mysqllocda = Nothing
            'mysqllocdacon = Nothing
            'mysqllocdasel = Nothing
            'mysqllocdadel = Nothing
            objdbtasks = Nothing
            objdbtaskcmd = Nothing

        End Try

    End Sub

    Private Sub frmDebDisp_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        FELD_SEP = "{<}"
        REC_SEP = "{>}"
        KSTKTR_SEP = "{-}"

        FELD_SEP_OUT = "{>}"
        REC_SEP_OUT = "{<}"

        Cursor = Cursors.WaitCursor
        UseWaitCursor = True

        Call InitDB()

        'Dim clImp As New ClassImport
        'clImp.FcDebitFill(intMandant)
        'clImp = Nothing

        BgWLoadDebi.RunWorkerAsync(intMandant)

        Do While BgWLoadDebi.IsBusy
            Threading.Thread.Sleep(1)
            Application.DoEvents()
            'Await Task.WhenAll(BgWLoadDebi)
        Loop

        'Tabellentyp darstellen
        Call FcReadFromSettingsIII("Buchh_RGTableType",
                                              intMandant,
                                              Me.TSLblDebType.Text)

        butCheckDeb.Enabled = True

        UseWaitCursor = False
        Cursor = Cursors.Default

        'MySQLdaDebitoren.Fill(dsDebitoren, "tblDebiHeadsFromUser")
        'MySQLdaDebitorenSub.Fill(dsDebitoren, "tblDebiSubsFromUser")


    End Sub

    'Friend Function FcDebiDisplay(intMandant As Int32,
    '                              LstMandnat As ListBox,
    '                              LstBPerioden As ListBox) As Int16

    '    Dim objdbConn As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString"))
    '    'Dim objdbtaskcmd As New MySqlCommand
    '    Dim objdbMSSQLConn As New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("SQLConnectionString"))
    '    Dim objdbSQLcommand As New SqlCommand

    '    Dim intFcReturns As Int16
    '    Dim strPeriode As String
    '    Dim strYearCh As String
    '    Dim BgWCheckDebitLocArgs As New BgWCheckDebitArgs
    '    'Dim objdbtasks As New DataTable

    '    'Dim intTeqNbr As Int32
    '    'Dim intTeqNbrLY As Int32
    '    'Dim intTeqNbrPLY As Int32
    '    'Dim strYear As String

    '    'Dim objFinanz As New SBSXASLib.AXFinanz
    '    'Dim objfiBuha As New SBSXASLib.AXiFBhg
    '    'Dim objdbBuha As New SBSXASLib.AXiDbBhg
    '    'Dim objdbPIFb As New SBSXASLib.AXiPlFin
    '    'Dim objFiBebu As New SBSXASLib.AXiBeBu
    '    'Dim objKrBuha As New SBSXASLib.AXiKrBhg


    '    Try

    '        Me.Cursor = Cursors.WaitCursor
    '        'Zuerst in tblImportTasks setzen
    '        'objdbtaskcmd.Connection = objdbConn
    '        'objdbtaskcmd.Connection.Open()
    '        'objdbtaskcmd.CommandText = "SELECT * FROM tblimporttasks WHERE IdentityName='" + frmImportMain.LblIdentity.Text + "' AND Type='D'"
    '        'objdbtasks.Load(objdbtaskcmd.ExecuteReader())
    '        'If objdbtasks.Rows.Count > 0 Then
    '        '    'update
    '        '    objdbtaskcmd.CommandText = "UPDATE tblimporttasks SET Mandant=" + Convert.ToString(LstMandnat.SelectedIndex) + ", Periode=" + Convert.ToString(LstBPerioden.SelectedIndex) + " WHERE IdentityName='" + frmImportMain.LblIdentity.Text + "' AND Type='D'"
    '        'Else
    '        '    'insert
    '        '    objdbtaskcmd.CommandText = "INSERT INTO tblimporttasks (IdentityName, Type, Mandant, Periode) VALUES ('" + frmImportMain.LblIdentity.Text + "', 'D', " + Convert.ToString(LstMandnat.SelectedIndex) + ", " + Convert.ToString(LstBPerioden.SelectedIndex) + ")"
    '        'End If
    '        'objdbtaskcmd.ExecuteNonQuery()
    '        'objdbtaskcmd.Connection.Close()

    '        'intMode = 0

    '        Me.butImport.Enabled = False


    '        'DGV Debitoren
    '        'dgvBookings.DataSource = Nothing
    '        'dgvBookingSub.DataSource = Nothing

    '        'dsDebitoren.Reset()
    '        'dsDebitoren.Clear()

    '        'Zuerst evtl. vorhandene DS löschen in Tabelle
    '        'MySQLdaDebitoren.DeleteCommand.Connection.Open()
    '        'MySQLdaDebitoren.DeleteCommand.ExecuteNonQuery()
    '        'MySQLdaDebitoren.DeleteCommand.Connection.Close()

    '        'MySQLdaDebitorenSub.DeleteCommand.Connection.Open()
    '        'MySQLdaDebitorenSub.DeleteCommand.ExecuteNonQuery()
    '        'MySQLdaDebitorenSub.DeleteCommand.Connection.Close()

    '        'Info neu erstellen
    '        dsDebitoren.Tables.Add("tblDebitorenInfo")
    '        Dim col1 As DataColumn = New DataColumn("strInfoT")
    '        col1.DataType = System.Type.GetType("System.String")
    '        col1.MaxLength = 50
    '        col1.Caption = "Info-Titel"
    '        dsDebitoren.Tables("tblDebitorenInfo").Columns.Add(col1)
    '        Dim col2 As DataColumn = New DataColumn("strInfoV")
    '        col2.DataType = System.Type.GetType("System.String")
    '        col2.MaxLength = 50
    '        col2.Caption = "Info-Wert"
    '        dsDebitoren.Tables("tblDebitorenInfo").Columns.Add(col2)

    '        dgvInfo.DataSource = dsDebitoren.Tables("tblDebitorenInfo")

    '        'Datums-Tabelle erstellen
    '        dsDebitoren.Tables.Add("tblDebitorenDates")
    '        Dim col7 As DataColumn = New DataColumn("intYear")
    '        col7.DataType = System.Type.GetType("System.Int16")
    '        col7.Caption = "Year"
    '        dsDebitoren.Tables("tblDebitorenDates").Columns.Add(col7)
    '        Dim col3 As DataColumn = New DataColumn("strDatType")
    '        col3.DataType = System.Type.GetType("System.String")
    '        col3.MaxLength = 50
    '        col3.Caption = "Datum-Typ"
    '        dsDebitoren.Tables("tblDebitorenDates").Columns.Add(col3)
    '        Dim col4 As DataColumn = New DataColumn("datFrom")
    '        col4.DataType = System.Type.GetType("System.DateTime")
    '        col4.Caption = "Von"
    '        dsDebitoren.Tables("tblDebitorenDates").Columns.Add(col4)
    '        Dim col5 As DataColumn = New DataColumn("datTo")
    '        col5.DataType = System.Type.GetType("System.DateTime")
    '        col5.Caption = "Bis"
    '        dsDebitoren.Tables("tblDebitorenDates").Columns.Add(col5)
    '        Dim col6 As DataColumn = New DataColumn("strStatus")
    '        col6.DataType = System.Type.GetType("System.String")
    '        col6.Caption = "S"
    '        dsDebitoren.Tables("tblDebitorenDates").Columns.Add(col6)
    '        dgvDates.DataSource = dsDebitoren.Tables("tblDebitorenDates")

    '        strPeriode = LstBPerioden.GetItemText(LstBPerioden.SelectedItem)

    '        Call FcLoginSage3(objdbConn,
    '                              objdbMSSQLConn,
    '                              objdbSQLcommand,
    '                              objFinanz,
    '                              objfiBuha,
    '                              objdbBuha,
    '                              objdbPIFb,
    '                              objFiBebu,
    '                              objKrBuha,
    '                              intMandant,
    '                              dsDebitoren.Tables("tblDebitorenInfo"),
    '                              dsDebitoren.Tables("tblDebitorenDates"),
    '                              strPeriode,
    '                              strYear,
    '                              intTeqNbr,
    '                              intTeqNbrLY,
    '                              intTeqNbrPLY,
    '                              datPeriodFrom,
    '                              datPeriodTo,
    '                              strPeriodStatus)

    '        'Gibt es mehr als ein Jahr?
    '        If LstBPerioden.Items.Count > 1 Then

    '            'Gibt es ein Vorjahr?
    '            If LstBPerioden.SelectedIndex + 1 > 1 Then
    '                strPeriode = LstBPerioden.Items(LstBPerioden.SelectedIndex - 1)
    '                'Peeriodendef holen
    '                Call FcLoginSage4(intMandant,
    '                                   dsDebitoren.Tables("tblDebitorenDates"),
    '                                   strPeriode)
    '            Else
    '                'Periode ezreugen und auf N stellen
    '                strYearCh = Convert.ToString(Val(strYear) - 1)
    '                dsDebitoren.Tables("tblDebitorenDates").Rows.Add(strYearCh, "GJ n/a", DateSerial(Convert.ToUInt16(strYearCh), 1, 1), DateSerial(Convert.ToUInt16(strYearCh), 12, 31), "N")
    '            End If

    '            'Gibt es ein Folgehahr?
    '            If LstBPerioden.SelectedIndex + 1 < LstBPerioden.Items.Count Then
    '                strPeriode = LstBPerioden.Items(LstBPerioden.SelectedIndex + 1)
    '                'Peeriodendef holen
    '                Call FcLoginSage4(intMandant,
    '                                   dsDebitoren.Tables("tblDebitorenDates"),
    '                                   strPeriode)
    '            Else
    '                'Periode ezreugen und auf N stellen
    '                strYearCh = Convert.ToString(Val(strYear) + 1)
    '                dsDebitoren.Tables("tblDebitorenDates").Rows.Add(strYearCh, "GJ n/a", DateSerial(Convert.ToUInt16(strYearCh), 1, 1), DateSerial(Convert.ToUInt16(strYearCh), 12, 31), "N")
    '            End If
    '        ElseIf LstBPerioden.Items.Count = 1 Then 'es gibt genau 1 Jahr
    '            'gewähltes Jahr checken
    '            Call FcLoginSage4(intMandant,
    '                                   dsDebitoren.Tables("tblDebitorenDates"),
    '                                   strPeriode)
    '            'VJ erzeugen
    '            strYearCh = Convert.ToString(Val(strYear) - 1)
    '            dsDebitoren.Tables("tblDebitorenDates").Rows.Add(strYearCh, "GJ n/a", DateSerial(Convert.ToUInt16(strYearCh), 1, 1), DateSerial(Convert.ToUInt16(strYearCh), 12, 31), "N")

    '            'FJ erzeugen
    '            strYearCh = Convert.ToString(Val(strYear) + 1)
    '            dsDebitoren.Tables("tblDebitorenDates").Rows.Add(strYearCh, "GJ n/a", DateSerial(Convert.ToUInt16(strYearCh), 1, 1), DateSerial(Convert.ToUInt16(strYearCh), 12, 31), "N")

    '        End If


    '        'Dim clImp As New ClassImport
    '        'clImp.FcDebitFill(intMandant)
    '        'clImp = Nothing


    '        'BgWLoadDebi.RunWorkerAsync(intMandant)

    '        'Do While BgWLoadDebi.IsBusy
    '        '    Application.DoEvents()
    '        'Loop

    '        ''Tabellentyp darstellen
    '        'Me.lblDB.Text = Main.FcReadFromSettingsII("Buchh_RGTableType", intMandant)


    '        'MySQLdaDebitoren.AcceptChangesDuringFill = False
    '        'MySQLdaDebitoren.Fill(dsDebitoren, "tblDebiHeadsFromUser")
    '        'Debug.Print("Changes nach Load Head " + dsDebitoren.Tables("tblDebiHeadsFromUser").GetChanges().Rows.Count.ToString)
    '        'MySQLdaDebitoren.Update(dsDebitoren, "tblDebiHeadsFromUser")
    '        'MySQLdaDebitorenSub.Fill(dsDebitoren, "tblDebiSubsFromUser")


    '        'Application.DoEvents()

    '        'Dim clCheck As New ClassCheck
    '        'clCheck.FcClCheckDebit(intMandant,
    '        '                       dsDebitoren,
    '        '                       Finanz,
    '        '                       FBhg,
    '        '                       DbBhg,
    '        '                       PIFin,
    '        '                       BeBu,
    '        '                       dsDebitoren.Tables("tblDebitorenInfo"),
    '        '                       dsDebitoren.Tables("tblDebitorenDates"),
    '        '                       frmImportMain.lstBoxMandant.Text,
    '        '                       intTeqNbr,
    '        '                       intTeqNbrLY,
    '        '                       intTeqNbrPLY,
    '        '                       strYear,
    '        '                       frmImportMain.chkValutaCorrect.Checked,
    '        '                       frmImportMain.dtpValutaCorrect.Value)
    '        'clCheck = Nothing

    '        'BgWCheckDebitLocArgs.intMandant = intMandant
    '        'BgWCheckDebitLocArgs.strMandant = frmImportMain.lstBoxMandant.GetItemText(frmImportMain.lstBoxMandant.SelectedItem)
    '        'BgWCheckDebitLocArgs.intTeqNbr = intTeqNbr
    '        'BgWCheckDebitLocArgs.intTeqNbrLY = intTeqNbrLY
    '        'BgWCheckDebitLocArgs.intTeqNbrPLY = intTeqNbrPLY
    '        'BgWCheckDebitLocArgs.strYear = strYear
    '        'BgWCheckDebitLocArgs.strPeriode = LstBPerioden.GetItemText(LstBPerioden.SelectedItem)
    '        'BgWCheckDebitLocArgs.booValutaCor = frmImportMain.chkValutaCorrect.Checked
    '        'BgWCheckDebitLocArgs.datValutaCor = frmImportMain.dtpValutaCorrect.Value

    '        'BgWCheckDebi.RunWorkerAsync(BgWCheckDebitLocArgs)

    '        'Do While BgWCheckDebi.IsBusy
    '        '    Application.DoEvents()
    '        'Loop

    '        'System.GC.Collect()

    '        Debug.Print("Vor Refresh DGV")
    '        'Debug.Print("Changes nach Check Head " + dsDebitoren.Tables("tblDebiHeadsFromUser").GetChanges().Rows.Count.ToString)

    '        'Grid neu aufbauen
    '        'dgvBookings.DataSource = Nothing
    '        'dgvBookingSub.DataSource = Nothing
    '        ''MySQLdaDebitoren.Update(dsDebitoren, "tblDebiHeadsFromUser")
    '        dgvBookings.DataSource = dsDebitoren.Tables("tblDebiHeadsFromUser")
    '        dgvBookingSub.DataSource = dsDebitoren.Tables("tblDebiSubsFromUser")

    '        Debug.Print("Vor Init DGV")
    '        intFcReturns = FcInitdgvInfo(dgvInfo)
    '        intFcReturns = FcInitdgvBookings(dgvBookings)
    '        intFcReturns = FcInitdgvDebiSub(dgvBookingSub)
    '        intFcReturns = FcInitdgvDate(dgvDates)

    '        'Anzahl schreiben
    '        txtNumber.Text = Me.dsDebitoren.Tables("tblDebiHeadsFromUser").Rows.Count.ToString
    '        Me.Cursor = Cursors.Default

    '        Me.butImport.Enabled = True
    '        Return 0

    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message, "Generelles Problem " + Convert.ToString(Err.Number) + "Check Debitoren", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    '        Err.Clear()
    '        Return 1

    '    Finally
    '        'objFinanz = Nothing
    '        'objfiBuha = Nothing
    '        'objdbBuha = Nothing
    '        'objdbPIFb = Nothing
    '        'objFiBebu = Nothing
    '        'objKrBuha = Nothing

    '        objdbConn = Nothing
    '        objdbMSSQLConn = Nothing
    '        objdbSQLcommand = Nothing
    '        'objdbtasks = Nothing

    '        BgWCheckDebitLocArgs = Nothing

    '        'System.GC.Collect()

    '    End Try



    'End Function

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

        dgvBookings.ShowCellToolTips = True
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

        dgvBookingSub.ShowCellToolTips = True
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


            Cursor = Cursors.WaitCursor
            UseWaitCursor = True
            'Button disablen damit er nicht noch einmal geklickt werden kann.
            Me.butImport.Enabled = False

            BgWImportDebi.RunWorkerAsync(BgWImportDebiLocArgs)

            Do While BgWImportDebi.IsBusy
                Application.DoEvents()
            Loop

            Cursor = Cursors.Default

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
            Err.Clear()

        Finally

            'Application.DoEvents()
            'Grid neu aufbauen, Daten von Mandant einlesen
            'Call butDebitoren.PerformClick()
            BgWImportDebiLocArgs = Nothing
            UseWaitCursor = False
            Me.Cursor = Cursors.Default
            'Me.butImport.Enabled = False
            'Me.Close()
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

            objmysqlcomdwritehead.Connection = objdbConnZHDB02
            objmysqlcomdwritesub.Connection = objdbConnZHDB02

            'Für den Save der Records
            strIdentityName = System.Security.Principal.WindowsIdentity.GetCurrent().Name
            strIdentityName = Strings.Replace(strIdentityName, "\", "/")

            intFcReturns = FcReadFromSettingsIII("Buchh_RGTableMDB",
                                                        intAccounting,
                                                        strMDBName)

            intFcReturns = FcReadFromSettingsIII("Buchh_SQLHead",
                                                 intAccounting,
                                                 strSQL)

            intFcReturns = FcReadFromSettingsIII("Buchh_RGTableType",
                                                         intAccounting,
                                                         strRGTableType)
            objdslocdebihead.EnforceConstraints = False
            'objdslocdebihead.AcceptChanges()

            If strRGTableType = "A" Then

                'Access
                Call FcInitAccessConnecation(objdbAccessConn,
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
            'objdslocdebisub.AcceptChanges()

            intFcReturns = FcReadFromSettingsIII("Buchh_SQLDetail",
                                                        intAccounting,
                                                        strSQLToParse)

            intFcReturns = FcInitInsCmdDHeads(objmysqlcomdwritehead)

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
                objmysqlcomdwritehead.Parameters("@strDebText").Value = FcDeleteNonAscii(IIf(IsDBNull(row("strDebText")), "", row("strDebText")))
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
                objmysqlcomdwritehead.Parameters("@lngLinkedGS").Value = row("lngLinkedGS")
                objmysqlcomdwritehead.Parameters("@strRGName").Value = FcDeleteNonAscii(IIf(IsDBNull(row("strRGName")), "", row("strRGName")))
                If row.Table.Columns.Contains("strDebIdentNbr2") Then
                    objmysqlcomdwritehead.Parameters("@strDebIdentNbr2").Value = row("strDebIdentNbr2")
                End If
                If row.Table.Columns.Contains("strRGBemerkung") Then
                    objmysqlcomdwritehead.Parameters("@strRGBemerkung").Value = row("strRGBemerkung")
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
                objmysqlcomdwritehead.Parameters("@intZKond").Value = row("intZKond")
                objmysqlcomdwritehead.ExecuteNonQuery()
                objmysqlcomdwritehead.Connection.Close()
                'objdtLocDebiHead.AcceptChanges()

                'Subs einlesen
                'Es muss der Weg über das DS genommen werden wegen den constraint-Verlethzungen
                strSQLSub = FcSQLParse2(strSQLToParse,
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

                'Subs schreiben
                intFcReturns = FcInitInscmdSubs(objmysqlcomdwritesub)
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

                    'objdtlocDebiSub.AcceptChanges()

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


        End Try

    End Sub

    Private Sub BgWCheckDebi_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BgWCheckDebi.DoWork

        Dim BgWCheckDebiArgsInProc As BgWCheckDebitArgs = e.Argument

        Dim strMandant As String
        Dim booAccOk As Boolean
        'Dim objFinanz As New SBSXASLib.AXFinanz
        'Dim objfiBuha As New SBSXASLib.AXiFBhg
        'Dim objdbBuha As New SBSXASLib.AXiDbBhg
        'Dim objdbPIFb As New SBSXASLib.AXiPlFin
        'Dim objFiBebu As New SBSXASLib.AXiBeBu

        Dim strBitLog As String = String.Empty
        Dim intReturnValue As Integer
        Dim strStatus As String = String.Empty
        Dim booAutoCorrect As Boolean
        Dim booCpyKSTToSub As Boolean
        Dim booSplittBill As Boolean
        Dim booLinkedGS As Boolean
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
        Dim booValutaEndCorrect As Boolean
        Dim datValutaEndCorrect As Date
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
        Dim booErfOPExt As Boolean
        Dim strDebiHeadText As String
        Dim strDebiSubText As String
        Dim selsubrow() As DataRow
        Dim nrow As DataRow
        Dim intFcReturns As Int16
        Dim strFcReturns As String
        Dim intActRGNbr As Int32
        Dim intTotRGs As Int32


        Try

            'objFinanz = Nothing
            'objFinanz = New SBSXASLib.AXFinanz
            'objfiBuha = Nothing

            'objdbBuha = Nothing

            'objdbPIFb = Nothing

            'objFiBebu = Nothing

            'Finanz-Obj init
            'Login
            Try
                Call objFinanz.ConnectSBSdb(System.Configuration.ConfigurationManager.AppSettings("OwnSageServer"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageDB"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageID"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSagePsw"), "")

            Catch inEx As Exception
                If inEx.HResult <> -2147473602 Then
                    MessageBox.Show(inEx.Message, "Connect to Sage - DB " + Err.Number.ToString)
                    Exit Sub
                End If

            End Try


            intFcReturns = FcReadFromSettingsIII("Buchh200_Name",
                                                BgWCheckDebiArgsInProc.intMandant,
                                                strMandant)

            booAccOk = objFinanz.CheckMandant(strMandant)
            'Open Mandantg
            objFinanz.OpenMandant(strMandant, BgWCheckDebiArgsInProc.strPeriode)

            objfiBuha = objFinanz.GetFibuObj()
            objdbBuha = objFinanz.GetDebiObj()
            objdbPIFb = objfiBuha.GetCheckObj()
            objFiBebu = objFinanz.GetBeBuObj()

            'Variablen einlesen
            booAutoCorrect = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettingsII("Buchh_HeadAutoCorrect", BgWCheckDebiArgsInProc.intMandant)))
            booCpyKSTToSub = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettingsII("Buchh_KSTHeadToSub", BgWCheckDebiArgsInProc.intMandant)))
            booSplittBill = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettingsII("Buchh_LinkedBookings", BgWCheckDebiArgsInProc.intMandant)))
            booLinkedGS = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettingsII("Buchh_LinkedGS", BgWCheckDebiArgsInProc.intMandant)))
            booCashSollCorrect = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettingsII("Buchh_CashSollKontoKorr", BgWCheckDebiArgsInProc.intMandant)))
            booGeneratePymentBooking = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettingsII("Buchh_GeneratePaymentBooking", BgWCheckDebiArgsInProc.intMandant)))
            booErfOPExt = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettingsII("Buchh_ErfOPExt", BgWCheckDebiArgsInProc.intMandant)))
            booDiffHeadText = IIf(FcReadFromSettingsII("Buchh_TextSpecial", BgWCheckDebiArgsInProc.intMandant) = "0", False, True)
            booDiffSubText = IIf(FcReadFromSettingsII("Buchh_SubTextSpecial", BgWCheckDebiArgsInProc.intMandant) = "0", False, True)
            intFcReturns = FcReadFromSettingsIII("Buchh_PKTable", BgWCheckDebiArgsInProc.intMandant, strFcReturns)
            booPKPrivate = IIf(strFcReturns = "t_customer", True, False)
            booValutaCorrect = BgWCheckDebiArgsInProc.booValutaCor
            datValutaCorrect = BgWCheckDebiArgsInProc.datValutaCor
            booValutaEndCorrect = BgWCheckDebiArgsInProc.booValutaEndCor
            datValutaEndCorrect = BgWCheckDebiArgsInProc.datValutaEndCor
            'dsDebitoren.Tables("tblDebiHeadsFromUser").AcceptChanges()

            intTotRGs = dsDebitoren.Tables("tblDebiHeadsFromUser").Rows.Count
            intActRGNbr = 0

            For Each row As DataRow In dsDebitoren.Tables("tblDebiHeadsFromUser").Rows

                'Progress Bar
                intActRGNbr += 1
                BgWCheckDebi.ReportProgress(100 / intTotRGs * intActRGNbr)

                'If row("strDebRGNbr") = "1502174" Then Stop
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
                intReturnValue = FcGetRefDebiNr(IIf(IsDBNull(row("lngDebNbr")), 0, row("lngDebNbr")),
                                                BgWCheckDebiArgsInProc.intMandant,
                                                intDebitorNew)
                If intReturnValue = 1 Then 'Neue Debi-Nr wurde angelegt
                    strStatus = "NDeb "
                End If
                If intDebitorNew <> 0 Or intReturnValue = 4 Then
                    intReturnValue = FcCheckDebitor(intDebitorNew,
                                                                row("intBuchungsart"))
                Else
                    intReturnValue = 2
                End If
                strBitLog = Trim(intReturnValue.ToString)

                'Kto 02
                'intReturnValue = FcCheckKonto(row("lngDebKtoNbr"), objfiBuha, row("dblDebMwSt"), 0)
                intReturnValue = 0
                strBitLog += Trim(intReturnValue.ToString)

                'Currency 03
                intReturnValue = FcCheckCurrency(row("strDebCur"))
                strBitLog += Trim(intReturnValue.ToString)

                'Sub 04
                'SplitBill oder GS
                If booSplittBill And IIf(IsDBNull(row("intRGArt")), 0, row("intRGArt")) = 10 Then
                    row("booLinked") = True
                Else
                    row("booLinked") = False
                End If
                If booLinkedGS And IIf(IsDBNull(row("intRGArt")), 0, row("intRGArt")) = 11 Then
                    row("booGS") = True
                Else
                    row("booGS") = False
                End If


                intReturnValue = FcCheckSubBookings(row("strDebRGNbr"),
                                                    dsDebitoren.Tables("tblDebiSubsFromUser"),
                                                    intSubNumber,
                                                    dblSubBrutto,
                                                    dblSubNetto,
                                                    dblSubMwSt,
                                                    row("datDebValDatum"),
                                                    row("intBuchungsart"),
                                                    booAutoCorrect,
                                                    booCpyKSTToSub,
                                                    IIf(IsDBNull(row("lngDebiKST")), 0, row("lngDebiKST")),
                                                    row("lngDebKtoNbr"),
                                                    booCashSollCorrect,
                                                    row("booLinked"),
                                                    row("booGS"))

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
                    Debug.Print("Zeile eingefügt in DebiSubs Bezahlbuchung " + row("strDebRGNbr"))
                    'dsDebitoren.Tables("tblDebiSubsFromUser").AcceptChanges()
                End If

                'Bei SplitBill - erste Rechnung evtl. Rückzahlung im Total nicht beachten
                If row("booLinked") And row("intRGArt") = 1 And IIf(IsDBNull(row("lngLinkedRG")), 0, row("lngLinkedRG")) > 0 Then
                    row("dblDebBrutto") = Decimal.Round(dblSubBrutto, 2, MidpointRounding.AwayFromZero) * -1
                    row("dblDebNetto") = Decimal.Round(dblSubNetto, 2, MidpointRounding.AwayFromZero) * -1
                    row("dblDebMwSt") = Decimal.Round(dblSubMwSt, 2, MidpointRounding.AwayFromZero) * -1
                End If

                'Bei GS wie SplitBill
                If row("booGS") And row("intRGArt") = 1 And IIf(IsDBNull(row("lngLinkedGS")), 0, row("lngLinkedGS")) > 0 Then
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
                            Debug.Print("Rundungsdifferenz eingefügt in SB " + row("strDebRGNbr"))
                            'dsDebitoren.Tables("tblDebiSubsFromUser").AcceptChanges()

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
                intReturnValue = FcCheckBelegHead(row("intBuchungsart"),
                                                  IIf(IsDBNull(row("dblDebBrutto")), 0, row("dblDebBrutto")),
                                                  IIf(IsDBNull(row("dblDebNetto")), 0, row("dblDebNetto")),
                                                  IIf(IsDBNull(row("dblDebMwSt")), 0, row("dblDebMwSt")),
                                                  dblRDiffBrutto)
                strBitLog += Trim(intReturnValue.ToString)

                'Referenz 08
                If IIf(IsDBNull(row("strDebReferenz")), "", row("strDebReferenz")) = "" And row("intBuchungsart") = 1 Then
                    intReturnValue = FcCreateDebRef(BgWCheckDebiArgsInProc.intMandant,
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
                            intReturnValue = FcIsPrivateDebitorCreatable(intDebitorNew,
                                                                         BgWCheckDebiArgsInProc.strMandant,
                                                                         BgWCheckDebiArgsInProc.intMandant)
                        Else
                            intReturnValue = FcIsDebitorCreatable(intDebitorNew,
                                                                  BgWCheckDebiArgsInProc.strMandant,
                                                                  BgWCheckDebiArgsInProc.intMandant)
                        End If
                        If intReturnValue = 0 Then
                            strStatus += " erstellt"
                            row("strDebBez") = FcReadDebitorName(intDebitorNew,
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
                    row("strDebBez") = FcReadDebitorName(intDebitorNew,
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
                intReturnValue = FcCheckOPDouble(row("lngDebNbr"),
                                                 IIf(IsDBNull(row("lngDebIdentNbr")), 0, row("lngDebIdentNbr")),
                                                 row("strOPNr"),
                                                 IIf(row("dblDebBrutto") > 0, "R", "G"),
                                                 row("strDebCur"),
                                                 booErfOPExt)
                strBitLog += Trim(intReturnValue.ToString)

                'PGV => Prüfung vor Valuta-Datum da Valuta-Datum verändert wird
                If Not IsDBNull(row("datPGVFrom")) Then
                    row("booPGV") = True
                End If

                'Bei Datum-Korrektur vorgängig Datum ersetzen um PGV-Buchungen zu verhindern
                If booValutaCorrect Then
                    If row("datDebRGDatum") < datValutaCorrect Then
                        row("datDebRGDatum") = datValutaCorrect.ToShortDateString
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "RgDCor"
                    End If
                    If row("datDebValDatum") < datValutaCorrect Then
                        row("datDebValDatum") = datValutaCorrect.ToShortDateString
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "ValDCor"
                    End If
                End If

                'Evtl. End-Datum Korrektur
                If booValutaEndCorrect Then
                    If row("datDebRGDatum") > datValutaEndCorrect Then
                        row("datDebRGDatum") = datValutaEndCorrect.ToShortDateString
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "RgDECor"
                    End If
                    If row("datDebValDatum") > datValutaEndCorrect Then
                        row("datDebValDatum") = datValutaEndCorrect.ToShortDateString
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "ValDECor"
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
                intReturnValue = FcCheckDate2(IIf(IsDBNull(row("datDebValDatum")), #1789-09-17#, row("datDebValDatum")),
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
                            intReturnValue = FcCheckDate2(row("datPGVTo"),
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
                                intReturnValue = FcCheckDate2(DateAndTime.DateAdd(DateInterval.Month, intMonthCounter, row("datPGVFrom")),
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
                intReturnValue = FcCheckDate2(IIf(IsDBNull(row("datDebRGDatum")), #1789-09-17#, row("datDebRGDatum")),
                                              BgWCheckDebiArgsInProc.strYear,
                                              dsDebitoren.Tables("tblDebitorenDates"),
                                              False)

                strBitLog += Trim(intReturnValue.ToString)

                'Interne Bank 12
                If IsDBNull(row("intPayType")) Then
                    row("intPayType") = 9
                End If
                intReturnValue = FcCheckDebiIntBank(BgWCheckDebiArgsInProc.intMandant,
                                                                IIf(IsDBNull(row("strDebiBank")), "", row("strDebiBank")),
                                                                row("intPayType"),
                                                                intiBankSage200)
                strBitLog += Trim(intReturnValue.ToString)

                'Bei SplittBill: Existiert verlinkter Beleg? 13
                If row("booLinked") Then
                    'dsDebitoren.Tables("tblDebiHeadsFromUser").AcceptChanges()
                    'Zuerst Debitor von erstem Beleg suchen
                    intDebitorNew = FcGetDebitorFromLinkedRG(IIf(IsDBNull(row("lngLinkedRG")), 0, row("lngLinkedRG")),
                                                                         BgWCheckDebiArgsInProc.intMandant,
                                                                         intLinkedDebitor,
                                                                         BgWCheckDebiArgsInProc.intTeqNbr,
                                                                         BgWCheckDebiArgsInProc.intTeqNbrLY,
                                                                         BgWCheckDebiArgsInProc.intTeqNbrPLY)
                    row("lngLinkedDeb") = intLinkedDebitor

                    intReturnValue = FcCheckLinkedRG(intLinkedDebitor,
                                                                 row("strDebCur"),
                                                                 row("lngLinkedRG"),
                                                                 row("dblDebBrutto"),
                                                                 BgWCheckDebiArgsInProc.strYear)
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

                    'dsDebitoren.Tables("tblDebiSubsFromUser").AcceptChanges()

                    For Each SBsubrow As DataRow In selSBrows
                        'Debug.Print("SB gelöscht da Linked RG")
                        If IIf(IsDBNull(SBsubrow("strArtikel")), "", SBsubrow("strArtikel")) <> "Rundungsdifferenz" Then
                            SBsubrow.Delete()
                        End If
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
                    dsDebitoren.Tables("tblDebiSubsFromUser").AcceptChanges()
                    drSBBuchung = Nothing
                    Debug.Print("SB eingefügt von Main ohne MWst bei Linked " + row("strDebRGNbr"))
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
                    intReturnValue = FcGetDZKondSageID(row("intZKond"),
                                                                   intDZKondS200)
                    row("intZKond") = intDZKondS200
                End If
                If row("intZKondT") = 1 And row("intZKond") = 0 Then
                    'Fall kein Privatekunde
                    If booPKPrivate = False Then
                        'Daten aus den Tab_Repbetriebe holen
                        intReturnValue = FcGetDZkondFromRep(row("lngDebNbr"),
                                                                    intDZKond,
                                                                    BgWCheckDebiArgsInProc.intMandant)
                    Else
                        'Daten aus der t_customer holen
                        intReturnValue = FcGetDZkondFromCust(row("lngDebNbr"),
                                                                         intDZKond,
                                                                         BgWCheckDebiArgsInProc.intMandant)
                    End If
                    row("intZKond") = intDZKond
                End If
                'Prüfem ob Zahlungs-Kondition - ID existiert in Sage 200 bei Mandant
                'strMandant = Main.FcReadFromSettingsII("Buchh200_Name",
                '                                BgWCheckDebiArgsInProc.intMandant)
                intReturnValue = FcCheckDZKond(strMandant,
                                                           row("intZKond"))
                strBitLog += Trim(intReturnValue.ToString)

                'GS Check 15
                If row("booGS") Then
                    'Zuerst Debitor von erstem Beleg suchen
                    intDebitorNew = FcGetDebitorFromLinkedRG(IIf(IsDBNull(row("lngLinkedGS")), 0, row("lngLinkedGS")),
                                                                         BgWCheckDebiArgsInProc.intMandant,
                                                                         intLinkedDebitor,
                                                                         BgWCheckDebiArgsInProc.intTeqNbr,
                                                                         BgWCheckDebiArgsInProc.intTeqNbrLY,
                                                                         BgWCheckDebiArgsInProc.intTeqNbrPLY)
                    row("lngLinkedGSDeb") = intLinkedDebitor

                    intReturnValue = FcCheckLinkedRG(intLinkedDebitor,
                                                                 row("strDebCur"),
                                                                 row("lngLinkedGS"),
                                                                 row("dblDebBrutto"),
                                                                 BgWCheckDebiArgsInProc.strYear)

                    'Falls erste Rechnung bezahlt, dann Flag setzen
                    If intReturnValue = 2 Then
                        row("booLinkedPayed") = True

                    Else
                        row("booLinkedPayed") = False

                    End If


                End If
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
                    row("strDebKtoBez") = FcReadDebitorKName(row("lngDebKtoNbr"))
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
                    ElseIf Mid(strBitLog, 13, 1) = "2" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "SplBBez1"
                    Else
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "SplBRG1<GS"
                    End If

                End If
                'Zahlungs-Kondition
                If Mid(strBitLog, 14, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "ZKond"
                End If
                'GS
                If Mid(strBitLog, 15, 1) <> "0" Then
                    If Mid(strBitLog, 15, 1) = "1" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "GSNoRG1"
                    ElseIf Mid(strBitLog, 15, 1) = "2" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "GSBezRG1"
                    Else
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "GSRG1<GS"
                    End If

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
                If Val(strBitLog) = 0 Or Val(strBitLog) = 10000022000 Or Val(strBitLog) = 22000 Or Val(strBitLog) = 10000000000 Then
                    row("booDebBook") = True
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "ok"
                End If
                row("strDebStatusText") = strStatus
                row("strDebStatusBitLog") = strBitLog

                'Wird ein anderer Text in der Head-Buchung gewünscht?
                'booDiffHeadText = IIf(Main.FcReadFromSettingsII("Buchh_TextSpecial", intAccounting) = "0", False, True)
                If booDiffHeadText Then
                    'dsDebitoren.Tables("tblDebiHeadsFromUser").AcceptChanges()
                    strDebiHeadText = FcSQLParse(FcReadFromSettingsII("Buchh_TextSpecialText",
                                                                                BgWCheckDebiArgsInProc.intMandant),
                                                             row("strDebRGNbr"),
                                                             dsDebitoren.Tables("tblDebiHeadsFromUser"),
                                                             "D")
                    row("strDebText") = strDebiHeadText
                End If

                'Wird ein anderer Text in den Sub-Buchung gewünscht?
                'booDiffSubText = IIf(Main.FcReadFromSettingsII("Buchh_SubTextSpecial", intAccounting) = "0", False, True)
                If booDiffSubText And Not row("booLinked") Then
                    strDebiSubText = FcSQLParse(FcReadFromSettingsII("Buchh_SubTextSpecialText",
                                                                               BgWCheckDebiArgsInProc.intMandant),
                                                            row("strDebRGNbr"),
                                                            dsDebitoren.Tables("tblDebiHeadsFromUser"),
                                                            "D")
                Else
                    strDebiSubText = IIf(IsDBNull(row("strDebText")), "NoText", row("strDebText"))
                End If
                'Falls nicht SB - Linked dann Text in SB ersetzen
                If Not row("booLinked") Then
                    'dsDebitoren.Tables("tblDebiSubsFromUser").AcceptChanges()
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
            Err.Clear()

        Finally
            'objFinanz = Nothing
            'objfiBuha = Nothing
            'objdbBuha = Nothing
            'objdbPIFb = Nothing
            'objFiBebu = Nothing
            selSBrows = Nothing
            selsubrow = Nothing

            dsDebitoren.Tables("tblDebiHeadsFromUser").AcceptChanges()

            'System.GC.Collect()

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
        Dim intFcReturns As Int16
        Dim strFcreturns As String
        Dim intActRGNbr As Int32
        Dim intTotRGs As Int32
        Dim intZV As Int32


        'Dim objFinanz As New SBSXASLib.AXFinanz
        'Dim objfiBuha As New SBSXASLib.AXiFBhg
        'Dim objdbBuha As New SBSXASLib.AXiDbBhg
        'Dim objdbPIFb As New SBSXASLib.AXiPlFin
        'Dim objFiBebu As New SBSXASLib.AXiBeBu
        'Dim objKrBuha As New SBSXASLib.AXiKrBhg

        Try

            'Me.Cursor = Cursors.WaitCursor
            'Button deaktivieren
            'Me.butImport.Enabled = False

            'Finanz-Obj init
            'objFinanz = Nothing
            'objFinanz = New SBSXASLib.AXFinanz

            ''Login
            'Call objFinanz.ConnectSBSdb(System.Configuration.ConfigurationManager.AppSettings("OwnSageServer"),
            '                        System.Configuration.ConfigurationManager.AppSettings("OwnSageDB"),
            '                        System.Configuration.ConfigurationManager.AppSettings("OwnSageID"),
            '                        System.Configuration.ConfigurationManager.AppSettings("OwnSagePsw"), "")

            'intFcReturns = FcReadFromSettingsIII("Buchh200_Name",
            '                                BgWImportDebiArgsInProc.intMandant,
            '                                strMandant)

            'booAccOk = objFinanz.CheckMandant(strMandant)
            ''Open Mandant
            'objFinanz.OpenMandant(strMandant, strPeriode)
            'objfiBuha = objFinanz.GetFibuObj()
            'objdbBuha = objFinanz.GetDebiObj()
            'objdbPIFb = objfiBuha.GetCheckObj()
            'objFiBebu = objFinanz.GetBeBuObj()
            'objKrBuha = objFinanz.GetKrediObj()


            'Start in Sync schreiben
            'intReturnValue = WFDBClass.FcWriteStartToSync(objdbConnZHDB02,
            '                                              BgWImportDebiArgsInProc.intMandant,
            '                                              1,
            '                                              dsDebitoren.Tables("tblDebiHeadsFromUser").Rows.Count)

            'Setting soll erfasste OP als externe Beleg-Nr. genommen werden und lngDebIdentNbr als Beleg-Nr.
            intFcReturns = FcReadFromSettingsIII("Buchh_ErfOPExt", BgWImportDebiArgsInProc.intMandant, strFcreturns)
            booErfOPExt = Convert.ToBoolean(Convert.ToInt16(strFcreturns))
            intFcReturns = FcReadFromSettingsIII("Buchh200_Name",
                                                 BgWImportDebiArgsInProc.intMandant,
                                                 strMandant)

            intTotRGs = dsDebitoren.Tables("tblDebiHeadsFromUser").Rows.Count
            intActRGNbr = 0

            objdbSQLcommand.Connection = objdbMSSQLConn

            'Kopfbuchung
            For Each row In dsDebitoren.Tables("tblDebiHeadsFromUser").Rows

                'Progress - Bar
                intActRGNbr += 1
                BgWImportDebi.ReportProgress(100 / intTotRGs * intActRGNbr)

                If IIf(IsDBNull(row("booDebBook")), False, row("booDebBook")) Then

                    'Für Err-Msg
                    strRGNbr = row("strDebRGNbr")

                    'Test ob OP - Buchung
                    If row("intBuchungsart") = 1 Then

                        'Verdopplung interne BelegsNummer verhindern
                        objdbBuha.CheckDoubleIntBelNbr = "J"

                        If row("dblDebBrutto") <0 Then
                            'Gutschrift
                            'Falls booGSToInv (Gutschrift zu Rechnung) dann OP-Nummer vorgeben, sonst hochzählen lassen
                            If row("booCrToInv") Then
                                'Beleg-Nummerierung desaktivieren
                                objdbBuha.IncrBelNbr = "N"
                                'Eingelesene OP-Nummer (=Verknüpfte OP-Nr.) = interne Beleg-Nummer
                                intDebBelegsNummer = FcCleanRGNrStrict(row("strOPNr"))
                                strExtBelegNbr = row("strDebRGNbr")
                            Else
                                'Zuerst Beleg-Nummerieungung aktivieren
                                objdbBuha.IncrBelNbr = "J"
                                'Belegsnummer abholen
                                intDebBelegsNummer = objdbBuha.GetNextBelNbr("G")
                                'Prüfen ob wirklich frei und falls nicht hochzählen
                                intReturnValue = FcCheckDebiExistance(intDebBelegsNummer,
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
                                intReturnValue = FcCheckDebiExistance(intDebBelegsNummer,
                                                                                  "R",
                                                                                  BgWImportDebiArgsInProc.intTeqNbr)
                            Else
                                If Strings.Len(FcCleanRGNrStrict(row("strOPNr"))) > 9 Then
                                    'Zahl zu gross
                                    objdbBuha.IncrBelNbr = "J"
                                    'Belegsnummer abholen
                                    intDebBelegsNummer = objdbBuha.GetNextBelNbr("R")
                                    intReturnValue = FcCheckDebiExistance(intDebBelegsNummer,
                                                                                      "R",
                                                                                      BgWImportDebiArgsInProc.intTeqNbr)
                                    strExtBelegNbr = row("strOPNr")
                                Else
                                    'Beleg-Nummerierung abschalten
                                    objdbBuha.IncrBelNbr = "N"
                                    'Gemäss Setting Erfasste OP-Nr. Nummern vergeben
                                    If Not booErfOPExt Then
                                        intDebBelegsNummer = FcCleanRGNrStrict(row("strOPNr"))
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
                            strDebiText = IIf(IsDBNull(row("strDebText")), "n/a", row("strDebText"))
                        End If
                        strCurrency = row("strDebCur")
                        If strCurrency <> "CHF" Then 'Muss ergänzt werden => Was ist Leitwährung auf dem Konto
                            dblKurs = FcGetKurs(strCurrency,
                                                     strValutaDatum)
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
                                    intReturnValue = FcGetSteuerFeld(strSteuerFeld,
                                                                         SubRow("lngKto"),
                                                                         SubRow("strDebSubText"),
                                                                         SubRow("dblBrutto") * -1,
                                                                         SubRow("strMwStKey"),
                                                                         SubRow("dblMwSt") * -1)
                                Else
                                    intReturnValue = FcGetSteuerFeld(strSteuerFeld,
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
                                        'ZV suchen
                                        intReturnValue = FcGetZV(objdbMSSQLConn,
                                                        objdbSQLcommand,
                                                        strMandant,
                                                        "SB",
                                                        intZV)

                                        Call objdbBuha.SetZahlung(intZV,
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
                                                              row("lngDebIdentNbr").ToString + ", TZ SB " + row("strDebRGNbr").ToString)

                                        Call objdbBuha.WriteTeilzahlung4(intLaufNbr.ToString,
                                                                     row("lngDebIdentNbr").ToString + ", TZ SB " + row("strDebRGNbr").ToString,
                                                                     "NOT_SET",
                                                                     ,
                                                                     "NOT_SET",
                                                                     "NOT_SET",
                                                                     "DEFAULT",
                                                                     "DEFAULT")

                                    End If


                                End If

                            End If

                            'Bei GS soeben gebuchte GS TZ, TZ auf RG1, Belege werden durch Dispatcher ausgebucht
                            If row("booGS") And Mid(row("strDebStatusBitLog"), 15, 1) = "0" Then 'Nur wenn Beleg in gleicher Buha
                                'Zuerst TZ auf GS
                                'Laufnummer von GS holen
                                intLaufNbr = objdbBuha.doesBelegExist2(intDebitorNbr.ToString,
                                                                       row("strDebCur"),
                                                                       intDebBelegsNummer.ToString,
                                                                       "NOT_SET",
                                                                       "G",
                                                                       "NOT_SET",
                                                                       "NOT_SET",
                                                                       "NOT_SET")

                                If intLaufNbr > 0 Then

                                    'ZV suchen
                                    intFcReturns = FcGetZV(objdbMSSQLConn,
                                                        objdbSQLcommand,
                                                        strMandant,
                                                        "GS",
                                                        intZV)

                                    Call objdbBuha.SetZahlung(intZV,
                                                          strBelegDatum,
                                                          strValutaDatum,
                                                          row("strDebCur"),
                                                          dblKurs,
                                                          "",
                                                          "",
                                                          row("lngDebNbr"),
                                                          (dblBetrag * -1).ToString,
                                                          row("strDebCur"),
                                                          ,
                                                          row("lngDebIdentNbr").ToString + ", TZ GS " + row("lngLinkedGS").ToString)

                                    Call objdbBuha.WriteTeilzahlung4(intLaufNbr.ToString,
                                                                     row("lngDebIdentNbr").ToString + ", TZ GS " + row("lngLinkedGS").ToString,
                                                                     "NOT_SET",
                                                                     ,
                                                                     "NOT_SET",
                                                                     "NOT_SET",
                                                                     "DEFAULT",
                                                                     "DEFAULT")

                                    'TZ Auf RG1
                                    intLaufNbr = objdbBuha.doesBelegExist2(row("lngLinkedGSDeb").ToString,
                                                                       row("strDebCur"),
                                                                       row("lngLinkedGS").ToString,
                                                                       "NOT_SET",
                                                                       "R",
                                                                       "NOT_SET",
                                                                       "NOT_SET",
                                                                       "NOT_SET")

                                    If intLaufNbr > 0 Then

                                        Call objdbBuha.SetZahlung(intZV,
                                                                  strBelegDatum,
                                                                  strValutaDatum,
                                                                  row("strDebCur"),
                                                                  dblKurs,
                                                                  "",
                                                                  "",
                                                                  row("lngLinkedGSDeb"),
                                                                  dblBetrag.ToString,
                                                                  row("strDebCur"),
                                                                  ,
                                                                  row("lngDebIdentNbr").ToString + ", TZ GS " + row("strDebRGNbr").ToString)

                                        Call objdbBuha.WriteTeilzahlung4(intLaufNbr.ToString,
                                                                     row("lngDebIdentNbr").ToString + ", TZ GS " + row("strDebRGNbr").ToString,
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
                            If (Err.Number And 65535) < 10000 Then
                                strErrMessage = "Belegerstellung RG " + strRGNbr + " Beleg " + intDebBelegsNummer.ToString + " NICHT möglich!"
                                MessageBox.Show(ex.Message + vbCrLf + strErrMessage, "Problem " + (Err.Number And 65535).ToString, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                booBooingok = False
                            Else
                                If (Err.Number And 65535) = 10030 Then
                                    'MwSt-7.7/8.1 überschneidung nichts machen
                                    booBooingok = True
                                Else
                                    strErrMessage = "Belegerstellung RG " + strRGNbr + " Beleg " + intDebBelegsNummer.ToString + " möglich mit Warnung"
                                    MessageBox.Show(ex.Message, "Warnung " + (Err.Number And 65535).ToString, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                    booBooingok = True
                                End If
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
                        strDebiText = row("strDebText")
                        strCurrency = row("strDebCur")
                        If strCurrency <> "CHF" Then 'Muss ergänzt werden => Was ist Leitwährung auf dem Konto
                            dblKurs = FcGetKurs(strCurrency,
                                                strValutaDatum)
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
                                    dblKursSoll = FcGetKurs(strCurrency,
                                                            strValutaDatum,
                                                            intSollKonto)
                                    'strSteuerInfo = Split(FBhg.GetKontoInfo(intSollKonto.ToString), "{>}")
                                    'Debug.Print("Konto-Info Soll: " + strSteuerInfo(26))
                                    dblSollBetrag = SubRow("dblNetto")
                                    strDebiTextSoll = SubRow("strDebSubText")
                                    If SubRow("dblMwSt") > 0 Then
                                        intReturnValue = FcGetSteuerFeld(strSteuerFeldSoll,
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
                                    dblKursHaben = FcGetKurs(strCurrency,
                                                             strValutaDatum,
                                                             intHabenKonto)
                                    'strSteuerInfo = Split(FBhg.GetKontoInfo(intSollKonto.ToString), "{>}")
                                    'Debug.Print("Konto-Info Haben: " + strSteuerInfo(26))
                                    dblHabenBetrag = SubRow("dblNetto") * -1
                                    'dblHabenBetrag = dblSollBetrag
                                    strDebiTextHaben = SubRow("strDebSubText")
                                    If SubRow("dblMwSt") * -1 > 0 Then
                                        intReturnValue = FcGetSteuerFeld(strSteuerFeldHaben,
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
                                If (Err.Number And 65535) < 10000 Then
                                    MessageBox.Show(ex.Message, "Schweres Problem " + (Err.Number And 65535).ToString + " Belegerstellung " + intDebBelegsNummer.ToString + ", RG " + strRGNbr, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                    booBooingok = False
                                Else
                                    If (Err.Number And 65535) = 10030 Then
                                        'MwSt 7.7/8.1 Überschneidung
                                        booAccOk = True
                                    Else
                                        MessageBox.Show(ex.Message, "Buchbares Problem " + (Err.Number And 65535).ToString + " Belegerstellung " + intDebBelegsNummer.ToString + ", RG " + strRGNbr, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                        booBooingok = True
                                    End If

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
                                        dblKursSoll = 1 / FcGetKurs(strCurrency,
                                                                    strValutaDatum,
                                                                    intSollKonto)
                                        dblSollBetrag = SubRow("dblNetto")
                                        If SubRow("dblMwSt") > 0 Then
                                            intReturnValue = FcGetSteuerFeld(strSteuerFeldSoll,
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
                                        dblKursHaben = 1 / FcGetKurs(strCurrency,
                                                                     strValutaDatum,
                                                                     intHabenKonto)
                                        dblHabenBetrag = SubRow("dblNetto") * -1
                                        If (SubRow("dblMwSt") * -1) > 0 Then
                                            intReturnValue = FcGetSteuerFeld(strSteuerFeldHaben,
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
                                If (Err.Number And 65535) < 10000 Then
                                    MessageBox.Show(ex.Message, "Scxhweres Problem " + (Err.Number And 65535).ToString + " Belegerstellung " + intDebBelegsNummer.ToString + ", RG " + strRGNbr, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                    booBooingok = False
                                Else
                                    If (Err.Number And 65535) = 10030 Then
                                        'MwSt 7.7/8.1 Überschneidung Keine Meldung
                                        booBooingok = True
                                    Else
                                        MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Belegerstellung " + intDebBelegsNummer.ToString + ", RG " + strRGNbr, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                        booBooingok = True
                                    End If

                                End If
                            End Try


                        End If

                    End If

                    If booBooingok Then
                        If row("booPGV") Then
                            'Bei PGV Buchungen
                            If IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")) = "" Or
                                    (IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")) = "RV" And row("intPGVMthsAY") + row("intPGVMthsNY") > 1) Then

                                intReturnValue = FcPGVDTreatment(dsDebitoren.Tables("tblDebiSubsFromUser"),
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


                            Else
                                'TA
                                intReturnValue = FcPGVDTreatmentYC(dsDebitoren.Tables("tblDebiSubsFromUser"),
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
                        'row("strDebBookStatus") = row("strDebStatusBitLog")
                        'row("booBooked") = True
                        'row("datBooked") = Now()
                        'row("lngBelegNr") = intDebBelegsNummer
                        'dsDebitoren.Tables("tblDebiHeadsFromUser").AcceptChanges()
                        'Application.DoEvents()

                        'Status in File RG-Tabelle schreiben
                        intReturnValue = FcWriteToRGTable(BgWImportDebiArgsInProc.intMandant,
                                                                          row("strDebRGNbr"),
                                                                          Now(),
                                                                          intDebBelegsNummer,
                                                                          objdbAccessConn,
                                                                          objOracleConn,
                                                                          objdbConnZHDB02,
                                                                          row("booDatChanged"),
                                                                          row("datDebRGDatum"),
                                                                          row("datDebValDatum"))
                        If intReturnValue <> 0 Then
                            'Throw an exception
                        End If

                        'Evtl.Query nach Buchung ausführen
                        Call FcExecuteAfterDebit(BgWImportDebiArgsInProc.intMandant)
                    End If


                End If

            Next

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Generelles Problem " + (Err.Number And 65535).ToString + " Belegerstellung " + intDebBelegsNummer.ToString + ", RG " + strRGNbr, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Err.Clear()

        Finally
            'Buhas freigeben
            'objKrBuha = Nothing
            'objFiBebu = Nothing
            'objdbPIFb = Nothing
            'objdbBuha = Nothing
            'objfiBuha = Nothing
            'objFinanz = Nothing

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

        Me.dsDebitoren.Reset()
        Me.dsDebitoren = Nothing
        'objKrBuha = Nothing
        'objFiBebu = Nothing
        'objdbPIFb = Nothing
        'objdbBuha = Nothing
        'objfiBuha = Nothing
        'objFinanz = Nothing

        Me.Dispose()
        'System.GC.Collect()
        'System.Diagnostics.Process.Start(Application.ExecutablePath)
        'Environment.Exit(0)
        'System.GC.Collect()
        'Application.Restart()

    End Sub


    Private Sub butCheckDeb_Click(sender As Object, e As EventArgs) Handles butCheckDeb.Click

        Dim objdbConn As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString"))
        Dim objdbtaskcmd As New MySqlCommand
        Dim objdbMSSQLConn As New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("SQLConnectionString"))
        Dim objdbSQLcommand As New SqlCommand

        Dim intFcReturns As Int16
        Dim strPeriode As String
        Dim strYearCh As String
        Dim BgWCheckDebitLocArgs As New BgWCheckDebitArgs
        Dim strErrC() As String
        Dim objtblErrC As New DataTable()
        Dim strErrFound() As DataRow
        Dim strErrDecoded As String
        Dim strErrResp As String
        'Dim objdbtasks As New DataTable

        'Dim intTeqNbr As Int32
        'Dim intTeqNbrLY As Int32
        'Dim intTeqNbrPLY As Int32
        'Dim strYear As String

        Dim objFinanzCopy As New SBSXASLib.AXFinanz
        'Dim objfiBuha As New SBSXASLib.AXiFBhg
        'Dim objdbBuha As New SBSXASLib.AXiDbBhg
        'Dim objdbPIFb As New SBSXASLib.AXiPlFin
        'Dim objFiBebu As New SBSXASLib.AXiBeBu
        'Dim objKrBuha As New SBSXASLib.AXiKrBhg


        Try

            objdbtaskcmd.Connection = objdbConn
            objdbtaskcmd.CommandText = "SELECT * FROM t_importer_errc"
            objdbtaskcmd.Connection.Open()
            objtblErrC.Load(objdbtaskcmd.ExecuteReader())
            objdbtaskcmd.Connection.Close()

            Me.Cursor = Cursors.WaitCursor
            UseWaitCursor = True

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

            strPeriode = lstBoxPerioden.GetItemText(lstBoxPerioden.SelectedItem)

            Call FcLoginSage3(objdbConn,
                                  objdbMSSQLConn,
                                  objdbSQLcommand,
                                  objFinanzCopy,
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
            If lstBoxPerioden.Items.Count > 1 Then

                'Gibt es ein Vorjahr?
                If lstBoxPerioden.SelectedIndex + 1 > 1 Then
                    strPeriode = lstBoxPerioden.Items(lstBoxPerioden.SelectedIndex - 1)
                    'Peeriodendef holen
                    Call FcLoginSage4(intMandant,
                                       dsDebitoren.Tables("tblDebitorenDates"),
                                       strPeriode)
                Else
                    'Periode ezreugen und auf N stellen
                    strYearCh = Convert.ToString(Val(strYear) - 1)
                    dsDebitoren.Tables("tblDebitorenDates").Rows.Add(strYearCh, "GJ n/a", DateSerial(Convert.ToUInt16(strYearCh), 1, 1), DateSerial(Convert.ToUInt16(strYearCh), 12, 31), "N")
                End If

                'Gibt es ein Folgehahr?
                If lstBoxPerioden.SelectedIndex + 1 < lstBoxPerioden.Items.Count Then
                    strPeriode = lstBoxPerioden.Items(lstBoxPerioden.SelectedIndex + 1)
                    'Peeriodendef holen
                    Call FcLoginSage4(intMandant,
                                       dsDebitoren.Tables("tblDebitorenDates"),
                                       strPeriode)
                Else
                    'Periode ezreugen und auf N stellen
                    strYearCh = Convert.ToString(Val(strYear) + 1)
                    dsDebitoren.Tables("tblDebitorenDates").Rows.Add(strYearCh, "GJ n/a", DateSerial(Convert.ToUInt16(strYearCh), 1, 1), DateSerial(Convert.ToUInt16(strYearCh), 12, 31), "N")
                End If
            ElseIf lstBoxPerioden.Items.Count = 1 Then 'es gibt genau 1 Jahr
                'gewähltes Jahr checken
                Call FcLoginSage4(intMandant,
                                       dsDebitoren.Tables("tblDebitorenDates"),
                                       strPeriode)
                'VJ erzeugen
                strYearCh = Convert.ToString(Val(strYear) - 1)
                dsDebitoren.Tables("tblDebitorenDates").Rows.Add(strYearCh, "GJ n/a", DateSerial(Convert.ToUInt16(strYearCh), 1, 1), DateSerial(Convert.ToUInt16(strYearCh), 12, 31), "N")

                'FJ erzeugen
                strYearCh = Convert.ToString(Val(strYear) + 1)
                dsDebitoren.Tables("tblDebitorenDates").Rows.Add(strYearCh, "GJ n/a", DateSerial(Convert.ToUInt16(strYearCh), 1, 1), DateSerial(Convert.ToUInt16(strYearCh), 12, 31), "N")

            End If

            MySQLdaDebitoren.Fill(dsDebitoren, "tblDebiHeadsFromUser")
            MySQLdaDebitorenSub.Fill(dsDebitoren, "tblDebiSubsFromUser")

            BgWCheckDebitLocArgs.intMandant = intMandant
            BgWCheckDebitLocArgs.strMandant = frmImportMain.lstBoxMandant.GetItemText(frmImportMain.lstBoxMandant.SelectedItem)
            BgWCheckDebitLocArgs.intTeqNbr = intTeqNbr
            BgWCheckDebitLocArgs.intTeqNbrLY = intTeqNbrLY
            BgWCheckDebitLocArgs.intTeqNbrPLY = intTeqNbrPLY
            BgWCheckDebitLocArgs.strYear = strYear
            BgWCheckDebitLocArgs.strPeriode = lstBoxPerioden.GetItemText(lstBoxPerioden.SelectedItem)
            BgWCheckDebitLocArgs.booValutaCor = frmImportMain.chkValutaCorrect.Checked
            BgWCheckDebitLocArgs.datValutaCor = frmImportMain.dtpValutaCorrect.Value
            BgWCheckDebitLocArgs.booValutaEndCor = frmImportMain.chkValutaEndCorrect.Checked
            BgWCheckDebitLocArgs.datValutaEndCor = frmImportMain.dtpValutaEndCorrect.Value

            BgWCheckDebi.RunWorkerAsync(BgWCheckDebitLocArgs)

            Do While BgWCheckDebi.IsBusy
                Application.DoEvents()
            Loop

            'Grid neu aufbauen
            dgvDates.DataSource = dsDebitoren.Tables("tblDebitorenDates")
            dgvInfo.DataSource = dsDebitoren.Tables("tblDebitorenInfo")
            dgvBookings.DataSource = dsDebitoren.Tables("tblDebiHeadsFromUser")
            dgvBookingSub.DataSource = dsDebitoren.Tables("tblDebiSubsFromUser")

            'Tooltip für Fehler
            For Each dgvr As DataGridViewRow In dgvBookings.Rows
                strErrDecoded = ""
                strErrC = Split(dgvr.Cells("strDebStatusText").Value, ",")
                For Each strErrCElement As String In strErrC
                    strErrResp = ""
                    strErrFound = objtblErrC.Select("code Like '" + Strings.Trim(strErrCElement) + "'")
                    If strErrFound.Length > 0 Then
                        If strErrFound(0).Item("resp_it") > 0 Then
                            strErrResp = "IT " + strErrFound(0).Item("resp_it").ToString
                        End If
                        If strErrFound(0).Item("resp_ac") > 0 Then
                            strErrResp += IIf(strErrResp <> "", ", ", "") + "AC " + strErrFound(0).Item("resp_ac").ToString
                        End If
                        If strErrFound(0).Item("resp_bs") > 0 Then
                            strErrResp += IIf(strErrResp <> "", ", ", "") + "BS " + strErrFound(0).Item("resp_bs").ToString
                        End If
                        If strErrFound(0).Item("resp_ab") > 0 Then
                            strErrResp += IIf(strErrResp <> "", ", ", "") + "AB " + strErrFound(0).Item("resp_ab").ToString
                        End If
                        strErrDecoded += strErrFound(0).Item("explained") + vbTab + strErrResp + vbCrLf
                    End If
                Next
                dgvr.Cells("strDebStatusText").ToolTipText = strErrDecoded
            Next

            intFcReturns = FcInitdgvInfo(dgvInfo)
            intFcReturns = FcInitdgvBookings(dgvBookings)
            intFcReturns = FcInitdgvDebiSub(dgvBookingSub)
            intFcReturns = FcInitdgvDate(dgvDates)
            'Anzahl schreiben
            Me.TSLblNmbr.Text = dsDebitoren.Tables("tblDebiHeadsFromUser").Rows.Count
            If dsDebitoren.Tables("tblDebiHeadsFromUser").Rows.Count > 0 Then
                butImport.Enabled = True
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Generelles Problem Check" + (Err.Number And 65535).ToString, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Err.Clear()


        Finally
            'objFinanz = Nothing
            'objfiBuha = Nothing
            'objdbBuha = Nothing
            'objdbPIFb = Nothing
            'objFiBebu = Nothing
            'objKrBuha = Nothing
            objdbConn = Nothing
            objdbMSSQLConn = Nothing
            objdbSQLcommand = Nothing
            butCheckDeb.Enabled = False
            UseWaitCursor = False
            Me.Cursor = Cursors.Default

        End Try


    End Sub

    Friend Function FcReadFromSettingsIII(strField As String,
                                                intMandant As Int16,
                                                ByRef strReturn As String) As Int16

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
            strReturn = objlocdtSetting.Rows(0).Item(0).ToString
            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Einstellung lesen")
            Err.Clear()
            Return 1

        Finally
            objlocdtSetting.Constraints.Clear()
            objlocdtSetting.Rows.Clear()
            objlocdtSetting.Columns.Clear()
            objlocdtSetting = Nothing
            objlocMySQLcmd = Nothing
            objdbconn = Nothing
            'System.GC.Collect()

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
            'dbPathAndFile = "\\sdlc.mssag.ch\Apps\Backends\" + strMDBName + ";Jet OLEDB:System Database=\\sdlc.mssag.ch\Apps\Backends\Workbench.mdw;User ID=HagerR;Persist Security Info=False;Connect Timeout=300;"
            dbPathAndFile = "\\sdlc.mssag.ch\Apps\Backends\" + strMDBName + ";Jet OLEDB:System Database=\\sdlc.mssag.ch\Apps\Backends\Workbench.mdw;User ID=HagerR;Persist Security Info=False;"
            objaccesscon.ConnectionString = dbProvider + dbSource + dbPathAndFile
            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message)
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
            inscmdFields += ", lngLinkedGS"
            inscmdValues += ", @lngLinkedGS"
            inscmdFields += ", strRGName"
            inscmdValues += ", @strRGName"
            inscmdFields += ", strDebIdentNbr2"
            inscmdValues += ", @strDebIdentNbr2"
            inscmdFields += ", strRGBemerkung"
            inscmdValues += ", @strRGBemerkung"
            inscmdFields += ", booCrToInv"
            inscmdValues += ", @booCrToInv"
            inscmdFields += ", datPGVFrom"
            inscmdValues += ", @datPGVFrom"
            inscmdFields += ", datPGVTo"
            inscmdValues += ", @datPGVTo"
            inscmdFields += ", intZKond"
            inscmdValues += ", @intZKond"



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
            mysqlinscmd.Parameters.Add("@lngLinkedGS", MySqlDbType.Int32).SourceColumn = "lngLinkedGS"
            mysqlinscmd.Parameters.Add("@strRGName", MySqlDbType.String).SourceColumn = "strRGName"
            mysqlinscmd.Parameters.Add("@strDebIdentNbr2", MySqlDbType.String).SourceColumn = "strDebIdentNbr2"
            mysqlinscmd.Parameters.Add("@strRGBemerkung", MySqlDbType.String).SourceColumn = "strRGBemerkung"
            mysqlinscmd.Parameters.Add("@booCrToInv", MySqlDbType.Int16).SourceColumn = "booCrToInv"
            mysqlinscmd.Parameters.Add("@datPGVFrom", MySqlDbType.Date).SourceColumn = "datPGVFrom"
            mysqlinscmd.Parameters.Add("@datPGVTo", MySqlDbType.Date).SourceColumn = "datPGVTo"
            mysqlinscmd.Parameters.Add("@intZKond", MySqlDbType.Int16).SourceColumn = "intZKond"

            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem HeadCommand Init", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 1

        End Try

    End Function

    Friend Function FcSQLParse2(ByVal strSQLToParse As String,
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
                            strField = FcGetKundenzeichen2(RowBooking(0).Item("lngDebIdentNbr"))
                        Case Else
                            strField = "unknown field"
                    End Select
                    strSQLToParse = Strings.Left(strSQLToParse, intPipePositionBegin - 1) & strField & Strings.Right(strSQLToParse, Len(strSQLToParse) - intPipePositionEnd)
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

    Friend Function FcLoginSage3(ByRef objdbconn As MySqlConnection,
                                       ByRef objsqlConn As SqlClient.SqlConnection,
                                       ByRef objsqlCom As SqlClient.SqlCommand,
                                       ByRef objFinanz As SBSXASLib.AXFinanz,
                                       ByVal intAccounting As Int16,
                                       ByRef objdtInfo As DataTable,
                                       ByRef objdtDates As DataTable,
                                       ByVal strPeriod As String,
                                       ByRef strYear As String,
                                       ByRef intTeqNbr As Int16,
                                       ByRef intTeqNbrLY As Int16,
                                       ByRef intTeqNbrPLY As Int16,
                                       ByRef datPeriodFrom As Date,
                                       ByRef datPeriodTo As Date,
                                       ByRef strPeriodStatus As String) As Int16

        '0=ok, 1=Fibu nicht ok, 2=Debi nicht ok, 3=Debi nicht ok
        Dim booAccOk As Boolean
        Dim strMandant As String
        Dim strLogonInfo() As String
        Dim strPeriode() As String
        Dim FcReturns As Int16
        Dim intPeriodenNr As Int16
        'Dim strPeriodenInfo As String
        Dim objdtPeriodeLY As New DataTable
        Dim strPeriodeLY As String
        Dim strPeriodePLY As String
        Dim objdbcmd As New MySqlCommand
        Dim dtPeriods As New DataTable


        Try

            'objFinanz = Nothing
            'objFinanz = New SBSXASLib.AXFinanz

            'Application.DoEvents()

            'Login
            Try
                Call objFinanz.ConnectSBSdb(System.Configuration.ConfigurationManager.AppSettings("OwnSageServer"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageDB"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageID"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSagePsw"), "")
            Catch inEx As Exception
                If inEx.HResult <> -2147473602 Then
                    MessageBox.Show(inEx.Message, "Connect to Sage - DB " + Err.Number.ToString)
                    Exit Function
                End If

            End Try

            'objdbconn.Open()
            FcReturns = FcReadFromSettingsIII("Buchh200_Name",
                                                intAccounting,
                                                strMandant)
            'objdbconn.Close()
            booAccOk = objFinanz.CheckMandant(strMandant)

            'Open Mandantg
            objFinanz.OpenMandant(strMandant, strPeriod)

            'Von Login aktuelle Periode auslesen
            strLogonInfo = Split(objFinanz.GetLogonInfo(), "{>}")
            objdtInfo.Rows.Add("Man/Periode", strMandant + "/" + strLogonInfo(7) + ", " + intAccounting.ToString)

            'Check Periode
            intPeriodenNr = objFinanz.ReadPeri(strMandant, strLogonInfo(7))
            strPeriodenInfo = objFinanz.GetPeriListe(0)

            strPeriode = Split(strPeriodenInfo, "{>}")

            'Teq-Nr von Vorjar lesen um in Suche nutzen zu können
            objdtPeriodeLY.Rows.Clear()
            strPeriodeLY = (Val(Strings.Left(strPeriode(4), 4)) - 1).ToString + Strings.Right(strPeriode(4), 4)
            objsqlCom.CommandText = "SELECT teqnbr FROM periode WHERE mandid='" + strMandant + "' AND dtebis='" + strPeriodeLY + "'"
            objsqlCom.Connection = objsqlConn
            objsqlConn.Open()
            objdtPeriodeLY.Load(objsqlCom.ExecuteReader)
            objsqlConn.Close()
            If objdtPeriodeLY.Rows.Count > 0 Then
                intTeqNbrLY = objdtPeriodeLY.Rows(0).Item("teqnbr")
            Else
                intTeqNbrLY = 0
            End If
            'Teq-Nr vom Vorvorjahr
            objdtPeriodeLY.Rows.Clear()
            strPeriodePLY = (Val(Strings.Left(strPeriode(4), 4)) - 2).ToString + Strings.Right(strPeriode(4), 4)
            objsqlCom.CommandText = "SELECT teqnbr FROM periode WHERE mandid='" + strMandant + "' AND dtebis='" + strPeriodePLY + "'"
            objsqlCom.Connection = objsqlConn
            objsqlConn.Open()
            objdtPeriodeLY.Load(objsqlCom.ExecuteReader)
            objsqlConn.Close()
            If objdtPeriodeLY.Rows.Count > 0 Then
                intTeqNbrPLY = objdtPeriodeLY.Rows(0).Item("teqnbr")
            Else
                intTeqNbrPLY = 0
            End If

            intTeqNbr = strPeriode(8)
            strYear = Strings.Left(strPeriode(4), 4)
            objdtInfo.Rows.Add("GeschäftsJ", strPeriode(3) + "-" + strPeriode(4) + ", teq: " + strPeriode(8).ToString + ", " + intTeqNbrLY.ToString + ", " + intTeqNbrPLY.ToString)
            objdtDates.Rows.Add(strYear, "GJ Mandant", Date.ParseExact(strPeriode(3), "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture), Date.ParseExact(strPeriode(4), "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture), "O")
            objdtInfo.Rows.Add("Buchungen/ Status", strPeriode(5) + "-" + strPeriode(6) + "/ " + strPeriode(2))
            objdtDates.Rows.Add(strYear, "Buchungen", Date.ParseExact(strPeriode(5), "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture), Date.ParseExact(strPeriode(6), "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture), strPeriode(2))


            FcReturns = FcReadPeriodenDef2(objsqlConn,
                                      objsqlCom,
                                      strPeriode(8),
                                      objdtInfo,
                                      objdtDates,
                                      strYear)

            'Perioden-Definition vom Tool einlesen
            objdbcmd.Connection = objdbconn
            objdbconn.Open()
            objdbcmd.CommandText = "SELECT * FROM t_sage_buchhaltungen_periods WHERE year=" + strYear + " AND refMandant=" + intAccounting.ToString
            dtPeriods.Load(objdbcmd.ExecuteReader)
            objdbconn.Close()
            If dtPeriods.Rows.Count > 0 Then
                datPeriodFrom = dtPeriods.Rows(0).Item("periodFrom")
                datPeriodTo = dtPeriods.Rows(0).Item("periodTo")
                strPeriodStatus = dtPeriods.Rows(0).Item("status")
            Else
                datPeriodFrom = Convert.ToDateTime(strYear + "-01-01 00:00:01")
                datPeriodTo = Convert.ToDateTime(strYear + "-12-31 23:59:59")
                strPeriodStatus = "O"
            End If
            objdtInfo.Rows.Add("Perioden", Format(datPeriodFrom, "dd.MM.yyyy hh:mm:ss") + " - " + Format(datPeriodTo, "dd.MM.yyyy hh:mm:ss") + "/ " + strPeriodStatus)

            'In Dates-Tabelle schreiben
            For Each dtperrow As DataRow In dtPeriods.Rows
                objdtDates.Rows.Add(strYear, "MSS Per " + Convert.ToString(dtperrow(2)), dtperrow(3), dtperrow(4), dtperrow(5))
            Next

            'Finanz Buha öffnen
            'If Not IsNothing(objfiBuha) Then
            '    objfiBuha = Nothing
            'End If
            'objfiBuha = New SBSXASLib.AXiFBhg
            'objfiBuha = objFinanz.GetFibuObj()
            'Debitor öffnen
            'If Not IsNothing(objdbBuha) Then
            '    objdbBuha = Nothing
            'End If
            'objdbBuha = New SBSXASLib.AXiDbBhg
            'objdbBuha = objFinanz.GetDebiObj()
            'If Not IsNothing(objdbPIFb) Then
            '    objdbPIFb = Nothing
            'End If
            'objdbPIFb = New SBSXASLib.AXiPlFin
            'objdbPIFb = objfiBuha.GetCheckObj()
            'If Not IsNothing(objFiBebu) Then
            '    objFiBebu = Nothing
            'End If
            'objFiBebu = New SBSXASLib.AXiBeBu
            'objFiBebu = objFinanz.GetBeBuObj()
            'Kreditor
            'If Not IsNothing(objkrBuha) Then
            '    objkrBuha = Nothing
            'End If
            'objkrBuha = New SBSXASLib.AXiKrBhg
            'objKrBuha = objFinanz.GetKrediObj

            'Application.DoEvents()

        Catch ex As Exception
            MsgBox("OpenMandant:" + vbCrLf + "Error" + vbCrLf + "Error # " + Str(Err.Number) + " was generated by " + Err.Source + vbCrLf + Err.Description + " Fehlernummer" & Str(Err.Number And 65535))
            Err.Clear()


        Finally
            objdtPeriodeLY = Nothing
            dtPeriods = Nothing
            'System.GC.Collect()

        End Try

    End Function

    Friend Function FcLoginSage4(ByVal intAccounting As Int16,
                                 ByRef objdtDates As DataTable,
                                 ByVal strPeriod As String) As Int16

        'wird gebaucht um das Vor- und Folge-Jahr in Sage zu prüfen

        Dim objdbConn As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString"))
        Dim objdbcmd As New MySqlCommand

        Dim objFinanz As New SBSXASLib.AXFinanz
        Dim strMandant As String
        Dim booAccOk As Boolean
        'Dim strPeriodenInfo As String
        Dim strArPeriode() As String
        Dim strArLogonInfo() As String
        Dim strLogonInfo() As String
        Dim strPeriodenInfo As String
        Dim strYear As String
        Dim intPeriodenNr As Int16
        Dim intFctReturns As Int16
        Dim dtPeriods As New DataTable

        Dim objFinanzCopy As New SBSXASLib.AXFinanz
        'Dim objfiBuhaCopy As New SBSXASLib.AXiFBhg

        Try

            'Login
            Try
                Call objFinanzCopy.ConnectSBSdb(System.Configuration.ConfigurationManager.AppSettings("OwnSageServer"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageDB"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageID"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSagePsw"), "")

            Catch inEx As Exception
                If inEx.HResult <> -2147473602 Then
                    MessageBox.Show(inEx.Message, "Connect to Sage - DB " + Err.Number.ToString)
                    Exit Function
                End If


            End Try

            intFctReturns = FcReadFromSettingsIII("Buchh200_Name",
                                                intAccounting,
                                                strMandant)

            booAccOk = objFinanzCopy.CheckMandant(strMandant)

            objFinanzCopy.OpenMandant(strMandant, strPeriod)

            'Von Login aktuelle Periode auslesen
            strLogonInfo = Split(objFinanzCopy.GetLogonInfo(), "{>}")
            'strArLogonInfo = Split(objFinanz.GetLogonInfo(), "{>}")

            'Check Periode
            intPeriodenNr = objFinanzCopy.ReadPeri(strMandant, strLogonInfo(7))
            strPeriodenInfo = objFinanzCopy.GetPeriListe(0)

            strArPeriode = Split(strPeriodenInfo, "{>}")

            strYear = Strings.Left(strArPeriode(4), 4)

            objdtDates.Rows.Add(strYear, "GJ Mandant", Date.ParseExact(strArPeriode(3), "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture), Date.ParseExact(strArPeriode(4), "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture), "O")
            objdtDates.Rows.Add(strYear, "Buchungen", Date.ParseExact(strArPeriode(5), "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture), Date.ParseExact(strArPeriode(6), "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture), strArPeriode(2))

            intFctReturns = FcReadPeriodenDef3(strArPeriode(8),
                                                    objdtDates,
                                                    strYear)

            'Perioden-Def vom Tool holen
            objdbcmd.Connection = objdbConn
            objdbcmd.CommandText = "SELECT * FROM t_sage_buchhaltungen_periods WHERE year=" + strYear + " AND refMandant=" + intAccounting.ToString
            objdbcmd.Connection.Open()
            dtPeriods.Load(objdbcmd.ExecuteReader)
            objdbcmd.Connection.Close()

            'In Dates-Tabelle schreiben
            For Each dtperrow As DataRow In dtPeriods.Rows
                objdtDates.Rows.Add(strYear, "MSS Per " + Convert.ToString(dtperrow(2)), dtperrow(3), dtperrow(4), dtperrow(5))
            Next


        Catch ex As Exception
            MsgBox("OpenMandant:" + vbCrLf + "Error" + vbCrLf + "Error # " + Str(Err.Number) + " was generated by " + Err.Source + vbCrLf + Err.Description + " Fehlernummer" & Str(Err.Number And 65535))
            Err.Clear()

        Finally
            objdbConn = Nothing
            objdbcmd = Nothing
            'objFinanz = Nothing
            strArPeriode = Nothing
            strArLogonInfo = Nothing
            dtPeriods = Nothing

        End Try

    End Function

    Friend Function FcReadPeriodenDef2(ByRef objSQLConnection As SqlClient.SqlConnection,
                                             ByRef objSQLCommand As SqlClient.SqlCommand,
                                             ByVal intPeriodenNr As Int32,
                                             ByRef objdtInfo As DataTable,
                                             ByRef objdtDates As DataTable,
                                             ByVal strYear As String) As Int16

        'Returns 0=definiert, 1=nicht defeniert, 9=Problem
        Dim objlocdtPeriDef As New DataTable
        Dim strPeriodenDef(4) As String


        Try

            objSQLConnection.Open()
            objSQLCommand.CommandText = "SELECT * FROM peridef WHERE teqnbr=" + intPeriodenNr.ToString
            objSQLCommand.Connection = objSQLConnection
            objlocdtPeriDef.Load(objSQLCommand.ExecuteReader)

            'info befüllen
            If objlocdtPeriDef.Rows.Count > 0 Then 'Perioden-Definition vorhanden

                strPeriodenDef(0) = IIf(IsDBNull(objlocdtPeriDef.Rows(0).Item(2)), "n/a", objlocdtPeriDef.Rows(0).Item(2)) 'Bezeichnung
                strPeriodenDef(1) = objlocdtPeriDef.Rows(0).Item(3).ToString  'Von
                strPeriodenDef(2) = objlocdtPeriDef.Rows(0).Item(4).ToString  'Bis
                strPeriodenDef(3) = objlocdtPeriDef.Rows(0).Item(5)  'Status

                objdtInfo.Rows.Add("Perioden S200", strPeriodenDef(0))
                objdtInfo.Rows.Add("Von - Bis/ Status", Format(Convert.ToDateTime(strPeriodenDef(1)), "dd.MM.yyyy hh:mm:ss") + " - " + Format(Convert.ToDateTime(strPeriodenDef(2)), "dd.MM.yyyy hh:mm:ss") + "/ " + strPeriodenDef(3))

                'Return 0
            Else

                objdtInfo.Rows.Add("Perioden S200", "keine")
                objdtInfo.Rows.Add("Von - Bis/ Status", Format(Convert.ToDateTime("01.01." + strYear + " 00:00:00"), "dd.MM.yyyy hh:mm:ss") + " - " + Format(Convert.ToDateTime("31.12." + strYear + " 23:59:59"), "dd.MM.yyyy hh:mm:ss") + "/ " + "O")

                Return 1

            End If

            'date Tabelle befüllen
            If objlocdtPeriDef.Rows.Count > 0 Then

                For Each perirow As DataRow In objlocdtPeriDef.Rows
                    objdtDates.Rows.Add(strYear, "PD " + perirow(2), perirow(3), perirow(4), perirow(5))
                Next

            End If
            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Periodendefinition lesen")
            Return 9

        Finally
            objSQLConnection.Close()
            objlocdtPeriDef.Constraints.Clear()
            objlocdtPeriDef.Clear()
            objlocdtPeriDef = Nothing
            strPeriodenDef = Nothing
            'System.GC.Collect()

        End Try

    End Function

    Friend Function FcGetRefDebiNr(lngDebiNbr As Int32,
                                          intAccounting As Int32,
                                          ByRef intDebiNew As Int32) As Int16

        'Return 0=ok, 1=Neue Debi genereiert und gesetzt, 2=Rep_Ref nicht definiert, 3=Nicht in Tab_Repbetriebe, 4=keine Angaben in Tab_Repbetriebe

        Dim strTableName, strTableType, strDebFieldName, strDebNewField As String
        'Dim intCreatable As Int16
        Dim objdtDebitor As New DataTable
        Dim intPKNewField As Int32
        Dim objdbConnDeb As New MySqlConnection
        Dim objsqlCommDeb As New MySqlCommand

        Dim objdbAccessConn As OleDb.OleDbConnection
        Dim objlocOLEdbcmd As New OleDb.OleDbCommand
        Dim strMDBName As String
        'Dim objOrcommand As New OracleClient.OracleCommand
        Dim strSQL As String
        Dim intFunctionReturns As Int16
        Dim strFcReturns As String

        Try

            intFunctionReturns = FcReadFromSettingsIII("Buchh_PKTableConnection",
                                                           intAccounting,
                                                           strMDBName)

            intFunctionReturns = FcReadFromSettingsIII("Buchh_PKTable",
                                                        intAccounting,
                                                        strTableName)
            intFunctionReturns = FcReadFromSettingsIII("Buchh_PKTableType",
                                                 intAccounting,
                                                 strTableType)
            intFunctionReturns = FcReadFromSettingsIII("Buchh_PKField",
                                                      intAccounting,
                                                      strDebFieldName)
            intFunctionReturns = FcReadFromSettingsIII("Buchh_PKNewField",
                                                     intAccounting,
                                                   strDebNewField)

            strSQL = "SELECT * " +
                 " FROM " + strTableName + " WHERE " + strDebFieldName + "=" + lngDebiNbr.ToString

            If strTableName <> "" And strDebFieldName <> "" Then

                If strTableType = "O" Then 'Oracle
                    Stop
                    'objOrdbconn.Open()
                    'objOrcommand.CommandText = "SELECT " + strDebFieldName + ", " + strDebNewField + ", " + strCompFieldName + ", " + strStreetFieldName + ", " + strZIPFieldName + ", " + strTownFieldName + ", " + strSageName + ", " + strDebiAccField +
                    '                            " FROM " + strTableName + " WHERE " + strDebFieldName + "=" + lngDebiNbr.ToString
                    'objOrcommand.CommandText = strSQL
                    'objdtDebitor.Load(objOrcommand.ExecuteReader)
                    'Ist DebiNrNew Linked oder Direkt
                    'If strDebNewFieldType = "D" Then

                    'objOrdbconn.Close()
                ElseIf strTableType = "M" Then 'MySQL
                    intDebiNew = 0
                    'MySQL - Tabelle einlesen
                    objdbConnDeb.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(FcReadFromSettingsII("Buchh_PKTableConnection", intAccounting))
                    objdbConnDeb.Open()
                    'objsqlCommDeb.CommandText = "SELECT " + strDebFieldName + ", " + strDebNewField + ", " + strCompFieldName + ", " + strStreetFieldName + ", " + strZIPFieldName + ", " + strTownFieldName + ", " + strSageName + ", " + strDebiAccField +
                    '                            " FROM " + strTableName + " WHERE " + strDebFieldName + "=" + lngDebiNbr.ToString
                    objsqlCommDeb.CommandText = strSQL
                    objsqlCommDeb.Connection = objdbConnDeb
                    objdtDebitor.Load(objsqlCommDeb.ExecuteReader)
                    objdbConnDeb.Close()

                ElseIf strTableType = "A" Then 'Access
                    'Access
                    Call FcInitAccessConnecation(objdbAccessConn, strMDBName)
                    objlocOLEdbcmd.CommandText = strSQL
                    objdbAccessConn.Open()
                    objlocOLEdbcmd.Connection = objdbAccessConn
                    objdtDebitor.Load(objlocOLEdbcmd.ExecuteReader)
                    objdbAccessConn.Close()

                End If

                If objdtDebitor.Rows.Count > 0 Then
                    'If IsDBNull(objdtDebitor.Rows(0).Item(strDebNewField)) And strTableName <> "Tab_Repbetriebe" Then 'Es steht nichts im Feld welches auf den Rep_Betrieb verweist oder wenn direkt
                    ' intDebiNew = 0
                    'Return 2
                    'Else

                    If strTableName <> "Tab_Repbetriebe" Then
                        'intPKNewField = objdtDebitor.Rows(0).Item(strDebNewField)
                        If strTableName = "t_customer" Then
                            intPKNewField = FcGetPKNewFromRep(IIf(IsDBNull(objdtDebitor.Rows(0).Item("ID")), 0, objdtDebitor.Rows(0).Item("ID")),
                                                           "P")
                        Else
                            'D.h. Neue PK-Nr. wird nie von anderer Tabelle gelesen als t_customer oder Repbetriebe, bei einem <> t_customer muss de Rep_Betiebnr mitgegeben werden
                            intPKNewField = FcGetPKNewFromRep(IIf(IsDBNull(objdtDebitor.Rows(0).Item(strDebNewField)), 0, objdtDebitor.Rows(0).Item(strDebNewField)),
                                                           "R")

                            'Stop
                        End If

                        If intPKNewField = 0 Then
                            'PK wurde nicht vergeben => Eine neue erzeugen und in der Tabelle Rep_Betriebe 
                            If strTableName = "t_customer" Then
                                intFunctionReturns = FcNextPrivatePKNr(objdtDebitor.Rows(0).Item("ID"),
                                                                            intDebiNew)
                                If intFunctionReturns = 0 And intDebiNew > 0 Then 'Vergabe hat geklappt
                                    intFunctionReturns = FcWriteNewPrivateDebToRepbetrieb(objdtDebitor.Rows(0).Item("ID"),
                                                                                               intDebiNew)
                                    If intFunctionReturns = 0 Then 'Schreiben hat geklappt
                                        Return 1
                                    End If
                                End If
                            Else
                                intFunctionReturns = FcNextPKNr(objdtDebitor.Rows(0).Item(strDebNewField),
                                                                     intDebiNew,
                                                                     intAccounting,
                                                                     "D")
                                If intFunctionReturns = 0 And intDebiNew > 0 Then 'Vergabe hat geklappt
                                    intFunctionReturns = FcWriteNewDebToRepbetrieb(objdtDebitor.Rows(0).Item(strDebNewField),
                                                                                        intDebiNew,
                                                                                        intAccounting,
                                                                                        "D")
                                    If intFunctionReturns = 0 Then 'Schreiben hat geklappt
                                        Return 1
                                    End If
                                End If
                                Stop
                            End If
                            'intDebiNew = 0
                            'Return 3
                        Else
                            intDebiNew = intPKNewField
                            Return 0
                        End If
                    Else 'Wenn Angaben nicht von anderer Tabelle kommen
                        'Prüfen ob Repbetrieb schon eine neue Nummer erhalten hat.
                        If Not IsDBNull(objdtDebitor.Rows(0).Item(strDebNewField)) Then
                            intDebiNew = objdtDebitor.Rows(0).Item(strDebNewField)
                        Else
                            intFunctionReturns = FcNextPKNr(lngDebiNbr,
                                                                 intDebiNew,
                                                                 intAccounting,
                                                                 "D")
                            If intFunctionReturns = 0 And intDebiNew > 0 Then 'Vergabe hat geklappt
                                intFunctionReturns = FcWriteNewDebToRepbetrieb(lngDebiNbr,
                                                                                    intDebiNew,
                                                                                    intAccounting,
                                                                                    "D")
                                If intFunctionReturns = 0 Then 'Schreiben hat geklappt
                                    Return 1
                                End If
                            End If
                        End If
                        Return 0
                    End If
                Else
                    intDebiNew = 0
                    Return 4
                End If
            Else
                'intDebiNew = 0
                'Return 4
            End If

            'End If

            Return intPKNewField

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Debitor - Suche", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 9

        Finally
            objdtDebitor = Nothing
            objdbConnDeb = Nothing
            objsqlCommDeb = Nothing
            objdbAccessConn = Nothing
            objlocOLEdbcmd = Nothing
            'System.GC.Collect()

        End Try


    End Function

    Friend Function FcGetPKNewFromRep(ByVal intPKRefField As Int32,
                                             ByVal strMode As String) As Int32

        'Aus Tabelle Rep_Betriebe auf ZHDB02 auslesen 
        Dim objdtRepBetrieb As New DataTable
        Dim objdbconnZHDB02 As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))
        Dim objsqlcommandZHDB02 As New MySqlCommand

        Try

            objsqlcommandZHDB02.Connection = objdbconnZHDB02
            objdbconnZHDB02.Open()
            objsqlcommandZHDB02.Connection = objdbconnZHDB02
            If strMode = "P" Then
                objsqlcommandZHDB02.CommandText = "SELECT PKNr From t_customer WHERE ID=" + intPKRefField.ToString
            Else
                objsqlcommandZHDB02.CommandText = "SELECT PKNr From tab_repbetriebe WHERE Rep_Nr=" + intPKRefField.ToString
            End If
            objdtRepBetrieb.Load(objsqlcommandZHDB02.ExecuteReader)
            If (objdtRepBetrieb.Rows.Count > 0) Then
                If Not IsDBNull(objdtRepBetrieb.Rows(0).Item("PKNr")) Then
                    Return objdtRepBetrieb.Rows(0).Item("PKNr")
                Else
                    Return 0
                End If
            Else
                Return 0
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem Neue PK-Nr.")
            Return 0

        Finally
            objdbconnZHDB02.Close()
            objdtRepBetrieb = Nothing
            objsqlcommandZHDB02 = Nothing
            objdbconnZHDB02 = Nothing

        End Try


    End Function

    Friend Function FcNextPrivatePKNr(ByVal intPersNr As Int32,
                                             ByRef intNewPKNr As Int32) As Int16

        '0=ok, 1=Rep - Nr. existiert nicht, 2=Bereich voll, 3=keine Bereichdefinition 9=Problem

        'PK - Nummer soll der Funktion gegeben werden, Funktion sucht sich dann die PK_Gruppe 
        'Konzept: Tabelle füllen und dann durchsteppen
        Dim objdbconnZHDB02 As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))
        Dim objsqlcommand As New MySqlCommand
        Dim objdtPKNr As New DataTable
        Dim intPKNrGuppenID As Int16
        Dim intRangeStart, intRangeEnd, i, intRecordCounter As Int32
        Dim objdsPKNbrs As New DataSet
        Dim objDAPKNbrs As New MySqlDataAdapter
        Dim objDAPersons As New MySqlDataAdapter
        Dim objdsPersons As New DataSet

        Try

            objdbconnZHDB02.Open()
            objsqlcommand.Connection = objdbconnZHDB02
            objsqlcommand.CommandText = "SELECT PKNrGruppeID FROM t_customer WHERE ID=" + intPersNr.ToString
            objDAPersons.SelectCommand = objsqlcommand
            objdsPersons.EnforceConstraints = False
            objDAPersons.Fill(objdsPersons)

            If objdsPersons.Tables(0).Rows.Count > 0 Then 'Person gefunden
                intPKNrGuppenID = objdsPersons.Tables(0).Rows(0).Item("PKNrGruppeID")
                'Start und End des Bereichs setzen
                objdtPKNr.Clear()
                objsqlcommand.CommandText = "SELECT RangeStart, RangeEnd " +
                                            "FROM tab_repbetriebe_pknrgruppe " +
                                            "WHERE ID=" + intPKNrGuppenID.ToString
                objdtPKNr.Load(objsqlcommand.ExecuteReader)
                If objdtPKNr.Rows.Count > 0 Then 'Bereichsdefinition gefunden
                    intRangeStart = objdtPKNr.Rows(0).Item("RangeStart")
                    intRangeEnd = objdtPKNr.Rows(0).Item("RangeEnd")
                    'PK - Bereich laden und durchsteppen und Lücke oder nächste PK-Nr suchen
                    'Muss über Dataset gehen da Datatable ein Fehler bringt
                    'objdtPKNr.Clear()

                    objsqlcommand.CommandText = "SELECT PKNr " +
                                                "FROM t_customer " +
                                                "WHERE PKNr BETWEEN " + intRangeStart.ToString + " AND " + intRangeEnd.ToString + " " +
                                                "ORDER BY PKNr"
                    'objdtPKNr.Load(objsqlcommand.ExecuteReader)
                    objDAPKNbrs.SelectCommand = objsqlcommand
                    objdsPKNbrs.EnforceConstraints = False
                    objDAPKNbrs.Fill(objdsPKNbrs)

                    intNewPKNr = 0
                    i = intRangeStart
                    If objdsPKNbrs.Tables(0).Rows.Count = 0 Then
                        intNewPKNr = i
                    Else
                        intRecordCounter = 0
                        Do Until intRecordCounter = objdsPKNbrs.Tables(0).Rows.Count
                            If Not objdsPKNbrs.Tables(0).Rows(intRecordCounter).Item("PKNr") = i Then
                                intNewPKNr = i
                                Return 0
                            End If
                            i += 1
                            intRecordCounter += 1
                        Loop
                        If i <= intRangeEnd Then
                            intNewPKNr = i
                        End If
                    End If
                    If intNewPKNr = 0 Then
                        Return 2
                    End If
                Else
                    Return 3
                End If
            Else
                Return 1
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return 9

        Finally

            objdbconnZHDB02.Close()
            objdbconnZHDB02 = Nothing
            objDAPKNbrs = Nothing
            objdsPKNbrs = Nothing
            objsqlcommand = Nothing
            objdtPKNr = Nothing
            objdsPersons = Nothing
            objDAPersons = Nothing
            objDAPKNbrs = Nothing

        End Try

    End Function

    Friend Function FcWriteNewPrivateDebToRepbetrieb(ByVal intPersNr As Int32,
                                                            intNewDebNr As Int32) As Int16

        '0=Update ok, 1=Update hat nicht geklappt, 9=Error

        Dim strSQL As String
        Dim objmysqlcmd As New MySqlCommand
        Dim intAffected As Int16
        Dim objdbconnZHDB02 As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))

        Try

            strSQL = "UPDATE t_customer SET PKNr=" + intNewDebNr.ToString + " WHERE ID=" + intPersNr.ToString
            objdbconnZHDB02.Open()
            objmysqlcmd.Connection = objdbconnZHDB02
            objmysqlcmd.CommandText = strSQL
            intAffected = objmysqlcmd.ExecuteNonQuery()
            If intAffected <> 1 Then
                Return 1
            Else
                Return 0
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return 9

        Finally

            objdbconnZHDB02.Close()
            objdbconnZHDB02 = Nothing
            objmysqlcmd = Nothing

        End Try

    End Function

    Friend Function FcNextPKNr(ByVal intRepNr As Int32,
                                      ByRef intNewPKNr As Int32,
                                      ByVal intAccounting As Int16,
                                      ByVal strMode As String) As Int16

        '0=ok, 1=Rep - Nr. existiert nicht, 2=Bereich voll, 3=keine Bereichdefinition 9=Problem

        'PK - Nummer soll der Funktion gegeben werden, Funktion sucht sich dann die PK_Gruppe 
        'Konzept: Tabelle füllen und dann durchsteppen
        'Dim objdbconnZHDB02 As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))
        Dim objsqlcommand As New MySqlCommand
        Dim objdtPKNr As New DataTable
        Dim intPKNrGuppenID As Int16
        Dim intRangeStart, intRangeEnd, i, intRecordCounter As Int32
        Dim objdsPKNbrs As New DataSet
        Dim objDAPKNbrs As New MySqlDataAdapter
        Dim objdbconn As New MySqlConnection
        Dim intFcReturns As Int16
        Dim strFcReturns As String

        Try

            'Wo ist die RepBetriebe?
            'objdbconnZHDB02.Open()
            If strMode = "D" Then
                'objdbconn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(Main.FcReadFromSettings(objdbconnZHDB02, "Buchh_PKTableConnection", intAccounting))
                intFcReturns = FcReadFromSettingsIII("Buch_TabRepConnection", intAccounting, strFcReturns)
                objdbconn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(strFcReturns)
            Else
                intFcReturns = FcReadFromSettingsIII("Buchh_PKKrediTableConnection", intAccounting, strFcReturns)
                objdbconn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(strFcReturns)
            End If

            objdbconn.Open()

            objsqlcommand.Connection = objdbconn
            objsqlcommand.CommandText = "SELECT PKNrGruppeID FROM tab_repbetriebe WHERE Rep_Nr=" + intRepNr.ToString
            objdtPKNr.Load(objsqlcommand.ExecuteReader)

            If objdtPKNr.Rows.Count > 0 Then 'Rep_Betrieb gefunden
                intPKNrGuppenID = IIf(IsDBNull(objdtPKNr.Rows(0).Item("PKNrGruppeID")), 2, objdtPKNr.Rows(0).Item("PKNrGruppeID"))
                'Start und End des Bereichs setzen
                objdtPKNr.Clear()
                objsqlcommand.CommandText = "SELECT RangeStart, RangeEnd " +
                                            "FROM tab_repbetriebe_pknrgruppe " +
                                            "WHERE ID=" + intPKNrGuppenID.ToString + " AND ID<5"
                objdtPKNr.Load(objsqlcommand.ExecuteReader)
                If objdtPKNr.Rows.Count > 0 Then 'Bereichsdefinition gefunden
                    intRangeStart = objdtPKNr.Rows(0).Item("RangeStart")
                    intRangeEnd = objdtPKNr.Rows(0).Item("RangeEnd")
                    'PK - Bereich laden und durchsteppen und Lücke oder nächste PK-Nr suchen
                    'Muss über Dataset gehen da Datatable ein Fehler bringt
                    'objdtPKNr.Clear()

                    objsqlcommand.CommandText = "SELECT PKNr " +
                                                "FROM tab_repbetriebe " +
                                                "WHERE PKNr BETWEEN " + intRangeStart.ToString + " AND " + intRangeEnd.ToString + " " +
                                                "ORDER BY PKNr"
                    'objdtPKNr.Load(objsqlcommand.ExecuteReader)
                    objDAPKNbrs.SelectCommand = objsqlcommand
                    objdsPKNbrs.EnforceConstraints = False
                    objDAPKNbrs.Fill(objdsPKNbrs)

                    intNewPKNr = 0
                    i = intRangeStart
                    If objdsPKNbrs.Tables(0).Rows.Count = 0 Then
                        intNewPKNr = i
                    Else
                        intRecordCounter = 0
                        Do Until intRecordCounter = objdsPKNbrs.Tables(0).Rows.Count
                            If Not objdsPKNbrs.Tables(0).Rows(intRecordCounter).Item("PKNr") = i Then
                                intNewPKNr = i
                                Return 0
                            End If
                            i += 1
                            intRecordCounter += 1
                        Loop
                        If i <= intRangeEnd Then
                            intNewPKNr = i
                        End If
                    End If
                    If intNewPKNr = 0 Then
                        Return 2
                    End If
                Else
                    Return 3
                End If
            Else
                Return 1
            End If

        Catch ex As InvalidCastException
            MessageBox.Show("Rep_Nr " + intRepNr.ToString + " ist keiner Gruppe zugewiesen. Erstellung nicht möglich.", "Gruppe fehlt", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 9

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem Debitoren-Nummer-Vergabe Rep_Nr " + intRepNr.ToString, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 9

        Finally
            'objdbconnZHDB02.Close()
            'objdbconnZHDB02 = Nothing
            objdbconn.Close()
            objdbconn = Nothing
            objDAPKNbrs = Nothing
            objdsPKNbrs = Nothing
            objsqlcommand = Nothing
            objdtPKNr = Nothing

        End Try


    End Function

    Friend Function FcWriteNewDebToRepbetrieb(ByVal intRepNr As Int32,
                                                     ByVal intNewDebNr As Int32,
                                                     ByVal intAccounting As Int16,
                                                     ByVal strMode As String) As Int16

        '0=Update ok, 1=Update hat nicht geklappt, 9=Error

        Dim strSQL As String
        'Dim objdbconnZHDB02 As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))
        Dim objmysqlcmd As New MySqlCommand
        Dim objdbconn As New MySqlConnection
        Dim intAffected As Int16
        Dim intFcReturns As Int16
        Dim strFcReturns As String

        Try

            'Wo ist die Rep_Betriebe?
            'objdbconnZHDB02.Open()
            If strMode = "D" Then
                'objdbconn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(Main.FcReadFromSettings(objdbconnZHDB02, "Buchh_PKTableConnection", intAccounting))
                intFcReturns = FcReadFromSettingsIII("Buch_TabRepConnection", intAccounting, strFcReturns)
                objdbconn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(strFcReturns)
            Else
                intFcReturns = FcReadFromSettingsIII("Buchh_PKKrediTableConnection", intAccounting, strFcReturns)
                objdbconn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(strFcReturns)
            End If
            objdbconn.Open()

            strSQL = "UPDATE tab_repbetriebe SET PKNr=" + intNewDebNr.ToString + " WHERE Rep_Nr=" + intRepNr.ToString
            objmysqlcmd.Connection = objdbconn
            objmysqlcmd.CommandText = strSQL
            intAffected = objmysqlcmd.ExecuteNonQuery()
            If intAffected <> 1 Then
                Return 1
            Else
                Return 0
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return 9

        Finally
            'objdbconnZHDB02.Close()
            'objdbconnZHDB02 = Nothing
            objdbconn.Close()
            objdbconn = Nothing
            objmysqlcmd = Nothing

        End Try

    End Function

    Friend Function FcReadPeriodenDef3(ByVal intPeriodenNr As Int32,
                                       ByRef objdtDates As DataTable,
                                       ByVal strYear As String) As Int16

        'Wird gebracuht um Pierodendefintionen vom Mandanten einzulesen und in die Dates-Tabelle zu schreiben
        '0=ok, 9=Problem

        Dim objSQLConnection As New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("SQLConnectionString"))
        Dim objSQLCommand As New SqlClient.SqlCommand
        Dim objlocdtPeriDef As New DataTable

        Try

            objSQLCommand.CommandText = "SELECT * FROM peridef WHERE teqnbr=" + intPeriodenNr.ToString
            objSQLCommand.Connection = objSQLConnection
            objSQLCommand.Connection.Open()
            objlocdtPeriDef.Load(objSQLCommand.ExecuteReader)
            objSQLCommand.Connection.Close()

            'date Tabelle befüllen
            If objlocdtPeriDef.Rows.Count > 0 Then

                For Each perirow As DataRow In objlocdtPeriDef.Rows
                    objdtDates.Rows.Add(strYear, "PD " + perirow(2), perirow(3), perirow(4), perirow(5))
                Next

            End If
            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Periodendefinition lesen")
            Return 9

        Finally
            objSQLConnection = Nothing
            objSQLCommand = Nothing
            objlocdtPeriDef = Nothing

        End Try


    End Function

    Friend Function FcGetKundenzeichen2(ByVal lngJournalNr As Int32) As String
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

    Friend Function FcGetKurs(strCurrency As String,
                              strDateValuta As String,
                              Optional intKonto As Integer = 0) As Double

        'Konzept: Falls ein Konto mitgegeben wird, wird überprüft ob auf dem Konto die mitgegebene Währung Leitwärhung ist. Falls ja wird der Kurs 1 zurück gegeben

        Dim strKursZeile As String = String.Empty
        Dim strKursZeileAr() As String
        Dim strKontoInfo() As String

        Try

            objfiBuha.ReadKurse(strCurrency, "", "J")

            Do While strKursZeile <> "EOF"
                strKursZeile = objfiBuha.GetKursZeile()
                If strKursZeile <> "EOF" Then
                    strKursZeileAr = Split(strKursZeile, "{>}")
                    If strKursZeileAr(0) = strCurrency Then
                        'If strKursZeileAr(0) = "EUR" Then Stop
                        'Prüfen ob Currency Leitwährung auf Konto. Falls ja Return 1
                        If intKonto <> 0 Then
                            strKontoInfo = Split(objfiBuha.GetKontoInfo(intKonto.ToString), "{>}")
                            If strKontoInfo(7) = strCurrency Then
                                Return 1
                            Else
                                Return strKursZeileAr(4)
                                Return 0
                            End If
                        Else
                            Return strKursZeileAr(4)
                        End If
                    End If
                Else
                    Return 1 'Kurs nicht gefunden
                End If
            Loop

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Currendy-Check " + Err.Number.ToString, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Err.Clear()

        Finally
            strKursZeileAr = Nothing
            strKontoInfo = Nothing

        End Try


    End Function

    Friend Function FcGetSteuerFeld(ByRef strSteuerFeld As String,
                                    lngKto As Long,
                                    strDebiSubText As String,
                                    dblBrutto As Double,
                                    strMwStKey As String,
                                    dblMwSt As Double) As Int16

        'Dim strSteuerFeld As String = String.Empty

        Try

            If dblMwSt <> 0 Then

                strSteuerFeld = objfiBuha.GetSteuerfeld(lngKto.ToString,
                                                      strDebiSubText,
                                                      dblBrutto.ToString,
                                                      strMwStKey,
                                                      dblMwSt.ToString)

            Else

                strSteuerFeld = objfiBuha.GetSteuerfeld(lngKto.ToString,
                                                      strDebiSubText,
                                                      dblBrutto.ToString,
                                                      strMwStKey)

            End If
            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return 1

        End Try


    End Function

    Friend Function FcGetSteuerFeld2(ByRef objFBhg As SBSXASLib.AXiFBhg,
                                            ByRef strSteuerFeld As String,
                                           lngKto As Long,
                                           strDebiSubText As String,
                                           dblBrutto As Double,
                                           strMwStKey As String,
                                           dblMwSt As Double,
                                           datValuta As Date) As Int16

        'Setzt Steuer-Feld mit Valuzta-Datum

        Try

            If dblMwSt <> 0 Then

                strSteuerFeld = objFBhg.GetSteuerfeld(lngKto.ToString,
                                                      strDebiSubText,
                                                      dblBrutto.ToString,
                                                      strMwStKey,
                                                      dblMwSt.ToString,
                                                      Format(datValuta, "yyyyMMdd"))

            Else

                strSteuerFeld = objFBhg.GetSteuerfeld(lngKto.ToString,
                                                      strDebiSubText,
                                                      dblBrutto.ToString,
                                                      strMwStKey)

            End If
            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return 1

        End Try


    End Function

    Friend Function FcCleanRGNrStrict(ByVal strRGNrToClean As String) As String

        Dim intCounter As Int16
        Dim strCleanRGNr As String = String.Empty

        Try

            For intCounter = 1 To Len(strRGNrToClean)
                If Mid(strRGNrToClean, intCounter, 1) = "0" Or Val(Mid(strRGNrToClean, intCounter, 1)) > 0 Then
                    strCleanRGNr += Mid(strRGNrToClean, intCounter, 1)
                End If

            Next

            Return strCleanRGNr

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem CleanString")

        End Try


    End Function

    Friend Function FcCheckDebiExistance(ByRef intBelegNbr As Int32,
                                                 ByVal strTyp As String,
                                                 ByVal intTeqNr As Int32) As Int16

        '0=ok, 1=Beleg existierte schon, 9=Problem

        'Prinzip: in Tabelle kredibuchung suchen da API - Funktion nur in spezifischen Kreditor sucht

        Dim intReturnvalue As Int32
        Dim intStatus As Int16
        Dim tblDebiBeleg As New DataTable
        Dim objdbMSSQLConn As New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("SQLConnectionString"))
        Dim objdbMSSQLCmd As New SqlCommand


        Try

            'Prüfung
            intReturnvalue = 10
            intStatus = 0

            objdbMSSQLCmd.Connection = objdbMSSQLConn
            objdbMSSQLConn.Open()

            Do Until intReturnvalue = 0

                'objdbMSSQLCmd.CommandText = "SELECT lfnbrk FROM kredibuchung WHERE teqnbr IN(" + intTeqNr.ToString + ", " + intTeqNrLY.ToString + ", " + intTeqNrPLY.ToString + ")" +
                '                                                        " AND typ='" + strTyp + "'" +
                '                                                        " AND belnbrint=" + intBelegNbr.ToString
                'Probehalber nur im aktuellen Jahr prüfen
                objdbMSSQLCmd.CommandText = "SELECT lfnbrd FROM debibuchung WHERE teqnbr IN(" + intTeqNr.ToString + ")" +
                                                                        " AND typ='" + strTyp + "'" +
                                                                        " AND belnbr=" + intBelegNbr.ToString

                tblDebiBeleg.Rows.Clear()
                tblDebiBeleg.Load(objdbMSSQLCmd.ExecuteReader)
                If tblDebiBeleg.Rows.Count > 0 Then
                    intReturnvalue = tblDebiBeleg.Rows(0).Item("lfnbrk")
                    intBelegNbr += 1
                Else
                    intReturnvalue = 0
                End If
            Loop

            Return intStatus


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Debitor - BelegExistenzprüfung Problem " + intBelegNbr.ToString)
            Err.Clear()
            Return 9

        Finally
            objdbMSSQLConn.Close()
            objdbMSSQLCmd = Nothing
            objdbMSSQLConn = Nothing
            tblDebiBeleg = Nothing

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
            objlocdtSetting.Constraints.Clear()
            objlocdtSetting.Rows.Clear()
            objlocdtSetting.Columns.Clear()
            objlocdtSetting = Nothing
            objlocMySQLcmd = Nothing
            objdbconn = Nothing
            'System.GC.Collect()

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
                        Case "rsDebiTemp.Fields([strRGBemerkung])"
                            strField = RowBooking(0).Item("strRGBemerkung")
                        'Case "rsDebiTemp.Fields(""strDebRGNbr"")"
                        '    strField = rsDebiTemp.Fields("strDebRGNbr")
                        'Case "rsDebiTemp.Fields([lngDebIdentNbr])"
                        '    strField = rsDebiTemp.Fields("lngDebIdentNbr")
                        Case "rsDebiTemp.Fields([strDebText])"
                            strField = RowBooking(0).Item("strDebText")
                        Case "KUNDENZEICHEN"
                            strField = FcGetKundenzeichen2(RowBooking(0).Item("lngDebIdentNbr"))
                        Case Else
                            strField = "unknown field"
                    End Select
                    strSQLToParse = Strings.Left(strSQLToParse, intPipePositionBegin - 1) + strField + Strings.Right(strSQLToParse, Len(strSQLToParse) - intPipePositionEnd)
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

    Friend Function FcReadDebitorKName(ByVal lngDebKtoNbr As Long) As String

        Dim strDebitorKName As String
        Dim strDebitorKAr() As String


        Try

            strDebitorKName = objfiBuha.GetKontoInfo(lngDebKtoNbr)

            strDebitorKAr = Split(strDebitorKName, "{>}")

            If strDebitorKName <> "EOF" Then
                Return strDebitorKAr(8)
            Else
                Return "EOF"
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Get Kundenzeichen " + Err.Number.ToString)

        Finally
            'Application.DoEvents()
            strDebitorKAr = Nothing

        End Try

    End Function

    Friend Function FcCheckDZKond(ByVal strMandant As String,
                                         ByVal intDZKond As Int16) As Int16

        'Return 0=definiert, 1=nicht definiert, 9=Problem

        Dim objSQLConnection As New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("SQLConnectionString"))
        Dim objSQLCommand As New SqlClient.SqlCommand
        Dim objdtDZKond As New DataTable

        Try

            objSQLConnection.Open()
            objSQLCommand.CommandText = "SELECT kondition.mandid, " +
                                               "kondition.kondnbr, " +
                                               "bezeichnung.langtext, " +
                                               "fi_kond_grp.status, " +
                                               "fi_kond_grp.valutatage, " +
                                               "fi_kond_grp.isdebi, " +
                                               "fi_kond_grp.iskredi, " +
                                               "kondition.verftage, " +
                                               "kondition.satz, " +
                                               "kondition.tolnbr, " +
                                               "kondition.akzttage " +
                                        "FROM   kondition INNER JOIN " +
                                               "fi_kond_grp ON kondition.mandid = fi_kond_grp.mandid AND kondition.kondnbr = fi_kond_grp.kondnbr INNER JOIN " +
                                               "bezeichnung ON kondition.mandid = bezeichnung.mandid AND kondition.beschrnr = bezeichnung.beschreibungnr " +
                                        "WHERE kondition.mandid='" + strMandant + "' AND " +
                                               "fi_kond_grp.isdebi='J' AND " +
                                               "status=1 AND " +
                                               "kondition.kondnbr=" + intDZKond.ToString

            objSQLCommand.Connection = objSQLConnection
            objdtDZKond.Load(objSQLCommand.ExecuteReader)

            If objdtDZKond.Rows.Count >= 1 Then 'Debitoren - Zahlungskondition gefunden
                Return 0
            Else
                Return 1
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Debitor - ZKondition lesen")
            Return 9

        Finally
            objSQLConnection.Close()
            objSQLConnection = Nothing
            objSQLCommand = Nothing
            objdtDZKond = Nothing

        End Try


    End Function

    Friend Function FcGetDZkondFromCust(ByVal lngDebiNbr As Long,
                                           ByRef intDZkond As Int16,
                                           ByVal intAccounting As Int16) As Int16

        'Returns 0=ok, 1=Repbetrieb nicht gefunden, 9=Problem; intDZKond wird abgefüllt

        Dim intDZKondDefault As Int16

        Dim objdbconnZHDB02 As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))
        Dim objsqlcommandZHDB02 As New MySqlCommand
        Dim objdsDebitor As New DataSet
        Dim objDADebitor As New MySqlDataAdapter
        Dim objdtDZKond As New DataTable("tbllocDZKond")

        Try

            objdbconnZHDB02.Open()
            objsqlcommandZHDB02.Connection = objdbconnZHDB02

            'Standard suchen auf Mandant
            objsqlcommandZHDB02.CommandText = "SELECT * " +
                                              "FROM t_payterms_client " +
                                              "INNER JOIN t_sage_zahlungskondition ON t_payterms_client.ZlgkID=t_sage_zahlungskondition.ID " +
                                              "WHERE t_payterms_client.MandantID = " + intAccounting.ToString + " " +
                                              "AND t_payterms_client.K_NR IS NULL " +
                                              "AND t_payterms_client.RepID IS NULL " +
                                              "AND t_payterms_client.CustomerID IS NULL " +
                                              "AND t_payterms_client.IsStandard = true " +
                                              "AND t_sage_zahlungskondition.IsKredi = false"

            objdtDZKond.Load(objsqlcommandZHDB02.ExecuteReader)
            If objdtDZKond.Rows.Count > 0 Then
                intDZKondDefault = objdtDZKond.Rows(0).Item("SageID")
            Else
                'Default MSS lesen
                'Es wird davon ausgegangen, dass der MSS - Standard auf jeden Fall existiert
                objsqlcommandZHDB02.CommandText = "SELECT * " +
                                              "FROM t_payterms_client " +
                                              "INNER JOIN t_sage_zahlungskondition ON t_payterms_client.ZlgkID=t_sage_zahlungskondition.ID " +
                                              "WHERE t_payterms_client.MandantID IS NULL " +
                                              "AND t_payterms_client.K_NR IS NULL " +
                                              "AND t_payterms_client.RepID IS NULL " +
                                              "AND t_payterms_client.CustomerID IS NULL " +
                                              "AND t_payterms_client.IsStandard = true " +
                                              "AND t_sage_zahlungskondition.IsKredi = false"
                objdtDZKond.Load(objsqlcommandZHDB02.ExecuteReader)
                intDZKondDefault = objdtDZKond.Rows(0).Item("SageID")

            End If

            'Zahlungsbedingung suchen
            'Es muss der Weg über ein Dataset genommen werden da sosnt constraint-Meldungen kommen
            objsqlcommandZHDB02.CommandText = "Select t_customer.PKNr, t_sage_zahlungskondition.SageID " +
                                                  "FROM t_customer INNER JOIN t_sage_zahlungskondition On t_customer.DebiZKonditionID = t_sage_zahlungskondition.ID " +
                                                  "WHERE t_customer.PKNr=" + lngDebiNbr.ToString
            objDADebitor.SelectCommand = objsqlcommandZHDB02
            objdsDebitor.EnforceConstraints = False
            objDADebitor.Fill(objdsDebitor)

            If objdsDebitor.Tables(0).Rows.Count > 0 Then

                'Rep-Betrieb existiert
                If Not IsDBNull(objdsDebitor.Tables(0).Rows(0).Item("SageID")) Then
                    intDZkond = objdsDebitor.Tables(0).Rows(0).Item("SageID")
                Else
                    intDZkond = intDZKondDefault
                End If
                Return 0

            Else

                'Kunde existiert nicht
                intDZkond = intDZKondDefault
                Return 1

            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Debitor - Z-Bedingung - von Cust lesen", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            intDZkond = intDZKondDefault
            Return 9

        Finally
            objdbconnZHDB02.Close()
            objdbconnZHDB02 = Nothing
            objsqlcommandZHDB02 = Nothing
            objdsDebitor = Nothing
            objDADebitor = Nothing
            objdtDZKond = Nothing

        End Try


    End Function

    Friend Function FcGetDZkondFromRep(ByVal lngDebiNbr As Long,
                                           ByRef intDZkond As Int16,
                                           ByVal intAccounting As Int16) As Int16

        'Returns 0=ok, 1=Repbetrieb nicht gefunden, 9=Problem; intDZKond wird abgefüllt

        Dim intDZKondDefault As Int16
        Dim objdbconnZHDB02 As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))
        Dim objsqlcommandZHDB02 As New MySqlCommand
        Dim objdsDebitor As New DataSet
        Dim objDADebitor As New MySqlDataAdapter
        Dim objdtDZKond As New DataTable("tbllocDZKond")

        Try

            objdbconnZHDB02.Open()
            objsqlcommandZHDB02.Connection = objdbconnZHDB02

            'Standard suchen auf Mandant
            objsqlcommandZHDB02.CommandText = "SELECT * " +
                                              "FROM t_payterms_client " +
                                              "INNER JOIN t_sage_zahlungskondition ON t_payterms_client.ZlgkID=t_sage_zahlungskondition.ID " +
                                              "WHERE t_payterms_client.MandantID = " + intAccounting.ToString + " " +
                                              "AND t_payterms_client.K_NR IS NULL " +
                                              "AND t_payterms_client.RepID IS NULL " +
                                              "AND t_payterms_client.CustomerID IS NULL " +
                                              "AND t_payterms_client.IsStandard = true " +
                                              "AND t_sage_zahlungskondition.IsKredi = false"

            objdtDZKond.Load(objsqlcommandZHDB02.ExecuteReader)
            If objdtDZKond.Rows.Count > 0 Then
                intDZKondDefault = objdtDZKond.Rows(0).Item("SageID")
            Else
                'Default MSS lesen
                'Es wird davon ausgegangen, dass der MSS - Standard auf jeden Fall existiert
                objsqlcommandZHDB02.CommandText = "SELECT * " +
                                              "FROM t_payterms_client " +
                                              "INNER JOIN t_sage_zahlungskondition ON t_payterms_client.ZlgkID=t_sage_zahlungskondition.ID " +
                                              "WHERE t_payterms_client.MandantID IS NULL " +
                                              "AND t_payterms_client.K_NR IS NULL " +
                                              "AND t_payterms_client.RepID IS NULL " +
                                              "AND t_payterms_client.CustomerID IS NULL " +
                                              "AND t_payterms_client.IsStandard = true " +
                                              "AND t_sage_zahlungskondition.IsKredi = false"
                objdtDZKond.Load(objsqlcommandZHDB02.ExecuteReader)
                intDZKondDefault = objdtDZKond.Rows(0).Item("SageID")

            End If

            'Zahlungsbedingung suchen
            'Es muss der Weg über ein Dataset genommen werden da sosnt constraint-Meldungen kommen
            objsqlcommandZHDB02.CommandText = "Select Tab_Repbetriebe.PKNr, t_sage_zahlungskondition.SageID " +
                                                  "FROM Tab_Repbetriebe INNER JOIN t_sage_zahlungskondition On Tab_Repbetriebe.Rep_DebiZKonditionID = t_sage_zahlungskondition.ID " +
                                                  "WHERE Tab_Repbetriebe.PKNr=" + lngDebiNbr.ToString
            objDADebitor.SelectCommand = objsqlcommandZHDB02
            objdsDebitor.EnforceConstraints = False
            objDADebitor.Fill(objdsDebitor)

            If objdsDebitor.Tables(0).Rows.Count > 0 Then

                'Rep-Betrieb existiert
                If Not IsDBNull(objdsDebitor.Tables(0).Rows(0).Item("SageID")) Then
                    intDZkond = objdsDebitor.Tables(0).Rows(0).Item("SageID")
                Else
                    'Es ist keine Definition vorgenommen worden
                    intDZkond = intDZKondDefault
                End If
                Return 0

            Else

                'Rep-Betrieb existiert nicht
                intDZkond = intDZKondDefault
                Return 1

            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Debitor - Z-Bedingung - von Rep lesen", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            intDZkond = intDZKondDefault
            Return 9

        Finally
            objdbconnZHDB02.Close()
            objdbconnZHDB02 = Nothing
            objsqlcommandZHDB02 = Nothing
            objdsDebitor = Nothing
            objDADebitor = Nothing
            objdtDZKond = Nothing

        End Try


    End Function

    Friend Function FcGetDZKondSageID(ByVal intDZkond As Int16,
                                              ByRef intDZKondS200 As Int16) As Int16

        'Returns 0=ok, 1=ZK nicht gefunden, 9=Problem; intDZKond wird mit Sage 200 ZK abgefüllt

        Dim objdsDebitor As New DataSet
        Dim objDADebitor As New MySqlDataAdapter
        Dim objdtDZKond As New DataTable("tbllocDZKond")
        Dim objdbconnZHDB02 As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))
        Dim objsqlcommandZHDB02 As New MySqlCommand

        Try

            objdbconnZHDB02.Open()

            objsqlcommandZHDB02.Connection = objdbconnZHDB02

            'Zahlungsbedingung suchen
            'Es muss der Weg über ein Dataset genommen werden da sosnt constraint-Meldungen kommen
            objsqlcommandZHDB02.CommandText = "Select t_sage_zahlungskondition.SageID " +
                                                  "FROM t_sage_zahlungskondition " +
                                                  "WHERE t_sage_zahlungskondition.ID=" + intDZkond.ToString
            objDADebitor.SelectCommand = objsqlcommandZHDB02
            objdsDebitor.EnforceConstraints = False
            objDADebitor.Fill(objdsDebitor)

            If objdsDebitor.Tables(0).Rows.Count > 0 Then

                'ZK existiert
                If Not IsDBNull(objdsDebitor.Tables(0).Rows(0).Item("SageID")) Then
                    intDZKondS200 = objdsDebitor.Tables(0).Rows(0).Item("SageID")
                Else
                    'ZK existiert, aber Sage ID nicht definiert
                    intDZKondS200 = 0
                End If
                Return 0

            Else

                'ZK existiert nicht
                intDZKondS200 = 0
                Return 1

            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Debitor - Z-Bedingung - von ZK-Tabelle lesen", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            intDZkond = 0
            Return 9

        Finally
            objdbconnZHDB02.Close()
            objdbconnZHDB02 = Nothing
            objsqlcommandZHDB02 = Nothing
            objdtDZKond = Nothing
            objdsDebitor = Nothing
            objDADebitor = Nothing


        End Try


    End Function

    Friend Function FcCheckLinkedRG(intNewDebiNbr As Int32,
                                    strDebiCur As String,
                                    intBelegNbr As Int32,
                                    dblBetragToBook As Double,
                                    strPeriod As String) As Int16

        'Returns 0=ok, 1=Beleg nicht existent, 2=Beleg existiert, ist aber bezahlt, 9=Problem

        Dim intLaufNbr As Int32
        Dim strBeleg As String
        Dim strBelegArr() As String
        Dim dblBetragOpen As Double
        Dim intLaufNbrTZ As Int32
        Dim dblTZPayed As Double = 0

        Try

            intLaufNbr = objdbBuha.doesBelegExist2(intNewDebiNbr.ToString,
                                                  strDebiCur,
                                                  intBelegNbr.ToString,
                                                  "NOT_SET",
                                                  "R",
                                                  "NOT_SET",
                                                  "NOT_SET",
                                                  "NOT_SET")

            If intLaufNbr > 0 Then
                'Prüfung ob Beleg bezahlt
                strBeleg = objdbBuha.GetBeleg(intNewDebiNbr.ToString,
                                             intLaufNbr.ToString)

                strBelegArr = Split(strBeleg, "{>}")
                dblBetragOpen = strBelegArr(19)
                If strBelegArr(4) = "B" Then
                    Return 2
                Else
                    'Teilzahlungen suchen
                    intLaufNbrTZ = objdbBuha.doesBelegExist2(intNewDebiNbr.ToString,
                                                             strDebiCur,
                                                             intBelegNbr.ToString,
                                                             "NOT_SET",
                                                             "T",
                                                             "NOT_SET",
                                                             "NOT_SET",
                                                             "NOT_SET")
                    If intLaufNbrTZ > 0 Then
                        'Zahlungen aufsummieren und prüfen ob Abbucnhung möglich
                        'Alle Belege des Debitors lesen und TZ mit gleicher Beleg-Nr. aufsummieren
                        Call objdbBuha.ReadBeleg2(intNewDebiNbr.ToString,
                                                  strPeriod,
                                                  strDebiCur,
                                                  "O")

                        strBeleg = objdbBuha.GetBelegZeile2()

                        Do Until strBeleg = "EOF"
                            strBelegArr = Split(strBeleg, "{>}")

                            If strBelegArr(1) = intBelegNbr And strBelegArr(5) = "T" Then
                                dblTZPayed += strBelegArr(19) * -1
                            End If

                            strBeleg = objdbBuha.GetBelegZeile2()
                        Loop

                        'Ist TZ von GS auf RG 1 möglich?
                        If dblBetragOpen - dblTZPayed < dblBetragToBook * -1 Then
                            Return 4
                        Else
                            Return 0
                        End If

                    Else
                        'Ist Abbuchung möglich?
                        If (dblBetragToBook * -1) > dblBetragOpen Then
                            Return 3
                        Else
                            Return 0
                        End If

                    End If

                End If

            Else
                Return 1
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Fehler Prüfen Splitt-Bill Bel " + intBelegNbr.ToString)
            Return 9

        Finally
            strBelegArr = Nothing

        End Try

    End Function

    Friend Function FcGetDebitorFromLinkedRG(ByVal lngRGNbr As Int32,
                                                    ByVal intAccounting As Int32,
                                                    ByRef intDebiNew As Int32,
                                                    ByVal intTeqNbr As Int16,
                                                    ByVal intTeqNbrLY As Int16,
                                                    ByVal intTeqNbrPLY As Int16) As Int16

        'Return 0=ok, 1=Neue Debi genereiert und gesetzt, 2=Rep_Ref nicht definiert, 3=Nicht in Tab_Repbetriebe, 4=keine Angaben in Tab_Repbetriebe

        ', , , strDebNewField, strDebNewFieldType, strCompFieldName, strStreetFieldName, strZIPFieldName, strTownFieldName, strSageName, strDebiAccField As String
        'Dim intCreatable As Int16
        Dim objdtDebitor As New DataTable
        Dim intPKNewField As Int32
        Dim objdbConnDeb As New MySqlConnection
        Dim objsqlCommDeb As New MySqlCommand
        Dim strTableName As String
        Dim strTableType As String
        Dim strDebFieldName As String
        Dim tblDebiBuchung As New DataTable
        Dim objOrcommand As New OracleClient.OracleCommand
        Dim objdbSQLConn As New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("SQLConnectionString"))
        Dim objdbSQLCmd As New SqlCommand

        'Dim objlocOLEdbcmd As New OleDb.OleDbCommand
        'Dim strMDBName As String = Main.FcReadFromSettings(objdbconn, "Buchh_PKTableConnection", intAccounting)
        Dim strSQL As String
        'Dim intFunctionReturns As Int16

        Try

            'Zuerst probieren vom Beleg zu holen
            objdbSQLConn.Open()

            objdbSQLCmd.CommandText = "SELECT * FROM debibuchung WHERE teqnbr IN (" + intTeqNbr.ToString + ", " + intTeqNbrLY.ToString + ", " + intTeqNbrPLY.ToString + ")" +
                                                                 " AND belnbr=" + lngRGNbr.ToString +
                                                                 " AND typ='R'"

            objdbSQLCmd.Connection = objdbSQLConn

            tblDebiBuchung.Load(objdbSQLCmd.ExecuteReader)

            If tblDebiBuchung.Rows.Count = 1 Then
                intDebiNew = tblDebiBuchung.Rows(0).Item("debinbr")
                Return 0
            Else
                'Sonst von RG holen
                strTableName = FcReadFromSettingsII("Buchh_TableDeb", intAccounting)
                strTableType = FcReadFromSettingsII("Buchh_RGTableType", intAccounting)
                strDebFieldName = "RGNr"

                strSQL = "SELECT PKNr " +
                     "FROM " + strTableName + " WHERE " + strDebFieldName + "=" + lngRGNbr.ToString

                If strTableName <> "" And strDebFieldName <> "" Then

                    If strTableType = "O" Then 'Oracle
                        objOrcommand.CommandText = strSQL
                        objdtDebitor.Load(objOrcommand.ExecuteReader)
                    ElseIf strTableType = "M" Then 'MySQL
                        intDebiNew = 0
                        'MySQL - Tabelle einlesen
                        objdbConnDeb.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(FcReadFromSettingsII("Buchh_RGTableMDB", intAccounting))
                        objdbConnDeb.Open()
                        objsqlCommDeb.CommandText = strSQL
                        objsqlCommDeb.Connection = objdbConnDeb
                        objdtDebitor.Load(objsqlCommDeb.ExecuteReader)
                        objdbConnDeb.Close()

                    End If

                    If objdtDebitor.Rows.Count > 0 Then
                        'If IsDBNull(objdtDebitor.Rows(0).Item(strDebNewField)) And strTableName <> "Tab_Repbetriebe" Then 'Es steht nichts im Feld welches auf den Rep_Betrieb verweist oder wenn direkt
                        ' intDebiNew = 0
                        'Return 2
                        'Else

                        'Prüfen ob Repbetrieb schon eine neue Nummer erhalten hat.
                        If Not IsDBNull(objdtDebitor.Rows(0).Item("PKNr")) Then
                            intDebiNew = objdtDebitor.Rows(0).Item("PKNr")
                            'Else
                            '    intFunctionReturns = Main.FcNextPKNr(objdbconnZHDB02, lngDebiNbr, intDebiNew)
                            '    If intFunctionReturns = 0 And intDebiNew > 0 Then 'Vergabe hat geklappt
                            '        intFunctionReturns = Main.FcWriteNewDebToRepbetrieb(objdbconnZHDB02, lngDebiNbr, intDebiNew)
                            '        If intFunctionReturns = 0 Then 'Schreiben hat geklappt
                            '            Return 1
                            '        End If
                            '    End If
                        End If
                        Return 0
                    End If
                Else
                    Return 1
                End If

            End If

            'Return intPKNewField

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Fehler Prüfen Splitt-Bill")
            Return 9

        Finally
            objdbSQLConn.Close()
            objdbSQLConn = Nothing
            objdbConnDeb = Nothing
            objsqlCommDeb = Nothing
            objdbSQLConn = Nothing
            objOrcommand = Nothing
            objdtDebitor = Nothing

        End Try

    End Function

    Friend Function FcCheckDebiIntBank(ByVal intAccounting As Integer,
                                            ByVal striBankS50 As String,
                                            ByVal intPayType As Int16,
                                            ByRef intIBankS200 As String) As Int16

        '0=ok, 1=Sage50 iBank nicht gefunden, 2=Kein Standard gesetzt, 3=Nichts angegeben, auf Standard gesetzt, 9=Problem

        Dim objdbconn As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))
        Dim objdbcommand As New MySqlCommand
        Dim objdtiBank As New DataTable

        Try
            'wurde i Bank definiert?
            If striBankS50 <> "" Then
                'Sage 50 - Bank suchen
                objdbcommand.Connection = objdbconn

                objdbconn.Open()

                If intPayType = 10 Then 'QR - Fall
                    objdbcommand.CommandText = "SELECT intSage200QR FROM t_sage_tblaccountingbank WHERE QRTNNR='" + striBankS50 + "' AND intAccountingID=" + intAccounting.ToString
                Else
                    objdbcommand.CommandText = "SELECT intSage200 FROM t_sage_tblaccountingbank WHERE strBank='" + striBankS50 + "' AND intAccountingID=" + intAccounting.ToString
                End If
                objdtiBank.Load(objdbcommand.ExecuteReader)
                'wurde DS gefunden?
                If objdtiBank.Rows.Count > 0 Then
                    If intPayType = 10 Then 'QR - Fall
                        'Wurde auch wirklich eine ZV definiert (= intSage200QR > 0)?
                        If objdtiBank.Rows(0).Item("intSage200QR") > 0 Then
                            intIBankS200 = objdtiBank.Rows(0).Item("intSage200QR")
                        Else
                            intIBankS200 = 0
                            Return 1
                        End If
                    Else
                        intIBankS200 = objdtiBank.Rows(0).Item("intSage200")
                    End If
                    Return 0
                Else
                    intIBankS200 = 0
                    Return 1
                End If
            Else
                'Standard nehmen
                objdbcommand.Connection = objdbconn
                'objdbconn.Open()
                objdbcommand.CommandText = "SELECT intSage200 FROM t_sage_tblaccountingbank WHERE booStandard=true AND intAccountingID=" + intAccounting.ToString
                objdtiBank.Load(objdbcommand.ExecuteReader)
                'wurde ein Standard definieren
                If objdtiBank.Rows.Count > 0 Then
                    intIBankS200 = objdtiBank.Rows(0).Item("intSage200")
                    Return 3
                Else
                    intIBankS200 = 0
                    Return 2
                End If

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Eigene Bank - Suche")
            Return 9

        Finally
            objdbconn.Close()
            objdbconn = Nothing
            objdbcommand = Nothing
            objdtiBank = Nothing

        End Try

    End Function

    Friend Function FcCheckDate2(datDateToCheck As Date,
                                 strSelYear As String,
                                 tblDates As DataTable,
                                 booYChngAllowed As Boolean) As Int16

        '0=ok, 1=Jahr <> Sel Jahr, 2=Blockiert, 9=Problem

        Dim selrelDates() As DataRow
        Dim booIsDateOk As Boolean

        Try

            If Not booYChngAllowed Then
                'Entspricht Jahr im Dateum dem selektierten Jahr?
                If DateAndTime.Year(datDateToCheck) <> Conversion.Val(strSelYear) Then
                    Return 1
                Else
                    'Ist etwas blockiert?
                    selrelDates = tblDates.Select("intYear=" + strSelYear)
                    booIsDateOk = True
                    For Each drselrelDates In selrelDates
                        If datDateToCheck >= drselrelDates("datFrom") And datDateToCheck <= drselrelDates("datTo") Then
                            'Ist Status <> O
                            If drselrelDates("strStatus") <> "O" Then
                                booIsDateOk = False
                            End If
                        End If
                    Next
                    If Not booIsDateOk Then
                        Return 2
                    Else
                        Return 0
                    End If
                End If

            Else
                selrelDates = tblDates.Select("intYear=" + Convert.ToString(DateAndTime.Year(datDateToCheck)))
                'Ist etwas blockiert?
                booIsDateOk = True
                For Each drselrelDates In selrelDates
                    If datDateToCheck >= drselrelDates("datFrom") And datDateToCheck <= drselrelDates("datTo") Then
                        'Ist Status <> O
                        If drselrelDates("strStatus") <> "O" Then
                            booIsDateOk = False
                        End If
                    End If
                Next
                If Not booIsDateOk Then
                    Return 3
                Else
                    Return 0
                End If

            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "PGV-Datumscheck")
            Return 9

        End Try

    End Function

    Friend Function FcCheckOPDouble(strDebitor As String,
                                    lngDebIdentNbr As Int32,
                                    strOPNr As String,
                                    strType As String,
                                    strCurrency As String,
                                    booErfOPExt As Boolean) As Int16

        'Return 0=ok, 1=Beleg existiert, 9=Problem

        Dim intBelegReturn As Int32

        Try
            If Not booErfOPExt Then
                intBelegReturn = objdbBuha.doesBelegExist(strDebitor,
                                                      strCurrency,
                                                      FcCleanRGNrStrict(strOPNr),
                                                      "NOT_SET",
                                                      strType,
                                                      "")
                If intBelegReturn = 0 Then
                    'Zusätzlich extern überprüfen
                    intBelegReturn = objdbBuha.doesBelegExistExtern(strDebitor,
                                                                strCurrency,
                                                                strOPNr,
                                                                strType,
                                                                "")
                    If intBelegReturn <> 0 Then
                        Return 1
                    Else
                        Return 0
                    End If
                Else
                    Return 1
                End If

            Else

                intBelegReturn = objdbBuha.doesBelegExist(strDebitor,
                                                          strCurrency,
                                                          lngDebIdentNbr.ToString,
                                                          "NOT_SET",
                                                          strType,
                                                          "")
                If intBelegReturn = 0 Then
                    'Zusätzlich extern überprüfen
                    intBelegReturn = objdbBuha.doesBelegExistExtern(strDebitor,
                                                                strCurrency,
                                                                strOPNr,
                                                                strType,
                                                                "")
                    If intBelegReturn <> 0 Then
                        Return 1
                    Else
                        Return 0
                    End If
                Else
                    Return 1
                End If


            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Check doppelte OP - Nr.")
            Return 9

        Finally
            'Application.DoEvents()

        End Try

    End Function

    Friend Function FcReadDebitorName(intDebiNbr As Int32,
                                      strCurrency As String) As String

        Dim strDebitorName As String
        Dim strDebitorAr() As String

        Try

            If strCurrency = "" Then

                strDebitorName = objdbBuha.ReadDebitor3(intDebiNbr * -1, strCurrency)

            Else

                strDebitorName = objdbBuha.ReadDebitor3(intDebiNbr, strCurrency)

            End If

            strDebitorAr = Split(strDebitorName, "{>}")

            Return strDebitorAr(0)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Debitoren-Daten-Lesen Problem", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        Finally
            'Application.DoEvents()

        End Try


    End Function

    Friend Function FcIsDebitorCreatable(lngDebiNbr As Long,
                                         strcmbBuha As String,
                                         intAccounting As Int16) As Int16

        'Return: 0=creatable und erstellt, 3=Sage - Suchtext nicht erfasst, 4=Betrieb nicht gefunden, 5=PK nicht geprüft, 9=Nicht hinterlegt

        Dim intCreatable As Int16
        Dim objdtDebitor As New DataTable
        Dim strLand As String
        Dim intLangauage As Int32
        'Dim intPKNewField As Int32
        Dim strSQL As String
        Dim intAffected As Int16
        Dim strIBANNr As String
        Dim strBankName As String = String.Empty
        Dim strBankAddress1 As String = String.Empty
        Dim strBankAddress2 As String = String.Empty
        Dim strBankPLZ As String = String.Empty
        Dim strBankOrt As String = String.Empty
        Dim strBankBIC As String = String.Empty
        Dim strBankCountry As String = String.Empty
        Dim strBankClearing As String = String.Empty
        Dim intReturnValue As Int16
        Dim intDebZB As Int16
        Dim objdsDebitor As New DataSet
        Dim objDADebitor As New MySqlDataAdapter
        Dim objdtSachB As New DataTable("tbliSachB")
        Dim strSachB As String
        Dim intPayType As Int16
        Dim intintBank As Int16
        Dim objdbConnDeb As New MySqlConnection
        Dim objsqlConnDeb As New MySqlCommand
        Dim objdbconnZHDB02 As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))
        Dim objsqlcommandZHDB02 As New MySqlCommand
        Dim strConnection As String

        Try

            'Angaben einlesen
            strConnection = FcReadFromSettingsII("Buch_TabRepConnection", intAccounting)
            objdbConnDeb.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(strConnection)

            objdbConnDeb.Open()

            objdbconnZHDB02.Open()

            objsqlConnDeb.Connection = objdbConnDeb
            objsqlConnDeb.CommandText = "Select Rep_Nr, " +
                                                      "Rep_Suchtext, " +
                                                      "Rep_Firma, " +
                                                      "Rep_Strasse, " +
                                                      "Rep_PLZ, " +
                                                      "Rep_Ort, " +
                                                      "Rep_DebiKonto, " +
                                                      "Rep_Gruppe, " +
                                                      "Rep_Vertretung, " +
                                                      "Rep_Ansprechpartner, " +
                                                      "If(Rep_Land Is NULL, 'Schweiz', Rep_Land) AS Rep_Land, " +
                                                      "Rep_Tel1, " +
                                                      "Rep_Fax, " +
                                                      "Rep_Mail, " +
                                                      "IF(Rep_Language Is NULL, 'D', Rep_Language) AS Rep_Language, " +
                                                      "Rep_Kredi_MWSTNr, " +
                                                      "Rep_Kreditlimite, " +
                                                      "Rep_Kred_Pay_Def, " +
                                                      "Rep_Kred_Bank_Name, " +
                                                      "Rep_Kred_Bank_PLZ, " +
                                                      "Rep_Kred_Bank_Ort, " +
                                                      "Rep_Kred_IBAN, " +
                                                      "Rep_Kred_Bank_BIC, " +
                                                      "IF(Rep_Kred_Currency Is NULL, 'CHF', Rep_Kred_Currency) AS Rep_Kred_Currency, " +
                                                      "Rep_Kred_PCKto, " +
                                                      "Rep_DebiErloesKonto, " +
                                                      "Rep_Kred_BankIntern, " +
                                                      "ReviewedOn " +
                                                      "FROM Tab_Repbetriebe WHERE PKNr=" + lngDebiNbr.ToString
            objdtDebitor.Load(objsqlConnDeb.ExecuteReader)

            'Gefunden?
            If objdtDebitor.Rows.Count > 0 Then
                'Debug.Print("Gefunden, kann e/rstellt werden")

                If IsDBNull(objdtDebitor.Rows(0).Item("ReviewedOn")) Then
                    'PK wurde nicht geprüft
                    Return 5

                Else

                    'Sachbearbeiter suchen
                    'Ist Ausnahme definiert?
                    If IsNothing(objsqlcommandZHDB02.Connection) Then
                        objsqlcommandZHDB02.Connection = objdbconnZHDB02
                    End If
                    objsqlcommandZHDB02.CommandText = "SELECT CustomerID FROM t_rep_sagesachbearbeiter WHERE Rep_Nr=" + objdtDebitor.Rows(0).Item("Rep_Nr").ToString + " And Buchh_Nr=" + intAccounting.ToString
                    objdtSachB.Load(objsqlcommandZHDB02.ExecuteReader)
                    If objdtSachB.Rows.Count > 0 Then 'Ausnahme definiert auf Rep-Betrieb
                        strSachB = Trim(objdtSachB.Rows(0).Item("CustomerID").ToString)
                    Else
                        'Default setzen
                        objsqlcommandZHDB02.CommandText = "SELECT CustomerID FROM t_rep_sagesachbearbeiter WHERE Rep_Nr=2535 And Buchh_Nr=" + intAccounting.ToString
                        objdtSachB.Load(objsqlcommandZHDB02.ExecuteReader)
                        If objdtSachB.Rows.Count > 0 Then 'Default ist definiert
                            strSachB = Trim(objdtSachB.Rows(0).Item("CustomerID").ToString)
                        Else
                            strSachB = String.Empty
                            MessageBox.Show("Kein Sachbearbeiter - Default gesetzt für Buha " + strcmbBuha, "Debitorenerstellung")
                        End If
                    End If

                    'interne Bank
                    intReturnValue = FcCheckDebiIntBank(intAccounting,
                                                             objdtDebitor.Rows(0).Item("Rep_Kred_BankIntern"),
                                                             intintBank)

                    'Zahlungsbedingung suchen
                    intReturnValue = FcGetDZkondFromRep(lngDebiNbr,
                                                        intDebZB,
                                                        intAccounting)


                    ''objdtKreditor.Clear()
                    ''Es muss der Weg über ein Dataset genommen werden da sosnt constraint-Meldungen kommen
                    'objsqlcommandZHDB02.CommandText = "Select Tab_Repbetriebe.PKNr, t_sage_zahlungskondition.SageID " +
                    '                                  "FROM Tab_Repbetriebe INNER JOIN t_sage_zahlungskondition On Tab_Repbetriebe.Rep_DebiZKonditionID = t_sage_zahlungskondition.ID " +
                    '                                  "WHERE Tab_Repbetriebe.PKNr=" + lngDebiNbr.ToString
                    'objDADebitor.SelectCommand = objsqlcommandZHDB02
                    'objdsDebitor.EnforceConstraints = False
                    'objDADebitor.Fill(objdsDebitor)

                    ''objdsKreditor.Load(objsqlcommandZHDB02.ExecuteReader)
                    ''objdtKreditor.Load(objsqlcommandZHDB02.ExecuteReader)
                    'If Not IsDBNull(objdsDebitor.Tables(0).Rows(0).Item("SageID")) Then
                    '    intDebZB = objdsDebitor.Tables(0).Rows(0).Item("SageID")
                    'Else
                    '    intDebZB = 1
                    'End If

                    'Land von Text auf Auto-Kennzeichen ändern
                    Select Case IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Land")), "Schweiz", objdtDebitor.Rows(0).Item("Rep_Land"))
                        Case "Schweiz"
                            strLand = "CH"
                        Case "Deutschland"
                            strLand = "DE"
                        Case "Frankreich"
                            strLand = "FR"
                        Case "Italien"
                            strLand = "IT"
                        Case "Österreich"
                            strLand = "AT"
                        Case "USA"
                            strLand = "US"
                        Case Else
                            strLand = "CH"
                    End Select

                    'Sprache zuweisen von 1-Stelligem String nach Sage 200 Regionen
                    Select Case Strings.UCase(IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Language")), "D", objdtDebitor.Rows(0).Item("Rep_Language")))
                        Case "D", "DE", ""
                            intLangauage = 2055
                        Case "F", "FR"
                            intLangauage = 4108
                        Case "I", "IT"
                            intLangauage = 2064
                        Case Else
                            intLangauage = 2057 'Englisch
                    End Select

                    'Variablen zuweisen für die Erstellung des Debitors
                    strIBANNr = IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Kred_IBAN")), "", objdtDebitor.Rows(0).Item("Rep_Kred_IBAN"))
                    strBankName = IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Kred_Bank_Name")), "", objdtDebitor.Rows(0).Item("Rep_Kred_Bank_Name"))
                    strBankAddress1 = String.Empty
                    strBankPLZ = IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Kred_Bank_PLZ")), "", objdtDebitor.Rows(0).Item("Rep_Kred_Bank_PLZ"))
                    strBankOrt = IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Kred_Bank_Ort")), "", objdtDebitor.Rows(0).Item("Rep_Kred_Bank_Ort"))
                    strBankAddress2 = strBankPLZ + " " + strBankOrt
                    strBankBIC = IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Kred_Bank_BIC")), "", objdtDebitor.Rows(0).Item("Rep_Kred_Bank_BIC"))
                    strBankClearing = IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Kred_PCKto")), "", objdtDebitor.Rows(0).Item("Rep_Kred_PCKto"))

                    If Len(strIBANNr) = 21 Then 'IBAN
                        'If intPayType <> 9 Then 'Type nicht IBAN angegeben aber IBAN - Nr. erfasst
                        intPayType = 9
                        'End If
                        intReturnValue = FcGetIBANDetails(strIBANNr,
                                                          strBankName,
                                                          strBankAddress1,
                                                          strBankAddress2,
                                                          strBankBIC,
                                                          strBankCountry,
                                                          strBankClearing)

                        'Kombinierte PLZ / Ort Feld trennen
                        strBankPLZ = Strings.Left(strBankAddress2, InStr(strBankAddress2, " "))
                        strBankOrt = Trim(Strings.Right(strBankAddress2, Len(strBankAddress2) - InStr(strBankAddress2, " ")))
                    End If

                    intCreatable = FcCreateDebitor(lngDebiNbr,
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Suchtext")), "", objdtDebitor.Rows(0).Item("Rep_Suchtext")),
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Firma")), "", objdtDebitor.Rows(0).Item("Rep_Firma")),
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Strasse")), "", objdtDebitor.Rows(0).Item("Rep_Strasse")),
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_PLZ")), "", objdtDebitor.Rows(0).Item("Rep_PLZ")),
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Ort")), "", objdtDebitor.Rows(0).Item("Rep_Ort")),
                                              objdtDebitor.Rows(0).Item("Rep_DebiKonto"),
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Gruppe")), "", objdtDebitor.Rows(0).Item("Rep_Gruppe")),
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Vertretung")), "", objdtDebitor.Rows(0).Item("Rep_Vertretung")),
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Ansprechpartner")), "", objdtDebitor.Rows(0).Item("Rep_Ansprechpartner")),
                                              strLand,
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Tel1")), "", objdtDebitor.Rows(0).Item("Rep_Tel1")),
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Fax")), "", objdtDebitor.Rows(0).Item("Rep_Fax")),
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Mail")), "", objdtDebitor.Rows(0).Item("Rep_Mail")),
                                              intLangauage,
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Kredi_MWStNr")), "", objdtDebitor.Rows(0).Item("Rep_Kredi_MWStNr")),
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Kreditlimite")), "", objdtDebitor.Rows(0).Item("Rep_Kreditlimite")),
                                              intPayType,
                                              strBankName,
                                              strBankPLZ,
                                              strBankOrt,
                                              strIBANNr,
                                              strBankBIC,
                                              strBankClearing,
                                              IIf(String.IsNullOrEmpty(objdtDebitor.Rows(0).Item("Rep_Kred_Currency")), "CHF", objdtDebitor.Rows(0).Item("Rep_Kred_Currency")),
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_DebiErloesKonto")), "3200", objdtDebitor.Rows(0).Item("Rep_DebiErloesKonto")),
                                              intDebZB,
                                              strSachB,
                                              intintBank,
                                              "")

                    If intCreatable = 0 Then
                        'MySQL
                        'strSQL = "INSERT INTO Tbl_RTFAutomail (RGNbr, MailCreateDate, MailCreateWho, MailTo, MailSender, MailTitle, MAilMsg, MailSent) VALUES (" +
                        ' intAccounting.ToString + lngDebiNbr.ToString + ", Date('" + Format(Today(), "yyyy-MM-dd").ToString + "'), 'Sage200Imp', " +
                        '                                     "'finance@mssag.ch', 'Sage200@mssag.ch', 'Debitor " +
                        'lngDebiNbr.ToString + " wurde erstell im Mandant " + strcmbBuha + "', 'Bitte kontrollieren und Daten erg&auml;nzen.', false)"
                        ' objlocMySQLRGConn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(strMDBName)
                        'objlocMySQLRGConn.Open()
                        'objlocMySQLRGcmd.Connection = objlocMySQLRGConn
                        'objsqlcommandZHDB02.CommandText = strSQL
                        'intAffected = objsqlcommandZHDB02.ExecuteNonQuery()

                    End If


                    Return 0
                End If

            Else

                Return 4

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Debitor - Erstellbar - Abklärung", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 9

        Finally
            objdbconnZHDB02.Close()
            objdbconnZHDB02 = Nothing
            objdbConnDeb.Close()
            objdbConnDeb = Nothing
            objsqlcommandZHDB02 = Nothing
            objdsDebitor = Nothing
            objDADebitor = Nothing

        End Try

    End Function

    Friend Function FcCheckDebiIntBank(ByVal intAccounting As Integer,
                                           ByVal striBankS50 As String,
                                           ByRef intIBankS200 As String) As Int16

        '0=ok, 1=Sage50 iBank nicht gefunden, 2=Kein Standard gesetzt, 3=Nichts angegeben, auf Standard gesetzt, 9=Problem

        Dim objdbconn As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString"))
        Dim objdbcommand As New MySqlCommand
        Dim objdtiBank As New DataTable

        Try

            objdbconn.Open()
            'wurde i Bank definiert?
            If striBankS50 <> "" Then
                'Sage 50 - Bank suchen
                objdbcommand.Connection = objdbconn
                'objdbconn.Open()
                objdbcommand.CommandText = "SELECT intSage200 FROM t_sage_tblaccountingbank WHERE strBank='" + striBankS50 + "' AND intAccountingID=" + intAccounting.ToString
                objdtiBank.Load(objdbcommand.ExecuteReader)
                'wurde DS gefunden?
                If objdtiBank.Rows.Count > 0 Then
                    intIBankS200 = objdtiBank.Rows(0).Item("intSage200")
                    Return 0
                Else
                    intIBankS200 = 0
                    Return 1
                End If
            Else
                'Standard nehmen
                objdbcommand.Connection = objdbconn
                'objdbconn.Open()
                objdbcommand.CommandText = "SELECT intSage200 FROM t_sage_tblaccountingbank WHERE booStandard=true AND intAccountingID=" + intAccounting.ToString
                objdtiBank.Load(objdbcommand.ExecuteReader)
                'wurde ein Standard definieren
                If objdtiBank.Rows.Count > 0 Then
                    intIBankS200 = objdtiBank.Rows(0).Item("intSage200")
                    Return 3
                Else
                    intIBankS200 = 0
                    Return 2
                End If

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Eigene Bank - Suche")
            Return 9

        Finally
            objdbconn.Close()
            objdbconn = Nothing
            objdtiBank = Nothing

        End Try

    End Function

    Friend Function FcGetIBANDetails(ByVal strIBAN As String,
                                        ByRef strBankName As String,
                                        ByRef strBankAddress1 As String,
                                        ByRef strBankAddress2 As String,
                                        ByRef strBankBIC As String,
                                        ByRef strBankCountry As String,
                                        ByRef strBankClearing As String) As Int16

        Dim objdbconn As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))
        Dim objIBANReq As HttpWebRequest
        Dim objdtIBAN As New DataTable
        'Dim striBANURI As New Uri("https://rest.sepatools.eu/validate_iban_dummy/AL90208110080000001039531801")
        Dim strIBANURI As New Uri("https://ssl.ibanrechner.de/http.html?function=validate_iban&iban=" + strIBAN + "&user=MSSAGSchweiz&password=6ux!mCXiS6EmCiA")
        Dim strResponse As String
        Dim objResponse As HttpWebResponse
        Dim objXMLDoc As New XmlDocument
        Dim objXMLNodeList As XmlNodeList
        Dim strXMLTag(10) As String
        Dim strXMLText(10) As String
        Dim strXMLAddress() As String
        Dim strBalance As String

        Dim objmysqlcom As New MySqlCommand

        Dim intRecAffected As Integer

        Try

            'Zuerst prüfen ob IBAN nicht schon in der Tabelle der bekannten existiert
            objdbconn.Open()
            objmysqlcom.Connection = objdbconn
            objmysqlcom.CommandText = "SELECT * FROM t_sage_tbliban WHERE strIBANNr='" + strIBAN + "'"
            objdtIBAN.Load(objmysqlcom.ExecuteReader)
            If objdtIBAN.Rows.Count = 0 Then

                objIBANReq = DirectCast(HttpWebRequest.Create(strIBANURI), HttpWebRequest)
                If (objIBANReq.GetResponse().ContentLength > 0) Then
                    objResponse = objIBANReq.GetResponse()
                    'Dim objStreamReader As New StreamReader(objIBANReq.GetResponse().GetResponseStream())
                    Dim objStreamReader As New StreamReader(objResponse.GetResponseStream())
                    'strResponse = objStreamReader.ReadToEnd()
                    objXMLDoc.LoadXml(objStreamReader.ReadToEnd())
                    'Antwort der Funktion
                    objXMLNodeList = objXMLDoc.SelectNodes("/result")
                    For Each objXMLNode As XmlNode In objXMLNodeList
                        'result
                        strXMLTag(0) = objXMLNode.ChildNodes.Item(1).Name
                        strXMLText(0) = objXMLNode.ChildNodes.Item(1).InnerText
                        'return code
                        strXMLTag(1) = objXMLNode.ChildNodes.Item(2).Name
                        strXMLText(1) = objXMLNode.ChildNodes.Item(2).InnerText
                        'country
                        strXMLTag(2) = objXMLNode.ChildNodes.Item(6).Name
                        strXMLText(2) = objXMLNode.ChildNodes.Item(6).InnerText
                        'bank-code
                        strXMLTag(3) = objXMLNode.ChildNodes.Item(7).Name
                        strXMLText(3) = objXMLNode.ChildNodes.Item(7).InnerText
                        'bank
                        strXMLTag(4) = objXMLNode.ChildNodes.Item(8).Name
                        strXMLText(4) = objXMLNode.ChildNodes.Item(8).InnerText
                        'bank address
                        strXMLTag(5) = objXMLNode.ChildNodes.Item(9).Name
                        strXMLAddress = Split(objXMLNode.ChildNodes.Item(9).InnerText, vbLf)
                        If strXMLAddress.Count = 2 Then
                            strXMLText(5) = strXMLAddress(0)
                            strXMLTag(6) = "bank_address2"
                            strXMLText(6) = strXMLAddress(1)
                        ElseIf strXMLAddress.Count = 3 Then
                            strXMLText(5) = strXMLAddress(1)
                            strXMLTag(6) = "bank_address2"
                            strXMLText(6) = strXMLAddress(2)
                        End If
                        strXMLTag(7) = objXMLNode.ChildNodes.Item(39).Name
                        strXMLText(7) = objXMLNode.ChildNodes.Item(39).InnerText
                    Next
                    'BIC
                    objXMLNodeList = objXMLDoc.SelectNodes("/result/bic_candidates-list/bic_candidates")
                    For Each objXMLNode As XmlNode In objXMLNodeList
                        'result
                        strXMLTag(8) = objXMLNode.ChildNodes.Item(0).Name
                        strXMLText(8) = objXMLNode.ChildNodes.Item(0).InnerText
                    Next

                    'objXMLDoc.Load(strResponse)
                    objStreamReader.Close()
                    objResponse.Close()
                    strBankName = Trim(strXMLText(4))
                    strBankAddress1 = Trim(strXMLText(5))
                    strBankAddress2 = Trim(strXMLText(6))
                    strBankCountry = Trim(strXMLText(2))
                    strBankClearing = Trim(strXMLText(3))
                    strBankBIC = Trim(strXMLText(8))

                    'in IBAN-Tabelle schreiben
                    objmysqlcom.CommandText = "INSERT INTO t_sage_tbliban (strIBANNr, 
                                                                        strIBANBankName, 
                                                                        strIBANBankAddress1, 
                                                                        strIBANBankAddress2, 
                                                                        strIBANBankBIC, 
                                                                        strIBANBankCountry, 
                                                                        strIBANBankClearing) " +
                                                            "VALUES('" + strIBAN + "', '" +
                                                            Replace(strBankName, "'", "`") + "', '" +
                                                            Replace(strBankAddress1, "'", "`") + "', '" +
                                                            Replace(strBankAddress2, "'", "`") + "', '" +
                                                            strBankBIC + "', '" +
                                                            strBankCountry + "', '" +
                                                            strBankClearing + "')"
                    intRecAffected = objmysqlcom.ExecuteNonQuery()

                    Return 0

                End If
            Else
                'Aus Tabelle zurückgeben
                strBankName = objdtIBAN.Rows(0).Item("strIBANBankName")
                strBankAddress1 = objdtIBAN.Rows(0).Item("strIBANBankAddress1")
                strBankAddress2 = objdtIBAN.Rows(0).Item("strIBANBankAddress2")
                strBankCountry = objdtIBAN.Rows(0).Item("strIBANBankCountry")
                strBankClearing = objdtIBAN.Rows(0).Item("strIBANBankClearing")
                strBankBIC = objdtIBAN.Rows(0).Item("strIBANBankBIC")

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Fehler auf IBAN-Check " + strIBAN)
            Return 9

        Finally
            objdbconn.Close()
            objdbconn = Nothing
            objmysqlcom = Nothing
            objdtIBAN = Nothing
            objmysqlcom = Nothing
            objXMLDoc = Nothing
            objResponse = Nothing
            objXMLNodeList = Nothing

        End Try

    End Function

    Friend Function FcCreateDebitor(intDebitorNewNbr As Int32,
                                    strSuchtext As String,
                                    strDebName As String,
                                    strDebStreet As String,
                                    strDebPLZ As String,
                                    strDebOrt As String,
                                    intDebSammelKto As Int32,
                                    strGruppe As String,
                                    strVertretung As String,
                                    strAnsprechpartner As String,
                                    strLand As String,
                                    strTel As String,
                                    strFax As String,
                                    strMail As String,
                                    intLangauage As Int32,
                                    strMwStNr As String,
                                    strKreditLimite As String,
                                    intPayDefault As Int16,
                                    strZVBankName As String,
                                    strZVBankPLZ As String,
                                    strZVBankOrt As String,
                                    strZVIBAN As String,
                                    strZVBIC As String,
                                    strZVClearing As String,
                                    strCurrency As String,
                                    intDebErlKto As Int16,
                                    intDebZB As Int16,
                                    strSachB As String,
                                    intintBank As Int16,
                                    strFirtName As String) As Int16

        Dim strDebCountry As String = strLand
        Dim strDebCurrency As String = strCurrency
        Dim strDebSprachCode As String = intLangauage.ToString
        Dim strDebSperren As String = "N"
        'Dim intDebErlKto As Integer = 3200
        Dim shrDebZahlK As Short = 1 'Wird für EE fix auf 30 Tage Netto gesetzt
        Dim intDebToleranzNbr As Integer = 1
        Dim intDebMahnGroup As Integer = 1
        Dim strDebWerbung As String = "N"
        Dim strText As String = String.Empty
        Dim strTelefon1 As String
        Dim strTelefax As String

        strText = IIf(strGruppe = "", "", "Gruppe: " + strGruppe) + IIf(strVertretung = "" Or "0", "", strText + vbCrLf + "Vertretung: " + strVertretung)
        strTelefon1 = IIf(strTel = "" Or strTel = "0", "", strTel)
        strTelefax = IIf(strFax = "" Or strFax = "0", "", strFax)

        'Evtl. falsch gesetztes Sammelkonto ändern
        If strCurrency <> "CHF" Then
            If strCurrency = "EUR" And intDebSammelKto <> 1105 Then
                intDebSammelKto = 1105
            End If
            If strCurrency = "USD" And intDebSammelKto <> 1102 Then
                intDebSammelKto = 1102
            End If
        End If

        'Debitor erstellen

        Try

            Call objdbBuha.SetCommonInfo2(intDebitorNewNbr,
                                         strDebName,
                                         strFirtName,
                                         "",
                                         strDebStreet,
                                         "",
                                         "",
                                         strDebCountry,
                                         strDebPLZ,
                                         strDebOrt,
                                         strTelefon1,
                                         "",
                                         strTelefax,
                                         strMail,
                                         "",
                                         strDebCurrency,
                                         "",
                                         "",
                                         strAnsprechpartner,
                                         strDebSprachCode,
                                         strText)

            Call objdbBuha.SetExtendedInfo8(strDebSperren,
                                           strKreditLimite,
                                           intDebSammelKto.ToString,
                                           intDebErlKto.ToString,
                                           strSachB,
                                           "",
                                           "",
                                           shrDebZahlK.ToString,
                                           intDebToleranzNbr.ToString,
                                           intDebMahnGroup.ToString,
                                           "",
                                           "",
                                           strDebWerbung,
                                           "",
                                           "",
                                           strMwStNr)

            'Suchtext in Indivual-Feld schreiben
            If Not String.IsNullOrEmpty(strSuchtext) Then
                Call objdbBuha.SetIndividInfoText(1,
                                                 strSuchtext)
            End If

            If intPayDefault = 9 Then 'IBAN
                If Len(strZVIBAN) > 15 Then
                    Call objdbBuha.SetZahlungsverbindung("B",
                                                        strZVIBAN,
                                                        strZVBankName,
                                                        "",
                                                        "",
                                                        strZVBankPLZ.ToString,
                                                        strZVBankOrt,
                                                        Strings.Left(strZVIBAN, 2),
                                                        strZVClearing,
                                                        "J",
                                                        strZVBIC,
                                                        "",
                                                        "",
                                                        "",
                                                        strZVIBAN,
                                                        "")
                End If
            End If
            Call objdbBuha.WriteDebitor3(0, intintBank.ToString)

            'Mail über Erstellung absetzen


            Return 0
            'intDebAdrLaufN = DbBhg.GetAdressLaufnr()
            'intDebBankLaufNr = DbBhg.GetZahlungsverbLaufnr()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Debitor - Erstellung " + intDebitorNewNbr.ToString + ", " + strDebName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Err.Clear()
            Return 1

        End Try

    End Function

    Friend Function FcIsPrivateDebitorCreatable(lngDebiNbr As Long,
                                              strcmbBuha As String,
                                              intAccounting As Int16) As Int16

        'Return: 0=creatable und erstellt, 3=Sage - Suchtext nicht erfasst, 4=Betrieb nicht gefunden, 9=Nicht hinterlegt

        Dim intCreatable As Int16
        Dim objdtDebitor As New DataTable
        Dim strLand As String
        Dim intLangauage As Int32
        'Dim intPKNewField As Int32
        Dim strSQL As String
        Dim intAffected As Int16
        Dim strIBANNr As String
        Dim strBankName As String = String.Empty
        Dim strBankAddress1 As String = String.Empty
        Dim strBankAddress2 As String = String.Empty
        Dim strBankPLZ As String = String.Empty
        Dim strBankOrt As String = String.Empty
        Dim strBankBIC As String = String.Empty
        Dim strBankCountry As String = String.Empty
        Dim strBankClearing As String = String.Empty
        Dim intReturnValue As Int16
        Dim intDebZB As Int16
        Dim objdbconnZHDB02 As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))
        Dim objsqlcommandZHDB02 As New MySqlCommand
        Dim objdsDebitor As New DataSet
        Dim objDADebitor As New MySqlDataAdapter
        Dim objdtSachB As New DataTable("tbliSachB")
        Dim strSachB As String
        Dim intPayType As Int16
        Dim strCurrency As String
        Dim intintBank As Int16

        Try

            'Angaben einlesen
            objdbconnZHDB02.Open()
            objsqlcommandZHDB02.Connection = objdbconnZHDB02
            objsqlcommandZHDB02.CommandText = "SELECT Lastname, " +
                                              "Firstname, " +
                                              "Street, " +
                                              "ZipCode, " +
                                              "City, " +
                                              "DebiGegenKonto, " +
                                              "'Privatperson' AS Gruppe, " +
                                              "IF(Country Is NULL, 'CH', country) AS country, " +
                                              "Phone, " +
                                              "Fax, " +
                                              "Email, " +
                                              "IF(Language Is NULL, 'DE',Language) AS Language, " +
                                              "BankName, " +
                                              "BankZipCode, " +
                                              "BankCountry, " +
                                              "IBAN, " +
                                              "BankBIC, " +
                                              "IF(Currency Is NULL, 'CHF', Currency) AS Currency, " +
                                              "DebiGegenKonto AS SammelKonto, " +
                                              "DebiErloesKonto AS ErloesKonto, " +
                                              "BankIntern, " +
                                              "DebiZKonditionID, " +
                                              "ReviewedOn " +
                                              "FROM t_customer WHERE PKNr=" + lngDebiNbr.ToString
            objdtDebitor.Load(objsqlcommandZHDB02.ExecuteReader)

            'Gefunden?
            If objdtDebitor.Rows.Count > 0 Then
                'Debug.Print("Gefunden, kann erstellt werden")

                If IsDBNull(objdtDebitor.Rows(0).Item("ReviewedOn")) Then
                    'PK wurde nicht geprüft

                    Return 5

                Else

                    'Sachbearbeiter suchen
                    'Default setzen
                    objsqlcommandZHDB02.CommandText = "SELECT CustomerID FROM t_rep_sagesachbearbeiter WHERE Rep_Nr=2535 And Buchh_Nr=" + intAccounting.ToString
                    objdtSachB.Load(objsqlcommandZHDB02.ExecuteReader)
                    If objdtSachB.Rows.Count > 0 Then 'Default ist definiert
                        strSachB = Trim(objdtSachB.Rows(0).Item("CustomerID").ToString)
                    Else
                        strSachB = String.Empty
                        MessageBox.Show("Kein Sachbearbeiter - Default gesetzt für Buha " + strcmbBuha, "Debitorenerstellung")
                    End If

                    'interne Bank
                    intReturnValue = FcCheckDebiIntBank(intAccounting,
                                                         objdtDebitor.Rows(0).Item("BankIntern"),
                                                         intintBank)


                    'Zahlungsbedingung suchen
                    intReturnValue = FcGetDZkondFromCust(lngDebiNbr,
                                                     intDebZB,
                                                     intAccounting)

                    'objdtKreditor.Clear()
                    'Es muss der Weg über ein Dataset genommen werden da sosnt constraint-Meldungen kommen
                    'objsqlcommandZHDB02.CommandText = "Select Tab_Repbetriebe.PKNr, t_sage_zahlungskondition.SageID " +
                    '                                  "FROM Tab_Repbetriebe INNER JOIN t_sage_zahlungskondition On Tab_Repbetriebe.Rep_DebiZKonditionID = t_sage_zahlungskondition.ID " +
                    '                                  "WHERE Tab_Repbetriebe.PKNr=" + lngDebiNbr.ToString
                    'objDADebitor.SelectCommand = objsqlcommandZHDB02
                    'objdsDebitor.EnforceConstraints = False
                    'objDADebitor.Fill(objdsDebitor)

                    ''objdsKreditor.Load(objsqlcommandZHDB02.ExecuteReader)
                    ''objdtKreditor.Load(objsqlcommandZHDB02.ExecuteReader)
                    'If Not IsDBNull(objdsDebitor.Tables(0).Rows(0).Item("SageID")) Then
                    'If IIf(IsDBNull(objdtDebitor.Rows(0).Item("DebiZKonditionID")), 0, objdtDebitor.Rows(0).Item("DebiZKonditionID")) <> 0 Then
                    '    intDebZB = objdtDebitor.Rows(0).Item("DebiZKonditionID")
                    'Else
                    '    intDebZB = 1
                    'End If

                    ''Land von Text auf Auto-Kennzeichen ändern
                    'Select Case IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Land")), "Schweiz", objdtDebitor.Rows(0).Item("Rep_Land"))
                    '    Case "Schweiz"
                    strLand = objdtDebitor.Rows(0).Item("country")
                    '    Case "Deutschland"
                    '        strLand = "DE"
                    '    Case "Frankreich"
                    '        strLand = "FR"
                    '    Case "Italien"
                    '        strLand = "IT"
                    '    Case "Österreich"
                    '        strLand = "AT"
                    '    Case Else
                    '        strLand = "CH"
                    'End Select

                    'Sprache zuweisen von 1-Stelligem String nach Sage 200 Regionen
                    Select Case IIf(IsDBNull(objdtDebitor.Rows(0).Item("Language")), "DE", objdtDebitor.Rows(0).Item("Language").ToUpper())
                        Case "DE", ""
                            intLangauage = 2055
                        Case "FR"
                            intLangauage = 4108
                        Case "IT"
                            intLangauage = 2064
                        Case Else
                            intLangauage = 2057 'Englisch
                    End Select

                    'Variablen zuweisen für die Erstellung des Debitors
                    strIBANNr = IIf(IsDBNull(objdtDebitor.Rows(0).Item("IBAN")), "", objdtDebitor.Rows(0).Item("IBAN"))
                    strBankName = IIf(IsDBNull(objdtDebitor.Rows(0).Item("BankName")), "", objdtDebitor.Rows(0).Item("BankName"))
                    strBankAddress1 = String.Empty
                    strBankPLZ = IIf(IsDBNull(objdtDebitor.Rows(0).Item("BankZipCode")), "", objdtDebitor.Rows(0).Item("BankZipCode"))
                    strBankOrt = String.Empty
                    strBankAddress2 = strBankPLZ + " " + strBankOrt
                    strBankBIC = IIf(IsDBNull(objdtDebitor.Rows(0).Item("BankBIC")), "", objdtDebitor.Rows(0).Item("BankBIC"))
                    strBankClearing = String.Empty

                    If Len(strIBANNr) >= 21 Then 'IBAN
                        'If intPayType <> 9 Then 'Type nicht IBAN angegeben aber IBAN - Nr. erfasst
                        intPayType = 9
                        'End If
                        intReturnValue = FcGetIBANDetails(strIBANNr,
                                                      strBankName,
                                                      strBankAddress1,
                                                      strBankAddress2,
                                                      strBankBIC,
                                                      strBankCountry,
                                                      strBankClearing)

                        'Kombinierte PLZ / Ort Feld trennen
                        strBankPLZ = Strings.Left(strBankAddress2, InStr(strBankAddress2, " "))
                        strBankOrt = Trim(Strings.Right(strBankAddress2, Len(strBankAddress2) - InStr(strBankAddress2, " ")))
                    End If

                    'Currency - Check
                    If objdtDebitor.Rows(0).Item("DebiGegenKonto") = 1105 And lngDebiNbr >= 40000 Then
                        strCurrency = "EUR"
                    Else
                        strCurrency = "CHF"
                    End If

                    intCreatable = FcCreateDebitor(lngDebiNbr,
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("LastName")), "", objdtDebitor.Rows(0).Item("LastName")) + IIf(IsDBNull(objdtDebitor.Rows(0).Item("FirstName")), "", objdtDebitor.Rows(0).Item("FirstName")),
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("LastName")), "", objdtDebitor.Rows(0).Item("LastName")),
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("Street")), "", objdtDebitor.Rows(0).Item("Street")),
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("ZipCode")), "", objdtDebitor.Rows(0).Item("ZipCode")),
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("City")), "", objdtDebitor.Rows(0).Item("City")),
                                              objdtDebitor.Rows(0).Item("SammelKonto"),
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("Gruppe")), "", objdtDebitor.Rows(0).Item("Gruppe")),
                                              "",
                                              "",
                                              strLand,
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("Phone")), "", objdtDebitor.Rows(0).Item("Phone")),
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("Fax")), "", objdtDebitor.Rows(0).Item("Fax")),
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("Email")), "", objdtDebitor.Rows(0).Item("Email")),
                                              intLangauage,
                                              "",
                                              0,
                                              intPayType,
                                              strBankName,
                                              strBankPLZ,
                                              strBankOrt,
                                              strIBANNr,
                                              strBankBIC,
                                              strBankClearing,
                                              strCurrency,
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("ErloesKonto")), "3200", objdtDebitor.Rows(0).Item("ErloesKonto")),
                                              intDebZB,
                                              strSachB,
                                              intintBank,
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("Firstname")), "", objdtDebitor.Rows(0).Item("Firstname")))

                    If intCreatable = 0 Then
                        'MySQL
                        'strSQL = "INSERT INTO Tbl_RTFAutomail (RGNbr, MailCreateDate, MailCreateWho, MailTo, MailSender, MailTitle, MAilMsg, MailSent) VALUES (" +
                        '                                     intAccounting.ToString + lngDebiNbr.ToString + ", Date('" + Format(Today(), "yyyy-MM-dd").ToString + "'), 'Sage200Imp', " +
                        '                                     "'finance@mssag.ch', 'Sage200@mssag.ch', 'Debitor " +
                        '                                     lngDebiNbr.ToString + " wurde erstell im Mandant " + strcmbBuha + "', 'Bitte kontrollieren und Daten erg&auml;nzen.', false)"
                        '' objlocMySQLRGConn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(strMDBName)
                        ''objlocMySQLRGConn.Open()
                        ''objlocMySQLRGcmd.Connection = objlocMySQLRGConn
                        'objsqlcommandZHDB02.CommandText = strSQL
                        'intAffected = objsqlcommandZHDB02.ExecuteNonQuery()

                        intCreatable = FcWriteDatetoPrivate(lngDebiNbr,
                                                             intAccounting,
                                                             0)


                    End If

                    Return 0

                End If

            Else

                Return 4

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Debitor - Erstellbar - Abklärung", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 9

        Finally
            objdbconnZHDB02.Close()
            objdbconnZHDB02 = Nothing
            objsqlcommandZHDB02 = Nothing
            objdsDebitor = Nothing
            objDADebitor = Nothing
            objdtDebitor = Nothing
            objdtSachB = Nothing

        End Try

    End Function

    Friend Function FcWriteDatetoPrivate(ByVal intNewPKNr As Int32,
                                                ByVal intAccounting As Int16,
                                                ByVal intDebitKredit As Int16) As Int16

        '0=ok, 1=PKNr nicht existent, 2=DS konnte nicht erstellt werden, 9=Problem

        Dim objdbCmd As New MySqlCommand
        Dim intAffected As Int16
        Dim strSQL As String
        Dim intRepNr As Int32
        Dim objdtPrivate As New DataTable
        Dim strDebiCreatedField As String
        Dim objdbcon As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))

        Try

            If intDebitKredit = 0 Then
                strDebiCreatedField = "DebiCreatedPKOn"
            Else
                strDebiCreatedField = "CrediCreatedPKON"
            End If

            'Zuerst CustomerID suchen

            objdbcon.Open()

            objdbCmd.Connection = objdbcon
            objdbCmd.CommandText = "SELECT ID FROM t_customer WHERE PKNr=" + intNewPKNr.ToString
            objdtPrivate.Load(objdbCmd.ExecuteReader)

            If objdtPrivate.Rows.Count > 0 Then 'Gefunden
                intRepNr = objdtPrivate.Rows(0).Item("ID")
                'Nun in t_customer_sagepknrcreation UPDATE probieren
                strSQL = "UPDATE t_customer_sagepkcreating SET " + strDebiCreatedField + " = CURRENT_DATE WHERE CustomerID=" + intRepNr.ToString + " AND Buchh_Nr=" + intAccounting.ToString
                objdbCmd.CommandText = strSQL
                intAffected = objdbCmd.ExecuteNonQuery()
                If intAffected <> 1 Then
                    'DS muss angelegt werden
                    strSQL = "INSERT INTO t_customer_sagepkcreating (CustomerID, Buchh_Nr, " + strDebiCreatedField + ", CreatedBy) VALUES(" + intRepNr.ToString + ", " + intAccounting.ToString + ", CURRENT_DATE, 'Sage 50 Transfer')"
                    objdbCmd.CommandText = strSQL
                    intAffected = objdbCmd.ExecuteNonQuery()
                    If intAffected <> 1 Then
                        Return 2
                    Else
                        Return 0
                    End If
                Else
                    'DS war schon da und konnte geupdated werden
                    Return 0
                End If

            Else
                Return 1
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Fehler Scrheiben t_rep_sagepknrcreation")
            Return 9

        Finally
            objdbcon.Close()
            objdbcon = Nothing
            objdbCmd = Nothing
            objdtPrivate = Nothing

        End Try

    End Function

    Friend Function FcCreateDebRef(ByVal intAccounting As Integer,
                                          ByVal strBank As String,
                                          ByVal strRGNr As String,
                                          ByRef strOPNr As String,
                                          ByVal intBuchungsArt As Integer,
                                          ByRef strReferenz As String,
                                          ByVal intPayType As Integer) As Integer

        'Return 0=ok oder nicht nötig, 1=keine Angaben hinterlegt, 2=Berechnung hat nicht geklappt

        Dim strTLNNr As String
        Dim strCleanedNr As String = String.Empty
        Dim strRefFrom As String
        Dim intLengthWTlNr As Int16

        Try

            If intBuchungsArt = 1 Then
                'Checken ob Referenz aus OP - Nr. oder aus Rechnung erstellt werden soll

                strRefFrom = FcReadFromSettingsII("Buchh_ESRNrFrom", intAccounting)
                If strRefFrom = "" Then
                    strRefFrom = "R"
                End If

                Select Case strRefFrom
                    Case "R"
                        strCleanedNr = strRGNr
                        strOPNr = strRGNr
                    Case "O"
                        strCleanedNr = strOPNr

                End Select

                strTLNNr = FcReadBankSettings(intAccounting,
                                              intPayType,
                                              strBank)

                'Bei HW Mandant an TLNr anhängen
                If intAccounting = 29 Then
                    strTLNNr += Strings.Left(strCleanedNr, 3)
                End If

                strCleanedNr = FcCleanRGNrStrict(strCleanedNr)

                intLengthWTlNr = 26 - Len(strTLNNr)

                strReferenz = strTLNNr + StrDup(intLengthWTlNr - Len(strCleanedNr), "0") + strCleanedNr + Trim(CStr(FcModulo10(strTLNNr + StrDup(intLengthWTlNr - Len(strCleanedNr), "0") + strCleanedNr)))

                Return 0

            Else
                Return 0
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem Referenzerstellung")
            Return 1

        Finally


        End Try


    End Function

    Friend Function FcReadBankSettings(ByVal intAccounting As Int16,
                                           ByVal intPayType As Int16,
                                           ByVal strBank As String) As String

        Dim objlocdtBank As New DataTable("tbllocBank")
        Dim objlocMySQLcmd As New MySqlCommand
        Dim objdbconn As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))

        Try


            If intPayType = 10 Then
                objlocMySQLcmd.CommandText = "SELECT strBLZ FROM t_sage_tblaccountingbank WHERE intAccountingID=" + intAccounting.ToString + " AND QRTNNR='" + strBank + "'"
            Else
                objlocMySQLcmd.CommandText = "SELECT strBLZ FROM t_sage_tblaccountingbank WHERE intAccountingID=" + intAccounting.ToString + " AND strBank='" + strBank + "'"
            End If

            objdbconn.Open()
            objlocMySQLcmd.Connection = objdbconn
            objlocdtBank.Load(objlocMySQLcmd.ExecuteReader)

            If objlocdtBank.Rows.Count > 0 Then
                Return objlocdtBank.Rows(0).Item(0).ToString
            Else
                Return "7777777"
            End If




        Catch ex As Exception
            MessageBox.Show(ex.Message, "Bankleitzahl suchen.")

        Finally
            objdbconn.Close()
            objdbconn = Nothing
            objlocdtBank = Nothing
            objlocMySQLcmd = Nothing

        End Try


    End Function

    Friend Function FcModulo10(ByVal strNummer As String) As Integer

        'strNummer darf nur Ziffern zwischen 0 und 9 enthalten!

        Dim intTabelle(0 To 9) As Integer
        Dim intÜbertrag As Integer
        Dim intIndex As Integer

        Try

            intTabelle(0) = 0 : intTabelle(1) = 9
            intTabelle(2) = 4 : intTabelle(3) = 6
            intTabelle(4) = 8 : intTabelle(5) = 2
            intTabelle(6) = 7 : intTabelle(7) = 1
            intTabelle(8) = 3 : intTabelle(9) = 5

            For intIndex = 1 To Len(strNummer)
                intÜbertrag = intTabelle((intÜbertrag + Mid(strNummer, intIndex, 1)) Mod 10)
            Next

            Return (10 - intÜbertrag) Mod 10

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem Modulo10")
            Err.Clear()

        End Try


    End Function

    Friend Function FcCheckBelegHead(intBuchungsArt As Int16,
                                     dblBrutto As Double,
                                     dblNetto As Double,
                                     dblMwSt As Double,
                                     dblRDiff As Double) As Int16

        'Returns 0=ok oder nicht wichtig, 1=Brutto, 2=Netto, 3=Beide, 4=Diff

        Try

            If intBuchungsArt = 1 Then
                If dblBrutto = 0 And dblNetto = 0 Then
                    Return 3
                ElseIf dblBrutto = 0 Then
                    Return 1
                ElseIf dblNetto = 0 Then
                    'Return 2
                ElseIf Math.Abs(Decimal.Round(dblBrutto - dblNetto - dblMwSt - dblRDiff, 2, MidpointRounding.AwayFromZero)) > 0 Then 'Math.Round(dblBrutto - dblRDiff - dblMwSt, 2, MidpointRounding.AwayFromZero) <> Math.Round(dblNetto, 2, MidpointRounding.AwayFromZero) Then
                    Return 4
                Else
                    Return 0
                End If
            Else
                Return 0
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem Check-Head")

        End Try

    End Function

    Friend Function FcCheckSubBookings(strDebRgNbr As String,
                                              ByRef objDtDebiSub As DataTable,
                                              ByRef intSubNumber As Int16,
                                              ByRef dblSubBrutto As Double,
                                              ByRef dblSubNetto As Double,
                                              ByRef dblSubMwSt As Double,
                                              datValuta As Date,
                                              intBuchungsArt As Int32,
                                              booAutoCorrect As Boolean,
                                              booCpyKSTToSub As Boolean,
                                              strKST As String,
                                              ByRef lngDebKonto As Int32,
                                              booCashSollKorrekt As Boolean,
                                              booSplittBill As Boolean,
                                              booLinkedGS As Boolean) As Int16

        'Functin Returns 0=ok, 1=Problem sub, 2=OP Diff zu Kopf, 3=OP nicht 0, 9=keine Subs

        'BitLog in Sub
        '1: Konto
        '2: KST
        '3: MwST
        '4: Brutto, Netto + MwSt 0
        '5: Netto 0
        '6: Brutto 0
        '7: Brutto - MwsT <> Netto

        Dim intReturnValue As Int32
        Dim strBitLog As String
        Dim strStatusText As String
        Dim strStrStCodeSage200 As String = String.Empty
        Dim dblStrStCodeSage As Double
        Dim strKstKtrSage200 As String = String.Empty
        Dim selsubrow() As DataRow
        Dim strStatusOverAll As String = "0000000"
        Dim strSteuer() As String
        Dim intSollKonto As Int32 = lngDebKonto
        Dim strProjDesc As String

        Try

            'Summen bilden und Angaben prüfen
            intSubNumber = 0
            dblSubNetto = 0
            dblSubMwSt = 0
            dblSubBrutto = 0

            selsubrow = objDtDebiSub.Select("strRGNr='" + strDebRgNbr + "'")

            For Each subrow As DataRow In selsubrow

                'Debug.Print("In Subrow Check")
                'If subrow("lngKto") = 3409 Then
                '    Stop
                'End If

                strBitLog = String.Empty

                'DB- Null Kto auf 0 setzen
                If IsDBNull(subrow("lngKto")) Then
                    subrow("lngKto") = 0
                End If

                'Runden
                If IsDBNull(subrow("dblNetto")) Then
                    subrow("dblNetto") = 0
                Else
                    subrow("dblNetto") = Decimal.Round(subrow("dblNetto"), 4, MidpointRounding.AwayFromZero)
                End If
                If IsDBNull(subrow("dblMwst")) Then
                    subrow("dblMwst") = 0
                Else
                    subrow("dblMwst") = Decimal.Round(subrow("dblMwst"), 2, MidpointRounding.AwayFromZero)
                End If
                If IsDBNull(subrow("dblBrutto")) Then
                    subrow("dblBrutto") = 0
                Else
                    subrow("dblBrutto") = Decimal.Round(subrow("dblBrutto"), 4, MidpointRounding.AwayFromZero)
                End If
                If IsDBNull(subrow("dblMwStSatz")) Then
                    subrow("dblMwStSatz") = 0
                Else
                    subrow("dblMwStSatz") = Decimal.Round(subrow("dblMwStSatz"), 1, MidpointRounding.AwayFromZero)
                End If

                'Falls KTRToSub dann kopieren
                If booCpyKSTToSub Then
                    subrow("lngKST") = strKST
                End If

                'Zuerst key auf 'ohne' setzen wenn MwSt-Satz = 0 und Mwst-Betrag = 0
                If subrow("dblMwStSatz") = 0 And subrow("dblMwst") = 0 And IIf(IsDBNull(subrow("strMwStKey")), "", subrow("strMwStKey")) <> "ohne" And
                    IIf(IsDBNull(subrow("strMwStKey")), "", subrow("strMwStKey")) <> "null" Then
                    'Stop
                    If IIf(IsDBNull(subrow("strMwStKey")), "", subrow("strMwStKey")) <> "AUSL0" And IIf(IsDBNull(subrow("strMwStKey")), "", subrow("strMwStKey")) <> "frei" Then
                        subrow("strMwStKey") = "ohne"
                    End If
                End If

                'Zuerst evtl. falsch gesetzte KTR oder Steuer - Sätze prüfen
                If (subrow("lngKto") >= 10000 Or subrow("lngKto") < 3000) Then 'Or subrow("strMwStKey") = "ohne" Then
                    Select Case subrow("lngKto")
                        Case 1120, 1121, 1200, 1500 To 1599, 1600 To 1699, 1700 To 1799, 2030
                            'Nur KST - Buchung resetten
                            subrow("lngKST") = 0
                        Case Else
                            subrow("strMwStKey") = Nothing
                            subrow("lngKST") = 0
                    End Select
                End If

                'MwSt prüfen 01
                If Not IsDBNull(subrow("strMwStKey")) And IIf(IsDBNull(subrow("strMwStKey")), "", subrow("strMwStKey")) <> "null" Then
                    dblStrStCodeSage = IIf(IsDBNull(subrow("dblMwStSatz")), 0, subrow("dblMwStSatz"))
                    intReturnValue = FcCheckMwSt(subrow("strMwStKey"),
                                                 dblStrStCodeSage,
                                                 strStrStCodeSage200,
                                                 subrow("lngKto"))
                    If intReturnValue = 0 Then
                        subrow("strMwStKey") = strStrStCodeSage200
                        subrow("dblMwStSatz") = dblStrStCodeSage
                        'Check ob korrekt berechnet
                        'Falsche Steueersätze abfangen
                        Try

                            strSteuer = Split(objfiBuha.GetSteuerfeld2(subrow("lngKto").ToString,
                                                                  "Zum Rechnen",
                                                                  subrow("dblBrutto").ToString,
                                                                  strStrStCodeSage200,
                                                                  "",
                                                                  Format(datValuta, "yyyyMMdd"),
                                                                  Convert.ToString(subrow("dblMwStSatz"))), "{<}")

                        Catch ex As Exception
                            'Debug.Print(ex.Message + ", " + (Err.Number And 65535).ToString)
                            If (Err.Number And 65535) = 525 Then
                                strSteuer = Split(objfiBuha.GetSteuerfeld2(subrow("lngKto").ToString,
                                                                  "Zum Rechnen",
                                                                  subrow("dblBrutto").ToString,
                                                                  strStrStCodeSage200), "{<}")
                            End If

                        End Try
                        If Val(strSteuer(2)) <> IIf(IsDBNull(subrow("dblMwst")), 0, subrow("dblMwst")) Then
                            'Im Fall von Auto-Korrekt anpassen wenn Toleranz
                            'Stop
                            '                            If booAutoCorrect Then 'And Val(strSteuer(2)) - subrow("dblMwst") <= 1.5 Then
                            'Falls MwSt-Betrag nur in 3 und 4 Stelle anders, dann erfassten Betrag nehmen.
                            If Math.Abs(Val(strSteuer(2)) - IIf(IsDBNull(subrow("dblMwst")), 0, subrow("dblMwst"))) >= 0.01 Then
                                strStatusText += "MwSt " + subrow("dblMwst").ToString
                                subrow("dblMwst") = Val(strSteuer(2))
                                'subrow("dblMwStSatz") = Val(strSteuer(3))
                                'subrow("dblBrutto") = Decimal.Round(subrow("dblNetto") + subrow("dblMwSt"), 2, MidpointRounding.AwayFromZero)
                                'subrow("dblNetto") = Decimal.Round(subrow("dblBrutto") + subrow("dblMwSt"), 2, MidpointRounding.AwayFromZero)
                                strStatusText += " cor -> " + subrow("dblMwst").ToString + ", "
                                '                           Else
                                '                          If Val(strSteuer(2)) - subrow("dblMwst") > 10 Then
                                '                         strStatusText += " -> " + strSteuer(2).ToString + ", "
                                '                        intReturnValue = 1
                                '                   Else
                                '                      strStatusText += " Tol -> " + strSteuer(2).ToString + ", "
                                '                 End If
                                '                End If
                            End If
                            subrow("dblBrutto") = Decimal.Round(subrow("dblNetto") + subrow("dblMwSt"), 4, MidpointRounding.AwayFromZero)
                            If Val(strSteuer(3)) <> 0 Then 'Wurde ein anderer Steuersatz gewählt?
                                subrow("dblMwStSatz") = Val(strSteuer(3))
                                subrow("strMwStKey") = strSteuer(0)
                            End If

                        End If
                    Else
                        subrow("strMwStKey") = "n/a"
                    End If
                Else
                    subrow("strMwStKey") = "null"
                    subrow("dblMwst") = 0
                    intReturnValue = 0

                End If
                strBitLog += Trim(intReturnValue.ToString)

                'If subrow("intSollHaben") <> 2 Then
                intSubNumber += 1
                If subrow("intSollHaben") = 0 Then
                    dblSubNetto += IIf(IsDBNull(subrow("dblNetto")), 0, subrow("dblNetto"))
                    dblSubMwSt += IIf(IsDBNull(subrow("dblMwSt")), 0, subrow("dblMwSt"))
                    dblSubBrutto += IIf(IsDBNull(subrow("dblBrutto")), 0, subrow("dblBrutto"))
                    If Strings.Left(subrow("lngKto").ToString, 1) = "1" Then
                        intSollKonto = subrow("lngKto") 'Für Sollkonto - Korretkur
                    End If
                Else
                    dblSubNetto -= IIf(IsDBNull(subrow("dblNetto")), 0, subrow("dblNetto"))
                    subrow("dblNetto") = subrow("dblNetto") * -1
                    dblSubMwSt -= IIf(IsDBNull(subrow("dblMwSt")), 0, subrow("dblMwSt"))
                    subrow("dblMwSt") = subrow("dblMwSt") * -1
                    dblSubBrutto -= IIf(IsDBNull(subrow("dblBrutto")), 0, subrow("dblBrutto"))
                    subrow("dblBrutto") = subrow("dblBrutto") * -1
                End If

                'Runden
                dblSubNetto = Decimal.Round(dblSubNetto, 4, MidpointRounding.AwayFromZero)
                dblSubMwSt = Decimal.Round(dblSubMwSt, 4, MidpointRounding.AwayFromZero)
                dblSubBrutto = Decimal.Round(dblSubBrutto, 4, MidpointRounding.AwayFromZero)

                'Konto prüfen 02
                If IIf(IsDBNull(subrow("lngKto")), 0, subrow("lngKto")) > 0 Then
                    'Falls KSt nicht gültig, entfernen
                    If CInt(Strings.Left(subrow("lngKto").ToString, 1)) < 3 Then
                        subrow("lngKST") = 0
                        subrow("strKtoBez") = "K<3KST ->"
                    End If
                    intReturnValue = FcCheckKonto(subrow("lngKto"),
                                                  IIf(IsDBNull(subrow("dblMwSt")), 0, subrow("dblMwSt")),
                                                  IIf(IsDBNull(subrow("lngKST")), 0, subrow("lngKST")),
                                                  False)
                    If intReturnValue = 0 Then
                        subrow("strKtoBez") += FcReadDebitorKName(subrow("lngKto"))
                    ElseIf intReturnValue = 2 Then
                        subrow("strKtoBez") += FcReadDebitorKName(subrow("lngKto")) + " MwSt!"
                    ElseIf intReturnValue = 3 Then
                        subrow("strKtoBez") += FcReadDebitorKName(subrow("lngKto")) + " NoKST"
                        'ElseIf intReturnValue = 4 Then
                        '    subrow("strKtoBez") = FcReadDebitorKName(objFiBhg, subrow("lngKto")) + " K<3KST"
                        '    subrow("lngKST") = 0
                        '    intReturnValue = 0
                    ElseIf intReturnValue = 5 Then
                        subrow("strKtoBez") += FcReadDebitorKName(subrow("lngKto")) + " K<3MwSt"
                    Else
                        subrow("strKtoBez") = "n/a"

                    End If
                Else
                    subrow("strKtoBez") = "null"
                    subrow("lngKto") = 0
                    intReturnValue = 1

                End If
                strBitLog += Trim(intReturnValue.ToString)

                'Kst/Ktr prüfen 03
                If IIf(IsDBNull(subrow("lngKST")), 0, subrow("lngKST")) > 0 Then
                    intReturnValue = FcCheckKstKtr(IIf(IsDBNull(subrow("lngKST")), 0, subrow("lngKST")),
                                                   subrow("lngKto"),
                                                   strKstKtrSage200)
                    If intReturnValue = 0 Then
                        subrow("strKstBez") = strKstKtrSage200
                    ElseIf intReturnValue = 1 Then
                        subrow("strKstBez") = "KoArt"
                    ElseIf intReturnValue = 3 Then
                        subrow("strKstBez") = "NoKST"
                        subrow("lngKST") = 0
                        intReturnValue = 0 'Kein Fehler
                    Else
                        subrow("strKstBez") = "n/a"

                    End If
                Else
                    subrow("lngKST") = 0
                    subrow("strKstBez") = "null"
                    intReturnValue = 0

                End If
                strBitLog += Trim(intReturnValue.ToString)

                'Projekt prüfen 04
                If IIf(IsDBNull(subrow("lngProj")), 0, subrow("lngProj")) > 0 Then
                    intReturnValue = FcCheckProj(subrow("lngProj"),
                                                 strProjDesc)
                    If intReturnValue = 0 Then
                        subrow("strProjBez") = strProjDesc
                    ElseIf intReturnValue = 9 Then
                        subrow("strProjBez") = "n/a"
                    End If

                Else
                    subrow("lngProj") = 0
                    subrow("strProjBez") = "null"
                    intReturnValue = 0
                End If
                strBitLog += Trim(intReturnValue.ToString)

                ''MwSt prüfen
                'If Not IsDBNull(subrow("strMwStKey")) Then
                '    intReturnValue = FcCheckMwSt(objdbconn, objFiBhg, subrow("strMwStKey"), subrow("lngMwStSatz"), strStrStCodeSage200)
                '    If intReturnValue = 0 Then
                '        subrow("strMwStKey") = strStrStCodeSage200
                '        'Check of korrekt berechnet
                '        strSteuer = Split(objFiBhg.GetSteuerfeld(subrow("lngKto").ToString, "Zum Rechnen", subrow("dblBrutto").ToString, strStrStCodeSage200), "{<}")
                '        If Val(strSteuer(2)) <> subrow("dblMwst") Then
                '            'Im Fall von Auto-Korrekt anpassen
                '            Stop
                '        End If
                '    Else
                '        subrow("strMwStKey") = "n/a"

                '    End If
                'Else
                '    subrow("strMwStKey") = "null"
                '    intReturnValue = 0

                'End If
                'strBitLog += Trim(intReturnValue.ToString)

                'Brutto + MwSt + Netto = 0; 05
                If IIf(IsDBNull(subrow("dblBrutto")), 0, subrow("dblBrutto")) = 0 And IIf(IsDBNull(subrow("dblMwSt")), 0, subrow("dblMwSt")) = 0 And IIf(IsDBNull(subrow("dblNetto")), 0, subrow("dblNetto")) Then
                    strBitLog += "1"

                Else
                    strBitLog += "0"
                End If

                'Netto = 0; 06
                If (IIf(IsDBNull(subrow("dblNetto")), 0, subrow("dblNetto")) = 0 And Not booSplittBill) Or (IIf(IsDBNull(subrow("dblNetto")), 0, subrow("dblNetto")) = 0 And Not booLinkedGS) Then
                    strBitLog += "1"

                Else
                    strBitLog += "0"
                End If

                'Brutto = 0; 07
                If IIf(IsDBNull(subrow("dblBrutto")), 0, subrow("dblBrutto")) = 0 Then
                    strBitLog += "1"

                Else
                    strBitLog += "0"
                End If

                'Brutto - MwSt <> Netto; 08
                If Decimal.Round(subrow("dblNetto") + subrow("dblMwSt"), 4, MidpointRounding.AwayFromZero) <> subrow("dblBrutto") Then
                    strBitLog += "1"

                Else
                    strBitLog += "0"
                End If


                'Statustext zusammen setzten
                'strStatusText = ""
                'MwSt
                If Strings.Left(strBitLog, 1) <> "0" Then
                    strStatusText += IIf(strStatusText <> "", ", ", "") + "MwSt"
                End If
                'Konto
                If Mid(strBitLog, 2, 1) <> "0" Then
                    If Mid(strBitLog, 2, 1) = "2" Then
                        strStatusText = "Kto MwSt"
                    ElseIf Mid(strBitLog, 2, 1) = "5" Then
                        strStatusText = "MwstK<3K"
                    ElseIf Mid(strBitLog, 2, 1) = "3" Then
                        strStatusText = "NoKST"
                    Else
                        strStatusText = "Kto"
                    End If
                End If
                'Kst/Ktr
                If Mid(strBitLog, 3, 1) <> "0" Then
                    strStatusText += IIf(strStatusText <> "", ", ", "") + "KST"
                End If
                'Projekt 
                If Mid(strBitLog, 4, 1) <> "0" Then
                    strStatusText += IIf(strStatusText <> "", ", ", "") + "Proj"
                End If
                'Alles 0
                If Mid(strBitLog, 5, 1) <> "0" Then
                    strStatusText += IIf(strStatusText <> "", ", ", "") + "All0"
                End If
                'Netto 0
                If Mid(strBitLog, 6, 1) <> "0" Then
                    strStatusText += IIf(strStatusText <> "", ", ", "") + "Nett0"
                End If
                'Brutto 0
                If Mid(strBitLog, 7, 1) <> "0" Then
                    strStatusText += IIf(strStatusText <> "", ", ", "") + "Brut0"
                End If
                'Diff 0
                If Mid(strBitLog, 8, 1) <> "0" Then
                    strStatusText += IIf(strStatusText <> "", ", ", "") + "Diff"
                End If

                If Val(strBitLog) = 0 Then
                    strStatusText += " ok"
                End If

                'BitLog und Text schreiben
                subrow("strStatusUBBitLog") = strBitLog
                subrow("strStatusUBText") = strStatusText
                strStatusText = String.Empty

                strStatusOverAll = strStatusOverAll Or strBitLog
                'Application.DoEvents()

            Next

            'Falls Soll-Konto-Korretkur gesetzt hier Konto ändern
            If booCashSollKorrekt And intBuchungsArt = 4 Then
                lngDebKonto = intSollKonto
                'Debug.Print("Konto in SB geändert")
            End If

            'Rückgabe der ganzen Funktion Sub-Prüfung
            If intSubNumber = 0 Then 'keine Subs
                Return 9
            Else
                If Val(strStatusOverAll) > 0 Then
                    Return 1
                Else
                    Return 0
                    'If intBuchungsArt = 1 Then
                    '    'OP - Buchung
                    '    'If dblSubNetto <> 0 Or dblSubBrutto <> 0 Or dblSubMwSt <> 0 Then 'Diff
                    '    'Return 2
                    '    'Else
                    '    Return 0
                    '    'End If
                    'Else
                    '    'Belegsbuchung 'Nur Brutto 0 - Test
                    '    If dblSubBrutto <> 0 Then
                    '        Return 3
                    '    Else
                    '        Return 0
                    '    End If
                    'End If
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Fehler Debi-Subbuchungen", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        Finally
            selsubrow = Nothing
            'objDtDebiSub.AcceptChanges()

        End Try

    End Function

    Friend Function FcCheckMwSt(strStrCode As String,
                                ByRef dblStrWert As Double,
                                ByRef strStrCode200 As String,
                                intKonto As Int32) As Integer

        'returns 0=ok, 1=nicht gefunden

        Dim objdbconn As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))
        Dim objlocdtMwSt As New DataTable("tbllocMwSt")
        Dim objlocMySQLcmd As New MySqlCommand
        Dim strSteuerRec As String = String.Empty
        'Dim strSteuerRecAr() As String
        Dim intLooper As Int16 = 0

        Try

            'Falls MwStKey 'ohne' und Konto >= 3000 und 3999 dann ohne = frei
            If strStrCode = "ohne" Then
                If intKonto >= 3000 And intKonto <= 3999 Then
                    strStrCode = "frei"
                End If
            ElseIf strStrCode = "null" Then
                strStrCode200 = "00"
                Return 0
            End If

            'Besprechung mit Muhi 20201209 => Es soll eine fixe Vergabe des MStSchlüssels passieren 
            objlocMySQLcmd.CommandText = "SELECT  * FROM t_sage_sage50mwst WHERE strKey='" + strStrCode + "'"

            objdbconn.Open()
            objlocMySQLcmd.Connection = objdbconn
            objlocdtMwSt.Load(objlocMySQLcmd.ExecuteReader)

            If objlocdtMwSt.Rows.Count = 0 Then
                MessageBox.Show("MwSt " + strStrCode + " ist nicht definiert für Sage 50 MsSt-Key " + strStrCode + ".", "MwSt-Key Check S50 " + strStrCode)
                Return 1
            Else
                'Wert von Tabelle übergeben
                If Not IsDBNull(objlocdtMwSt.Rows(0).Item("intSage200Key")) Then
                    strStrCode200 = objlocdtMwSt.Rows(0).Item("intSage200Key")
                    'Evtl falsch gesetzte MwSt-Satz korrigieren
                    If objlocdtMwSt.Rows(0).Item("dblProzent") <> dblStrWert Then
                        dblStrWert = objlocdtMwSt.Rows(0).Item("dblProzent")
                    End If
                    Return 0
                Else
                    strStrCode200 = "00"
                    Return 2
                End If

            End If

            'Besprechung mit Muhi 20201209 => Es soll eine fixe Vergabe des MStSchlüssels passieren 
            'objlocMySQLcmd.CommandText = "SELECT  * FROM t_sage_sage50mwst WHERE strKey='" + strStrCode + "' AND dblProzent=" + dblStrWert.ToString

            'objlocMySQLcmd.Connection = objdbconn
            'objlocdtMwSt.Load(objlocMySQLcmd.ExecuteReader)

            'If objlocdtMwSt.Rows.Count = 0 Then
            '    MessageBox.Show("MwSt " + strStrCode + " ist nicht definiert für " + dblStrWert.ToString + ".")
            '    Return 1
            'Else
            '    'In Sage 200 suchen
            '    Do Until strSteuerRec = "EOF"
            '        strSteuerRec = objFiBhg.GetStIDListe(intLooper)
            '        If strSteuerRec <> "EOF" Then
            '            strSteuerRecAr = Split(strSteuerRec, "{>}")
            '            'Gefunden?
            '            If strSteuerRecAr(3) = dblStrWert And strSteuerRecAr(6) = objlocdtMwSt.Rows(0).Item("strBruttoNetto") And strSteuerRecAr(7) = objlocdtMwSt.Rows(0).Item("strGegenKonto") Then
            '                'Debug.Print("Found " + strSteuerRecAr(0).ToString)
            '                strStrCode200 = strSteuerRecAr(0)
            '                Return 0
            '            End If
            '        Else
            '            Return 1
            '        End If
            '        intLooper += 1
            '    Loop
            'End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "MwSt-Key Check")
            Return 9

        Finally
            objdbconn.Close()
            objdbconn = Nothing
            objlocdtMwSt = Nothing
            objlocMySQLcmd = Nothing

        End Try


    End Function

    Friend Function FcCheckKonto(lngKtoNbr As Long,
                                 dblMwSt As Double,
                                 lngKST As Int32,
                                 booExistanceOnly As Boolean) As Integer

        'Returns 0=ok, 1=existiert nicht, 2=existiert aber keine KST erlaubt, 3=KST nicht auf Konto definiert, 4=KST auf Konto > 3

        Dim strReturn As String
        Dim strKontoInfo() As String

        Try

            'If lngKtoNbr = 1173 Then Stop

            strReturn = objfiBuha.GetKontoInfo(lngKtoNbr.ToString)
            If strReturn = "EOF" Then
                Return 1
            Else
                'If dblMwSt = 0 Then
                'Return 0
                strKontoInfo = Split(objfiBuha.GetKontoInfo(lngKtoNbr.ToString), "{>}")
                If booExistanceOnly Then
                    Return 0
                End If
                'KST?
                'strKontoInfo = Split(objfiBuha.GetKontoInfo(lngKtoNbr.ToString), "{>}")
                If lngKST > 0 Then
                    If CInt(Strings.Left(lngKtoNbr.ToString, 1)) >= 3 Then
                        'strKontoInfo = Split(objfiBuha.GetKontoInfo(lngKtoNbr.ToString), "{>}")
                        If strKontoInfo(22) = "" Then
                            Return 3
                        Else
                            If dblMwSt <> 0 Then
                                If strKontoInfo(26) = "" Then
                                    'Gemäss Andy 5.12.2023 falsch
                                    'Return 5
                                    Return 0
                                Else
                                    Return 0
                                End If
                            Else
                                Return 0
                            End If

                            'Else
                            'Steuerpflichtig?
                            'strKontoInfo = Split(objfiBuha.GetKontoInfo(lngKtoNbr.ToString), "{>}")
                            'If strKontoInfo(26) = "" Then
                            'Return 2
                            'Else
                            'Return 0
                            'End If
                            'End If
                        End If
                    Else
                        Return 4
                    End If
                Else
                    'Ist keine KST erlaubt?
                    If strKontoInfo(22) <> "" Then
                        Return 3
                    End If
                    If dblMwSt <> 0 Then
                        If strKontoInfo(26) = "" Then
                            Return 5
                        Else
                            Return 0
                        End If
                    Else
                        Return 0
                    End If
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem Konto")
            Return 9

        End Try


    End Function

    Friend Function FcCheckKstKtr(lngKST As Long,
                                  lngKonto As Long,
                                  ByRef strKstKtrSage200 As String) As Int16

        'return 0=ok, 1=Kst existiert kene Kostenart, 2=Kst nicht defniert, 3=nicht auf Konto anwendbar 1000 - 2999

        Dim strReturn As String
        Dim strReturnAr() As String
        Dim booKstKAok As Boolean
        Dim strKst, strKA As String

        booKstKAok = False
        'objFiPI = Nothing
        'objFiPI = objFiBhg.GetCheckObj

        Try
            'If CInt(Left(lngKonto.ToString, 1)) >= 3 Then
            strReturn = objfiBuha.GetKstKtrInfo(lngKST.ToString)
            If strReturn = "EOF" Then
                Return 2
            Else
                strReturnAr = Split(strReturn, "{>}")
                strKstKtrSage200 = strReturnAr(1)
                strKst = Convert.ToString(lngKST)
                strKA = Convert.ToString(lngKonto)
                'Ist Kst auf Kostenbart definiert?
                booKstKAok = objdbPIFb.CheckKstKtr(strKst, strKA)

                If booKstKAok Then
                    Return 0
                Else
                    Return 1
                End If
            End If
            'Else
            'Return 3
            'End If

        Catch ex As Exception
            Return 1

        End Try

    End Function

    Friend Function FcCheckProj(intProj As Int32,
                                ByRef strProjDesc As String) As Int16

        'Returns 0=ok, 9=Problem

        Dim strLine As String
        Dim booFoundProject As Boolean
        Dim strProjectAr() As String

        Try

            booFoundProject = False
            strLine = String.Empty

            Call objFiBebu.ReadProjektTree(0)

            strLine = objFiBebu.GetProjektLine()
            Do While strLine <> "EOF"
                strProjectAr = Split(strLine, "{>}")
                Debug.Print("Aktuelle Line: " + strLine)
                'Projekt gefunden?
                If Val(strProjectAr(0)) = intProj Then
                    booFoundProject = True
                    strProjDesc = strProjectAr(1)
                End If
                strLine = objFiBebu.GetProjektLine()
            Loop

            If booFoundProject Then
                Return 0
            Else
                Return 9
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem Check-Project")

        End Try


    End Function

    Friend Function FcCheckCurrency(strCurrency As String) As Integer

        Dim strReturn As String
        Dim booFoundCurrency As Boolean

        Try

            booFoundCurrency = False
            strReturn = String.Empty

            Call objfiBuha.ReadWhg()

            'If strCurrency = "EUR" Then Stop

            strReturn = objfiBuha.GetWhgZeile()
            Do While strReturn <> "EOF"
                If Strings.Left(strReturn, 3) = strCurrency Then
                    'If strCurrency = "EUR" Then Stop
                    booFoundCurrency = True
                End If
                strReturn = objfiBuha.GetWhgZeile()
                'Application.DoEvents()
            Loop

            If booFoundCurrency Then
                Return 0
            Else
                Return 1
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem Currency")
            Return 9

        End Try

    End Function

    Friend Function FcCheckDebitor(lngDebitor As Long,
                                   intBuchungsart As Integer) As Integer

        Dim strReturn As String

        Try

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

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem Currency")
            Return 9

        End Try

    End Function

    Friend Function FcPGVDTreatment(tblDebiB As DataTable,
                                    strDRGNbr As String,
                                    intDBelegNr As Int32,
                                    strCur As String,
                                    datValuta As Date,
                                    strIType As String,
                                    datPGVStart As Date,
                                    datPGVEnd As Date,
                                    intITotal As Int16,
                                    intITY As Int16,
                                    intINY As Int16,
                                    intAcctTY As Int16,
                                    intAcctNY As Int16,
                                    strPeriode As String,
                                    objdbcon As MySqlConnection,
                                    objsqlcon As SqlConnection,
                                    objsqlcmd As SqlCommand,
                                    intAccounting As Int16,
                                    ByRef objdtInfo As DataTable,
                                    ByRef strYear As String,
                                    ByRef intTeqNbr As Int16,
                                    ByRef intTeqNbrLY As Int16,
                                    ByRef intTeqNbrPLY As Int16,
                                    ByRef strPGVType As String,
                                    ByRef datPeriodFrom As Date,
                                    ByRef datPeriodTo As Date,
                                    ByRef strPeriodStatus As String) As Int16

        Dim dblNettoBetrag As Double
        Dim intSollKonto As Int16
        Dim strBelegDatum As String
        Dim strDebiTextSoll As String
        Dim strDebiCurrency As String
        Dim dblKursD As Double
        Dim strSteuerFeldSoll As String
        Dim intHabenKonto As Int16
        Dim strDebiTextHaben As String
        Dim dblKursH As Double
        Dim strValutaDatum As String
        Dim strSteuerFeldHaben As String
        Dim drDebiSub() As DataRow
        Dim strBebuEintragSoll As String
        Dim strBebuEintragHaben As String
        Dim strPeriodenInfoA() As String
        Dim strPeriodenInfo As String
        Dim intReturnValue As Int32
        Dim strActualYear As String
        Dim datPGVEndSave As Date
        Dim datValutaSave As Date
        Dim strLogonInfo() As String

        Dim objFinanzCopy As New SBSXASLib.AXFinanz
        Dim objfiBuhaCopy As New SBSXASLib.AXiFBhg

        Try

            objFinanzCopy = objFinanz.DuplicateObjekt(2)
            objfiBuhaCopy = objFinanzCopy.GetFibuObj()


            'Jahr retten
            strActualYear = strYear
            'Zuerst betroffene Buchungen selektieren
            drDebiSub = tblDebiB.Select("strRGNr='" + strDRGNbr + "' AND dblNetto<>0")

            'Durch die Buchungen steppen
            For Each drDSubrow As DataRow In drDebiSub
                'Auflösung
                '=========

                datValuta = datValutaSave
                If strPGVType = "RV" Then
                    datPGVStart = frmImportMain.strNexY + "-01-01"
                End If

                'Evtl. Aufteilen auf 2 Jahre
                For intYearLooper As Int16 = 0 To Year(DateAdd(DateInterval.Month, intITotal, datPGVStart)) - Year(datPGVStart)

                    If intYearLooper = 0 And intITotal > 1 Then
                        dblNettoBetrag = drDSubrow("dblNetto") * -1 / intITotal * intITY
                        intHabenKonto = intAcctTY
                    Else
                        dblNettoBetrag = drDSubrow("dblNetto") * -1 / intITotal * intINY
                        intHabenKonto = intAcctNY
                    End If

                    If dblNettoBetrag <> 0 Then 'Falls in einem Jahr nichts zu buchen ist

                        strBelegDatum = Format(datValuta, "yyyyMMdd").ToString

                        strDebiTextHaben = drDSubrow("strDebSubText") + ", PGV Auflösung"

                        strSteuerFeldHaben = "STEUERFREI"

                        intSollKonto = drDSubrow("lngKto")

                        strDebiTextSoll = drDSubrow("strDebSubText") + ", PGV Auflösung"

                        strSteuerFeldSoll = "STEUERFREI"
                        strValutaDatum = Format(datValuta, "yyyyMMdd").ToString

                        'Falls nicht CHF dann umrechnen und auf CHF setzen
                        If strCur <> "CHF" Then
                            dblKursD = FcGetKurs(strCur,
                                                 strValutaDatum,
                                                 drDSubrow("lngKto"))
                            strDebiCurrency = "CHF"
                        Else
                            dblKursD = 1.0#
                            strDebiCurrency = strCur
                        End If
                        dblKursH = dblKursD

                        'KORE
                        If drDSubrow("lngKST") > 0 Then

                            If drDSubrow("intSollHaben") = 0 Then 'Soll
                                strBebuEintragHaben = drDSubrow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
                                strBebuEintragSoll = Nothing
                            Else
                                strBebuEintragHaben = Nothing
                                strBebuEintragSoll = drDSubrow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
                            End If
                        Else
                            strBebuEintragHaben = Nothing
                            strBebuEintragSoll = Nothing

                        End If

                        'Buchen
                        Call objfiBuhaCopy.WriteBuchung(0,
                           intDBelegNr,
                           strBelegDatum,
                           intSollKonto.ToString,
                           strDebiTextSoll,
                           strDebiCurrency,
                           dblKursD.ToString,
                           (dblNettoBetrag * dblKursD).ToString,
                           strSteuerFeldSoll,
                           intHabenKonto.ToString,
                           strDebiTextHaben,
                           strDebiCurrency,
                           dblKursH.ToString,
                           (dblNettoBetrag * dblKursH).ToString,
                           strSteuerFeldHaben,
                           strDebiCurrency,
                           dblKursH.ToString,
                           (dblNettoBetrag * dblKursH).ToString,
                           dblNettoBetrag.ToString,
                           strBebuEintragSoll,
                           strBebuEintragHaben,
                           strValutaDatum)

                    End If

                Next

                'Falls FY dann 2312 auf 2311
                'Gab es eine Neutralisierung fürs FJ?
                If intINY > 0 And intITotal > 1 Then
                    'Was ist die aktuelle angemeldete Periode ?
                    'strPeriodenInfo = objFinanz.GetPeriListe(0)
                    'strPeriodenInfoA = Split(strPeriodenInfo, "{>}")
                    strLogonInfo = Split(objFinanzCopy.GetLogonInfo(), "{>}")

                    'Ist aktuell angemeldete Periode = FJ
                    If Year(datPGVEnd) <> Val(Strings.Left(strPeriodenInfo, 4)) Then
                        'Zuerst Info-Table löschen
                        'objdtInfo.Clear()
                        'Application.DoEvents()
                        'Login ins FJ
                        intReturnValue = FcLoginSage2(objdbcon,
                                                      objsqlcon,
                                                      objsqlcmd,
                                                      objFinanzCopy,
                                                      objfiBuhaCopy,
                                                      intAccounting,
                                                      Year(datPGVEnd).ToString,
                                                      strActualYear)

                        'Application.DoEvents()

                        '2311 -> 2312
                        datValuta = frmImportMain.strNexY + "-01-01" 'Achtung provisorisch
                        strValutaDatum = Format(datValuta, "yyyyMMdd").ToString
                        strBelegDatum = strValutaDatum
                        intHabenKonto = intAcctTY
                        intSollKonto = intAcctNY
                        strDebiTextHaben = drDSubrow("strDebSubText") + ", PGV AJ / FJ"
                        strDebiTextSoll = drDSubrow("strDebSubText") + ", PGV AJ / FJ"
                        strBebuEintragHaben = Nothing
                        strBebuEintragSoll = Nothing

                        'Buchen
                        Call objfiBuhaCopy.WriteBuchung(0,
                           intDBelegNr,
                           strBelegDatum,
                           intSollKonto.ToString,
                           strDebiTextSoll,
                           strDebiCurrency,
                           dblKursD.ToString,
                           (dblNettoBetrag * dblKursD).ToString,
                           strSteuerFeldSoll,
                           intHabenKonto.ToString,
                           strDebiTextHaben,
                           strDebiCurrency,
                           dblKursH.ToString,
                           (dblNettoBetrag * dblKursH).ToString,
                           strSteuerFeldHaben,
                           strDebiCurrency,
                           dblKursH.ToString,
                           (dblNettoBetrag * dblKursH).ToString,
                           dblNettoBetrag.ToString,
                           strBebuEintragSoll,
                           strBebuEintragHaben,
                           strValutaDatum)


                    End If

                End If

                'Einzelene Monate buchen
                For intMonthLooper As Int16 = 0 To intITotal - 1
                    datValuta = DateAdd(DateInterval.Month, intMonthLooper, datPGVStart)
                    strValutaDatum = Format(datValuta, "yyyyMMdd").ToString
                    strBelegDatum = strValutaDatum
                    intHabenKonto = drDSubrow("lngKto")
                    strDebiTextHaben = drDSubrow("strDebSubText") + ", PGV M " + (intMonthLooper + 1).ToString + "/ " + intITotal.ToString
                    dblNettoBetrag = drDSubrow("dblNetto") * -1 / intITotal
                    If intITotal = 1 Then
                        intSollKonto = intAcctNY
                    Else
                        intSollKonto = intAcctTY
                    End If

                    strDebiTextSoll = strDebiTextHaben

                    If drDSubrow("intSollHaben") = 0 Then 'Haben
                        strBebuEintragHaben = Nothing
                        strBebuEintragSoll = drDSubrow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
                    Else
                        strBebuEintragHaben = drDSubrow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
                        strBebuEintragSoll = Nothing
                    End If

                    If Year(datValuta) = frmImportMain.intCurY And Year(datValuta) <> Val(strActualYear) Then 'Achtung provisorisch
                        'Zuerst Info-Table löschen
                        'objdtInfo.Clear()
                        'Application.DoEvents()
                        'Im 2023 anmelden
                        intReturnValue = FcLoginSage2(objdbcon,
                                                      objsqlcon,
                                                      objsqlcmd,
                                                      objFinanzCopy,
                                                      objfiBuhaCopy,
                                                      intAccounting,
                                                      frmImportMain.strCurY,
                                                      strActualYear)
                        'Application.DoEvents()

                    ElseIf Year(datValuta) = frmImportMain.intNexY And Year(datValuta) <> Val(strActualYear) Then
                        'Zuerst Info-Table löschen
                        'objdtInfo.Clear()
                        'Application.DoEvents()
                        'Im 2023 anmelden
                        intReturnValue = FcLoginSage2(objdbcon,
                                                      objsqlcon,
                                                      objsqlcmd,
                                                      objFinanzCopy,
                                                      objfiBuhaCopy,
                                                      intAccounting,
                                                      frmImportMain.strNexY,
                                                      strActualYear)
                        'Application.DoEvents()

                    End If

                    'Buchen
                    Call objfiBuhaCopy.WriteBuchung(0,
                           intDBelegNr,
                           strBelegDatum,
                           intSollKonto.ToString,
                           strDebiTextSoll,
                           strDebiCurrency,
                           dblKursD.ToString,
                           (dblNettoBetrag * dblKursD).ToString,
                           strSteuerFeldSoll,
                           intHabenKonto.ToString,
                           strDebiTextHaben,
                           strDebiCurrency,
                           dblKursH.ToString,
                           (dblNettoBetrag * dblKursH).ToString,
                           strSteuerFeldHaben,
                           strDebiCurrency,
                           dblKursH.ToString,
                           (dblNettoBetrag * dblKursH).ToString,
                           dblNettoBetrag.ToString,
                           strBebuEintragSoll,
                           strBebuEintragHaben,
                           strValutaDatum)

                Next

            Next
            'Für weitere Buchungen ins ursprüngliche Jahr anmelden 
            'If strYear <> strActualYear Then
            '    'Zuerst Info-Table löschen
            '    objdtInfo.Clear()
            '    'Application.DoEvents()
            '    'Im Aufrufjahr anmelden
            '    intReturnValue = FcLoginSage2(objdbcon,
            '                                  objsqlcon,
            '                                  objsqlcmd,
            '                                  objFinanz,
            '                                  objFBhg,
            '                                  objDbBhg,
            '                                  objPiFin,
            '                                  objBebu,
            '                                  objKrBhg,
            '                                  intAccounting,
            '                                  objdtInfo,
            '                                  strActualYear,
            '                                  strYear,
            '                                  intTeqNbr,
            '                                  intTeqNbrLY,
            '                                  intTeqNbrPLY,
            '                                  datPeriodFrom,
            '                                  datPeriodTo,
            '                                  strPeriodStatus)
            '    'Application.DoEvents()
            'End If
            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem PGV - Buchung Debitoren", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 9

        Finally
            drDebiSub = Nothing
            strPeriodenInfoA = Nothing

        End Try

    End Function

    Friend Function FcLoginSage2(ByRef objdbconn As MySqlConnection,
                                       ByRef objsqlConn As SqlClient.SqlConnection,
                                       ByRef objsqlCom As SqlClient.SqlCommand,
                                       ByRef objFinanzCopy As SBSXASLib.AXFinanz,
                                       ByRef objfiBuhaCopy As SBSXASLib.AXiFBhg,
                                       ByVal intAccounting As Int16,
                                       ByVal strPeriod As String,
                                       ByRef strYear As String) As Int16

        '0=ok, 1=Fibu nicht ok, 2=Debi nicht ok, 3=Debi nicht ok
        Dim booAccOk As Boolean
        Dim strMandant As String
        Dim strLogonInfo() As String
        Dim strPeriode() As String
        Dim FcReturns As Int16
        Dim intPeriodenNr As Int16
        Dim strPeriodenInfo As String
        Dim objdtPeriodeLY As New DataTable
        Dim strPeriodeLY As String
        Dim strPeriodePLY As String
        Dim objdbcmd As New MySqlCommand
        Dim dtPeriods As New DataTable


        Try

            objfiBuhaCopy = Nothing
            objFinanzCopy = Nothing
            objFinanzCopy = New SBSXASLib.AXFinanz

            'Application.DoEvents()

            'Login
            Call objFinanzCopy.ConnectSBSdb(System.Configuration.ConfigurationManager.AppSettings("OwnSageServer"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageDB"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageID"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSagePsw"), "")

            'objdbconn.Open()
            strMandant = FcReadFromSettingsII("Buchh200_Name",
                                            intAccounting)
            'objdbconn.Close()
            booAccOk = objFinanz.CheckMandant(strMandant)

            'Open Mandantg
            objFinanzCopy.OpenMandant(strMandant, strPeriod)

            'Von Login aktuelle Periode auslesen
            strLogonInfo = Split(objFinanzCopy.GetLogonInfo(), "{>}")
            'objdtInfo.Rows.Add("Man/Periode", strMandant + "/" + strLogonInfo(7) + ", " + intAccounting.ToString)

            'Check Periode
            intPeriodenNr = objFinanzCopy.ReadPeri(strMandant, strLogonInfo(7))
            strPeriodenInfo = objFinanzCopy.GetPeriListe(0)

            strPeriode = Split(strPeriodenInfo, "{>}")

            ''Teq-Nr von Vorjar lesen um in Suche nutzen zu können
            'objdtPeriodeLY.Rows.Clear()
            'strPeriodeLY = (Val(Strings.Left(strPeriode(4), 4)) - 1).ToString + Strings.Right(strPeriode(4), 4)
            'objsqlCom.CommandText = "SELECT teqnbr FROM periode WHERE mandid='" + strMandant + "' AND dtebis='" + strPeriodeLY + "'"
            'objsqlCom.Connection = objsqlConn
            'objsqlConn.Open()
            'objdtPeriodeLY.Load(objsqlCom.ExecuteReader)
            'objsqlConn.Close()
            'If objdtPeriodeLY.Rows.Count > 0 Then
            '    intTeqNbrLY = objdtPeriodeLY.Rows(0).Item("teqnbr")
            'Else
            '    intTeqNbrLY = 0
            'End If
            ''Teq-Nr vom Vorvorjahr
            'objdtPeriodeLY.Rows.Clear()
            'strPeriodePLY = (Val(Strings.Left(strPeriode(4), 4)) - 2).ToString + Strings.Right(strPeriode(4), 4)
            'objsqlCom.CommandText = "SELECT teqnbr FROM periode WHERE mandid='" + strMandant + "' AND dtebis='" + strPeriodePLY + "'"
            'objsqlCom.Connection = objsqlConn
            'objsqlConn.Open()
            'objdtPeriodeLY.Load(objsqlCom.ExecuteReader)
            'objsqlConn.Close()
            'If objdtPeriodeLY.Rows.Count > 0 Then
            '    intTeqNbrPLY = objdtPeriodeLY.Rows(0).Item("teqnbr")
            'Else
            '    intTeqNbrPLY = 0
            'End If

            'intTeqNbr = strPeriode(8)
            'objdtInfo.Rows.Add("GeschäftsJ", strPeriode(3) + "-" + strPeriode(4) + ", teq: " + strPeriode(8).ToString + ", " + intTeqNbrLY.ToString + ", " + intTeqNbrPLY.ToString)
            'objdtInfo.Rows.Add("Buchungen/ Status", strPeriode(5) + "-" + strPeriode(6) + "/ " + strPeriode(2))
            strYear = Strings.Left(strPeriode(4), 4)

            'FcReturns = FcReadPeriodenDef(objsqlConn,
            '                          objsqlCom,
            '                          strPeriode(8),
            '                          objdtInfo,
            '                          strYear)

            ''Perioden-Definition vom Tool einlesen
            ''In einer ersten Phase nur erster DS einlesen
            'objdbcmd.Connection = objdbconn
            'objdbconn.Open()
            'objdbcmd.CommandText = "SELECT * FROM t_sage_buchhaltungen_periods WHERE year=" + strYear + " AND refMandant=" + intAccounting.ToString
            'dtPeriods.Load(objdbcmd.ExecuteReader)
            'objdbconn.Close()
            'If dtPeriods.Rows.Count > 0 Then
            '    datPeriodFrom = dtPeriods.Rows(0).Item("periodFrom")
            '    datPeriodTo = dtPeriods.Rows(0).Item("periodTo")
            '    strPeriodStatus = dtPeriods.Rows(0).Item("status")
            'Else
            '    datPeriodFrom = Convert.ToDateTime(strYear + "-01-01 00:00:01")
            '    datPeriodTo = Convert.ToDateTime(strYear + "-12-31 23:59:59")
            '    strPeriodStatus = "O"
            'End If
            'objdtInfo.Rows.Add("Perioden", Format(datPeriodFrom, "dd.MM.yyyy hh:mm:ss") + " - " + Format(datPeriodTo, "dd.MM.yyyy hh:mm:ss") + "/ " + strPeriodStatus)

            ''Finanz Buha öffnen
            'If Not IsNothing(objfiBuha) Then
            '    objfiBuha = Nothing
            'End If
            'objfiBuha = New SBSXASLib.AXiFBhg
            objfiBuhaCopy = objFinanzCopy.GetFibuObj()
            'Debitor öffnen
            'If Not IsNothing(objdbBuha) Then
            '    objdbBuha = Nothing
            'End If
            'objdbBuha = New SBSXASLib.AXiDbBhg
            'objdbBuha = objFinanz.GetDebiObj()
            'If Not IsNothing(objdbPIFb) Then
            '    objdbPIFb = Nothing
            'End If
            'objdbPIFb = New SBSXASLib.AXiPlFin
            'objdbPIFb = objfiBuha.GetCheckObj()
            'If Not IsNothing(objFiBebu) Then
            '    objFiBebu = Nothing
            'End If
            'objFiBebu = New SBSXASLib.AXiBeBu
            'objFiBebu = objFinanz.GetBeBuObj()
            ''Kreditor
            'If Not IsNothing(objKrBuha) Then
            '    objKrBuha = Nothing
            'End If
            'objKrBuha = New SBSXASLib.AXiKrBhg
            'objKrBuha = objFinanz.GetKrediObj

            'Application.DoEvents()

        Catch ex As Exception
            MsgBox("OpenMandant:" + vbCrLf + "Error" + vbCrLf + "Error # " + Str(Err.Number) + " was generated by " + Err.Source + vbCrLf + Err.Description + " Fehlernummer" & Str(Err.Number And 65535))
            Err.Clear()

        Finally
            objdtPeriodeLY = Nothing
            dtPeriods = Nothing

        End Try

    End Function

    Friend Function FcReadPeriodenDef(ByRef objSQLConnection As SqlClient.SqlConnection,
                                      ByRef objSQLCommand As SqlClient.SqlCommand,
                                      intPeriodenNr As Int32,
                                      ByRef objdtInfo As DataTable,
                                      strYear As String) As Int16

        'Returns 0=definiert, 1=nicht defeniert, 9=Problem
        Dim objlocdtPeriDef As New DataTable
        Dim strPeriodenDef(4) As String


        Try

            objSQLConnection.Open()
            objSQLCommand.CommandText = "SELECT * FROM peridef WHERE teqnbr=" + intPeriodenNr.ToString
            objSQLCommand.Connection = objSQLConnection
            objlocdtPeriDef.Load(objSQLCommand.ExecuteReader)

            If objlocdtPeriDef.Rows.Count > 0 Then 'Perioden-Definition vorhanden

                strPeriodenDef(0) = objlocdtPeriDef.Rows(0).Item(2) 'Bezeichnung
                strPeriodenDef(1) = objlocdtPeriDef.Rows(0).Item(3).ToString  'Von
                strPeriodenDef(2) = objlocdtPeriDef.Rows(0).Item(4).ToString  'Bis
                strPeriodenDef(3) = objlocdtPeriDef.Rows(0).Item(5)  'Status

                objdtInfo.Rows.Add("Perioden S200", strPeriodenDef(0))
                objdtInfo.Rows.Add("Von - Bis/ Status", Format(Convert.ToDateTime(strPeriodenDef(1)), "dd.MM.yyyy hh:mm:ss") + " - " + Format(Convert.ToDateTime(strPeriodenDef(2)), "dd.MM.yyyy hh:mm:ss") + "/ " + strPeriodenDef(3))

                Return 0
            Else

                objdtInfo.Rows.Add("Perioden S200", "keine")
                objdtInfo.Rows.Add("Von - Bis/ Status", Format(Convert.ToDateTime("01.01." + strYear + " 00:00:00"), "dd.MM.yyyy hh:mm:ss") + " - " + Format(Convert.ToDateTime("31.12." + strYear + " 23:59:59"), "dd.MM.yyyy hh:mm:ss") + "/ " + "O")

                Return 1

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Periodendefinition lesen")
            Return 9

        Finally
            objSQLConnection.Close()
            objlocdtPeriDef.Constraints.Clear()
            objlocdtPeriDef.Clear()
            objlocdtPeriDef.Dispose()
            strPeriodenDef = Nothing

        End Try

    End Function

    Friend Function FcPGVDTreatmentYC(tblDebiB As DataTable,
                                      strDRGNbr As String,
                                      intDBelegNr As Int32,
                                      strCur As String,
                                      datValuta As Date,
                                      strIType As String,
                                      datPGVStart As Date,
                                      datPGVEnd As Date,
                                      intITotal As Int16,
                                      intITY As Int16,
                                      intINY As Int16,
                                      intAcctTY As Int16,
                                      intAcctNY As Int16,
                                      strPeriode As String,
                                      objdbcon As MySqlConnection,
                                      objsqlcon As SqlConnection,
                                      objsqlcmd As SqlCommand,
                                      intAccounting As Int16,
                                      ByRef objdtInfo As DataTable,
                                      ByRef strYear As String,
                                      ByRef intTeqNbr As Int16,
                                      ByRef intTeqNbrLY As Int16,
                                      ByRef intTeqNbrPLY As Int16,
                                      ByRef strPGVType As String,
                                      ByRef datPeriodFrom As Date,
                                      ByRef datPeriodTo As Date,
                                      ByRef strPeriodStatus As String) As Int16

        Dim dblNettoBetrag As Double
        Dim intSollKonto As Int16
        Dim strBelegDatum As String
        Dim strDebiTextSoll As String
        Dim strDebiCurrency As String
        Dim dblKursD As Double
        Dim strSteuerFeldSoll As String
        Dim intHabenKonto As Int16
        Dim strDebiTextHaben As String
        Dim dblKursH As Double
        Dim strValutaDatum As String
        Dim strSteuerFeldHaben As String
        Dim drDebiSub() As DataRow
        Dim strBebuEintragSoll As String
        Dim strBebuEintragHaben As String
        Dim strPeriodenInfoA() As String
        Dim strPeriodenInfo As String
        Dim intReturnValue As Int32
        Dim strActualYear As String
        Dim datPGVEndSave As Date
        Dim datValutaSave As Date
        Dim strLogonInfo() As String

        Dim objFinanzCopy As New SBSXASLib.AXFinanz
        Dim objfiBuhaCopy As New SBSXASLib.AXiFBhg


        Try

            Try
                objFinanzCopy = objFinanz.DuplicateObjekt(2)

            Catch inEx As Exception
                If inEx.HResult <> -2147473602 Then
                    MessageBox.Show(inEx.Message, "Connect to Sage - DB " + Err.Number.ToString)
                    Exit Function
                End If


            End Try

            objfiBuhaCopy = objFinanzCopy.GetFibuObj()

            'Jahr retten
            strActualYear = strYear
            'Valuta saven
            datValutaSave = datValuta
            'Zuerst betroffene Buchungen selektieren
            drDebiSub = tblDebiB.Select("strRGNr='" + strDRGNbr + "' AND dblNetto<>0")

            'Durch die Buchungen steppen
            For Each drDSubrow As DataRow In drDebiSub

                'Auflösung
                '=========
                If intITotal = 1 Then
                    If strPGVType = "VR" Then
                        'Falls VR dann PGVEnd saven
                        datValuta = datValutaSave
                        datPGVEndSave = datPGVEnd
                        datPGVEnd = datValuta
                        intINY = 1
                        intITY = 0
                    ElseIf strPGVType = "RV" Then
                        'Damit die Periodenbuchung auf den ersten gebucht wird.
                        datPGVStart = frmImportMain.strNexY + "-01-01"
                        datValuta = datValutaSave
                        intITY = 1
                        intINY = 0
                        intAcctTY = 1312
                    End If
                End If

                'Evtl. Aufteilen auf 2 Jahre
                For intYearLooper As Int16 = Year(datValuta) To Year(datPGVEnd)

                    If intYearLooper = frmImportMain.intCurY Then
                        dblNettoBetrag = drDSubrow("dblNetto") * -1 / intITotal * intITY
                        intHabenKonto = intAcctTY
                    Else
                        dblNettoBetrag = drDSubrow("dblNetto") * -1 / intITotal * intINY
                        intHabenKonto = intAcctNY
                    End If

                    If dblNettoBetrag <> 0 Then 'Falls in einem Jahr nichts zu buchen ist

                        strBelegDatum = Format(datValuta, "yyyyMMdd").ToString

                        If intITotal = 1 Then
                            If Year(datValuta) = frmImportMain.intCurY Then
                                strDebiTextHaben = drDSubrow("strDebSubText") + ", TA"
                            Else
                                strDebiTextHaben = drDSubrow("strDebSubText") + ", TA Auflösung"
                            End If
                        Else
                            strDebiTextHaben = drDSubrow("strDebSubText") + ", PGV Auflösung"
                        End If

                        strSteuerFeldHaben = "STEUERFREI"

                        intSollKonto = drDSubrow("lngKto")

                        If intITotal = 1 Then
                            strDebiTextSoll = strDebiTextHaben
                            If strPGVType = "VR" Then
                                'Valuta - Datum auf 01.01.24 legen, Achtung provisorisch
                                strValutaDatum = frmImportMain.strNexY + "0101"
                                strBelegDatum = frmImportMain.strNexY + "0101"
                            Else
                                'strValutaDatum = Format(datValuta, "yyyyMMdd").ToString
                                strValutaDatum = Format(datValuta, "yyyyMMdd").ToString
                                strBelegDatum = Format(datValuta, "yyyyMMdd").ToString
                            End If
                        Else
                            strDebiTextSoll = drDSubrow("strDebSubText") + ", PGV Auflösung"
                            strValutaDatum = Format(datValuta, "yyyyMMdd").ToString
                        End If

                        strSteuerFeldSoll = "STEUERFREI"

                        'Falls nicht CHF dann umrechnen und auf CHF setzen
                        If strCur <> "CHF" Then
                            dblKursD = FcGetKurs(strCur,
                                                 strValutaDatum,
                                                 drDSubrow("lngKto"))
                            strDebiCurrency = "CHF"
                        Else
                            dblKursD = 1.0#
                            strDebiCurrency = strCur
                        End If
                        dblKursH = dblKursD

                        'KORE
                        If drDSubrow("lngKST") > 0 Then

                            If drDSubrow("intSollHaben") = 0 Then 'Soll
                                strBebuEintragHaben = drDSubrow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
                                strBebuEintragSoll = Nothing
                            Else
                                strBebuEintragHaben = Nothing
                                strBebuEintragSoll = drDSubrow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
                            End If
                        Else
                            strBebuEintragHaben = Nothing
                            strBebuEintragSoll = Nothing

                        End If

                        If Year(datValuta) = frmImportMain.intCurY And Year(datValuta) <> Val(strActualYear) Then 'Achtung provisorisch
                            'Zuerst Info-Table löschen
                            'objdtInfo.Clear()
                            'Application.DoEvents()
                            'Im 2022 anmelden
                            intReturnValue = FcLoginSage2(objdbcon,
                                                          objsqlcon,
                                                          objsqlcmd,
                                                          objFinanzCopy,
                                                          objfiBuhaCopy,
                                                          intAccounting,
                                                          frmImportMain.strCurY,
                                                          strActualYear)
                            ''Application.DoEvents()

                        ElseIf Year(datValuta) = frmImportMain.intNexY And Year(datValuta) <> Val(strActualYear) Then
                            'Zuerst Info-Table löschen
                            'objdtInfo.Clear()
                            'Application.DoEvents()
                            'Im 2023 anmelden
                            intReturnValue = FcLoginSage2(objdbcon,
                                                          objsqlcon,
                                                          objsqlcmd,
                                                          objFinanzCopy,
                                                          objfiBuhaCopy,
                                                          intAccounting,
                                                          frmImportMain.strNexY,
                                                          strActualYear)
                            'Application.DoEvents()

                        End If

                        'doppelte Beleg-Nummern zulassen in HB
                        objfiBuhaCopy.CheckDoubleIntBelNbr = "N"

                        'Buchen
                        Call objfiBuhaCopy.WriteBuchung(0,
                               intDBelegNr,
                               strBelegDatum,
                               intSollKonto.ToString,
                               strDebiTextSoll,
                               strDebiCurrency,
                               dblKursD.ToString,
                               (dblNettoBetrag * dblKursD).ToString,
                               strSteuerFeldSoll,
                               intHabenKonto.ToString,
                               strDebiTextHaben,
                               strDebiCurrency,
                               dblKursH.ToString,
                               (dblNettoBetrag * dblKursH).ToString,
                               strSteuerFeldHaben,
                               strDebiCurrency,
                               dblKursH.ToString,
                               (dblNettoBetrag * dblKursH).ToString,
                               dblNettoBetrag.ToString,
                               strBebuEintragSoll,
                               strBebuEintragHaben,
                               strValutaDatum)

                    End If

                Next

                If strPGVType = "VR" Then
                    'Falls VR dann PGVEnd zurück
                    datPGVEnd = datPGVEndSave
                End If

                'Falls FY dann 2312 auf 2311
                'Gab es eine Neutralisierung fürs FJ?
                If intINY > 0 And intITotal > 1 Then
                    'Was ist die aktuelle angemeldete Periode ?
                    'strPeriodenInfo = objFinanz.GetPeriListe(0)
                    'strPeriodenInfoA = Split(strPeriodenInfo, "{>}")
                    strLogonInfo = Split(objFinanzCopy.GetLogonInfo(), "{>}")

                    'Ist aktuell angemeldete Periode = FJ
                    If Year(datPGVEnd) <> Val(Strings.Left(strLogonInfo(7), 4)) Then
                        'Zuerst Info-Table löschen
                        'objdtInfo.Clear()
                        'Application.DoEvents()
                        'Login ins FJ
                        intReturnValue = FcLoginSage2(objdbcon,
                                                          objsqlcon,
                                                          objsqlcmd,
                                                          objFinanzCopy,
                                                          objfiBuhaCopy,
                                                          intAccounting,
                                                          Year(datPGVEnd).ToString,
                                                          strActualYear)

                        'Application.DoEvents()

                        '2311 -> 2312
                        datValuta = frmImportMain.strNexY + "-01-01" 'Achtung provisorisch
                        strValutaDatum = Format(datValuta, "yyyyMMdd").ToString
                        strBelegDatum = strValutaDatum
                        intHabenKonto = intAcctTY
                        intSollKonto = intAcctNY
                        If intITotal = 1 Then
                            strDebiTextHaben = drDSubrow("strDebSubText") + ", TA AJ / FJ"
                            strDebiTextSoll = drDSubrow("strDebSubText") + ", TA AJ / FJ"
                        Else
                            strDebiTextHaben = drDSubrow("strDebSubText") + ", PGV AJ / FJ"
                            strDebiTextSoll = drDSubrow("strDebSubText") + ", PGV AJ / FJ"
                        End If
                        strBebuEintragHaben = Nothing
                        strBebuEintragSoll = Nothing

                        'doppelte Beleg-Nummern zulassen in HB
                        objfiBuhaCopy.CheckDoubleIntBelNbr = "N"

                        'Buchen
                        Call objfiBuhaCopy.WriteBuchung(0,
                               intDBelegNr,
                               strBelegDatum,
                               intSollKonto.ToString,
                               strDebiTextSoll,
                               strDebiCurrency,
                               dblKursD.ToString,
                               (dblNettoBetrag * dblKursD).ToString,
                               strSteuerFeldSoll,
                               intHabenKonto.ToString,
                               strDebiTextHaben,
                               strDebiCurrency,
                               dblKursH.ToString,
                               (dblNettoBetrag * dblKursH).ToString,
                               strSteuerFeldHaben,
                               strDebiCurrency,
                               dblKursH.ToString,
                               (dblNettoBetrag * dblKursH).ToString,
                               dblNettoBetrag.ToString,
                               strBebuEintragSoll,
                               strBebuEintragHaben,
                               strValutaDatum)


                    End If

                End If

                'Einzelene Monate buchen
                For intMonthLooper As Int16 = 0 To intITotal - 1
                    datValuta = DateAdd(DateInterval.Month, intMonthLooper, datPGVStart)
                    strValutaDatum = Format(datValuta, "yyyyMMdd").ToString
                    strBelegDatum = strValutaDatum
                    intHabenKonto = drDSubrow("lngKto")
                    If intITotal = 1 Then
                        If Year(datValuta) = frmImportMain.intCurY Then
                            strDebiTextHaben = drDSubrow("strDebSubText") + ", TA"
                        Else
                            strDebiTextHaben = drDSubrow("strDebSubText") + ", TA Auflösung"
                        End If
                    Else
                        strDebiTextHaben = drDSubrow("strDebSubText") + ", PGV M " + (intMonthLooper + 1).ToString + "/ " + intITotal.ToString
                    End If

                    dblNettoBetrag = drDSubrow("dblNetto") * -1 / intITotal
                    If intITotal = 1 Then
                        intSollKonto = intAcctNY
                    Else
                        intSollKonto = intAcctTY
                    End If

                    strDebiTextSoll = strDebiTextHaben

                    If drDSubrow("intSollHaben") = 0 Then 'Haben
                        strBebuEintragHaben = Nothing
                        strBebuEintragSoll = drDSubrow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
                    Else
                        strBebuEintragHaben = drDSubrow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
                        strBebuEintragSoll = Nothing
                    End If

                    If Year(datValuta) = frmImportMain.intCurY And Year(datValuta) <> Val(strActualYear) Then 'Achtung provisorisch
                        'Zuerst Info-Table löschen
                        'objdtInfo.Clear()
                        'Application.DoEvents()
                        'Im 2022 anmelden
                        intReturnValue = FcLoginSage2(objdbcon,
                                                          objsqlcon,
                                                          objsqlcmd,
                                                          objFinanzCopy,
                                                          objfiBuhaCopy,
                                                          intAccounting,
                                                          frmImportMain.strCurY,
                                                          strActualYear)
                        'Application.DoEvents()

                    ElseIf Year(datValuta) = frmImportMain.intNexY And Year(datValuta) <> Val(strActualYear) Then
                        'Zuerst Info-Table löschen
                        'objdtInfo.Clear()
                        'Application.DoEvents()
                        'Im 2023 anmelden
                        intReturnValue = FcLoginSage2(objdbcon,
                                                          objsqlcon,
                                                          objsqlcmd,
                                                          objFinanzCopy,
                                                          objfiBuhaCopy,
                                                          intAccounting,
                                                          frmImportMain.strNexY,
                                                          strActualYear)
                        'Application.DoEvents()


                    End If

                    'doppelte Beleg-Nummern zulassen in HB
                    objfiBuhaCopy.CheckDoubleIntBelNbr = "N"

                    'Buchen
                    Call objfiBuhaCopy.WriteBuchung(0,
                               intDBelegNr,
                               strBelegDatum,
                               intSollKonto.ToString,
                               strDebiTextSoll,
                               strDebiCurrency,
                               dblKursD.ToString,
                               (dblNettoBetrag * dblKursD).ToString,
                               strSteuerFeldSoll,
                               intHabenKonto.ToString,
                               strDebiTextHaben,
                               strDebiCurrency,
                               dblKursH.ToString,
                               (dblNettoBetrag * dblKursH).ToString,
                               strSteuerFeldHaben,
                               strDebiCurrency,
                               dblKursH.ToString,
                               (dblNettoBetrag * dblKursH).ToString,
                               dblNettoBetrag.ToString,
                               strBebuEintragSoll,
                               strBebuEintragHaben,
                               strValutaDatum)

                Next

            Next

            'Für weitere Buchungen ins ursprüngliche Jahr anmelden 
            'If strYear <> strActualYear Then
            '    'Zuerst Info-Table löschen
            '    objdtInfo.Clear()
            '    'Application.DoEvents()
            '    'Im Aufrufjahr anmelden
            '    intReturnValue = FcLoginSage2(objdbcon,
            '                                      objsqlcon,
            '                                      objsqlcmd,
            '                                      objFinanz,
            '                                      objFBhg,
            '                                      objDbBhg,
            '                                      objPiFin,
            '                                      objBebu,
            '                                      objKrBhg,
            '                                      intAccounting,
            '                                      objdtInfo,
            '                                      strActualYear,
            '                                      strYear,
            '                                      intTeqNbr,
            '                                      intTeqNbrLY,
            '                                      intTeqNbrPLY,
            '                                      datPeriodFrom,
            '                                      datPeriodTo,
            '                                      strPeriodStatus)
            '    'Application.DoEvents()
            'End If
            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem PGV - Buchung Debitoren", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 9

        Finally
            drDebiSub = Nothing
            strPeriodenInfoA = Nothing

        End Try

    End Function

    Friend Function FcWriteToRGTable(intMandant As Int32,
                                     strRGNbr As String,
                                     datDate As Date,
                                     intBelegNr As Int32,
                                     ByRef objdbAccessConn As OleDb.OleDbConnection,
                                     ByRef objOracleConn As OracleConnection,
                                     ByRef objMySQLConn As MySqlConnection,
                                     booDatChanged As Boolean,
                                     datDebRGDatum As Date,
                                     datDebValDatum As Date) As Int16

        'Returns 0=ok, 1=Problem

        Dim strSQL As String
        Dim intAffected As Int16
        Dim objlocOLEdbcmd As New OleDb.OleDbCommand
        Dim objlocOracmd As New OracleCommand
        Dim objlocMySQLRGConn As New MySqlConnection
        Dim objlocMySQLRGcmd As New MySqlCommand
        Dim strNameRGTable As String
        Dim strBelegNrName As String
        Dim strRGNbrFieldName As String
        Dim strRGTableType As String
        Dim strMDBName As String
        Dim strBookedFieldName As String
        Dim strBookedDateFieldName As String
        Dim strDebRGFieldName As String
        Dim strDebValFieldName As String

        objMySQLConn.Open()

        strMDBName = FcReadFromSettingsII("Buchh_RGTableMDB", intMandant)
        strRGTableType = FcReadFromSettingsII("Buchh_RGTableType", intMandant)
        strNameRGTable = FcReadFromSettingsII("Buchh_TableDeb", intMandant)
        strBelegNrName = FcReadFromSettingsII("Buchh_TableRGBelegNrName", intMandant)
        strRGNbrFieldName = FcReadFromSettingsII("Buchh_TableRGNbrFieldName", intMandant)
        strDebRGFieldName = FcReadFromSettingsII("Buchh_DRGDateField", intMandant)
        strDebValFieldName = FcReadFromSettingsII("Buchh_DValDateField", intMandant)

        Try

            If strRGTableType = "A" Then
                'Access
                Call FcInitAccessConnecation(objdbAccessConn, strMDBName)

                strSQL = "UPDATE " + strNameRGTable + " Set gebucht=True, gebuchtDatum=#" + Format(datDate, "yyyy-MM-dd").ToString + "#, " +
                                                            strBelegNrName + "=" + intBelegNr.ToString +
                                                      " WHERE " + strRGNbrFieldName + "=" + strRGNbr
                objdbAccessConn.Open()
                objlocOLEdbcmd.CommandText = strSQL
                objlocOLEdbcmd.Connection = objdbAccessConn
                intAffected = objlocOLEdbcmd.ExecuteNonQuery()
                'Falls Datum changed, dann geänderte Daten in RG - Tabelle schreiben
                If booDatChanged Then
                    strSQL = "UPDATE " + strNameRGTable + " Set " + strDebRGFieldName + "=#" + Format(datDebRGDatum, "yyyy-MM-dd").ToString + "#, " +
                                                                    strDebValFieldName + "=#" + Format(datDebValDatum, "yyyy-MM-dd").ToString + "# " +
                                                        " WHERE " + strRGNbrFieldName + "=" + strRGNbr
                    objlocOLEdbcmd.CommandText = strSQL
                    intAffected = objlocOLEdbcmd.ExecuteNonQuery()
                End If

            ElseIf strRGTableType = "M" Then
                'MySQL
                'Bei IG Felnamen anders
                If intMandant = 25 Then
                    strBookedFieldName = "IGBooked"
                    strBookedDateFieldName = "IGDBDate"
                Else
                    strBookedFieldName = "gebucht"
                    strBookedDateFieldName = "gebuchtDatum"
                End If
                strSQL = "UPDATE " + strNameRGTable + " Set " + strBookedFieldName + "=True, " +
                                                                strBookedDateFieldName + "=Date('" + Format(datDate, "yyyy-MM-dd").ToString + "'), " +
                                                                strBelegNrName + "=" + intBelegNr.ToString +
                                                    " WHERE " + strRGNbrFieldName + "=" + strRGNbr
                objlocMySQLRGConn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(strMDBName)
                objlocMySQLRGConn.Open()
                objlocMySQLRGcmd.Connection = objlocMySQLRGConn
                objlocMySQLRGcmd.CommandText = strSQL
                intAffected = objlocMySQLRGcmd.ExecuteNonQuery()
                'Falls Datum-Changed dann geänderte Daten in RG-Tabelle schreiben
                If booDatChanged Then
                    strSQL = "UPDATE " + strNameRGTable + " SET " + strDebRGFieldName + "=DATE('" + Format(datDebRGDatum, "yyyy-MM-dd").ToString + "'), " +
                                                                    strDebValFieldName + "=DATE('" + Format(datDebValDatum, "yyyy-MM-dd").ToString + "')" +
                                                        " WHERE " + strRGNbrFieldName + "=" + strRGNbr
                    objlocMySQLRGcmd.CommandText = strSQL
                    intAffected = objlocMySQLRGcmd.ExecuteNonQuery()
                End If

            End If

            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Status in Debitor-RG-Tabelle schreiben")
            Return 1

        Finally
            If objdbAccessConn.State = ConnectionState.Open Then
                objdbAccessConn.Close()
            End If

            If objlocMySQLRGConn.State = ConnectionState.Open Then
                objlocMySQLRGConn.Close()
            End If

            If objMySQLConn.State = ConnectionState.Open Then
                objMySQLConn.Close()
            End If

        End Try

    End Function

    Friend Function FcExecuteAfterDebit(ByVal intMandant As Integer) As Int16

        Dim strSQL As String
        Dim strAfterDebiRunType As String
        Dim strMDBName As String
        Dim objlocMySQLRGConn As New MySqlConnection
        Dim objlocMySQLRGcmd As New MySqlCommand
        Dim intAffected As Int16


        Try

            'objMySQLConn.Open()
            strSQL = FcReadFromSettingsII("Buchh_SQLafterDebiRun", intMandant)
            strAfterDebiRunType = FcReadFromSettingsII("Buchh_SQLafterDebiType", intMandant)
            strMDBName = FcReadFromSettingsII("Buchh_SQLafterDebiMDB", intMandant)

            If strSQL <> "" Then

                If strAfterDebiRunType = "A" Then
                    Stop
                    'Access
                    'strdbProvider = "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
                    'strdbSource = "Data Source="
                    'strdbPathAndFile = "\\sdlc.mssag.ch\Apps\Backends\" + strMDBName + ";Jet OLEDB:System Database=\\sdlc.mssag.ch\Apps\Backends\Workbench.mdw;User ID=HagerR;"
                    'strSQL = "UPDATE " + strNameKRGTable + " SET Kredigebucht=true, KredigebuchtDatum=#" + Format(datDate, "yyyy-MM-dd").ToString + "#, " + strBelegNrName + "='" + intBelegNr.ToString + "' WHERE " + strKRGNbrFieldName + "=" + lngKredID.ToString

                    'objdbAccessConn.ConnectionString = strdbProvider + strdbSource + strdbPathAndFile
                    'objdbAccessConn.Open()

                    'objlocOLEdbcmd.CommandText = strSQL

                    'objlocOLEdbcmd.Connection = objdbAccessConn
                    'intAffected = objlocOLEdbcmd.ExecuteNonQuery()

                ElseIf strAfterDebiRunType = "M" Then
                    'MySQL
                    objlocMySQLRGConn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(strMDBName)
                    objlocMySQLRGConn.Open()
                    objlocMySQLRGcmd.Connection = objlocMySQLRGConn
                    objlocMySQLRGcmd.CommandText = strSQL
                    intAffected = objlocMySQLRGcmd.ExecuteNonQuery()

                End If

            End If

            Return 0


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Nach Debitor - Ausführung")
            Return 1

        Finally

            If objlocMySQLRGConn.State = ConnectionState.Open Then
                objlocMySQLRGConn.Close()
            End If
            objlocMySQLRGConn = Nothing
            objlocMySQLRGcmd = Nothing

        End Try

    End Function

    Friend Function FcDeleteNonAscii(strTexttoClean As String) As String

        Dim I As Long
        Dim J As Long
        Dim strChar As String

        I = 1
        For J = 1 To Len(strTexttoClean)
            strChar = Mid$(strTexttoClean, J, 1)
            If (AscW(strChar) And &HFFFF&) <= &H7F& Then
                Mid$(strTexttoClean, I, 1) = strChar
                I = I + 1
            End If
        Next
        strTexttoClean = Strings.Left$(strTexttoClean, I - 1)
        Return strTexttoClean

    End Function

    Friend Function FcGetZV(ByRef objSQLConn As SqlClient.SqlConnection,
                            ByRef objSQLCmd As SqlClient.SqlCommand,
                            ByVal strMandant As String,
                            ByVal strType As String,
                            ByRef intZV As Int32) As Int16

        Dim tblZV As New DataTable

        Try

            If objSQLConn.State = ConnectionState.Closed Then
                objSQLConn.Open()
            End If

            If strType = "SB" Then

                objSQLCmd.CommandText = "SELECT lfnbr FROM bankpost WHERE mandid='" + strMandant + "'" +
                                                                        " AND typ='E'" +
                                                                        " AND ktofibu='1092'"
            ElseIf strType = "GS" Then

                objSQLCmd.CommandText = "SELECT lfnbr FROM bankpost WHERE mandid='" + strMandant + "'" +
                                                                        " AND typ='E'" +
                                                                        " AND ktofibu='1093'"

            End If

            'Auslesen
            tblZV.Load(objSQLCmd.ExecuteReader)
            If tblZV.Rows.Count = 1 Then
                'Gefunden
                intZV = tblZV.Rows(0).Item("lfnbr")
            Else
                intZV = 0
            End If
            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem ZV - Suche")
            intZV = 0
            Return 9

        Finally
            objSQLConn.Close()

        End Try


    End Function

    Private Sub ButDeselect_Click(sender As Object, e As EventArgs) Handles ButDeselect.Click

        'Alle selektierten Records werden deselektiert

        For Each row As DataRow In dsDebitoren.Tables("tblDebiHeadsFromUser").Rows
            If Not IsDBNull(row("booDebBook")) Then
                If row("booDebBook") Then
                    row("booDebBook") = False
                End If
            End If
        Next
        dsDebitoren.Tables("tblDebiHeadsFromUser").AcceptChanges()
        'Me.Refresh()


    End Sub



    Private Sub BgWCheckDebi_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BgWCheckDebi.ProgressChanged

        Me.PRDebi.Value = e.ProgressPercentage

    End Sub

    Private Sub BgWCheckDebi_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgWCheckDebi.RunWorkerCompleted

        Me.PRDebi.Value = 0

    End Sub

    Private Sub BgWImportDebi_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BgWImportDebi.ProgressChanged

        Me.PRDebi.Value = e.ProgressPercentage

    End Sub

    Private Sub BgWImportDebi_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgWImportDebi.RunWorkerCompleted

        Me.PRDebi.Value = 0

    End Sub
End Class