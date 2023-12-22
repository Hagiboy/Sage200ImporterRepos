Option Strict Off
Option Explicit On

Imports MySql.Data.MySqlClient
Imports System.Data.OracleClient
Imports System.Data.SqlClient
Imports CLClassSage200.WFSage200Import
'Imports System.IO
'Imports Google.Protobuf.WellKnownTypes
'Imports Org.BouncyCastle.Crypto.Prng
Imports Sage200Importer.frmDebDisp
Imports System.IO
Imports System.Net
Imports System.Xml

Public Class frmKredDisp

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



    Public Sub InitDB()

        Dim strIdentityName As String
        Dim objdbtaskcmd As New MySqlCommand
        Dim objdbtasks As New DataTable


        Try

            strIdentityName = System.Security.Principal.WindowsIdentity.GetCurrent().Name
            strIdentityName = Strings.Replace(strIdentityName, "\", "/")

            frmImportMain.LblIdentity.Text = strIdentityName
            frmImportMain.LblTaskID.Text = Process.GetCurrentProcess().Id.ToString

            mysqlconn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString")

            'Zuert alle evtl. vorhandenen DS des Users löschen
            mysqlcmdKredDel.CommandText = "DELETE FROM tblkreditorenhead WHERE IdentityName='" + strIdentityName + "'"
            mysqlcmdKredDel.Connection.Open()
            mysqlcmdKredDel.ExecuteNonQuery()
            mysqlcmdKredDel.Connection.Close()

            mysqlcmdKredSubDel.CommandText = "DELETE FROM tblkreditorensub WHERE IdentityName='" + strIdentityName + "'"
            mysqlcmdKredSubDel.Connection.Open()
            mysqlcmdKredSubDel.ExecuteNonQuery()
            mysqlcmdKredSubDel.Connection.Close()


            'Read cmd DebiHead
            mysqlcmdKredRead.CommandText = "SELECT * FROM tblkreditorenhead WHERE IdentityName='" + strIdentityName + "' AND ProcessID= " + Process.GetCurrentProcess().Id.ToString

            'Del cmd DebiHead
            mysqlcmdKredDel.CommandText = "DELETE FROM tblkreditorenhead WHERE IdentityName='" + strIdentityName + "' AND ProcessID= " + Process.GetCurrentProcess().Id.ToString


            'Debitoren Sub
            'Read
            mysqlcmdKredSubRead.CommandText = "Select * FROM tblkreditorensub WHERE IdentityName='" + strIdentityName + "' AND ProcessID= " + Process.GetCurrentProcess().Id.ToString

            'Del cmd Debi Sub
            mysqlcmdKredSubDel.CommandText = "DELETE FROM tblkreditorensub WHERE IdentityName='" + strIdentityName + "' AND ProcessID= " + Process.GetCurrentProcess().Id.ToString

            'Mandant holen
            objdbtaskcmd.Connection = objdbConn
            objdbtaskcmd.Connection.Open()
            objdbtaskcmd.CommandText = "SELECT * FROM tblimporttasks WHERE IdentityName='" + strIdentityName + "' AND Type='C'"
            objdbtasks.Load(objdbtaskcmd.ExecuteReader())
            If objdbtasks.Rows.Count > 0 Then
                intMandant = objdbtasks.Rows(0).Item("Mandant")
            Else
                intMandant = 1
                MessageBox.Show("Mandant konnte nicht gelesen werden. => Setzen auf AHZ")
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Generelles Problem " + Convert.ToString(Err.Number) + "Init Kreditoren", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Err.Clear()

        Finally
            objdbtasks = Nothing
            objdbtaskcmd = Nothing

        End Try

    End Sub

    Private Sub frmKredDisp_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        FELD_SEP = "{<}"
        REC_SEP = "{>}"
        KSTKTR_SEP = "{-}"

        FELD_SEP_OUT = "{>}"
        REC_SEP_OUT = "{<}"

        Me.Cursor = Cursors.WaitCursor

        Call InitDB()

        BgWLoadKredi.RunWorkerAsync(intMandant)

        Do While BgWLoadKredi.IsBusy
            Threading.Thread.Sleep(1)
            Application.DoEvents()
        Loop

        'Tabellentyp darstellen
        Call FcReadFromSettingsIII("Buchh_RGTableType",
                                              intMandant,
                                              Me.lblDB.Text)

        Me.Cursor = Cursors.Default


    End Sub


    'Friend Function FcKrediDisplay(intMandant As Int32,
    '                               LstMandant As ListBox,
    '                               LstBPerioden As ListBox) As Int16

    '    Dim objdbConn As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString"))
    '    Dim objdbtaskcmd As New MySqlCommand
    '    Dim objdbMSSQLConn As New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("SQLConnectionString"))
    '    Dim objdbSQLcommand As New SqlCommand

    '    Dim intFcReturns As Int16
    '    Dim strPeriode As String
    '    Dim strYearCh As String
    '    Dim BgWCheckKrediLocArgs As New BgWCheckDebitArgs
    '    Dim objdbtasks As New DataTable

    '    'Dim objFinanz As New SBSXASLib.AXFinanz
    '    'Dim objfiBuha As New SBSXASLib.AXiFBhg
    '    'Dim objdbBuha As New SBSXASLib.AXiDbBhg
    '    'Dim objdbPIFb As New SBSXASLib.AXiPlFin
    '    'Dim objFiBebu As New SBSXASLib.AXiBeBu
    '    'Dim objKrBuha As New SBSXASLib.AXiKrBhg


    '    Try

    '        Me.Cursor = Cursors.WaitCursor

    '        'Zuerst in tblImportTasks setzen
    '        objdbtaskcmd.Connection = objdbConn
    '        objdbtaskcmd.Connection.Open()
    '        objdbtaskcmd.CommandText = "SELECT * FROM tblimporttasks WHERE IdentityName='" + frmImportMain.LblIdentity.Text + "' AND Type='C'"
    '        objdbtasks.Load(objdbtaskcmd.ExecuteReader())
    '        If objdbtasks.Rows.Count > 0 Then
    '            'update
    '            objdbtaskcmd.CommandText = "UPDATE tblimporttasks SET Mandant=" + Convert.ToString(LstMandant.SelectedIndex) + ", Periode=" + Convert.ToString(LstBPerioden.SelectedIndex) + " WHERE IdentityName='" + frmImportMain.LblIdentity.Text + "' AND Type='C'"
    '        Else
    '            'insert
    '            objdbtaskcmd.CommandText = "INSERT INTO tblimporttasks (IdentityName, Type, Mandant, Periode) VALUES ('" + frmImportMain.LblIdentity.Text + "', 'C', " + Convert.ToString(LstMandant.SelectedIndex) + ", " + Convert.ToString(LstBPerioden.SelectedIndex) + ")"
    '        End If
    '        objdbtaskcmd.ExecuteNonQuery()
    '        objdbtaskcmd.Connection.Close()

    '        'DGVs
    '        dgvBookings.DataSource = Nothing
    '        dgvBookingSub.DataSource = Nothing

    '        Me.butImport.Enabled = False

    '        'Zuerst evtl. vorhandene DS löschen in Tabelle
    '        MySQLdaKreditoren.DeleteCommand.Connection.Open()
    '        MySQLdaKreditoren.DeleteCommand.ExecuteNonQuery()
    '        MySQLdaKreditoren.DeleteCommand.Connection.Close()

    '        MySQLdaKreditorenSub.DeleteCommand.Connection.Open()
    '        MySQLdaKreditorenSub.DeleteCommand.ExecuteNonQuery()
    '        MySQLdaKreditorenSub.DeleteCommand.Connection.Close()

    '        'Info neu erstellen
    '        dsKreditoren.Tables.Add("tblKreditorenInfo")
    '        Dim col1 As DataColumn = New DataColumn("strInfoT")
    '        col1.DataType = System.Type.GetType("System.String")
    '        col1.MaxLength = 50
    '        col1.Caption = "Info-Titel"
    '        dsKreditoren.Tables("tblKreditorenInfo").Columns.Add(col1)
    '        Dim col2 As DataColumn = New DataColumn("strInfoV")
    '        col2.DataType = System.Type.GetType("System.String")
    '        col2.MaxLength = 50
    '        col2.Caption = "Info-Wert"
    '        dsKreditoren.Tables("tblKreditorenInfo").Columns.Add(col2)

    '        dgvInfo.DataSource = dsKreditoren.Tables("tblKreditorenInfo")

    '        'Datums-Tabelle erstellen
    '        dsKreditoren.Tables.Add("tblKreditorenDates")
    '        Dim col7 As DataColumn = New DataColumn("intYear")
    '        col7.DataType = System.Type.GetType("System.Int16")
    '        col7.Caption = "Year"
    '        dsKreditoren.Tables("tblKreditorenDates").Columns.Add(col7)
    '        Dim col3 As DataColumn = New DataColumn("strDatType")
    '        col3.DataType = System.Type.GetType("System.String")
    '        col3.MaxLength = 50
    '        col3.Caption = "Datum-Typ"
    '        dsKreditoren.Tables("tblKreditorenDates").Columns.Add(col3)
    '        Dim col4 As DataColumn = New DataColumn("datFrom")
    '        col4.DataType = System.Type.GetType("System.DateTime")
    '        col4.Caption = "Von"
    '        dsKreditoren.Tables("tblKreditorenDates").Columns.Add(col4)
    '        Dim col5 As DataColumn = New DataColumn("datTo")
    '        col5.DataType = System.Type.GetType("System.DateTime")
    '        col5.Caption = "Bis"
    '        dsKreditoren.Tables("tblKreditorenDates").Columns.Add(col5)
    '        Dim col6 As DataColumn = New DataColumn("strStatus")
    '        col6.DataType = System.Type.GetType("System.String")
    '        col6.Caption = "S"
    '        dsKreditoren.Tables("tblKreditorenDates").Columns.Add(col6)
    '        dgvDates.DataSource = dsKreditoren.Tables("tblKreditorenDates")

    '        strPeriode = LstBPerioden.GetItemText(LstBPerioden.SelectedItem)

    '        Call Main.FcLoginSage3(objdbConn,
    '                              objdbMSSQLConn,
    '                              objdbSQLcommand,
    '                              objFinanz,
    '                              objfiBuha,
    '                              objdbBuha,
    '                              objdbPIFb,
    '                              objFiBebu,
    '                              objKrBuha,
    '                              intMandant,
    '                              dsKreditoren.Tables("tblKreditorenInfo"),
    '                              dsKreditoren.Tables("tblDebitorenDates"),
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
    '                Call Main.FcLoginSage4(intMandant,
    '                                   dsKreditoren.Tables("tblDebitorenDates"),
    '                                   strPeriode)
    '            Else
    '                'Periode ezreugen und auf N stellen
    '                strYearCh = Convert.ToString(Val(strYear) - 1)
    '                dsKreditoren.Tables("tblDebitorenDates").Rows.Add(strYearCh, "GJ n/a", DateSerial(Convert.ToUInt16(strYearCh), 1, 1), DateSerial(Convert.ToUInt16(strYearCh), 12, 31), "N")
    '            End If

    '            'Gibt es ein Folgehahr?
    '            If LstBPerioden.SelectedIndex + 1 < LstBPerioden.Items.Count Then
    '                strPeriode = LstBPerioden.Items(LstBPerioden.SelectedIndex + 1)
    '                'Peeriodendef holen
    '                Call Main.FcLoginSage4(intMandant,
    '                                   dsKreditoren.Tables("tblDebitorenDates"),
    '                                   strPeriode)
    '            Else
    '                'Periode ezreugen und auf N stellen
    '                strYearCh = Convert.ToString(Val(strYear) + 1)
    '                dsKreditoren.Tables("tblDebitorenDates").Rows.Add(strYearCh, "GJ n/a", DateSerial(Convert.ToUInt16(strYearCh), 1, 1), DateSerial(Convert.ToUInt16(strYearCh), 12, 31), "N")
    '            End If

    '        ElseIf LstBPerioden.Items.Count = 1 Then 'es gibt genau 1 Jahr
    '            'gewähltes Jahr checken
    '            Call Main.FcLoginSage4(intMandant,
    '                                   dsKreditoren.Tables("tblDebitorenDates"),
    '                                   strPeriode)
    '            'VJ erzeugen
    '            strYearCh = Convert.ToString(Val(strYear) - 1)
    '            dsKreditoren.Tables("tblDebitorenDates").Rows.Add(strYearCh, "GJ n/a", DateSerial(Convert.ToUInt16(strYearCh), 1, 1), DateSerial(Convert.ToUInt16(strYearCh), 12, 31), "N")

    '            'FJ erzeugen
    '            strYearCh = Convert.ToString(Val(strYear) + 1)
    '            dsKreditoren.Tables("tblDebitorenDates").Rows.Add(strYearCh, "GJ n/a", DateSerial(Convert.ToUInt16(strYearCh), 1, 1), DateSerial(Convert.ToUInt16(strYearCh), 12, 31), "N")

    '        End If


    '        'Dim clImp As New ClassImport
    '        'clImp.FcKreditFill(intMandant)
    '        'clImp = Nothing

    '        BgWLoadKredi.RunWorkerAsync(intMandant)

    '        Do While BgWLoadKredi.IsBusy
    '            Application.DoEvents()
    '        Loop


    '        'Tabellentyp darstellen
    '        Me.lblDB.Text = Main.FcReadFromSettingsII("Buchh_KRGTableType", intMandant)


    '        MySQLdaKreditoren.Fill(dsKreditoren, "tblKrediHeadsFromUser")
    '        MySQLdaKreditorenSub.Fill(dsKreditoren, "tblKrediSubsFromUser")


    '        'Application.DoEvents()

    '        'Dim clCheck As New ClassCheck
    '        'clCheck.FcCheckKredit(intMandant,
    '        '                  dsKreditoren,
    '        '                  Finanz,
    '        '                  FBhg,
    '        '                  KrBhg,
    '        '                  BeBu,
    '        '                  dsKreditoren.Tables("tblKreditorenInfo"),
    '        '                  dsKreditoren.Tables("tblDebitorenDates"),
    '        '                  frmImportMain.lstBoxMandant.Text,
    '        '                  strYear,
    '        '                  strPeriode,
    '        '                  datPeriodFrom,
    '        '                  datPeriodTo,
    '        '                  strPeriodStatus,
    '        '                  frmImportMain.chkValutaCorrect.Checked,
    '        '                  frmImportMain.dtpValutaCorrect.Value)

    '        'clCheck = Nothing

    '        BgWCheckKrediLocArgs.intMandant = intMandant
    '        BgWCheckKrediLocArgs.strMandant = frmImportMain.lstBoxMandant.GetItemText(frmImportMain.lstBoxMandant.SelectedItem)
    '        BgWCheckKrediLocArgs.intTeqNbr = intTeqNbr
    '        BgWCheckKrediLocArgs.intTeqNbrLY = intTeqNbrLY
    '        BgWCheckKrediLocArgs.intTeqNbrPLY = intTeqNbrPLY
    '        BgWCheckKrediLocArgs.strYear = strYear
    '        BgWCheckKrediLocArgs.strPeriode = LstBPerioden.GetItemText(LstBPerioden.SelectedItem)
    '        BgWCheckKrediLocArgs.booValutaCor = frmImportMain.chkValutaCorrect.Checked
    '        BgWCheckKrediLocArgs.datValutaCor = frmImportMain.dtpValutaCorrect.Value

    '        BgWCheckKredi.RunWorkerAsync(BgWCheckKrediLocArgs)

    '        Do While BgWCheckKredi.IsBusy
    '            Application.DoEvents()
    '        Loop

    '        Debug.Print("Vor Refresh DGV")

    '        'Grid neu aufbauen
    '        dgvBookings.DataSource = dsKreditoren.Tables("tblKrediHeadsFromUser")
    '        dgvBookingSub.DataSource = dsKreditoren.Tables("tblKrediSubsFromUser")

    '        intFcReturns = FcInitdgvInfo(dgvInfo)
    '        intFcReturns = FcInitdgvKreditoren(dgvBookings)
    '        intFcReturns = FcInitdgvKrediSub(dgvBookingSub)
    '        intFcReturns = FcInitdgvDate(dgvDates)


    '        'Anzahl schreiben
    '        txtNumber.Text = Me.dsKreditoren.Tables("tblKrediHeadsFromUser").Rows.Count.ToString

    '        Me.Cursor = Cursors.Default

    '        Me.butImport.Enabled = True
    '        Return 0

    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message, "Generelles Problem Kredi-Check" + Err.Number.ToString)
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
    '        objdbtaskcmd = Nothing
    '        objdbtasks = Nothing

    '        BgWCheckKrediLocArgs = Nothing

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

    Friend Function FcInitdgvKreditoren(ByRef dgvBookings As DataGridView) As Int16

        Try

            dgvBookings.ShowCellToolTips = False
            dgvBookings.AllowUserToAddRows = False
            dgvBookings.AllowUserToDeleteRows = False
            Dim okcol As New DataGridViewCheckBoxColumn
            okcol.DataPropertyName = "booKredBook"
            okcol.HeaderText = "ok"
            okcol.DisplayIndex = 0
            okcol.Width = 40
            dgvBookings.Columns.Add(okcol)
            dgvBookings.Columns("booKredBook").Visible = False
            'dgvBookings.Columns("booKredBook").DisplayIndex = 0
            'dgvBookings.Columns("booKredBook").HeaderText = "ok"
            'dgvBookings.Columns("booKredBook").Width = 40
            'dgvBookings.Columns("booKredBook").ValueType = System.Type.[GetType]("System.Boolean")
            'dgvBookings.Columns("booKredBook").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            'dgvBookings.Columns("booKredBook").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
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
            dgvBookings.Columns("lngKredNbr").Width = 50
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
            dgvBookings.Columns("intSubBookings").Width = 30
            dgvBookings.Columns("intSubBookings").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvBookings.Columns("dblSumSubBookings").DisplayIndex = 11
            dgvBookings.Columns("dblSumSubBookings").DefaultCellStyle.Format = "N2"
            dgvBookings.Columns("dblSumSubBookings").HeaderText = "Sub-Summe"
            dgvBookings.Columns("dblSumSubBookings").Width = 80
            dgvBookings.Columns("dblSumSubBookings").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvBookings.Columns("lngKredIdentNbr").DisplayIndex = 12
            dgvBookings.Columns("lngKredIdentNbr").HeaderText = "Ident"
            dgvBookings.Columns("lngKredIdentNbr").Width = 80
            'Dim cmbBuchungsart As New DataGridViewComboBoxColumn()
            'Dim objdtBA As New DataTable("objidtBA")
            'Dim objlocMySQLcmd As New MySqlCommand
            ''objdbConn.Open()
            'objlocMySQLcmd.CommandText = "Select * FROM t_sage_tblbuchungsarten"
            'objlocMySQLcmd.Connection = objdbConn
            'objdtBA.Load(objlocMySQLcmd.ExecuteReader)
            'cmbBuchungsart.DataSource = objdtBA
            ''objdbConn.Close()
            'cmbBuchungsart.DisplayMember = "strBuchungsart"
            'cmbBuchungsart.ValueMember = "idBuchungsart"
            'cmbBuchungsart.HeaderText = "BA"
            'cmbBuchungsart.Name = "intBuchungsart"
            'cmbBuchungsart.DataPropertyName = "intBuchungsart"
            'cmbBuchungsart.DisplayIndex = 13
            'cmbBuchungsart.Width = 70
            'dgvBookings.Columns.Add(cmbBuchungsart)
            dgvBookings.Columns("intBuchungsart").DisplayIndex = 13
            dgvBookings.Columns("intBuchungsart").DisplayIndex = 13
            dgvBookings.Columns("intBuchungsart").HeaderText = "BA"
            dgvBookings.Columns("intBuchungsart").Width = 30
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
            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error " + Err.Number.ToString)
            Return 1

        End Try

    End Function

    Friend Function FcInitdgvKrediSub(ByRef dgvBookingSub As DataGridView) As Int16

        Try

            dgvBookingSub.ShowCellToolTips = False
            dgvBookingSub.AllowUserToAddRows = False
            dgvBookingSub.AllowUserToDeleteRows = False
            'dgvBookingSub.Columns("strRGNr").DisplayIndex = 0
            'dgvBookingSub.Columns("strRGNr").Width = 60
            'dgvBookingSub.Columns("strRGNr").HeaderText = "RG-Nr"
            dgvBookingSub.Columns("lngKredID").DisplayIndex = 0
            dgvBookingSub.Columns("lngKredID").Width = 60
            dgvBookingSub.Columns("lngKredID").HeaderText = "Kred-ID"
            dgvBookingSub.Columns("intSollHaben").DisplayIndex = 1
            dgvBookingSub.Columns("intSollHaben").Width = 30
            dgvBookingSub.Columns("intSollHaben").HeaderText = "S/H"
            dgvBookingSub.Columns("lngKto").DisplayIndex = 2
            dgvBookingSub.Columns("lngKto").Width = 50
            dgvBookingSub.Columns("lngKto").HeaderText = "Konto"
            dgvBookingSub.Columns("strKtoBez").DisplayIndex = 3
            dgvBookingSub.Columns("strKtoBez").HeaderText = "Bezeichnung"
            dgvBookingSub.Columns("lngKST").DisplayIndex = 4
            dgvBookingSub.Columns("lngKST").Width = 50
            dgvBookingSub.Columns("lngKST").HeaderText = "KST"
            dgvBookingSub.Columns("strKSTBez").DisplayIndex = 5
            dgvBookingSub.Columns("strKSTBez").Width = 80
            dgvBookingSub.Columns("strKSTBez").HeaderText = "Bezeichnung"
            dgvBookingSub.Columns("dblNetto").DisplayIndex = 6
            dgvBookingSub.Columns("dblNetto").Width = 70
            dgvBookingSub.Columns("dblNetto").HeaderText = "Netto"
            dgvBookingSub.Columns("dblNetto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            'dgvBookingSub.Columns("dblNetto").DefaultCellStyle.Format = "N2"
            dgvBookingSub.Columns("dblMwSt").DisplayIndex = 7
            dgvBookingSub.Columns("dblMwSt").Width = 60
            dgvBookingSub.Columns("dblMwSt").HeaderText = "MwSt"
            dgvBookingSub.Columns("dblMwSt").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            'dgvBookingSub.Columns("dblMwSt").DefaultCellStyle.Format = "N2"
            dgvBookingSub.Columns("dblBrutto").DisplayIndex = 8
            dgvBookingSub.Columns("dblBrutto").Width = 70
            dgvBookingSub.Columns("dblBrutto").HeaderText = "Brutto"
            dgvBookingSub.Columns("dblBrutto").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            'dgvBookingSub.Columns("dblBrutto").DefaultCellStyle.Format = "N2"
            dgvBookingSub.Columns("dblMwStSatz").DisplayIndex = 9
            dgvBookingSub.Columns("dblMwStSatz").Width = 40
            dgvBookingSub.Columns("dblMwStSatz").HeaderText = "MwStS"
            dgvBookingSub.Columns("dblMwStSatz").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvBookingSub.Columns("dblMwStSatz").DefaultCellStyle.Format = "N1"
            dgvBookingSub.Columns("strMwStKey").DisplayIndex = 10
            dgvBookingSub.Columns("strMwStKey").Width = 60
            dgvBookingSub.Columns("strMwStKey").HeaderText = "MwStK"
            dgvBookingSub.Columns("strKredSubText").DisplayIndex = 11
            dgvBookingSub.Columns("strKredSubText").Width = 100
            dgvBookingSub.Columns("booRebilling").DisplayIndex = 12
            dgvBookingSub.Columns("booRebilling").Width = 40
            dgvBookingSub.Columns("booRebilling").HeaderText = "Rebill"
            dgvBookingSub.Columns("strStatusUBText").DisplayIndex = 13
            dgvBookingSub.Columns("strStatusUBText").HeaderText = "Status"
            dgvBookingSub.Columns("strStatusUBText").Width = 120
            'dgvBookingSub.Columns("lngID").Visible = False
            'dgvBookingSub.Columns("strArtikel").Visible = False
            'dgvBookingSub.Columns("strStatusUBBitLog").Visible = False
            'dgvBookingSub.Columns("strDebSubText").Visible = False
            'dgvBookingSub.Columns("strDebBookStatus").Visible = False

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error " + Err.Number.ToString)

        End Try

    End Function

    Private Sub frmKredDisp_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        'DS in tabelle löschen
        MySQLdaKreditoren.DeleteCommand.Connection.Open()
        MySQLdaKreditoren.DeleteCommand.ExecuteNonQuery()
        MySQLdaKreditoren.DeleteCommand.Connection.Close()

        MySQLdaKreditorenSub.DeleteCommand.Connection.Open()
        MySQLdaKreditorenSub.DeleteCommand.ExecuteNonQuery()
        MySQLdaKreditorenSub.DeleteCommand.Connection.Close()


    End Sub

    Private Sub dgvBookings_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvBookings.CellContentClick

        Dim intFctReturns As Int16

        Try

            If e.RowIndex >= 0 Then

                dgvBookingSub.DataSource = dsKreditoren.Tables("tblKrediSubsFromUser").Select("lngKredID=" + dgvBookings.Rows(e.RowIndex).Cells("lngKredID").Value.ToString).CopyToDataTable
                intFctReturns = FcInitdgvKrediSub(dgvBookingSub)

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)

        Finally
            'dgvBookingSub.Update()

        End Try


    End Sub

    Private Sub dgvBookings_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvBookings.CellValueChanged


        Try


            If dgvBookings.Columns(e.ColumnIndex).HeaderText = "ok" And e.RowIndex >= 0 Then


                If IIf(IsDBNull(dgvBookings.Rows(e.RowIndex).Cells("booKredBook").Value), False, dgvBookings.Rows(e.RowIndex).Cells("booKredBook").Value) Then

                    'MsgBox("Geändert zu checked " + dgvDebitoren.Rows(e.RowIndex).Cells("strDebRGNbr").Value + ", " + dgvDebitoren.Rows(e.RowIndex).Cells("booDebBook").Value.ToString + Val(dgvDebitoren.Rows(e.RowIndex).Cells("strDebStatusBitLog").Value).ToString)
                    'Zulassen? = keine Fehler
                    If Val(dgvBookings.Rows(e.RowIndex).Cells("strKredStatusBitLog").Value) <> 0 Then
                        MsgBox("Kredi-Rechnung ist nicht buchbar.", vbOKOnly + vbExclamation, "Nicht buchbar")
                        dgvBookings.Rows(e.RowIndex).Cells("booKredBook").Value = False
                    End If

                End If

            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Generelles Problem " + Err.Number.ToString)

        End Try


    End Sub

    Private Sub butImport_Click(sender As Object, e As EventArgs) Handles butImport.Click

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

        Dim BgWImportKrediLocArgs As New BgWCheckDebitArgs

        Try

            'Variablen zuweisen
            BgWImportKrediLocArgs.intMandant = frmImportMain.lstBoxMandant.SelectedValue
            BgWImportKrediLocArgs.intTeqNbr = intTeqNbr
            BgWImportKrediLocArgs.intTeqNbrLY = intTeqNbrLY
            BgWImportKrediLocArgs.intTeqNbrPLY = intTeqNbrPLY
            BgWImportKrediLocArgs.strYear = strYear
            BgWImportKrediLocArgs.strPeriode = frmImportMain.lstBoxPerioden.GetItemText(frmImportMain.lstBoxPerioden.SelectedItem)


            Cursor = Cursors.WaitCursor
            'Application.DoEvents()
            'Button disablen damit er nicht noch einmal geklickt werden kann.
            Me.butImport.Enabled = False

            BgWImportKredi.RunWorkerAsync(BgWImportKrediLocArgs)

            Do While BgWImportKredi.IsBusy
                Application.DoEvents()
            Loop

            'Me.Cursor = Cursors.Default


            ''Start in Sync schreiben
            'intReturnValue = WFDBClass.FcWriteStartToSync(objdbConn,
            '                                              frmImportMain.lstBoxMandant.SelectedValue,
            '                                              2,
            '                                              dsKreditoren.Tables("tblKrediHeadsFromUser").Rows.Count)

            ''intTeqNbr = Conversion.Val(Strings.Right(objdtInfo.Rows(1).Item(1), 3))

            ''Kopfbuchung
            'For Each row As DataRow In Me.dsKreditoren.Tables("tblKrediHeadsFromUser").Rows

            '    If IIf(IsDBNull(row("booKredBook")), False, row("booKredBook")) Then

            '        'Test ob OP - Buchung
            '        If row("intBuchungsart") = 1 Then

            '            'Immer zuerst Belegs-Nummerierung aktivieren, falls vorhanden externe Nummer = OP - Nr. Rg
            '            'Führt zu Problemen beim Ausbuchen des DTA - Files
            '            'Resultat Besprechnung 17.09.20 mit Muhi/ Andy
            '            'DbBhg.IncrBelNbr = "J"
            '            'Belegsnummer abholen
            '            'intDebBelegsNummer = DbBhg.GetNextBelNbr("R")

            '            'Auf Provisorisch setzen
            '            Call KrBhg.SetBuchMode("P")

            '            'Automatische ESR - Zahlungsverbindung
            '            KrBhg.EnableAutoESRZlgVerb = "J"

            '            'Eindeutigkeit der internen Beleg-Nummer setzen
            '            KrBhg.CheckDoubleIntBelNbr = "N"

            '            'Eindeutigkeit externer Beleg-Nummer setzen
            '            KrBhg.CheckDoubleExtBelNbr = "J"

            '            'If IsDBNull(row("strOPNr")) Or row("StrOPNr") = "" Then
            '            'strExtBelegNbr = row("strOPNr")

            '            'Zuerst Beleg-Nummerieungung aktivieren
            '            'KrBhg.IncrBelNbr = "N"
            '            'Belegsnummer abholen
            '            'intKredBelegsNummer = KrBhg.GetNextBelNbr("R")
            '            'Else
            '            'Beleg-Nummerierung abschalten
            '            'KrBhg.IncrBelNbr = "N"
            '            'intKredBelegsNummer = row("strOPNr")
            '            'strExtBelegNbr = row("strOPNr")
            '            'End If
            '            strExtKredBelegsNummer = row("strKredRGNbr")

            '            'Variablen zuweisen
            '            intKreditorNbr = row("lngKredNbr")
            '            If row("dblKredBrutto") < 0 Then
            '                strBuchType = "G"
            '                'strZahlSperren = "J"
            '                row("dblKredBrutto") = row("dblKredBrutto") * -1
            '                'Belegsnummer abholen
            '                KrBhg.IncrBelNbr = "J"
            '                intKredBelegsNummer = KrBhg.GetNextBelNbr("G")
            '                KrBhg.IncrBelNbr = "N"

            '                intReturnValue = MainKreditor.FcCheckKrediExistance(objdbMSSQLConn,
            '                                                                    objdbSQLcommand,
            '                                                                    intKredBelegsNummer,
            '                                                                    "G",
            '                                                                    intTeqNbr,
            '                                                                    intTeqNbrLY,
            '                                                                    intTeqNbrPLY,
            '                                                                    KrBhg)

            '            Else
            '                strBuchType = "R"
            '                'strZahlSperren = "N"
            '                'Belegsnummer abholen
            '                KrBhg.IncrBelNbr = "J"
            '                intKredBelegsNummer = KrBhg.GetNextBelNbr("R")
            '                'Muss auf Nicht hochzählen gesetzt werden da Sage 200 nicht merkt, dass Beleg-Nr. schon vergeben worden sind. => In den Einstellungen muss von Zeit zu Zeit der Zähler geändert werden
            '                KrBhg.IncrBelNbr = "N"

            '                intReturnValue = MainKreditor.FcCheckKrediExistance(objdbMSSQLConn,
            '                                                                    objdbSQLcommand,
            '                                                                    intKredBelegsNummer,
            '                                                                    "R",
            '                                                                    intTeqNbr,
            '                                                                    intTeqNbrLY,
            '                                                                    intTeqNbrPLY,
            '                                                                    KrBhg)

            '            End If

            '            strValutaDatum = Format(row("datKredValDatum"), "yyyyMMdd").ToString
            '            strBelegDatum = Format(row("datKredRGDatum"), "yyyyMMdd").ToString
            '            strVerfallDatum = String.Empty
            '            'strReferenz = IIf(IsDBNull(row("strKredRef")), "", row("strKredRef"))
            '            'If IsDBNull(row("strKrediBank")) Then
            '            'intTeilnehmer = 0
            '            'Else
            '            'Teilnehmer nur bei ESR setzen
            '            If row("intPayType") <> 9 Then 'nicht IBAN
            '                'QR-Referenz
            '                strReferenz = IIf(IsDBNull(row("strKredRef")), "", row("strKredRef"))
            '                If row("intPayType") = 10 Then
            '                    strTeilnehmer = row("strKrediBank")
            '                Else
            '                    strTeilnehmer = Val(row("strKrediBank"))
            '                End If
            '                intBankNbr = 0
            '            Else
            '                'IBAN
            '                strReferenz = IIf(IsDBNull(row("strKredRef")), "", row("strKredRef"))
            '                intBankNbr = IIf(IsDBNull(row("intEBank")), 0, row("intEBank"))
            '                strTeilnehmer = String.Empty
            '            End If
            '            'End If
            '            strMahnerlaubnis = String.Empty 'Format(row("datDebRGDatum"), "yyyyMMdd").ToString
            '            'Sachbearbeiter aus Debitor auslesen
            '            strDebiLine = KrBhg.ReadKreditor3(row("lngKredNbr") * -1, "")
            '            strDebitor = Split(strDebiLine, "{>}")
            '            strSachBID = strDebitor(29)

            '            dblBetrag = row("dblKredBrutto")
            '            strKrediText = IIf(IsDBNull(row("strKredText")), "", row("strKredText"))
            '            strCurrency = row("strKredCur")
            '            'intBankNbr = 0
            '            intKondition = IIf(IsDBNull(row("intZKond")), 1, row("intZKond"))
            '            'LN 0=automatsich ersterfasste Kondition, -1=Schlechteste Kondition, -2=Beste Kondition
            '            intKonditionLN = 0
            '            intEigeneBank = row("intintBank")

            '            If strCurrency <> "CHF" Then 'Muss ergänzt werden => Was ist Leitwährung auf dem Konto
            '                dblKurs = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg)
            '            Else
            '                dblKurs = 1.0#
            '            End If

            '            Try
            '                booBookingok = True
            '                Call KrBhg.SetBelegKopf2(intKredBelegsNummer,
            '                                         strValutaDatum,
            '                                         intKreditorNbr,
            '                                         strExtKredBelegsNummer,
            '                                         strBelegDatum,
            '                                         strVerfallDatum,
            '                                         strKrediText,
            '                                         intBankNbr,
            '                                         strBuchType,
            '                                         strZahlSperren,
            '                                         strVorausZahlung,
            '                                         intKondition,
            '                                         intKonditionLN,
            '                                         strSachBID,
            '                                         strReferenz,
            '                                         strSkonto,
            '                                         dblBetrag.ToString,
            '                                         strErfassungsArt,
            '                                         dblKurs.ToString,
            '                                         strCurrency,
            '                                         "",
            '                                         strTeilnehmer,
            '                                         intEigeneBank.ToString)

            '            Catch ex As Exception
            '                MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Belegkopf ")
            '                If (Err.Number And 65535) < 10000 Then
            '                    booBookingok = False
            '                Else
            '                    booBookingok = True
            '                End If

            '            End Try

            '            selKrediSub = dsKreditoren.Tables("tblKrediSubsFromUser").Select("lngKredID=" + row("lngKredID").ToString)

            '            For Each SubRow As DataRow In selKrediSub

            '                intGegenKonto = SubRow("lngKto")
            '                strFibuText = SubRow("strKredSubText")
            '                'Soll auf Minus setzen
            '                'If SubRow("intSollHaben") = 1 Then
            '                'dblNettoBetrag = SubRow("dblNetto") * -1
            '                'dblMwStBetrag = SubRow("dblMwSt") * -1
            '                'dblBruttoBetrag = SubRow("dblBrutto") * -1
            '                'Else
            '                If intGegenKonto <> 6906 Then
            '                    If strBuchType = "R" Then
            '                        dblNettoBetrag = SubRow("dblNetto")
            '                        dblMwStBetrag = SubRow("dblMwSt")
            '                        dblBruttoBetrag = SubRow("dblBrutto")
            '                    Else
            '                        dblNettoBetrag = SubRow("dblNetto") * -1
            '                        dblMwStBetrag = SubRow("dblMwSt") * -1
            '                        dblBruttoBetrag = SubRow("dblBrutto") * -1
            '                    End If
            '                Else 'Rundungsdifferenzen
            '                    If strBuchType = "R" Then
            '                        dblNettoBetrag = SubRow("dblBrutto")
            '                        dblMwStBetrag = SubRow("dblMwSt")
            '                        dblBruttoBetrag = SubRow("dblBrutto")
            '                    Else
            '                        dblNettoBetrag = SubRow("dblBrutto") * -1
            '                        dblMwStBetrag = SubRow("dblMwSt") * -1
            '                        dblBruttoBetrag = SubRow("dblBrutto") * -1
            '                    End If

            '                End If

            '                'If strBuchType = "R" Then
            '                '    dblNettoBetrag = SubRow("dblNetto")
            '                '    dblMwStBetrag = SubRow("dblMwSt")
            '                '    dblBruttoBetrag = SubRow("dblBrutto")
            '                'Else
            '                '    dblNettoBetrag = SubRow("dblNetto") * -1
            '                '    dblMwStBetrag = SubRow("dblMwSt") * -1
            '                '    dblBruttoBetrag = SubRow("dblBrutto") * -1
            '                'End If
            '                'End If
            '                'dblBebuBetrag = 1000.0#
            '                If SubRow("lngKST") > 0 Then
            '                    strBeBuEintrag = SubRow("lngKST").ToString + "{<}" + SubRow("strKredSubText") + "{<}" + "CALCULATE" + "{>}"    '"PROD{<}BebuText{<}" + dblBebuBetrag.ToString + "{>}"
            '                Else
            '                    'strBeBuEintrag = "00" + "{<}" + SubRow("strKredSubText") + "{<}" + "0" + "{>}"
            '                End If
            '                If Not IsDBNull(SubRow("strMwStKey")) And SubRow("strMwStKey") <> "null" Then ' And SubRow("strMwStKey") <> "25" Then
            '                    strSteuerFeld = Main.FcGetSteuerFeld(FBhg, SubRow("lngKto"), SubRow("strKredSubText"), dblBruttoBetrag, SubRow("strMwStKey"), dblMwStBetrag)     '"25{<}DEBI D Bhg Export MwSt{<}0{>}"
            '                Else
            '                    strSteuerFeld = "STEUERFREI"
            '                End If

            '                'strSteuerInfo = Split(FBhg.GetKontoInfo(intGegenKonto.ToString), "{>}")
            '                'Debug.Print("Konto-Info: " + strSteuerInfo(26))

            '                Try
            '                    booBookingok = True
            '                    Call KrBhg.SetVerteilung(intGegenKonto.ToString, strFibuText, dblNettoBetrag.ToString, strSteuerFeld, strBeBuEintrag)

            '                Catch ex As Exception
            '                    MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Verteilung " + SubRow("strKredSubText") + ", Konto " + SubRow("lngKto").ToString)
            '                    If (Err.Number And 65535) < 10000 Then
            '                        booBookingok = False
            '                    Else
            '                        booBookingok = True
            '                    End If

            '                End Try


            '                strSteuerFeld = String.Empty
            '                strBeBuEintrag = String.Empty

            '                'Status Sub schreiben
            '                'Application.DoEvents()

            '            Next


            '            Try
            '                booBookingok = True
            '                Call KrBhg.WriteBuchung()

            '            Catch ex As Exception
            '                If (Err.Number And 65535) < 10000 Then
            '                    MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Belegerstellung nicht möglich")
            '                    booBookingok = False
            '                Else
            '                    MessageBox.Show(ex.Message, "Warnung " + (Err.Number And 65535).ToString + " Belegerstellung")
            '                    booBookingok = True
            '                End If

            '            End Try

            '            'Application.DoEvents()

            '            strBeBuEintrag = String.Empty
            '            strSteuerFeld = String.Empty
            '            dblNettoBetrag = 0
            '            dblMwStBetrag = 0
            '            dblBruttoBetrag = 0

            '        Else

            '            'Buchung nur in Fibu
            '            'Prinzip Funktion WriteBuchung() anwenden mit allen Parametern
            '            'Beleg-Nummerierung aktivieren
            '            'DbBhg.IncrBelNbr = "J"
            '            'Belegsnummer abholen
            '            intKredBelegsNummer = FBhg.GetNextBelNbr()

            '            'Prüfen, ob wirklich frei
            '            intReturnValue = 10
            '            Do Until intReturnValue = 0
            '                intReturnValue = FBhg.doesBelegExist(intKredBelegsNummer,
            '                                                     "NOT_SET",
            '                                                     "NOT_SET",
            '                                                     Strings.Left(frmImportMain.lstBoxPerioden.Text, 4) + "0101",
            '                                                     Strings.Left(frmImportMain.lstBoxPerioden.Text, 4) + "1231")
            '                If intReturnValue <> 0 Then
            '                    intKredBelegsNummer += 1
            '                End If
            '            Loop

            '            booBookingok = True

            '            'Variablen zuweisen
            '            strBelegDatum = Format(row("datKredRGDatum"), "yyyyMMdd").ToString
            '            strValutaDatum = Format(row("datKredValDatum"), "yyyyMMdd").ToString
            '            'strDebiText = row("strDebText")
            '            strCurrency = row("strKredCur")
            '            If strCurrency <> "CHF" Then 'Muss ergänzt werden => Was ist Leitwährung auf dem Konto
            '                dblKurs = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg)
            '            Else
            '                dblKurs = 1.0#
            '            End If

            '            selKrediSub = dsKreditoren.Tables("tblKrediSubsFromUser").Select("lngKredID=" + row("lngKredID").ToString)

            '            If selKrediSub.Length = 2 Then

            '                For Each SubRow As DataRow In selKrediSub

            '                    If SubRow("intSollHaben") = 0 Then 'Soll

            '                        intSollKonto = SubRow("lngKto")
            '                        dblKursSoll = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg, intSollKonto)
            '                        'strSteuerInfo = Split(FBhg.GetKontoInfo(intSollKonto.ToString), "{>}")
            '                        'Debug.Print("Konto-Info Soll: " + strSteuerInfo(26))
            '                        dblSollBetrag = SubRow("dblNetto")
            '                        strKrediTextSoll = SubRow("strKredSubText")
            '                        If SubRow("dblMwSt") > 0 Then
            '                            strSteuerFeldSoll = Main.FcGetSteuerFeld(FBhg,
            '                                                                     SubRow("lngKto"),
            '                                                                     strKrediTextSoll,
            '                                                                     SubRow("dblBrutto") * dblKursSoll,
            '                                                                     SubRow("strMwStKey"),
            '                                                                     SubRow("dblMwSt"))
            '                        End If
            '                        If SubRow("lngKST") > 0 Then
            '                            strBeBuEintragSoll = SubRow("lngKST").ToString + "{<}" + strKrediTextSoll + "{<}" + "CALCULATE" + "{>}"
            '                        End If


            '                    ElseIf SubRow("intSollHaben") = 1 Then 'Haben

            '                        intHabenKonto = SubRow("lngKto")
            '                        dblKursHaben = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg, intHabenKonto)
            '                        'strSteuerInfo = Split(FBhg.GetKontoInfo(intSollKonto.ToString), "{>}")
            '                        'Debug.Print("Konto-Info Haben: " + strSteuerInfo(26))
            '                        dblHabenBetrag = SubRow("dblNetto") * -1
            '                        'dblHabenBetrag = dblSollBetrag
            '                        strKrediTextHaben = SubRow("strKredSubText")
            '                        If SubRow("dblMwSt") > 0 Then
            '                            strSteuerFeldHaben = Main.FcGetSteuerFeld(FBhg,
            '                                                                      SubRow("lngKto"),
            '                                                                      strKrediTextHaben,
            '                                                                      SubRow("dblBrutto") * dblKursHaben * -1,
            '                                                                      SubRow("strMwStKey"),
            '                                                                      SubRow("dblMwSt") * -1)
            '                        End If
            '                        If SubRow("lngKST") > 0 Then
            '                            strBeBuEintragHaben = SubRow("lngKST").ToString + "{<}" + strKrediTextHaben + "{<}" + "CALCULATE" + "{>}"
            '                        End If

            '                    Else

            '                        'Sammelbeleg
            '                        MsgBox("Nicht definierter Wert Sub-Buchungs-SollHaben: " + SubRow("intSollHaben").ToString)
            '                        'strKrediText = IIf(IsDBNull(row("strKredText")), "", row("strKredText"))

            '                    End If

            '                Next

            '                'Tieferer Betrag für die Gesamt-Buchung herausfinden
            '                If dblSollBetrag <= dblHabenBetrag Then
            '                    dblNettoBetrag = dblSollBetrag
            '                ElseIf dblHabenBetrag < dblSollBetrag Then
            '                    dblNettoBetrag = dblHabenBetrag
            '                End If

            '                'Buchung ausführen
            '                Call FBhg.WriteBuchung(0,
            '                                       intKredBelegsNummer,
            '                                       strBelegDatum,
            '                                       intSollKonto.ToString,
            '                                       strKrediTextSoll,
            '                                       strCurrency,
            '                                       dblKursSoll.ToString,
            '                                       (dblNettoBetrag * dblKursSoll).ToString,
            '                                       strSteuerFeldSoll,
            '                                       intHabenKonto.ToString,
            '                                       strKrediTextHaben,
            '                                       strCurrency,
            '                                       dblKursHaben.ToString,
            '                                       (dblNettoBetrag * dblKursHaben).ToString,
            '                                       strSteuerFeldHaben,
            '                                       strCurrency,
            '                                       dblKurs.ToString,
            '                                       dblNettoBetrag.ToString,
            '                                       (dblNettoBetrag * dblKurs).ToString,
            '                                       strBeBuEintragSoll,
            '                                       strBeBuEintragHaben,
            '                                       strValutaDatum)

            '            Else
            '                MsgBox("Nicht 2 Subbuchungen.")
            '            End If

            '        End If

            '        If booBookingok Then
            '            If row("booPGV") Then
            '                'Bei PGV Buchungen
            '                If IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")) = "" Or
            '                   (IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")) = "RV" And row("intPGVMthsAY") + row("intPGVMthsNY") > 1) Then

            '                    intReturnValue = MainKreditor.FcPGVKTreatment(FBhg,
            '                                                           Finanz,
            '                                                           DbBhg,
            '                                                           PIFin,
            '                                                           BeBu,
            '                                                           KrBhg,
            '                                                           dsKreditoren.Tables("tblKrediSubsFromUser"),
            '                                                           row("lngKredID"),
            '                                                           intKredBelegsNummer,
            '                                                           row("strKredCur"),
            '                                                           row("datKredValDatum"),
            '                                                           "M",
            '                                                           row("datPGVFrom"),
            '                                                           row("datPGVTo"),
            '                                                           row("intPGVMthsAY") + row("intPGVMthsNY"),
            '                                                           row("intPGVMthsAY"),
            '                                                           row("intPGVMthsNY"),
            '                                                           2311,
            '                                                           2312,
            '                                                           frmImportMain.lstBoxPerioden.Text,
            '                                                           objdbConn,
            '                                                           objdbMSSQLConn,
            '                                                           objdbSQLcommand,
            '                                                           frmImportMain.lstBoxMandant.SelectedValue,
            '                                                           dsKreditoren.Tables("tblKreditorenInfo"),
            '                                                           strYear,
            '                                                           intTeqNbr,
            '                                                           intTeqNbrLY,
            '                                                           intTeqNbrPLY,
            '                                                           IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")),
            '                                                           datPeriodFrom,
            '                                                           datPeriodTo,
            '                                                           strPeriodStatus)

            '                Else

            '                    'TP
            '                    intReturnValue = MainKreditor.FcPGVKTreatmentYC(FBhg,
            '                                                           Finanz,
            '                                                           DbBhg,
            '                                                           PIFin,
            '                                                           BeBu,
            '                                                           KrBhg,
            '                                                           dsKreditoren.Tables("tblKrediSubsFromUser"),
            '                                                           row("lngKredID"),
            '                                                           intKredBelegsNummer,
            '                                                           row("strKredCur"),
            '                                                           row("datKredValDatum"),
            '                                                           "M",
            '                                                           row("datPGVFrom"),
            '                                                           row("datPGVTo"),
            '                                                           row("intPGVMthsAY") + row("intPGVMthsNY"),
            '                                                           row("intPGVMthsAY"),
            '                                                           row("intPGVMthsNY"),
            '                                                           2311,
            '                                                           2312,
            '                                                           frmImportMain.lstBoxPerioden.Text,
            '                                                           objdbConn,
            '                                                           objdbMSSQLConn,
            '                                                           objdbSQLcommand,
            '                                                           frmImportMain.lstBoxMandant.SelectedValue,
            '                                                           dsKreditoren.Tables("tblKreditorenInfo"),
            '                                                           strYear,
            '                                                           intTeqNbr,
            '                                                           intTeqNbrLY,
            '                                                           intTeqNbrPLY,
            '                                                           IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")),
            '                                                           datPeriodFrom,
            '                                                           datPeriodTo,
            '                                                           strPeriodStatus)


            '                End If


            '            End If

            '            'Status Head schreiben
            '            row("strKredBookStatus") = row("strKredStatusBitLog")
            '            row("booBooked") = True
            '            row("datBooked") = Now()
            '            row("lngBelegNr") = intKredBelegsNummer

            '            dsKreditoren.Tables("tblKrediHeadsFromUser").AcceptChanges()
            '            'strKRGReferTo = "lngKredID"
            '            'strKRGReferTo = "strKredRGNbr"
            '            If objdbConn.State = ConnectionState.Closed Then
            '                objdbConn.Open()
            '            End If
            '            strKRGReferTo = Main.FcReadFromSettings(objdbConn, "Buchh_TableKRGReferTo", frmImportMain.lstBoxMandant.SelectedValue)
            '            If objdbConn.State = ConnectionState.Open Then
            '                objdbConn.Close()
            '            End If
            '            'Status in File RG-Tabelle schreiben
            '            intReturnValue = MainKreditor.FcWriteToKrediRGTable(frmImportMain.lstBoxMandant.SelectedValue,
            '                                                            row(strKRGReferTo),
            '                                                            row("datBooked"),
            '                                                            row("lngBelegNr"),
            '                                                            objOracleConn,
            '                                                            objdbConn)
            '            If intReturnValue <> 0 Then
            '                'Throw an exception
            '                MessageBox.Show("Achtung, Beleg-Nummer: " + row("lngBelegNr").ToString + " konnte nicht In die RG-Tabelle geschrieben werden auf RG-ID: " + row("lngKredID").ToString + ".", "RG-Table Update nicht möglich", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            '            End If

            '        End If

            '    End If

            'Next

            ''In sync-Tabelle schreiben
            'intReturnValue = WFDBClass.FcWriteEndToSync(objdbConn,
            '                                            frmImportMain.lstBoxMandant.SelectedValue,
            '                                            2,
            '                                            0,
            '                                            IIf(booBookingok, "ok", "Probleme"))


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString)
            Err.Clear()

        Finally
            'Neu aufbauen
            'butKreditoren_Click(butDebitoren, EventArgs.Empty)
            BgWImportKrediLocArgs = Nothing
            Cursor = Cursors.Default
            'Me.butImportK.Enabled = True
            'Me.Close()

        End Try


    End Sub


    Private Sub BgWLoadKredi_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BgWLoadKredi.DoWork

        Dim strIdentityName As String
        Dim strMDBName As String
        Dim strSQL As String
        Dim strSQLSub As String
        Dim strKRGTableType As String
        Dim objdtLocKrediHead As New DataTable
        Dim objdtLocKrediSubs As New DataTable
        Dim objdaolelocKrediSubs As New OleDb.OleDbDataAdapter
        Dim objdaolelocKrediHeads As New OleDb.OleDbDataAdapter
        Dim objdaolesubsselcomd As New OleDb.OleDbCommand
        Dim objdaoleheadselcomd As New OleDb.OleDbCommand
        Dim objdalocKrediSubs As New MySqlDataAdapter
        Dim objdalocKrediHeads As New MySqlDataAdapter
        Dim objdasubselcomd As New MySqlCommand
        Dim objdaheadselcomd As New MySqlCommand
        Dim objdslocKredisub As New DataSet
        Dim objdslocKredihead As New DataSet
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

            intFcReturns = FcReadFromSettingsIII("Buchh_KRGTableMDB",
                                              intAccounting,
                                              strMDBName)

            intFcReturns = FcReadFromSettingsIII("Buchh_SQLHeadKred",
                                          intAccounting,
                                          strSQL)

            intFcReturns = FcReadFromSettingsIII("Buchh_KRGTableType",
                                                   intAccounting,
                                                   strKRGTableType)

            objdslocKredihead.EnforceConstraints = False


            If strKRGTableType = "A" Then

                'Access
                Call FcInitAccessConnecation(objdbAccessConn,
                                                  strMDBName)
                objdaolesubsselcomd.Connection = objdbAccessConn
                objdaolesubsselcomd.CommandText = strSQL
                objdaolelocKrediHeads.SelectCommand = objdaolesubsselcomd
                objdaolelocKrediHeads.SelectCommand.Connection.Open()
                objdaolelocKrediHeads.Fill(objdslocKredihead, "tblkredihead")
                objdaolelocKrediHeads.SelectCommand.Connection.Close()
                'objdbAccessConn.Open()
                'objOLEdbcmdLoc.CommandText = strSQL
                'objOLEdbcmdLoc.Connection = objdbAccessConn
                'objdtLocKrediHead.Load(objOLEdbcmdLoc.ExecuteReader)
                'objdbAccessConn.Close()
            ElseIf strKRGTableType = "M" Then

                strConnection = System.Configuration.ConfigurationManager.AppSettings(strMDBName)
                objRGMySQLConn.ConnectionString = strConnection
                objdaheadselcomd.Connection = objRGMySQLConn
                objdaheadselcomd.CommandText = strSQL
                objdalocKrediHeads.SelectCommand = objdaheadselcomd
                objdalocKrediHeads.SelectCommand.Connection.Open()
                objdalocKrediHeads.Fill(objdslocKredihead, "tblkredihead")
                objdalocKrediHeads.SelectCommand.Connection.Close()
                'objlocMySQLcmd.Connection = objRGMySQLConn
                'objlocMySQLcmd.CommandText = strSQL
                'objRGMySQLConn.Open()
                'objdtLocKrediHead.Load(objlocMySQLcmd.ExecuteReader)
                'objRGMySQLConn.Close()
            Else
                MessageBox.Show("Tabletype not A or M")
                Exit Sub
            End If
            'objdslocKredihead.AcceptChanges()

            objdslocKredisub.EnforceConstraints = False

            intFcReturns = FcReadFromSettingsIII("Buchh_SQLDetailKred",
                                                    intAccounting,
                                                    strSQLToParse)

            intFcReturns = FcInitInsCmdKHeads(objmysqlcomdwritehead)

            For Each row As DataRow In objdslocKredihead.Tables("tblkredihead").Rows

                objmysqlcomdwritehead.Connection.Open()
                objmysqlcomdwritehead.Parameters("@IdentityName").Value = strIdentityName
                objmysqlcomdwritehead.Parameters("@ProcessID").Value = Process.GetCurrentProcess().Id
                objmysqlcomdwritehead.Parameters("@intBuchhaltung").Value = intAccounting
                objmysqlcomdwritehead.Parameters("@strKredRGNbr").Value = row("strKredRGNbr")
                objmysqlcomdwritehead.Parameters("@intBuchungsart").Value = row("intBuchungsart")
                objmysqlcomdwritehead.Parameters("@lngKredID").Value = row("lngKredID")
                objmysqlcomdwritehead.Parameters("@strOPNr").Value = row("strOPNr")
                objmysqlcomdwritehead.Parameters("@lngKredNbr").Value = row("lngKredNbr")
                objmysqlcomdwritehead.Parameters("@lngKredKtoNbr").Value = row("lngKredKtoNbr")
                objmysqlcomdwritehead.Parameters("@strKredCur").Value = row("strKredCur")
                objmysqlcomdwritehead.Parameters("@lngKrediKST").Value = row("lngKrediKST")
                objmysqlcomdwritehead.Parameters("@dblKredNetto").Value = row("dblKredNetto")
                objmysqlcomdwritehead.Parameters("@dblKredMwSt").Value = row("dblKredMwSt")
                objmysqlcomdwritehead.Parameters("@dblKredBrutto").Value = row("dblKredBrutto")
                objmysqlcomdwritehead.Parameters("@lngKredIdentNbr").Value = row("lngKredIdentNbr")
                objmysqlcomdwritehead.Parameters("@strKredText").Value = row("strKredText")
                objmysqlcomdwritehead.Parameters("@strKredRef").Value = row("strKredRef")
                objmysqlcomdwritehead.Parameters("@datKredRGDatum").Value = row("datKredRGDatum")
                objmysqlcomdwritehead.Parameters("@datKredValDatum").Value = row("datKredValDatum")
                objmysqlcomdwritehead.Parameters("@intPayType").Value = row("intPayType")
                objmysqlcomdwritehead.Parameters("@strKrediBank").Value = row("strKrediBank")
                objmysqlcomdwritehead.Parameters("@strKrediBankInt").Value = row("strKrediBankInt")
                objmysqlcomdwritehead.Parameters("@strRGName").Value = row("strRGName")
                objmysqlcomdwritehead.Parameters("@strRGBemerkung").Value = row("strRGBemerkung")
                objmysqlcomdwritehead.Parameters("@intZKond").Value = row("intZKond")
                If row.Table.Columns.Contains("datPGVFrom") Then
                    objmysqlcomdwritehead.Parameters("@datPGVFrom").Value = row("datPGVFrom")
                End If
                If row.Table.Columns.Contains("datPGVTo") Then
                    objmysqlcomdwritehead.Parameters("@datPGVTo").Value = row("datPGVTo")
                End If

                objmysqlcomdwritehead.ExecuteNonQuery()
                objmysqlcomdwritehead.Connection.Close()


                'Subs einlesen
                strSQLSub = FcSQLParseKredi(strSQLToParse,
                                                row("lngKredID"),
                                                objdslocKredihead.Tables("tblkredihead"))

                If strKRGTableType = "A" Then
                    objdaolesubsselcomd.CommandText = strSQLSub
                    objdaolesubsselcomd.Connection = objdbAccessConn
                    objdaolelocKrediSubs.SelectCommand = objdaolesubsselcomd
                    objdaolelocKrediSubs.SelectCommand.Connection.Open()
                    objdaolelocKrediSubs.Fill(objdslocKredisub, "tblkredisubs")
                    objdaolelocKrediSubs.SelectCommand.Connection.Close()
                    'objdbAccessConn.Open()
                    'objOLEdbcmdLoc.CommandText = strSQLSub
                    'objdtLocKrediSubs.Load(objOLEdbcmdLoc.ExecuteReader)
                    'objdbAccessConn.Close()
                ElseIf strKRGTableType = "M" Then
                    objdasubselcomd.CommandText = strSQLSub
                    objdasubselcomd.Connection = objRGMySQLConn
                    objdalocKrediSubs.SelectCommand = objdasubselcomd
                    objdalocKrediSubs.SelectCommand.Connection.Open()
                    objdalocKrediSubs.Fill(objdslocKredisub, "tblkredisubs")
                    objdalocKrediSubs.SelectCommand.Connection.Close()
                    'objlocMySQLcmd.CommandText = strSQLSub
                    'objRGMySQLConn.Open()
                    'objdtLocKrediSubs.Load(objlocMySQLcmd.ExecuteReader)
                    'objRGMySQLConn.Close()
                End If

            Next
            objdslocKredisub.AcceptChanges()

            If Not IsNothing(objdslocKredisub.Tables("tblkredisubs")) Then

                'Subs schreiben
                intFcReturns = Main.FcInitInscmdKSubs(objmysqlcomdwritesub)
                For Each drsub As DataRow In objdslocKredisub.Tables("tblkredisubs").Rows

                    objmysqlcomdwritesub.Connection.Open()
                    objmysqlcomdwritesub.Parameters("@IdentityName").Value = strIdentityName
                    objmysqlcomdwritesub.Parameters("@ProcessID").Value = Process.GetCurrentProcess().Id
                    objmysqlcomdwritesub.Parameters("@lngKredID").Value = drsub("lngKredID")
                    objmysqlcomdwritesub.Parameters("@lngKto").Value = drsub("lngKto")
                    objmysqlcomdwritesub.Parameters("@lngKST").Value = drsub("lngKST")
                    objmysqlcomdwritesub.Parameters("@dblNetto").Value = IIf(IsDBNull(drsub("dblNetto")), 0, drsub("dblNetto"))
                    objmysqlcomdwritesub.Parameters("@dblMwSt").Value = IIf(IsDBNull(drsub("dblMwSt")), 0, drsub("dblMwSt"))
                    objmysqlcomdwritesub.Parameters("@dblBrutto").Value = IIf(IsDBNull(drsub("dblBrutto")), 0, drsub("dblBrutto"))
                    objmysqlcomdwritesub.Parameters("@dblMwStSatz").Value = drsub("dblMwStSatz")
                    objmysqlcomdwritesub.Parameters("@strMwStKey").Value = drsub("strMwStKey")
                    objmysqlcomdwritesub.Parameters("@intSollHaben").Value = drsub("intSollHaben")
                    If objdtLocKrediSubs.Columns.Contains("strKredSubText") Then
                        objmysqlcomdwritesub.Parameters("@strKredSubText").Value = drsub("strKredSubText")
                    End If
                    objmysqlcomdwritesub.Parameters("@booRebilling").Value = drsub("booRebilling")
                    objmysqlcomdwritesub.ExecuteNonQuery()
                    objmysqlcomdwritesub.Connection.Close()

                Next

            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Kreditoren-Daten-Lesen Problem", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Err.Clear()

        Finally
            objdbAccessConn = Nothing
            objRGMySQLConn = Nothing

            objdslocKredihead = Nothing
            objdslocKredisub = Nothing

            objdalocKrediHeads = Nothing
            objdalocKrediSubs = Nothing

            objdaolelocKrediHeads = Nothing
            objdaolelocKrediSubs = Nothing

            objdaoleheadselcomd = Nothing
            objdaheadselcomd = Nothing
            objdaolesubsselcomd = Nothing
            objdasubselcomd = Nothing

            strConnection = Nothing

        End Try


    End Sub

    Private Sub BgWCheckKredi_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BgWCheckKredi.DoWork

        Dim BgWCheckKrediArgsInProc As BgWCheckDebitArgs = e.Argument

        Dim strMandant As String
        Dim booAccOk As Boolean
        'Dim objFinanz As New SBSXASLib.AXFinanz
        'Dim objfiBuha As New SBSXASLib.AXiFBhg
        'Dim objKrBuha As New SBSXASLib.AXiKrBhg
        'Dim objFiBebu As New SBSXASLib.AXiBeBu

        Dim booAutoCorrect As Boolean
        Dim booCpyKSTToSub As Boolean
        Dim intReturnValue As Int16
        Dim intKreditorNew As Int32
        Dim strKreditorNew As String
        Dim strBitLog As String
        Dim intSubNumber As Int32
        Dim dblSubBrutto As Decimal
        Dim dblSubNetto As Decimal
        Dim dblSubMwSt As Decimal
        Dim dblRDiffBrutto As Decimal
        Dim dblRDiffMwSt As Decimal
        Dim dblRDiffNetto As Decimal
        Dim strCleanOPNbr As String
        Dim strKredTyp As String
        Dim strStatus As String
        Dim datValutaSave As Date
        Dim intMonthsAJ As Int16
        Dim intMonthsNJ As Int16
        Dim intPGVMonths As Int16
        Dim intintBank As Int32
        Dim intPayType As Int16
        Dim booPKPrivate As Boolean
        Dim strIBANToPass As String
        Dim booDiffHeadText As Boolean
        Dim strKrediHeadText As String
        Dim booLeaveSubText As Boolean
        Dim booDiffSubText As Boolean
        Dim strKrediSubText As String
        Dim selsubrow() As DataRow
        Dim intFcReturns As Int16
        Dim strFcReturns As String

        Try

            'Finanz-Obj init
            'Login
            Call objFinanz.ConnectSBSdb(System.Configuration.ConfigurationManager.AppSettings("OwnSageServer"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageDB"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageID"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSagePsw"), "")

            intFcReturns = FcReadFromSettingsIII("Buchh200_Name",
                                            BgWCheckKrediArgsInProc.intMandant,
                                            strMandant)

            booAccOk = objFinanz.CheckMandant(strMandant)
            'Open Mandantg
            objFinanz.OpenMandant(strMandant, BgWCheckKrediArgsInProc.strPeriode)

            objfiBuha = objFinanz.GetFibuObj()
            objKrBuha = objFinanz.GetKrediObj()
            objFiBebu = objFinanz.GetBeBuObj()

            'Variablen einesen
            booAutoCorrect = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettingsII("Buchh_HeadKAutoCorrect", BgWCheckKrediArgsInProc.intMandant)))
            booCpyKSTToSub = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettingsII("Buchh_KKSTHeadToSub", BgWCheckKrediArgsInProc.intMandant)))
            booPKPrivate = IIf(FcReadFromSettingsII("Buchh_PKKrediTable", BgWCheckKrediArgsInProc.intMandant) = "t_customer", True, False)
            booDiffHeadText = IIf(FcReadFromSettingsII("Buchh_KTextSpecial", BgWCheckKrediArgsInProc.intMandant) = "0", False, True)
            booLeaveSubText = IIf(FcReadFromSettingsII("Buchh_KSubLeaveText", BgWCheckKrediArgsInProc.intMandant) = "0", False, True)
            booDiffSubText = IIf(FcReadFromSettingsII("Buchh_KSubTextSpecial", BgWCheckKrediArgsInProc.intMandant) = "0", False, True)

            For Each row As DataRow In dsKreditoren.Tables("tblKrediHeadsFromUser").Rows

                'If row("lngKredID") = "117383" Then Stop
                'Runden
                row("dblKredNetto") = Decimal.Round(row("dblKredNetto"), 2, MidpointRounding.AwayFromZero)
                row("dblKredMwSt") = Decimal.Round(row("dblKredMwst"), 2, MidpointRounding.AwayFromZero)
                row("dblKredBrutto") = Decimal.Round(row("dblKredBrutto"), 2, MidpointRounding.AwayFromZero)

                'Status-String erstellen
                'Kreditor 01
                intReturnValue = FcGetRefKrediNr(IIf(IsDBNull(row("lngKredNbr")), 0, row("lngKredNbr")),
                                                            BgWCheckKrediArgsInProc.intMandant,
                                                            intKreditorNew)

                If intKreditorNew <> 0 Then
                    intReturnValue = FcCheckKreditor(intKreditorNew,
                                                                  row("intBuchungsart"))
                Else
                    intReturnValue = 2
                End If
                strBitLog = Trim(intReturnValue.ToString)

                'Kto 02
                'intReturnValue = FcCheckKonto(row("lngKredKtoNbr"), objfiBuha, row("dblKredMwSt"), 0)
                intReturnValue = 0
                strBitLog += Trim(intReturnValue.ToString)

                'Currency 03
                intReturnValue = FcCheckCurrency(row("strKredCur"))
                strBitLog += Trim(intReturnValue.ToString)

                'Sub 04
                intReturnValue = FcCheckKrediSubBookings2(row("lngKredID"),
                                                         dsKreditoren.Tables("tblKrediSubsFromUser"),
                                                         intSubNumber,
                                                         dblSubBrutto,
                                                         dblSubNetto,
                                                         dblSubMwSt,
                                                         IIf(IsDBNull(row("datKredValDatum")), #1789-09-17#, row("datKredValDatum")),
                                                         row("intBuchungsart"),
                                                         booAutoCorrect,
                                                         booCpyKSTToSub,
                                                         row("lngKrediKST"),
                                                         row("intPayType"),
                                                         IIf(IsDBNull(row("strKrediBank")), "", row("strKrediBank")))

                strBitLog += Trim(intReturnValue.ToString)

                'Autokorrektur 05
                If booAutoCorrect Then
                    'Git es etwas zu korrigieren?
                    If Math.Abs(IIf(IsDBNull(row("dblKredBrutto")), 0, row("dblKredBrutto")) - dblSubBrutto) < 0.1 Then
                        If IIf(IsDBNull(row("dblKredNetto")), 0, row("dblKredNetto")) <> dblSubNetto Or
                            IIf(IsDBNull(row("dblKredMwSt")), 0, row("dblKredMwSt")) <> dblSubMwSt Then
                            'IIf(IsDBNull(row("dblKredBrutto")), 0, row("dblKredBrutto")) <> dblSubBrutto Or
                            'row("dblKredBrutto") = Decimal.Round(dblSubBrutto, 2, MidpointRounding.AwayFromZero)
                            'Limit korrektur setzen 1 Fr.
                            'If Math.Abs(IIf(IsDBNull(row("dblKredNetto")), 0, row("dblKredNetto")) - dblSubNetto) > 1 Or
                            '   Math.Abs(IIf(IsDBNull(row("dblKredMwSt")), 0, row("dblKredMwSt")) - dblSubMwSt) > 1 Then
                            '    'Nicht korrigieren
                            '    strBitLog += "3"
                            'Else
                            row("dblKredBrutto") = Decimal.Round(dblSubBrutto, 2, MidpointRounding.AwayFromZero)
                            row("dblKredNetto") = Decimal.Round(dblSubNetto, 2, MidpointRounding.AwayFromZero)
                            row("dblKredMwSt") = Decimal.Round(dblSubMwSt, 2, MidpointRounding.AwayFromZero)
                            strBitLog += "1"
                            'End If
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
                        strBitLog += "3"
                    End If
                    dsKreditoren.Tables("tblKrediHeadsFromUser").AcceptChanges()
                Else
                    If row("intBuchungsart") = 1 Then

                        dblRDiffBrutto = 0
                        If IIf(IsDBNull(row("dblKredMwSt")), 0, row("dblKredMwSt")) <> dblSubMwSt Then
                            row("dblKredMwSt") = dblSubMwSt
                        End If
                        If IIf(IsDBNull(row("dblKredNetto")), 0, row("dblKredNetto")) <> dblSubNetto Then
                            row("dblKredNetto") = dblSubNetto
                        End If

                        'Für evtl. Rundungsdifferenzen einen Datensatz in die Sub-Tabelle hinzufügen
                        If IIf(IsDBNull(row("dblKredBrutto")), 0, row("dblKredBrutto")) - dblSubBrutto <> 0 Then

                            dblRDiffBrutto = Decimal.Round(dblSubBrutto - row("dblKredBrutto"), 2, MidpointRounding.AwayFromZero) * -1
                            dblRDiffMwSt = 0
                            dblRDiffNetto = 0

                            'Zu Sub-Tabelle hinzufügen
                            Dim objdrKrediSub As DataRow = dsKreditoren.Tables("tblKrediSubsFromUser").NewRow
                            objdrKrediSub("lngKredID") = row("lngKredID")
                            objdrKrediSub("intSollHaben") = 1
                            objdrKrediSub("lngKto") = 6906
                            objdrKrediSub("strKtoBez") = "Rundungsdifferenzen"
                            objdrKrediSub("lngKST") = 40
                            objdrKrediSub("strKstBez") = "SystemKST"
                            objdrKrediSub("dblNetto") = dblRDiffNetto
                            objdrKrediSub("dblMwSt") = dblRDiffMwSt
                            objdrKrediSub("dblBrutto") = dblRDiffBrutto
                            objdrKrediSub("dblMwStSatz") = 0
                            objdrKrediSub("strMwStKey") = "null"
                            objdrKrediSub("strArtikel") = "Rundungsdifferenz"
                            objdrKrediSub("strKredSubText") = "Rundung"
                            objdrKrediSub("booRebilling") = True
                            objdrKrediSub("strStatusUBBitLog") = "00000000"
                            If Math.Abs(dblRDiffBrutto) > 1 Then
                                objdrKrediSub("strStatusUBText") = "Rund > 1"
                            Else
                                objdrKrediSub("strStatusUBText") = "ok"
                            End If
                            dsKreditoren.Tables("tblKrediSubsFromUser").Rows.Add(objdrKrediSub)

                            'dsKreditoren.Tables("tblKrediSubsFromUser").AcceptChanges()

                            'Summe SubBuchung anpassen
                            dblSubBrutto = Decimal.Round(dblSubBrutto + dblRDiffBrutto, 2, MidpointRounding.AwayFromZero)
                            If Math.Abs(dblRDiffBrutto) > 1 Then
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
                    'strBitLog += "0"
                End If

                'Diff Kopf - Sub? 06
                If row("intBuchungsart") = 1 Then 'OP
                    If IIf(IsDBNull(row("dblKredBrutto")), 0, row("dblKredBrutto")) - dblSubBrutto <> 0 _
                        Or IIf(IsDBNull(row("dblKredMwSt")), 0, row("dblKredMwSt")) - dblSubMwSt <> 0 _
                        Or IIf(IsDBNull(row("dblKredNetto")), 0, row("dblKredNetto")) - dblSubNetto <> 0 Then
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
                                                  IIf(IsDBNull(row("dblKredBrutto")), 0, row("dblKredBrutto")),
                                                  IIf(IsDBNull(row("dblKredNetto")), 0, row("dblKredNetto")),
                                                  IIf(IsDBNull(row("dblKredMwSt")), 0, row("dblKredMwSt")),
                                                  dblRDiffBrutto)
                strBitLog += Trim(intReturnValue.ToString)

                'OP - Nummer prüfen 08
                'intReturnValue = FcCreateDebRef(objdbconn, intAccounting, row("strDebiBank"), row("strDebRGNbr"), row("intBuchungsart"), strDebiReferenz)
                strCleanOPNbr = IIf(IsDBNull(row("strOPNr")), "", row("strOPNr"))
                intReturnValue = FcChCeckKredOP(strCleanOPNbr,
                                                IIf(IsDBNull(row("strKredRGNbr")), "", row("strKredRGNbr")))
                row("strOPNr") = strCleanOPNbr
                strBitLog += Trim(intReturnValue.ToString)
                'OP - Verdopplung 09
                If row("dblKredBrutto") < 0 Then
                    strKredTyp = "G"
                Else
                    strKredTyp = "R"
                End If
                intReturnValue = FcCheckKrediOPDouble(intKreditorNew,
                                                      row("strKredRGNbr"),
                                                      row("strKredCur"),
                                                      strKredTyp)

                strBitLog += Trim(intReturnValue.ToString)

                'PGV => Prüfung vor Valuta-Datum da Valuta-Datum verändert wird. PGV soll nur möglich sein wenn rebilled
                If Not IsDBNull(row("datPGVFrom")) And FcIsAllKrediRebilled(dsKreditoren.Tables("tblKrediSubsFromUser"), row("lngKredID")) = 0 Then
                    row("booPGV") = True
                ElseIf Not IsDBNull(row("datPGVFrom")) And FcIsAllKrediRebilled(dsKreditoren.Tables("tblKrediSubsFromUser"), row("lngKredID")) = 1 Then
                    row("strPGVType") = "XX"
                End If

                'Bei Datum-Korrekur vorgängig Datum ersetzen um PGV-Buchung zu verhindern
                If BgWCheckKrediArgsInProc.booValutaCor Then
                    If row("datKredRGDatum") < BgWCheckKrediArgsInProc.datValutaCor Then
                        row("datKredRGDatum") = BgWCheckKrediArgsInProc.datValutaCor.ToShortDateString
                        strStatus = "RgDCor"
                    End If
                    If IIf(IsDBNull(row("datKredValDatum")), #1789-09-17#, row("datKredValDatum")) < BgWCheckKrediArgsInProc.datValutaCor Then
                        row("datKredValDatum") = BgWCheckKrediArgsInProc.datValutaCor.ToShortDateString
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "ValDCor"
                    End If
                End If

                'Jahresübergreifend RG- / Valuta-Datum
                If Year(row("datKredRGDatum")) <> Year(IIf(IsDBNull(row("datKredValDatum")), #1789-09-17#, row("datKredValDatum"))) And Year(IIf(IsDBNull(row("datKredValDatum")), #1789-09-17#, row("datKredValDatum"))) >= 2023 Then

                    row("booPGV") = True
                    'datValutaPGV = row("datKredValDatum")
                    'Bei Valuta-Datum in einem anderen Jahr Valuta-Datum ändern
                    If Year(row("datKredRGDatum")) < Year(row("datKredValDatum")) Then
                        row("strPGVType") = "RV"
                    Else
                        row("strPGVType") = "VR"
                    End If
                    datValutaSave = row("datKredValDatum")

                    If IsDBNull(row("datPGVFrom")) Then
                        If row("strPGVType") = "VR" Then
                            row("datPGVFrom") = Year(datValutaSave).ToString + "-" + Month(datValutaSave).ToString + "-" + DateAndTime.Day(datValutaSave).ToString
                            row("datPGVTo") = Year(datValutaSave).ToString + "-" + Month(datValutaSave).ToString + "-" + DateAndTime.Day(datValutaSave).ToString
                            row("datKredValDatum") = "2024-01-01" ' Year(row("datKredRGDatum")).ToString + "-01-01"
                        ElseIf row("strPGVType") = "RV" Then
                            row("datPGVFrom") = Year(datValutaSave).ToString + "-" + Month(datValutaSave).ToString + "-" + DateAndTime.Day(datValutaSave).ToString
                            row("datPGVTo") = Year(datValutaSave).ToString + "-" + Month(datValutaSave).ToString + "-" + DateAndTime.Day(datValutaSave).ToString
                            row("datKredValDatum") = row("datKredRGDatum")
                        End If
                    Else
                        If row("strPGVType") = "RV" Then
                            row("datKredValDatum") = row("datKredRGDatum")
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
                        If Year(DateAdd(DateInterval.Month, intMonthCounter, row("datPGVFrom"))) > Convert.ToInt32(strYear) Then
                            intMonthsNJ += 1
                        Else
                            intMonthsAJ += 1
                        End If
                    Next
                    row("intPGVMthsAY") = intMonthsAJ
                    row("intPGVMthsNY") = intMonthsNJ

                End If

                'Valuta - Datum 10
                'Falls nichts ausgefüllt, dann 
                If IsDBNull(row("datKredValDatum")) Then
                    row("datKredValDatum") = row("datKredRGDatum")
                End If
                intReturnValue = FcCheckDate2(IIf(IsDBNull(row("datKredValDatum")), row("datKredRGDatum"), row("datKredValDatum")),
                                              strYear,
                                              dsKreditoren.Tables("tblKreditorenDates"),
                                              False)

                'Bei PGV checken ob PGV-Startdatum in blockierter Periode
                If row("booPGV") And intReturnValue = 0 Then
                    'Ist TP ?
                    If intMonthsAJ + intMonthsNJ = 1 Then
                        'Ist Differenz Jahre grösser 1?
                        If Math.Abs(Convert.ToInt16(strYear) - Year(row("datPGVTo"))) > 1 Then
                            intReturnValue = 4
                        Else
                            intReturnValue = FcCheckDate2(row("datPGVTo"),
                                                      strYear,
                                                      dsKreditoren.Tables("tblKreditorenDates"),
                                                      True)
                        End If
                    Else
                        'mehrere Monate PGV
                        For intMonthCounter = 0 To intPGVMonths - 1
                            'Ist Differenz Jahre grösser 1?
                            If Math.Abs(Convert.ToInt16(strYear) - Year(row("datPGVFrom"))) > 1 Then
                                intReturnValue = 4
                            Else
                                intReturnValue = FcCheckDate2(DateAndTime.DateAdd(DateInterval.Month, intMonthCounter, row("datPGVFrom")),
                                                          strYear,
                                                          dsKreditoren.Tables("tblKreditorenDates"),
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
                intReturnValue = FcCheckDate2(IIf(IsDBNull(row("datKredRGDatum")), #1789-09-17#, row("datKredRGDatum")),
                                              strYear,
                                              dsKreditoren.Tables("tblKreditorenDates"),
                                              False)

                strBitLog += Trim(intReturnValue.ToString)

                ''Referenz 12
                If IsDBNull(row("strKredRef")) Then
                    row("strKredRef") = ""
                    intReturnValue = 1
                Else
                    If (Not String.IsNullOrEmpty(row("strKredRef"))) And (row("intPayType") = 3 Or row("intPayType") = 10) Then
                        If Val(Strings.Left(row("strKredRef"), Len(row("strKredRef")) - 1)) > 0 Then

                            'Prüfziffer korrekt?
                            If Strings.Right(row("strKredRef"), 1) <> FcModulo10(Strings.Left(row("strKredRef"), Len(row("strKredRef")) - 1)) Then
                                intReturnValue = 2
                            Else
                                intReturnValue = 0
                            End If

                        Else
                            intReturnValue = 3
                        End If
                    Else
                        intReturnValue = 0
                    End If

                End If
                'Debug.Print("Erfasste Prüfziffer " + Right(row("strKredRef"), 1) + ", kalkuliert " + Main.FcModulo10(Left(row("strKredRef"), Len(row("strKredRef")) - 1)).ToString)
                strBitLog += Trim(intReturnValue.ToString)

                'interne Bank 13
                intReturnValue = FcCheckDebiIntBank(BgWCheckKrediArgsInProc.intMandant,
                                                         row("strKrediBankInt"),
                                                         intintBank)
                row("intintBank") = intintBank
                strBitLog += Trim(intReturnValue.ToString)

                'Buchungstext 14
                If IIf(IsDBNull(row("strKredText")), "", row("strKredText")) = "" Then
                    strBitLog += "1"
                Else
                    strBitLog += "0"
                End If

                'Zalungstyp logisch 15
                intPayType = IIf(IsDBNull(row("intPayType")), 0, row("intPayType"))
                intReturnValue = FcCheckPayType(intPayType,
                                                     IIf(IsDBNull(row("strKredRef")), "", row("strKredRef")),
                                                     IIf(IsDBNull(row("strKrediBank")), "", row("strKrediBank")))
                row("intPayType") = intPayType
                If intReturnValue >= 4 Then
                    strBitLog += Trim(intReturnValue.ToString)
                Else
                    strBitLog += "0"
                End If

                'Status-String auswerten
                'booPKPrivate = IIf(Main.FcReadFromSettingsII("Buchh_PKKrediTable", BgWCheckKrediArgsInProc.intMandant) = "t_customer", True, False)
                'Kreditor 1
                If Strings.Left(strBitLog, 1) <> "0" Then
                    strStatus += "Kred"
                    If Strings.Left(strBitLog, 1) <> "2" Then
                        If booPKPrivate Then
                            intReturnValue = FcIsPrivateKreditorCreatable(intKreditorNew,
                                                                                        IIf(IsDBNull(row("intPayType")), 3, row("intPayType")),
                                                                                        IIf(IsDBNull(row("strKredRef")), "", row("strKredRef")),
                                                                                        intintBank,
                                                                                        IIf(IsDBNull(row("strKrediBank")), "", row("strKrediBank")),
                                                                                        BgWCheckKrediArgsInProc.strMandant,
                                                                                        BgWCheckKrediArgsInProc.intMandant)
                        Else
                            intReturnValue = FcIsKreditorCreatable(intKreditorNew,
                                                                            BgWCheckKrediArgsInProc.strMandant,
                                                                            IIf(IsDBNull(row("intPayType")), 9, row("intPayType")),
                                                                            IIf(IsDBNull(row("strKredRef")), "", row("strKredRef")),
                                                                            intintBank,
                                                                            IIf(IsDBNull(row("strKrediBank")), "", row("strKrediBank")),
                                                                            BgWCheckKrediArgsInProc.intMandant)

                        End If
                        If intReturnValue = 0 Then
                            strStatus += " erstellt"
                            intReturnValue = FcReadKreditorName(strKreditorNew,
                                                                                intKreditorNew,
                                                                                row("strKredCur"))

                            row("strKredBez") = strKreditorNew
                        ElseIf intReturnValue = 5 Then
                            strStatus += " not approved"
                            row("strKredBez") = "nap"
                        ElseIf intReturnValue = 6 Then
                            strStatus += " AufwKto n/a"
                            row("strKredBez") = "Aufwandskonto n/a"
                        Else
                            strStatus += " nicht erstellt."
                            row("strKredBez") = "n/a"
                        End If
                        row("lngKredNbr") = intKreditorNew
                    Else
                        strStatus += " keine Ref"
                        row("strKredBez") = "n/a"
                    End If
                Else
                    intReturnValue = FcReadKreditorName(strKreditorNew,
                                                                        intKreditorNew,
                                                                        row("strKredCur"))
                    row("strKredBez") = strKreditorNew
                    row("lngKredNbr") = intKreditorNew
                    row("intEBank") = 0
                    If row("intPayType") = 9 Then
                        strIBANToPass = row("strKredRef")
                    ElseIf row("intPayType") = 10 Then
                        strIBANToPass = IIf(IsDBNull(row("strKrediBank")), "", row("strKrediBank"))
                    End If
                    If (row("intPayType") = 9 Or row("intPayType") = 10) And Len(strIBANToPass) > 0 Then
                        intReturnValue = FcCheckKreditBank(intKreditorNew,
                                                       IIf(IsDBNull(row("intPayType")), 9, row("intPayType")),
                                                       strIBANToPass,
                                                       IIf(IsDBNull(row("strKrediBank")), "", row("strKrediBank")),
                                                       row("strKredCur"),
                                                       row("intEBank"))
                    End If
                End If
                'Konto 2
                If Mid(strBitLog, 2, 1) <> "0" Then
                    If Mid(strBitLog, 2, 1) <> 2 Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Kto"
                    Else
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Kto MwSt"
                    End If
                    row("strKredKtoBez") = "n/a"
                Else
                    row("strKredKtoBez") = FcReadDebitorKName(row("lngKredKtoNbr"))
                End If
                'Währung 3
                If Mid(strBitLog, 3, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Cur"
                End If
                'Subbuchungen 4
                'Totale in Head schreiben
                row("intSubBookings") = intSubNumber.ToString
                row("dblSumSubBookings") = dblSubBrutto.ToString
                If Mid(strBitLog, 4, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Sub"
                End If
                'Autokorretkur 5
                If Mid(strBitLog, 5, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "AutoC"
                    If Mid(strBitLog, 5, 1) = "3" Then
                        strStatus += " >1"
                    End If
                End If
                'Diff zu Subbuchungen 6
                If Mid(strBitLog, 6, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "DiffS"
                End If
                'OP Kopf 7
                If Mid(strBitLog, 7, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "BelK"
                End If
                'OP Nummer 8
                If Mid(strBitLog, 8, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "OPNbr"
                End If
                'OP Doppelt 9
                If Mid(strBitLog, 9, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "OPDbl"
                    'Else
                    '   row("strDebRef") = strDebiReferenz
                End If
                'Valuta Datum 10
                If Mid(strBitLog, 10, 1) <> "0" Then
                    If Mid(strBitLog, 10, 1) = "1" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "ValD"
                    ElseIf Mid(strBitLog, 10, 1) = "2" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "VDBlck"
                        'Korrektur hat geklappt, Wert wieder auf 0 setzen
                        'strBitLog = Left(strBitLog, 9) + "0" + Right(strBitLog, Len(strBitLog) - 10)
                    ElseIf Mid(strBitLog, 10, 1) = "3" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "PGVBlck"
                    ElseIf Mid(strBitLog, 10, 1) = "4" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "PGVYear>1"
                        'ElseIf Mid(strBitLog, 10, 1) = "5" Then
                        '    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "PGVVDCor"
                        '    'Korrektur hat geklappt, Wert wieder auf 0 setzen
                        '    strBitLog = Left(strBitLog, 9) + "0" + Right(strBitLog, Len(strBitLog) - 10)
                    End If
                End If
                'RG Datum 11
                If Mid(strBitLog, 11, 1) <> "0" Then
                    If Mid(strBitLog, 11, 1) = "1" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "RgD"
                    ElseIf Mid(strBitLog, 11, 1) = "2" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "RgBlck"
                        'Korrektur hat geklappt, Wert wieder auf 0 setzen
                        'strBitLog = Left(strBitLog, 10) + "0" + Right(strBitLog, Len(strBitLog) - 11)
                        'ElseIf Mid(strBitLog, 11, 1) = "3" Then
                        '    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "RgDCorNok"
                        'ElseIf Mid(strBitLog, 11, 1) = "4" Then
                        '    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "PGVDblck"
                    End If
                End If
                'Referenz 12
                If Mid(strBitLog, 12, 1) <> "0" Then
                    If Mid(strBitLog, 12, 1) = "1" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "NoRef "
                    ElseIf Mid(strBitLog, 12, 1) = "2" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "RefChkD "
                    ElseIf Mid(strBitLog, 12, 1) = "3" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Ref "
                    End If
                End If
                'Int Bank 13
                If Mid(strBitLog, 13, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "IBank "
                End If
                'Keinen Text 14
                If Mid(strBitLog, 14, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Text "
                End If
                'PayType 15
                If Mid(strBitLog, 15, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "PType "
                    If Mid(strBitLog, 14, 1) = "4" Then
                        strStatus += "NoR"
                    ElseIf Mid(strBitLog, 14, 1) = "6" Then
                        strStatus += "BRef"
                    ElseIf Mid(strBitLog, 14, 1) = "7" Then
                        strStatus += "QIBAN"
                    ElseIf Mid(strBitLog, 14, 1) = "5" Then
                        strStatus += "BNoQ"
                    Else
                        strStatus += Mid(strBitLog, 14, 1)
                    End If
                End If
                'PGV keine Ziffer
                If row("booPGV") Then
                    If row("intPGVMthsAY") + row("intPGVMthsNY") = 1 Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "TP " + row("strPGVType")
                    Else
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "PGV " + row("strPGVType")
                    End If
                End If

                'Status schreiben
                If Val(strBitLog) = 0 Or Val(strBitLog) = 10000000000 Then
                    row("booKredBook") = True
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "ok"
                End If
                row("strKredStatusText") = strStatus
                row("strKredStatusBitLog") = strBitLog

                'Wird ein anderer Text in der Head-Buchung gewünscht?
                'booDiffHeadText = IIf(Main.FcReadFromSettingsII("Buchh_KTextSpecial", BgWCheckKrediArgsInProc.intMandant) = "0", False, True)
                If booDiffHeadText Then
                    strKrediHeadText = FcSQLParse(FcReadFromSettingsII("Buchh_KTextSpecialText",
                                                                                BgWCheckKrediArgsInProc.intMandant),
                                                                                row("strKredRGNbr"),
                                                                            dsKreditoren.Tables("tblKrediHeadsFromUser"),
                                                                            "C")
                    row("strKredText") = strKrediHeadText
                End If

                'Wird ein anderer Text in den Sub-Buchung gewünscht?
                'Soll der Gelesene Sub-Text bleiben?
                'booLeaveSubText = IIf(Main.FcReadFromSettingsII("Buchh_KSubLeaveText", BgWCheckKrediArgsInProc.intMandant) = "0", False, True)
                If Not booLeaveSubText Then
                    'booDiffSubText = IIf(Main.FcReadFromSettingsII("Buchh_KSubTextSpecial", BgWCheckKrediArgsInProc.intMandant) = "0", False, True)
                    If booDiffSubText Then
                        strKrediSubText = FcSQLParse(FcReadFromSettingsII("Buchh_KSubTextSpecialText",
                                                                                BgWCheckKrediArgsInProc.intMandant),
                                                                                row("strKredRGNbr"),
                                                                           dsKreditoren.Tables("tblKrediHeadsFromUser"),
                                                                           "C")
                    Else
                        strKrediSubText = row("strKredText")
                    End If
                    selsubrow = dsKreditoren.Tables("tblKrediSubsFromUser").Select("lngKredID=" + row("lngKredID").ToString)
                    For Each subrow As DataRow In selsubrow
                        subrow("strKredSubText") = strKrediSubText
                    Next
                    dsKreditoren.Tables("tblKrediSubsFromUser").AcceptChanges()
                End If

                'Init
                strBitLog = String.Empty
                strStatus = String.Empty
                intSubNumber = 0
                dblSubBrutto = 0
                dblSubNetto = 0
                dblSubMwSt = 0
                intKreditorNew = 0

                'Application.DoEvents()
                dsKreditoren.Tables("tblKrediHeadsFromUser").AcceptChanges()

            Next



        Catch ex As Exception
            MessageBox.Show(ex.Message, "Check-Kredit " + intKreditorNew.ToString)
            Err.Clear()

        Finally
            'objFinanz = Nothing
            'objfiBuha = Nothing
            'objKrBuha = Nothing
            'objFiBebu = Nothing
            selsubrow = Nothing
            BgWCheckKrediArgsInProc = Nothing


        End Try


    End Sub

    Private Sub frmKredDisp_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed

        Me.mysqlcmdKredDel = Nothing
        Me.mysqlcmdKredRead = Nothing
        Me.mysqlcmdKredSubDel = Nothing
        Me.mysqlcmdKredSubRead = Nothing
        Me.mysqlconn = Nothing
        Me.MySQLdaKreditoren = Nothing
        Me.MySQLdaKreditorenSub = Nothing

        Me.dsKreditoren.Reset()
        Me.dsKreditoren = Nothing
        Me.Dispose()
        'Application.Restart()

    End Sub

    Private Sub BgWImportKredi_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BgWImportKredi.DoWork

        Dim BgWImportKrediArgsInProc As BgWCheckDebitArgs = e.Argument
        Dim objdbMSSQLConn As New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("SQLConnectionString"))
        Dim objdbConnZHDB02 As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))
        Dim objdbSQLcommand As New SqlCommand
        Dim objOracleConn As New OracleConnection("Data Source=(DESCRIPTION=" _
                        + "(ADDRESS=(PROTOCOL=TCP)(HOST=10.0.0.29)(PORT=1521))" _
                        + "(CONNECT_DATA=(SERVICE_NAME=CISNEW)));" _
                        + "User Id=cis;Password=sugus;")


        Dim intReturnValue As Int16
        Dim strMandant As String
        Dim booAccOk As Boolean
        Dim strPeriode As String = BgWImportKrediArgsInProc.strPeriode
        Dim strExtKredBelegsNummer As String
        Dim intKreditorNbr As Int32
        Dim strBuchType As String
        Dim intKredBelegsNummer As Int32
        Dim strValutaDatum As String
        Dim strBelegDatum As String
        Dim strVerfallDatum As String
        Dim strReferenz As String
        Dim strTeilnehmer As String
        Dim intBankNbr As Int32
        Dim strMahnerlaubnis As String
        Dim strDebiLine As String
        Dim strDebitor() As String
        Dim strSachBID As String
        Dim dblBetrag As Decimal
        Dim strKrediText As String
        Dim strCurrency As String
        Dim intKondition As Int32
        Dim intKonditionLN As Int32
        Dim intEigeneBank As Int32
        Dim dblKurs As Double
        Dim booBookingok As Boolean
        Dim strZahlSperren As String = "N"
        Dim strVorausZahlung As String = "N"
        Dim strErfassungsArt As String = "K"
        Dim strSkonto As String = String.Empty
        Dim selKrediSub() As DataRow
        Dim intGegenKonto As Int32
        Dim strFibuText As String
        Dim dblNettoBetrag As Decimal
        Dim dblMwStBetrag As Decimal
        Dim dblBruttoBetrag As Decimal
        Dim strBeBuEintrag As String
        Dim strSteuerFeld As String
        Dim intSollKonto As Int32
        Dim dblKursSoll As Double
        Dim dblSollBetrag As Decimal
        Dim strKrediTextSoll As String
        Dim strSteuerFeldSoll As String
        Dim strBeBuEintragSoll As String
        Dim intHabenKonto As Int32
        Dim dblKursHaben As Double
        Dim dblHabenBetrag As Decimal
        Dim strKrediTextHaben As String
        Dim strSteuerFeldHaben As String
        Dim strBeBuEintragHaben As String
        Dim strKRGReferTo As String

        'Dim objFinanz As New SBSXASLib.AXFinanz
        'Dim objfiBuha As New SBSXASLib.AXiFBhg
        'Dim objdbBuha As New SBSXASLib.AXiDbBhg
        'Dim objdbPIFb As New SBSXASLib.AXiPlFin
        'Dim objFiBebu As New SBSXASLib.AXiBeBu
        'Dim objKrBuha As New SBSXASLib.AXiKrBhg

        Try

            'objdbSQLcommand.Connection = objdbMSSQLConn

            ''Start in Sync schreiben
            'intReturnValue = WFDBClass.FcWriteStartToSync(objdbConnZHDB02,
            '                                              BgWImportKrediArgsInProc.intMandant,
            '                                              2,
            '                                              dsKreditoren.Tables("tblKrediHeadsFromUser").Rows.Count)

            ''Finanz-Obj init
            ''Login
            'Call objFinanz.ConnectSBSdb(System.Configuration.ConfigurationManager.AppSettings("OwnSageServer"),
            '                        System.Configuration.ConfigurationManager.AppSettings("OwnSageDB"),
            '                        System.Configuration.ConfigurationManager.AppSettings("OwnSageID"),
            '                        System.Configuration.ConfigurationManager.AppSettings("OwnSagePsw"), "")

            'strMandant = Main.FcReadFromSettingsII("Buchh200_Name",
            '                                BgWImportKrediArgsInProc.intMandant)

            'booAccOk = objFinanz.CheckMandant(strMandant)
            ''Open Mandant
            'objFinanz.OpenMandant(strMandant, strPeriode)
            'objfiBuha = objFinanz.GetFibuObj()
            'objdbBuha = objFinanz.GetDebiObj()
            'objdbPIFb = objfiBuha.GetCheckObj()
            'objFiBebu = objFinanz.GetBeBuObj()
            'objKrBuha = objFinanz.GetKrediObj()

            'Kopfbuchung
            For Each row As DataRow In Me.dsKreditoren.Tables("tblKrediHeadsFromUser").Rows

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
                        Call objKrBuha.SetBuchMode("P")

                        'Automatische ESR - Zahlungsverbindung
                        objKrBuha.EnableAutoESRZlgVerb = "J"

                        'Eindeutigkeit der internen Beleg-Nummer setzen
                        objKrBuha.CheckDoubleIntBelNbr = "N"

                        'Eindeutigkeit externer Beleg-Nummer setzen
                        objKrBuha.CheckDoubleExtBelNbr = "J"

                        strExtKredBelegsNummer = row("strKredRGNbr")

                        'Variablen zuweisen
                        intKreditorNbr = row("lngKredNbr")
                        If row("dblKredBrutto") < 0 Then
                            strBuchType = "G"
                            'strZahlSperren = "J"
                            row("dblKredBrutto") = row("dblKredBrutto") * -1
                            'Belegsnummer abholen
                            objKrBuha.IncrBelNbr = "J"
                            intKredBelegsNummer = objKrBuha.GetNextBelNbr("G")
                            objKrBuha.IncrBelNbr = "N"

                            intReturnValue = FcCheckKrediExistance(intKredBelegsNummer,
                                                                                "G",
                                                                                BgWImportKrediArgsInProc.intTeqNbr,
                                                                                BgWImportKrediArgsInProc.intTeqNbrLY,
                                                                                BgWImportKrediArgsInProc.intTeqNbrPLY)

                        Else
                            strBuchType = "R"
                            'strZahlSperren = "N"
                            'Belegsnummer abholen
                            objKrBuha.IncrBelNbr = "J"
                            intKredBelegsNummer = objKrBuha.GetNextBelNbr("R")
                            'Muss auf Nicht hochzählen gesetzt werden da Sage 200 nicht merkt, dass Beleg-Nr. schon vergeben worden sind. => In den Einstellungen muss von Zeit zu Zeit der Zähler geändert werden
                            objKrBuha.IncrBelNbr = "N"

                            intReturnValue = FcCheckKrediExistance(intKredBelegsNummer,
                                                                                "R",
                                                                                BgWImportKrediArgsInProc.intTeqNbr,
                                                                                BgWImportKrediArgsInProc.intTeqNbrLY,
                                                                                BgWImportKrediArgsInProc.intTeqNbrPLY)

                        End If

                        strValutaDatum = Format(row("datKredValDatum"), "yyyyMMdd").ToString
                        strBelegDatum = Format(row("datKredRGDatum"), "yyyyMMdd").ToString
                        strVerfallDatum = String.Empty
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
                            strTeilnehmer = String.Empty
                        End If
                        strMahnerlaubnis = String.Empty 'Format(row("datDebRGDatum"), "yyyyMMdd").ToString
                        'Sachbearbeiter aus Debitor auslesen
                        strDebiLine = objKrBuha.ReadKreditor3(row("lngKredNbr") * -1, "")
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
                            dblKurs = FcGetKurs(strCurrency,
                                                     strValutaDatum)
                        Else
                            dblKurs = 1.0#
                        End If

                        Try
                            booBookingok = True
                            Call objKrBuha.SetBelegKopf2(intKredBelegsNummer,
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

                        selKrediSub = dsKreditoren.Tables("tblKrediSubsFromUser").Select("lngKredID=" + row("lngKredID").ToString)

                        For Each SubRow As DataRow In selKrediSub
                            intGegenKonto = SubRow("lngKto")
                            strFibuText = SubRow("strKredSubText")

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

                            If SubRow("lngKST") > 0 Then
                                strBeBuEintrag = SubRow("lngKST").ToString + "{<}" + SubRow("strKredSubText") + "{<}" + "CALCULATE" + "{>}"    '"PROD{<}BebuText{<}" + dblBebuBetrag.ToString + "{>}"
                            Else
                                'strBeBuEintrag = "00" + "{<}" + SubRow("strKredSubText") + "{<}" + "0" + "{>}"
                            End If
                            If Not IsDBNull(SubRow("strMwStKey")) And SubRow("strMwStKey") <> "null" Then ' And SubRow("strMwStKey") <> "25" Then
                                intReturnValue = FcGetSteuerFeld2(strSteuerFeld,
                                                                     SubRow("lngKto"),
                                                                     SubRow("strKredSubText"),
                                                                     dblBruttoBetrag,
                                                                     SubRow("strMwStKey"),
                                                                     dblMwStBetrag,
                                                                     row("datKredValDatum"))
                            Else
                                strSteuerFeld = "STEUERFREI"
                            End If

                            Try
                                booBookingok = True
                                Call objKrBuha.SetVerteilung(intGegenKonto.ToString,
                                                             strFibuText,
                                                             dblNettoBetrag.ToString,
                                                             strSteuerFeld,
                                                             strBeBuEintrag)

                            Catch ex As Exception
                                MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Verteilung " + SubRow("strKredSubText") + ", Konto " + SubRow("lngKto").ToString)
                                If (Err.Number And 65535) < 10000 Then
                                    booBookingok = False
                                Else
                                    booBookingok = True
                                End If

                            End Try

                            strSteuerFeld = String.Empty
                            strBeBuEintrag = String.Empty

                        Next

                        Try
                            booBookingok = True
                            Call objKrBuha.WriteBuchung()

                        Catch ex As Exception
                            If (Err.Number And 65535) < 10000 Then
                                MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Belegerstellung nicht möglich")
                                booBookingok = False
                            Else
                                If (Err.Number And 65535) = 10030 Then
                                    'MwSt-7.7/8.1 überschneidung nichts machen
                                    booBookingok = True
                                Else
                                    MessageBox.Show(ex.Message, "Warnung " + (Err.Number And 65535).ToString + " Belegerstellung")
                                    booBookingok = True
                                End If
                            End If

                        End Try

                        'Application.DoEvents()

                        strBeBuEintrag = String.Empty
                        strSteuerFeld = String.Empty
                        dblNettoBetrag = 0
                        dblMwStBetrag = 0
                        dblBruttoBetrag = 0

                    Else

                        'Buchung nur in Fibu
                        'Prinzip Funktion WriteBuchung() anwenden mit allen Parametern
                        'Beleg-Nummerierung aktivieren
                        'DbBhg.IncrBelNbr = "J"
                        'Belegsnummer abholen
                        intKredBelegsNummer = objfiBuha.GetNextBelNbr()

                        'Prüfen, ob wirklich frei
                        intReturnValue = 10
                        Do Until intReturnValue = 0
                            intReturnValue = objfiBuha.doesBelegExist(intKredBelegsNummer,
                                                                 "NOT_SET",
                                                                 "NOT_SET",
                                                                 Strings.Left(BgWImportKrediArgsInProc.strPeriode, 4) + "0101",
                                                                 Strings.Left(BgWImportKrediArgsInProc.strPeriode, 4) + "1231")
                            If intReturnValue <> 0 Then
                                intKredBelegsNummer += 1
                            End If
                        Loop

                        booBookingok = True

                        'Variablen zuweisen
                        strBelegDatum = Format(row("datKredRGDatum"), "yyyyMMdd").ToString
                        strValutaDatum = Format(row("datKredValDatum"), "yyyyMMdd").ToString
                        'strDebiText = row("strDebText")
                        strCurrency = row("strKredCur")
                        If strCurrency <> "CHF" Then 'Muss ergänzt werden => Was ist Leitwährung auf dem Konto
                            dblKurs = FcGetKurs(strCurrency,
                                                     strValutaDatum)
                        Else
                            dblKurs = 1.0#
                        End If

                        selKrediSub = dsKreditoren.Tables("tblKrediSubsFromUser").Select("lngKredID=" + row("lngKredID").ToString)

                        If selKrediSub.Length = 2 Then

                            For Each SubRow As DataRow In selKrediSub

                                If SubRow("intSollHaben") = 0 Then 'Soll

                                    intSollKonto = SubRow("lngKto")
                                    dblKursSoll = Main.FcGetKurs(strCurrency, strValutaDatum, objfiBuha, intSollKonto)
                                    'strSteuerInfo = Split(FBhg.GetKontoInfo(intSollKonto.ToString), "{>}")
                                    'Debug.Print("Konto-Info Soll: " + strSteuerInfo(26))
                                    dblSollBetrag = SubRow("dblNetto")
                                    strKrediTextSoll = SubRow("strKredSubText")
                                    If SubRow("dblMwSt") > 0 Then
                                        intReturnValue = FcGetSteuerFeld(strSteuerFeldSoll,
                                                                                 SubRow("lngKto"),
                                                                                 strKrediTextSoll,
                                                                                 SubRow("dblBrutto") * dblKursSoll,
                                                                                 SubRow("strMwStKey"),
                                                                                 SubRow("dblMwSt"))
                                    End If
                                    If SubRow("lngKST") > 0 Then
                                        strBeBuEintragSoll = SubRow("lngKST").ToString + "{<}" + strKrediTextSoll + "{<}" + "CALCULATE" + "{>}"
                                    End If


                                ElseIf SubRow("intSollHaben") = 1 Then 'Haben

                                    intHabenKonto = SubRow("lngKto")
                                    dblKursHaben = Main.FcGetKurs(strCurrency, strValutaDatum, objfiBuha, intHabenKonto)
                                    'strSteuerInfo = Split(FBhg.GetKontoInfo(intSollKonto.ToString), "{>}")
                                    'Debug.Print("Konto-Info Haben: " + strSteuerInfo(26))
                                    dblHabenBetrag = SubRow("dblNetto") * -1
                                    'dblHabenBetrag = dblSollBetrag
                                    strKrediTextHaben = SubRow("strKredSubText")
                                    If SubRow("dblMwSt") > 0 Then
                                        intReturnValue = FcGetSteuerFeld(strSteuerFeldHaben,
                                                                                  SubRow("lngKto"),
                                                                                  strKrediTextHaben,
                                                                                  SubRow("dblBrutto") * dblKursHaben * -1,
                                                                                  SubRow("strMwStKey"),
                                                                                  SubRow("dblMwSt") * -1)
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
                            Call objfiBuha.WriteBuchung(0,
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
                        If row("booPGV") Then
                            'Bei PGV Buchungen
                            If IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")) = "" Or
                               (IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")) = "RV" And row("intPGVMthsAY") + row("intPGVMthsNY") > 1) Then

                                intReturnValue = FcPGVKTreatment(dsKreditoren.Tables("tblKrediSubsFromUser"),
                                                                       row("lngKredID"),
                                                                       intKredBelegsNummer,
                                                                       row("strKredCur"),
                                                                       row("datKredValDatum"),
                                                                       "M",
                                                                       row("datPGVFrom"),
                                                                       row("datPGVTo"),
                                                                       row("intPGVMthsAY") + row("intPGVMthsNY"),
                                                                       row("intPGVMthsAY"),
                                                                       row("intPGVMthsNY"),
                                                                       2311,
                                                                       2312,
                                                                       BgWImportKrediArgsInProc.strPeriode,
                                                                       objdbConnZHDB02,
                                                                       objdbMSSQLConn,
                                                                       objdbSQLcommand,
                                                                       BgWImportKrediArgsInProc.intMandant,
                                                                       dsKreditoren.Tables("tblKreditorenInfo"),
                                                                       BgWImportKrediArgsInProc.strYear,
                                                                       BgWImportKrediArgsInProc.intTeqNbr,
                                                                       BgWImportKrediArgsInProc.intTeqNbrLY,
                                                                       BgWImportKrediArgsInProc.intTeqNbrPLY,
                                                                       IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")),
                                                                       datPeriodFrom,
                                                                       datPeriodTo,
                                                                       strPeriodStatus)

                            Else

                                'TP
                                intReturnValue = FcPGVKTreatmentYC(dsKreditoren.Tables("tblKrediSubsFromUser"),
                                                                       row("lngKredID"),
                                                                       intKredBelegsNummer,
                                                                       row("strKredCur"),
                                                                       row("datKredValDatum"),
                                                                       "M",
                                                                       row("datPGVFrom"),
                                                                       row("datPGVTo"),
                                                                       row("intPGVMthsAY") + row("intPGVMthsNY"),
                                                                       row("intPGVMthsAY"),
                                                                       row("intPGVMthsNY"),
                                                                       2311,
                                                                       2312,
                                                                       BgWImportKrediArgsInProc.strPeriode,
                                                                       objdbConnZHDB02,
                                                                       objdbMSSQLConn,
                                                                       objdbSQLcommand,
                                                                       BgWImportKrediArgsInProc.intMandant,
                                                                       dsKreditoren.Tables("tblKreditorenInfo"),
                                                                       BgWImportKrediArgsInProc.strYear,
                                                                       BgWImportKrediArgsInProc.intTeqNbr,
                                                                       BgWImportKrediArgsInProc.intTeqNbrLY,
                                                                       BgWImportKrediArgsInProc.intTeqNbrPLY,
                                                                       IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")),
                                                                       datPeriodFrom,
                                                                       datPeriodTo,
                                                                       strPeriodStatus)


                            End If


                        End If

                        'Status Head schreiben
                        'row("strKredBookStatus") = row("strKredStatusBitLog")
                        'row("booBooked") = True
                        'row("datBooked") = Now()
                        'row("lngBelegNr") = intKredBelegsNummer

                        'dsKreditoren.Tables("tblKrediHeadsFromUser").AcceptChanges()
                        'strKRGReferTo = "lngKredID"
                        'strKRGReferTo = "strKredRGNbr"
                        'If objdbConn.State = ConnectionState.Closed Then
                        '    objdbConn.Open()
                        'End If
                        strKRGReferTo = FcReadFromSettingsII("Buchh_TableKRGReferTo", BgWImportKrediArgsInProc.intMandant)
                        'If objdbConn.State = ConnectionState.Open Then
                        '    objdbConn.Close()
                        'End If
                        'Status in File RG-Tabelle schreiben
                        intReturnValue = FcWriteToKrediRGTable(BgWImportKrediArgsInProc.intMandant,
                                                                        row(strKRGReferTo),
                                                                        Now(),
                                                                        intKredBelegsNummer)
                        If intReturnValue <> 0 Then
                            'Throw an exception
                            MessageBox.Show("Achtung, Beleg-Nummer: " + row("lngBelegNr").ToString + " konnte nicht In die RG-Tabelle geschrieben werden auf RG-ID: " + row("lngKredID").ToString + ".", "RG-Table Update nicht möglich", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        End If

                    End If

                End If

            Next

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Generelles Problem " + (Err.Number And 65535).ToString + " Belegerstellung " + intKredBelegsNummer.ToString + ", RG " + strExtKredBelegsNummer, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Err.Clear()


        Finally
            'objKrBuha = Nothing
            'objfiBuha = Nothing
            'objdbPIFb = Nothing
            'objdbBuha = Nothing
            'objFiBebu = Nothing
            'objFinanz = Nothing

            objdbConnZHDB02 = Nothing
            objdbMSSQLConn = Nothing
            objOracleConn = Nothing
            objdbSQLcommand = Nothing

            BgWImportKrediArgsInProc = Nothing

        End Try


    End Sub

    Private Sub butDeSeöect_Click(sender As Object, e As EventArgs) Handles butDeSeöect.Click

        'Alle selektierten Records werden deselektiert

        For Each row As DataRow In dsKreditoren.Tables("tblKrediHeadsFromUser").Rows
            If row("booKredBook") Then
                row("booKredBook") = False
            End If
        Next
        dsKreditoren.Tables("tblKrediHeadsFromUser").AcceptChanges()
        'Me.Refresh()


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

    Friend Function FcInitInsCmdKHeads(ByRef mysqlinscmd As MySqlCommand) As Int16

        'Dim strIdentityName As String

        'Kreditoren - Head
        Dim inscmdFields As String
        Dim inscmdValues As String

        Try

            inscmdFields = "IdentityName"
            inscmdValues = "@IdentityName"
            inscmdFields += ", ProcessID"
            inscmdValues += ", @ProcessID"
            inscmdFields += ", intBuchhaltung"
            inscmdValues += ", @intBuchhaltung"
            inscmdFields += ", lngKredID"
            inscmdValues += ", @lngKredID"
            inscmdFields += ", strKredRGNbr"
            inscmdValues += ", @strKredRGNbr"
            inscmdFields += ", intBuchungsart"
            inscmdValues += ", @intBuchungsart"
            inscmdFields += ", strOPNr"
            inscmdValues += ", @strOPNr"
            inscmdFields += ", lngKredNbr"
            inscmdValues += ", @lngKredNbr"
            inscmdFields += ", lngKredKtoNbr"
            inscmdValues += ", @lngKredKtoNbr"
            inscmdFields += ", strKredCur"
            inscmdValues += ", @strKredCur"
            inscmdFields += ", lngKrediKST"
            inscmdValues += ", @lngKrediKST"
            inscmdFields += ", dblKredNetto"
            inscmdValues += ", @dblKredNetto"
            inscmdFields += ", dblKredMwSt"
            inscmdValues += ", @dblKredMwSt"
            inscmdFields += ", dblKredBrutto"
            inscmdValues += ", @dblKredBrutto"
            inscmdFields += ", lngKredIdentNbr"
            inscmdValues += ", @lngKredIdentNbr"
            inscmdFields += ", strKredText"
            inscmdValues += ", @strKredText"
            inscmdFields += ", strKredRef"
            inscmdValues += ", @strKredRef"
            inscmdFields += ", datKredRGDatum"
            inscmdValues += ", @datKredRGDatum"
            inscmdFields += ", datKredValDatum"
            inscmdValues += ", @datKredValDatum"
            inscmdFields += ", intPayType"
            inscmdValues += ", @intPayType"
            inscmdFields += ", strKrediBank"
            inscmdValues += ", @strKrediBank"
            inscmdFields += ", strKrediBankInt"
            inscmdValues += ", @strKrediBankInt"
            inscmdFields += ", strRGBemerkung"
            inscmdValues += ", @strRGBemerkung"
            inscmdFields += ", strRGName"
            inscmdValues += ", @strRGName"
            inscmdFields += ", intZKond"
            inscmdValues += ", @intZKond"
            inscmdFields += ", datPGVFrom"
            inscmdValues += ", @datPGVFrom"
            inscmdFields += ", datPGVTo"
            inscmdValues += ", @datPGVTo"



            'Ins cmd KrediiHead
            mysqlinscmd.CommandText = "INSERT INTO tblkreditorenhead (" + inscmdFields + ") VALUES (" + inscmdValues + ")"
            mysqlinscmd.Parameters.Add("@IdentityName", MySqlDbType.String).SourceColumn = "IdentityName"
            mysqlinscmd.Parameters.Add("@ProcessID", MySqlDbType.Int16).SourceColumn = "ProcessID"
            mysqlinscmd.Parameters.Add("@intBuchhaltung", MySqlDbType.Int16).SourceColumn = "intBuchhaltung"
            mysqlinscmd.Parameters.Add("@lngKredID", MySqlDbType.Int32).SourceColumn = "lngKredID"
            mysqlinscmd.Parameters.Add("@strKredRGNbr", MySqlDbType.String).SourceColumn = "strKredRGNbr"
            mysqlinscmd.Parameters.Add("@intBuchungsart", MySqlDbType.Int16).SourceColumn = "intBuchungsart"
            mysqlinscmd.Parameters.Add("@strOPNr", MySqlDbType.String).SourceColumn = "strOPNr"
            mysqlinscmd.Parameters.Add("@lngKredNbr", MySqlDbType.Int32).SourceColumn = "lngKredNbr"
            mysqlinscmd.Parameters.Add("@lngKredKtoNbr", MySqlDbType.Int32).SourceColumn = "lngKredKtoNbr"
            mysqlinscmd.Parameters.Add("@strKredCur", MySqlDbType.String).SourceColumn = "strKredCur"
            mysqlinscmd.Parameters.Add("@lngKrediKST", MySqlDbType.Int32).SourceColumn = "lngKrediKST"
            mysqlinscmd.Parameters.Add("@dblKredNetto", MySqlDbType.Decimal).SourceColumn = "dblKredNetto"
            mysqlinscmd.Parameters.Add("@dblKredMwSt", MySqlDbType.Decimal).SourceColumn = "dblKredMwSt"
            mysqlinscmd.Parameters.Add("@dblKredBrutto", MySqlDbType.Decimal).SourceColumn = "dblKredBrutto"
            mysqlinscmd.Parameters.Add("@strKredText", MySqlDbType.String).SourceColumn = "strKredText"
            mysqlinscmd.Parameters.Add("@lngKredIdentNbr", MySqlDbType.Int32).SourceColumn = "lngKredIdentNbr"
            mysqlinscmd.Parameters.Add("@strKredRef", MySqlDbType.String).SourceColumn = "strKredRef"
            mysqlinscmd.Parameters.Add("@datKredRGDatum", MySqlDbType.Date).SourceColumn = "datKredRGDatum"
            mysqlinscmd.Parameters.Add("@datKredValDatum", MySqlDbType.Date).SourceColumn = "datKredValDatum"
            mysqlinscmd.Parameters.Add("@intPayType", MySqlDbType.Int16).SourceColumn = "intPayType"
            mysqlinscmd.Parameters.Add("@strKrediBank", MySqlDbType.String).SourceColumn = "strKrediBank"
            mysqlinscmd.Parameters.Add("@strKrediBankInt", MySqlDbType.String).SourceColumn = "strKrediBankInt"
            mysqlinscmd.Parameters.Add("@strRGName", MySqlDbType.String).SourceColumn = "strRGName"
            mysqlinscmd.Parameters.Add("@strRGBemerkung", MySqlDbType.String).SourceColumn = "strRGBemerkung"
            mysqlinscmd.Parameters.Add("@intZKond", MySqlDbType.Int16).SourceColumn = "intZKond"
            mysqlinscmd.Parameters.Add("@datPGVFrom", MySqlDbType.Date).SourceColumn = "datPGVFrom"
            mysqlinscmd.Parameters.Add("@datPGVTo", MySqlDbType.Date).SourceColumn = "datPGVTo"

            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem KHeadCommand Init", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 1

        End Try

    End Function

    Friend Function FcSQLParseKredi(ByVal strSQLToParse As String,
                                           ByVal lngKredID As Int32,
                                           ByVal objdtKredi As DataTable) As String

        'Funktion setzt in eingelesenem SQL wieder Variablen ein
        Dim intPipePositionBegin, intPipePositionEnd As Integer
        Dim strWork, strField As String
        Dim RowKredi() As DataRow
        Dim strFieldType As String

        Try

            'Zuerst Datensatz in Kredii-Head suchen
            RowKredi = objdtKredi.Select("lngKredID=" + lngKredID.ToString)

            '| suchen
            If InStr(strSQLToParse, "|") > 0 Then
                'Vorkommen gefunden
                intPipePositionBegin = InStr(strSQLToParse, "|")
                intPipePositionEnd = InStr(intPipePositionBegin + 1, strSQLToParse, "|")
                Do Until intPipePositionBegin = 0
                    strField = Mid(strSQLToParse, intPipePositionBegin + 1, intPipePositionEnd - intPipePositionBegin - 1)
                    Select Case strField
                        Case "rsKredi.Fields(""KrediID"")"
                            strField = RowKredi(0).Item("lngKredID")
                            strFieldType = "V"
                        Case "rsKredi.Fields(""KrediRGNr"")"
                            strField = RowKredi(0).Item("strKredRGNbr")
                            strFieldType = "T"
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
                    strSQLToParse = Strings.Left(strSQLToParse, intPipePositionBegin - 1) + IIf(strFieldType = "T", "'", "") + strField + IIf(strFieldType = "T", "'", "") + Strings.Right(strSQLToParse, Len(strSQLToParse) - intPipePositionEnd)
                    'Neuer Anfang suchen für evtl. weitere |
                    intPipePositionBegin = InStr(strSQLToParse, "|")
                    'intPipePositionBegin = InStr(intPipePositionEnd + 1, strSQLToParse, "|")
                    intPipePositionEnd = InStr(intPipePositionBegin + 1, strSQLToParse, "|")
                Loop
            End If

            Return strSQLToParse

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Fehler Parsing " + Err.Number.ToString)
            Err.Clear()

        End Try


    End Function

    Private Sub butCheclLred_Click(sender As Object, e As EventArgs) Handles butCheclLred.Click

        Dim objdbConn As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString"))
        'Dim objdbtaskcmd As New MySqlCommand
        Dim objdbMSSQLConn As New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("SQLConnectionString"))
        Dim objdbSQLcommand As New SqlCommand

        Dim intFcReturns As Int16
        Dim strPeriode As String
        Dim strYearCh As String
        Dim BgWCheckKrediLocArgs As New BgWCheckDebitArgs
        'Dim objdbtasks As New DataTable

        'Dim objFinanz As New SBSXASLib.AXFinanz
        'Dim objfiBuha As New SBSXASLib.AXiFBhg
        'Dim objdbBuha As New SBSXASLib.AXiDbBhg
        'Dim objdbPIFb As New SBSXASLib.AXiPlFin
        'Dim objFiBebu As New SBSXASLib.AXiBeBu
        'Dim objKrBuha As New SBSXASLib.AXiKrBhg


        Try

            'Info neu erstellen
            dsKreditoren.Tables.Add("tblKreditorenInfo")
            Dim col1 As DataColumn = New DataColumn("strInfoT")
            col1.DataType = System.Type.GetType("System.String")
            col1.MaxLength = 50
            col1.Caption = "Info-Titel"
            dsKreditoren.Tables("tblKreditorenInfo").Columns.Add(col1)
            Dim col2 As DataColumn = New DataColumn("strInfoV")
            col2.DataType = System.Type.GetType("System.String")
            col2.MaxLength = 50
            col2.Caption = "Info-Wert"
            dsKreditoren.Tables("tblKreditorenInfo").Columns.Add(col2)

            'Datums-Tabelle erstellen
            dsKreditoren.Tables.Add("tblKreditorenDates")
            Dim col7 As DataColumn = New DataColumn("intYear")
            col7.DataType = System.Type.GetType("System.Int16")
            col7.Caption = "Year"
            dsKreditoren.Tables("tblKreditorenDates").Columns.Add(col7)
            Dim col3 As DataColumn = New DataColumn("strDatType")
            col3.DataType = System.Type.GetType("System.String")
            col3.MaxLength = 50
            col3.Caption = "Datum-Typ"
            dsKreditoren.Tables("tblKreditorenDates").Columns.Add(col3)
            Dim col4 As DataColumn = New DataColumn("datFrom")
            col4.DataType = System.Type.GetType("System.DateTime")
            col4.Caption = "Von"
            dsKreditoren.Tables("tblKreditorenDates").Columns.Add(col4)
            Dim col5 As DataColumn = New DataColumn("datTo")
            col5.DataType = System.Type.GetType("System.DateTime")
            col5.Caption = "Bis"
            dsKreditoren.Tables("tblKreditorenDates").Columns.Add(col5)
            Dim col6 As DataColumn = New DataColumn("strStatus")
            col6.DataType = System.Type.GetType("System.String")
            col6.Caption = "S"
            dsKreditoren.Tables("tblKreditorenDates").Columns.Add(col6)

            Call FcLoginSage3(objdbConn,
                                  objdbMSSQLConn,
                                  objdbSQLcommand,
                                  objFinanz,
                                  objfiBuha,
                                  objdbBuha,
                                  objdbPIFb,
                                  objFiBebu,
                                  objKrBuha,
                                  intMandant,
                                  dsKreditoren.Tables("tblKreditorenInfo"),
                                  dsKreditoren.Tables("tblKreditorenDates"),
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
                                       dsKreditoren.Tables("tblKreditorenDates"),
                                       strPeriode)
                Else
                    'Periode ezreugen und auf N stellen
                    strYearCh = Convert.ToString(Val(strYear) - 1)
                    dsKreditoren.Tables("tblKreditorenDates").Rows.Add(strYearCh, "GJ n/a", DateSerial(Convert.ToUInt16(strYearCh), 1, 1), DateSerial(Convert.ToUInt16(strYearCh), 12, 31), "N")
                End If

                'Gibt es ein Folgehahr?
                If lstBoxPerioden.SelectedIndex + 1 < lstBoxPerioden.Items.Count Then
                    strPeriode = lstBoxPerioden.Items(lstBoxPerioden.SelectedIndex + 1)
                    'Peeriodendef holen
                    Call FcLoginSage4(intMandant,
                                       dsKreditoren.Tables("tblKreditorenDates"),
                                       strPeriode)
                Else
                    'Periode ezreugen und auf N stellen
                    strYearCh = Convert.ToString(Val(strYear) + 1)
                    dsKreditoren.Tables("tblKreditorenDates").Rows.Add(strYearCh, "GJ n/a", DateSerial(Convert.ToUInt16(strYearCh), 1, 1), DateSerial(Convert.ToUInt16(strYearCh), 12, 31), "N")
                End If

            ElseIf lstBoxPerioden.Items.Count = 1 Then 'es gibt genau 1 Jahr
                'gewähltes Jahr checken
                Call FcLoginSage4(intMandant,
                                       dsKreditoren.Tables("tblKreditorenDates"),
                                       strPeriode)
                'VJ erzeugen
                strYearCh = Convert.ToString(Val(strYear) - 1)
                dsKreditoren.Tables("tblKreditorenDates").Rows.Add(strYearCh, "GJ n/a", DateSerial(Convert.ToUInt16(strYearCh), 1, 1), DateSerial(Convert.ToUInt16(strYearCh), 12, 31), "N")

                'FJ erzeugen
                strYearCh = Convert.ToString(Val(strYear) + 1)
                dsKreditoren.Tables("tblKreditorenDates").Rows.Add(strYearCh, "GJ n/a", DateSerial(Convert.ToUInt16(strYearCh), 1, 1), DateSerial(Convert.ToUInt16(strYearCh), 12, 31), "N")

            End If

            MySQLdaKreditoren.Fill(dsKreditoren, "tblKrediHeadsFromUser")
            MySQLdaKreditorenSub.Fill(dsKreditoren, "tblKrediSubsFromUser")

            BgWCheckKrediLocArgs.intMandant = intMandant
            BgWCheckKrediLocArgs.strMandant = frmImportMain.lstBoxMandant.GetItemText(frmImportMain.lstBoxMandant.SelectedItem)
            BgWCheckKrediLocArgs.intTeqNbr = intTeqNbr
            BgWCheckKrediLocArgs.intTeqNbrLY = intTeqNbrLY
            BgWCheckKrediLocArgs.intTeqNbrPLY = intTeqNbrPLY
            BgWCheckKrediLocArgs.strYear = strYear
            BgWCheckKrediLocArgs.strPeriode = lstBoxPerioden.GetItemText(lstBoxPerioden.SelectedItem)
            BgWCheckKrediLocArgs.booValutaCor = frmImportMain.chkValutaCorrect.Checked
            BgWCheckKrediLocArgs.datValutaCor = frmImportMain.dtpValutaCorrect.Value

            BgWCheckKredi.RunWorkerAsync(BgWCheckKrediLocArgs)

            Do While BgWCheckKredi.IsBusy
                Application.DoEvents()
            Loop

            'Grid neu aufbauen
            dgvDates.DataSource = dsKreditoren.Tables("tblKreditorenDates")
            dgvInfo.DataSource = dsKreditoren.Tables("tblKreditorenInfo")
            dgvBookings.DataSource = dsKreditoren.Tables("tblKrediHeadsFromUser")
            dgvBookingSub.DataSource = dsKreditoren.Tables("tblKrediSubsFromUser")

            intFcReturns = FcInitdgvInfo(dgvInfo)
            intFcReturns = FcInitdgvKreditoren(dgvBookings)
            intFcReturns = FcInitdgvKrediSub(dgvBookingSub)
            intFcReturns = FcInitdgvDate(dgvDates)


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Generelles Problem Kredi-Check" + Err.Number.ToString)
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
            'Anzahl schreiben
            txtNumber.Text = Me.dsKreditoren.Tables("tblKrediHeadsFromUser").Rows.Count.ToString

        End Try


    End Sub

    Friend Function FcLoginSage3(ByRef objdbconn As MySqlConnection,
                                       ByRef objsqlConn As SqlClient.SqlConnection,
                                       ByRef objsqlCom As SqlClient.SqlCommand,
                                       ByRef objFinanz As SBSXASLib.AXFinanz,
                                       ByRef objfiBuha As SBSXASLib.AXiFBhg,
                                       ByRef objdbBuha As SBSXASLib.AXiDbBhg,
                                       ByRef objdbPIFb As SBSXASLib.AXiPlFin,
                                       ByRef objFiBebu As SBSXASLib.AXiBeBu,
                                       ByRef objkrBuha As SBSXASLib.AXiKrBhg,
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
            Call objFinanz.ConnectSBSdb(System.Configuration.ConfigurationManager.AppSettings("OwnSageServer"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageDB"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageID"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSagePsw"), "")

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
            objfiBuha = objFinanz.GetFibuObj()
            'Debitor öffnen
            'If Not IsNothing(objdbBuha) Then
            '    objdbBuha = Nothing
            'End If
            'objdbBuha = New SBSXASLib.AXiDbBhg
            objdbBuha = objFinanz.GetDebiObj()
            'If Not IsNothing(objdbPIFb) Then
            '    objdbPIFb = Nothing
            'End If
            'objdbPIFb = New SBSXASLib.AXiPlFin
            objdbPIFb = objfiBuha.GetCheckObj()
            'If Not IsNothing(objFiBebu) Then
            '    objFiBebu = Nothing
            'End If
            'objFiBebu = New SBSXASLib.AXiBeBu
            objFiBebu = objFinanz.GetBeBuObj()
            'Kreditor
            'If Not IsNothing(objkrBuha) Then
            '    objkrBuha = Nothing
            'End If
            'objkrBuha = New SBSXASLib.AXiKrBhg
            objkrBuha = objFinanz.GetKrediObj

            'Application.DoEvents()

        Catch ex As Exception
            MsgBox("OpenMandant:" + vbCrLf + "Error" + vbCrLf + "Error # " + Str(Err.Number) + " was generated by " + Err.Source + vbCrLf + Err.Description + " Fehlernummer" & Str(Err.Number And 65535))
            Err.Clear()
            End

        Finally
            objdtPeriodeLY = Nothing
            dtPeriods = Nothing
            'System.GC.Collect()

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

    Friend Function FcLoginSage4(ByVal intAccounting As Int16,
                                 ByRef objdtDates As DataTable,
                                 ByVal strPeriod As String) As Int16

        'wird gebaucht um das Vor- und Folge-Jahr in Sage zu prüfen

        Dim objdbConn As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString"))
        Dim objdbcmd As New MySqlCommand

        'Dim objFinanz As New SBSXASLib.AXFinanz
        Dim strMandant As String
        Dim booAccOk As Boolean
        'Dim strPeriodenInfo As String
        Dim strArPeriode() As String
        Dim strArLogonInfo() As String
        Dim strYear As String
        Dim intPeriodenNr As Int16
        Dim intFctReturns As Int16
        Dim dtPeriods As New DataTable

        Try

            'Login
            'Call objFinanz.ConnectSBSdb(System.Configuration.ConfigurationManager.AppSettings("OwnSageServer"),
            '                        System.Configuration.ConfigurationManager.AppSettings("OwnSageDB"),
            '                        System.Configuration.ConfigurationManager.AppSettings("OwnSageID"),
            '                        System.Configuration.ConfigurationManager.AppSettings("OwnSagePsw"), "")

            intFctReturns = FcReadFromSettingsIII("Buchh200_Name",
                                                intAccounting,
                                                strMandant)

            'booAccOk = objFinanz.CheckMandant(strMandant)

            'objFinanz.OpenMandant(strMandant, strPeriod)

            'Von Login aktuelle Periode auslesen
            'strArLogonInfo = Split(objFinanz.GetLogonInfo(), "{>}")

            'Check Periode
            'intPeriodenNr = objFinanz.ReadPeri(strMandant, strArLogonInfo(7))
            'strPeriodenInfo = objFinanz.GetPeriListe(0)

            strArPeriode = Split(strPeriodenInfo, "{>}")

            strYear = Strings.Left(strArPeriode(4), 4)

            objdtDates.Rows.Add(strYear, "GJ Mandant", Date.ParseExact(strArPeriode(3), "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture), Date.ParseExact(strArPeriode(4), "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture), "O")
            objdtDates.Rows.Add(strYear, "Buchungen", Date.ParseExact(strArPeriode(5), "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture), Date.ParseExact(strArPeriode(6), "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture), strArPeriode(2))

            intFctReturns = FcReadPeriodenDef3(intPeriodenNr,
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

    Friend Function FcGetRefKrediNr(lngKrediNbr As Int32,
                                    intAccounting As Int32,
                                    ByRef intKrediNew As Int32) As Int16

        'Return 0=ok, 1=noch nicht implementiert, 2=Rep_Ref nicht definiert, 3=Nicht in Tab_Repbetriebe, 4=keine Angaben in Tab_Repbetriebe, 9=Problem

        Dim strTableName, strTableType, strKredFieldName, strKredNewField, strKredNewFieldType As String
        'Dim intCreatable As Int16
        Dim objdtKreditor As New DataTable
        Dim intPKNewField As Int32
        Dim objdbConnKred As New MySqlConnection
        Dim objsqlCommKred As New MySqlCommand

        Dim objdbAccessConn As OleDb.OleDbConnection
        Dim objlocOLEdbcmd As New OleDb.OleDbCommand
        Dim strMDBName As String = Main.FcReadFromSettingsII("Buchh_PKKrediTableConnection", intAccounting)
        'Dim objOrcommand As OracleClient.OracleCommand
        Dim strSQL As String
        Dim intFunctionReturns As Int16

        Try

            strTableName = FcReadFromSettingsII("Buchh_PKKrediTable", intAccounting)
            strTableType = FcReadFromSettingsII("Buchh_PKKrediTableType", intAccounting)
            strKredFieldName = FcReadFromSettingsII("Buchh_PKKrediField", intAccounting)
            strKredNewField = FcReadFromSettingsII("Buchh_PKKrediNewField", intAccounting)
            strKredNewFieldType = FcReadFromSettingsII("Buchh_PKKrediNewFType", intAccounting)

            strSQL = "SELECT * " + 'strKredFieldName + ", " + strKredNewField + ", " + strCompFieldName + ", " + strStreetFieldName + ", " + strZIPFieldName + ", " + strTownFieldName + ", " + strSageName + ", " + strKredAccField +
                 " FROM " + strTableName + " WHERE " + strKredFieldName + "=" + lngKrediNbr.ToString

            If strTableName <> "" And strKredFieldName <> "" Then

                If strTableType = "O" Then 'Oracle
                    Stop
                    'objOrdbconn.Open()
                    'objOrcommand.CommandText = strSQL
                    'objdtKreditor.Load(objOrcommand.ExecuteReader)
                    'Ist DebiNrNew Linked oder Direkt
                    'If strDebNewFieldType = "D" Then

                    'objOrdbconn.Close()
                ElseIf strTableType = "M" Then 'MySQL
                    intKrediNew = 0
                    'MySQL - Tabelle einlesen
                    objdbConnKred.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(Main.FcReadFromSettingsII("Buchh_PKKrediTableConnection", intAccounting))
                    objdbConnKred.Open()
                    objsqlCommKred.CommandText = strSQL
                    objsqlCommKred.Connection = objdbConnKred
                    objdtKreditor.Load(objsqlCommKred.ExecuteReader)
                    objdbConnKred.Close()

                ElseIf strTableType = "A" Then 'Access
                    'Access
                    Call FcInitAccessConnecation(objdbAccessConn, strMDBName)
                    objlocOLEdbcmd.CommandText = strSQL
                    objdbAccessConn.Open()
                    objlocOLEdbcmd.Connection = objdbAccessConn
                    objdtKreditor.Load(objlocOLEdbcmd.ExecuteReader)
                    objdbAccessConn.Close()

                End If

                'If IsDBNull(objdtKreditor.Rows(0).Item(strKredNewField)) Then
                If objdtKreditor.Rows.Count > 0 Then
                    'If IsDBNull(objdtKreditor.Rows(0).Item(strKredNewField)) And strTableName <> "Tab_Repbetriebe" Then
                    '    intKrediNew = 0
                    '    Return 2
                    'Else

                    If strTableName <> "Tab_Repbetriebe" Then
                        'intPKNewField = objdtKreditor.Rows(0).Item(strKredNewField)
                        If strTableName = "t_customer" Then
                            intPKNewField = FcGetPKNewFromRep(IIf(IsDBNull(objdtKreditor.Rows(0).Item("ID")), 0, objdtKreditor.Rows(0).Item("ID")),
                                                                       "P")
                        Else
                            intPKNewField = FcGetPKNewFromRep(objdtKreditor.Rows(0).Item(strKredNewField),
                                                                        "R") 'Rep_Nr
                            Stop
                        End If

                        If intPKNewField = 0 Then
                            'PK wurde nicht vergeben => Eine neue erzeugen und in der Tabelle Rep_Betriebe 
                            If strTableName = "t_customer" Then
                                intFunctionReturns = FcNextPrivatePKNr(objdtKreditor.Rows(0).Item("ID"),
                                                                            intKrediNew)
                                If intFunctionReturns = 0 And intKrediNew > 0 Then 'Vergabe hat geklappt
                                    intFunctionReturns = FcWriteNewPrivateDebToRepbetrieb(objdtKreditor.Rows(0).Item("ID"),
                                                                                                   intKrediNew)
                                    If intFunctionReturns = 0 Then 'Schreiben hat geklappt
                                        Return 1
                                    End If
                                End If
                            Else
                                intFunctionReturns = FcNextPKNr(objdtKreditor.Rows(0).Item(strKredNewField),
                                                                         intKrediNew,
                                                                         intAccounting,
                                                                         "C")
                                If intFunctionReturns = 0 And intKrediNew > 0 Then 'Vergabe hat geklappt
                                    intFunctionReturns = FcWriteNewDebToRepbetrieb(objdtKreditor.Rows(0).Item("Rep_Nr"),
                                                                                           intKrediNew,
                                                                                           intAccounting,
                                                                                           "C")
                                    If intFunctionReturns = 0 Then 'Schreiben hat geklappt
                                        Return 1
                                    End If
                                End If
                                Stop
                            End If

                            'intKrediNew = 0
                            'Return 3
                        Else
                            intKrediNew = intPKNewField
                            Return 0
                        End If
                    Else 'Wenn Angaben nicht von anderer Tabelle kommen
                        'Prüfen ob Repbetrieb schon eine neue Nummer erhalten hat
                        If Not IsDBNull(objdtKreditor.Rows(0).Item(strKredNewField)) Then
                            intKrediNew = objdtKreditor.Rows(0).Item(strKredNewField)
                        Else
                            intFunctionReturns = FcNextPKNr(objdtKreditor.Rows(0).Item("Rep_Nr"),
                                                                    intKrediNew,
                                                                    intAccounting,
                                                                    "C")
                            If intFunctionReturns = 0 And intKrediNew > 0 Then 'Vergabe hat geklappt
                                intFunctionReturns = FcWriteNewDebToRepbetrieb(objdtKreditor.Rows(0).Item("Rep_Nr"),
                                                                                       intKrediNew,
                                                                                       intAccounting,
                                                                                       "C")
                                If intFunctionReturns = 0 Then 'Schreiben hat geklappt
                                    Return 1
                                End If
                            End If
                        End If
                        Return 0
                    End If
                End If
            Else
                intKrediNew = 0
                Return 4
            End If

            'End If

            Return intPKNewField

        Catch ex As Exception
            MessageBox.Show(ex.Message, "kreditor-Ref " + Err.Number.ToString)

        Finally
            objdtKreditor = Nothing
            objdbConnKred = Nothing
            objsqlCommKred = Nothing
            objdbAccessConn = Nothing
            objlocOLEdbcmd = Nothing

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

    Friend Function FcCheckKreditor(lngKreditor As Long,
                                    intBuchungsart As Integer) As Integer

        Dim strReturn As String

        Try

            If intBuchungsart = 1 Then 'OP Buchung

                strReturn = objKrBuha.ReadKreditor3(lngKreditor * -1, "")
                If strReturn = "EOF" Then
                    Return 1
                Else
                    Return 0
                End If
            Else
                Return 0

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Kreditor-Check" + Err.Number.ToString)
            Err.Clear()

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

    Friend Function FcCheckKrediSubBookings2(ByVal lngKredID As Int32,
                                              ByRef objDtKrediSub As DataTable,
                                              ByRef intSubNumber As Int16,
                                              ByRef dblSubBrutto As Double,
                                              ByRef dblSubNetto As Double,
                                              ByRef dblSubMwSt As Double,
                                              ByVal datValuta As Date,
                                              ByVal intBuchungsArt As Int32,
                                              ByVal booAutoCorrect As Boolean,
                                              ByVal booCpyKSTToSub As Boolean,
                                              ByVal lngKrediKST As Int32,
                                              ByVal intPayType As Int16,
                                              ByVal strKrediBank As String) As Int16

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
        Dim strKstKtrSage200 As String = String.Empty
        Dim selsubrow() As DataRow
        Dim strStatusOverAll As String = "0000000"
        Dim strSteuer() As String

        'Summen bilden und Angaben prüfen
        intSubNumber = 0
        dblSubNetto = 0
        dblSubMwSt = 0
        dblSubBrutto = 0

        selsubrow = objDtKrediSub.Select("lngKredID=" + lngKredID.ToString)

        Try

            For Each subrow As DataRow In selsubrow

                'Application.DoEvents()

                strBitLog = String.Empty
                'Runden
                'subrow("dblNetto") = IIf(IsDBNull(subrow("dblNetto")), 0, Decimal.Round(subrow("dblNetto"), 2, MidpointRounding.AwayFromZero))
                'subrow("dblMwSt") = IIf(IsDBNull(subrow("dblMwst")), 0, Decimal.Round(subrow("dblMwst"), 2, MidpointRounding.AwayFromZero))
                'subrow("dblBrutto") = IIf(IsDBNull(subrow("dblBrutto")), 0, Decimal.Round(subrow("dblBrutto"), 2, MidpointRounding.AwayFromZero))
                'subrow("dblMwStSatz") = IIf(IsDBNull(subrow("dblMwStSatz")), 0, Decimal.Round(subrow("dblMwStSatz"), 1, MidpointRounding.AwayFromZero))

                'Runden
                If IsDBNull(subrow("dblNetto")) Then
                    subrow("dblNetto") = 0
                Else
                    subrow("dblNetto") = Decimal.Round(subrow("dblNetto"), 2, MidpointRounding.AwayFromZero)
                End If
                If IsDBNull(subrow("dblMwst")) Then
                    subrow("dblMwst") = 0
                Else
                    subrow("dblMwst") = Decimal.Round(subrow("dblMwst"), 2, MidpointRounding.AwayFromZero)
                End If
                If IsDBNull(subrow("dblBrutto")) Then
                    subrow("dblBrutto") = 0
                Else
                    subrow("dblBrutto") = Decimal.Round(subrow("dblBrutto"), 2, MidpointRounding.AwayFromZero)
                End If
                If IsDBNull(subrow("dblMwStSatz")) Then
                    subrow("dblMwStSatz") = 0
                Else
                    subrow("dblMwStSatz") = Decimal.Round(subrow("dblMwStSatz"), 1, MidpointRounding.AwayFromZero)
                End If

                'Falls KTRToSub dann kopieren
                If booCpyKSTToSub Then
                    subrow("lngKST") = lngKrediKST
                End If

                'Zuerst evtl. falsch gesetzte KTR oder Steuer - Sätze prüfen
                If subrow("lngKto") < 3000 Then
                    If (subrow("lngKto") <> 1120) And (subrow("lngKto") <> 1121) Then 'Ausnahme AW24
                        subrow("strMwStKey") = Nothing
                    End If
                    subrow("lngKST") = 0
                End If

                'Falls IBAN und BankKonto nicht CH, dann MwSt-Satz und MwSt-Key ändern
                If intPayType = 9 Then
                    If Char.IsLetter(CChar(Strings.Left(strKrediBank, 1))) And Char.IsLetter(CChar(Strings.Mid(strKrediBank, 2, 1))) Then
                        'Nun da klar ist, dass es 2 Zeichen sind muss noch geklärt werden. ob es keine CH Bankv. ist
                        If Strings.Left(strKrediBank, 2) <> "CH" Or Strings.Left(strKrediBank, 2) <> "ch" Then
                            'TODO: Routine ausprogrammieren.
                            subrow("dblMwStSatz") = 0
                            subrow("strMwStKey") = Nothing
                            subrow("dblNetto") = subrow("dblBrutto")
                            subrow("dblMwSt") = 0
                            'If booAutoCorrect Then
                            '    strStatusText = "MwSt K " + subrow("dblMwst").ToString + " -> " + Val(strSteuer(2)).ToString
                            '    subrow("dblMwst") = Val(strSteuer(2))
                            '    subrow("dblBrutto") = subrow("dblNetto") + subrow("dblMwSt")
                            'Else
                            '    'Nur korrigieren wenn weniger als 1 Fr
                            '    strStatusText = "MwSt K " + subrow("dblMwSt").ToString + ", " + Val(strSteuer(2)).ToString
                            '    If Math.Abs(subrow("dblMwSt") - Val(strSteuer(2))) > 1 Then
                            '        strStatusText += " >1 "
                            '        intReturnValue = 1
                            '    Else
                            '        strStatusText += " <1 "
                            '        subrow("dblMwst") = Val(strSteuer(2))
                            '        subrow("dblBrutto") = Decimal.Round(subrow("dblNetto") + subrow("dblMwSt"), 2, MidpointRounding.AwayFromZero)
                            '    End If

                            'End If
                        End If
                    Else
                        'subrow("strMwStKey") = "n/a"
                    End If
                Else
                    'subrow("strMwStKey") = "null"
                    'subrow("dblMwst") = 0
                    'intReturnValue = 0

                End If

                'Falsch vergebener MwSt-Schlüssel zurücksetzen
                If subrow("dblMwStSatz") = 0 And subrow("dblMwSt") = 0 And Not IsDBNull(subrow("strMwStKey")) Then
                    subrow("strMwStKey") = Nothing
                End If
                If Not IsDBNull(subrow("strMwStKey")) Then
                    intReturnValue = FcCheckMwSt(subrow("strMwStKey"),
                                                 subrow("dblMwStSatz"),
                                                 strStrStCodeSage200,
                                                 subrow("lngKto"))
                    If intReturnValue = 0 Then
                        subrow("strMwStKey") = strStrStCodeSage200
                        'Check ob korrekt berechnet
                        'falsche Steuersätze abfangen
                        Try

                            strSteuer = Split(objfiBuha.GetSteuerfeld2(subrow("lngKto").ToString,
                                                                    "Zum Rechnen",
                                                                    subrow("dblBrutto").ToString,
                                                                    strStrStCodeSage200,
                                                                    "",
                                                                    Format(datValuta, "yyyyMMdd"),
                                                                    Convert.ToString(subrow("dblMwStSatz"))), "{<}")

                        Catch ex As Exception
                            If (Err.Number And 65535) = 525 Then
                                strSteuer = Split(objfiBuha.GetSteuerfeld2(subrow("lngKto").ToString,
                                                                 "Zum Rechnen",
                                                                 subrow("dblBrutto").ToString,
                                                                 strStrStCodeSage200), "{<}")
                            End If

                        End Try
                        If Val(strSteuer(2)) <> subrow("dblMwst") Then
                            'Im Fall von Auto-Korrekt anpassen
                            If booAutoCorrect Then
                                strStatusText = "MwSt K " + subrow("dblMwst").ToString + " -> " + Val(strSteuer(2)).ToString
                                subrow("dblMwst") = Val(strSteuer(2))
                                subrow("dblBrutto") = subrow("dblNetto") + subrow("dblMwSt")
                            Else
                                'Nur korrigieren wenn weniger als 1 Fr
                                strStatusText = "MwSt K " + subrow("dblMwSt").ToString + ", " + Val(strSteuer(2)).ToString
                                If Math.Abs(subrow("dblMwSt") - Val(strSteuer(2))) > 1 Then
                                    strStatusText += " >1 "
                                    intReturnValue = 1
                                Else
                                    strStatusText += " <1 "
                                    subrow("dblMwst") = Val(strSteuer(2))
                                    subrow("dblBrutto") = Decimal.Round(subrow("dblNetto") + subrow("dblMwSt"), 2, MidpointRounding.AwayFromZero)
                                End If

                            End If
                        End If
                    Else
                        subrow("strMwStKey") = "n/a"
                    End If
                Else
                    subrow("strMwStKey") = "null"
                    intReturnValue = 0
                End If

                strBitLog += Trim(intReturnValue.ToString)


                'If subrow("intSollHaben") <> 2 Then
                intSubNumber += 1
                If subrow("intSollHaben") = 0 Then
                    dblSubNetto += IIf(IsDBNull(subrow("dblNetto")), 0, subrow("dblNetto"))
                    dblSubMwSt += IIf(IsDBNull(subrow("dblMwSt")), 0, subrow("dblMwSt"))
                    dblSubBrutto += IIf(IsDBNull(subrow("dblBrutto")), 0, subrow("dblBrutto"))
                Else
                    dblSubNetto += IIf(IsDBNull(subrow("dblNetto")), 0, subrow("dblNetto")) * -1
                    subrow("dblNetto") = Math.Abs(subrow("dblNetto")) * -1
                    dblSubMwSt += IIf(IsDBNull(subrow("dblMwSt")), 0, subrow("dblMwSt")) * -1
                    subrow("dblMwSt") = Math.Abs(subrow("dblMwSt")) * -1
                    dblSubBrutto += IIf(IsDBNull(subrow("dblBrutto")), 0, subrow("dblBrutto")) * -1
                    subrow("dblBrutto") = Math.Abs(subrow("dblBrutto")) * -1
                End If
                dblSubNetto = Decimal.Round(dblSubNetto, 2, MidpointRounding.AwayFromZero)
                dblSubMwSt = Decimal.Round(dblSubMwSt, 2, MidpointRounding.AwayFromZero)
                dblSubBrutto = Decimal.Round(dblSubBrutto, 2, MidpointRounding.AwayFromZero)

                'Konto prüfen 02
                If IIf(IsDBNull(subrow("lngKto")), 0, subrow("lngKTo")) > 0 Then
                    intReturnValue = FcCheckKonto(subrow("lngKto"),
                                                  IIf(IsDBNull(subrow("dblMwSt")), 0, subrow("dblMwSt")),
                                                  IIf(IsDBNull(subrow("lngKST")), 0, subrow("lngKST")),
                                                  False)
                    If intReturnValue = 0 Then
                        subrow("strKtoBez") = FcReadDebitorKName(subrow("lngKto"))
                    ElseIf intReturnValue = 2 Then
                        subrow("strKtoBez") = FcReadDebitorKName(subrow("lngKto")) + " MwSt!"
                    ElseIf intReturnValue = 3 Then
                        subrow("strKtoBez") = FcReadDebitorKName(subrow("lngKto")) + " NoKST"
                        'Falls keine KST definiert KST auf 0 setzen
                        subrow("lngKST") = 0
                        'Error zurück setzen
                        intReturnValue = 0
                    Else
                        subrow("strKtoBez") = "n/a"

                    End If
                Else
                    subrow("strKtoBez") = "null"
                    subrow("lngKto") = 0
                    intReturnValue = 1

                End If
                strBitLog += Trim(intReturnValue.ToString)

                'Kst/Ktr prüfen
                If IIf(IsDBNull(subrow("lngKST")), 0, subrow("lngKST")) > 0 Then
                    intReturnValue = FcCheckKstKtr(IIf(IsDBNull(subrow("lngKST")), 0, subrow("lngKST")),
                                                   subrow("lngKto"),
                                                   strKstKtrSage200)
                    If intReturnValue = 0 Then
                        subrow("strKstBez") = strKstKtrSage200
                    ElseIf intReturnValue = 1 Then
                        subrow("strKstBez") = "KoArt"

                    Else
                        subrow("strKstBez") = "n/a"

                    End If
                Else
                    subrow("strKstBez") = "null"
                    subrow("lngKST") = 0
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

                'Brutto + MwSt + Netto = 0
                If IIf(IsDBNull(subrow("dblBrutto")), 0, subrow("dblBrutto")) = 0 And IIf(IsDBNull(subrow("dblMwSt")), 0, subrow("dblMwSt")) = 0 And IIf(IsDBNull(subrow("dblNetto")), 0, subrow("dblNetto")) Then
                    strBitLog += "1"

                Else
                    strBitLog += "0"
                End If

                'Netto = 0
                If IIf(IsDBNull(subrow("dblNetto")), 0, subrow("dblNetto")) = 0 Then
                    strBitLog += "1"

                Else
                    strBitLog += "0"
                End If

                'Brutto = 0
                If IIf(IsDBNull(subrow("dblBrutto")), 0, subrow("dblBrutto")) = 0 Then
                    strBitLog += "1"

                Else
                    strBitLog += "0"
                End If

                'Brutto - MwSt <> Netto
                If Math.Round(IIf(IsDBNull(subrow("dblBrutto")), 0, subrow("dblBrutto")) - IIf(IsDBNull(subrow("dblMwSt")), 0, subrow("dblMwSt")), 2, MidpointRounding.AwayFromZero) <> IIf(IsDBNull(subrow("dblNetto")), 0, subrow("dblNetto")) Then
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
                    If Strings.Left(strBitLog, 1) = "2" Then
                        strStatusText = "Kto MwSt"
                    ElseIf Mid(strBitLog, 2, 1) = "3" Then
                        strStatusText = "Kto nKST"
                    Else
                        strStatusText = "Kto"
                    End If
                End If
                'Kst/Ktr
                If Mid(strBitLog, 3, 1) <> "0" Then
                    strStatusText += IIf(strStatusText <> "", ", ", "") + "KST"
                End If
                'Alles 0
                If Mid(strBitLog, 4, 1) <> "0" Then
                    strStatusText += IIf(strStatusText <> "", ", ", "") + "All0"
                End If
                'Netto 0
                If Mid(strBitLog, 5, 1) <> "0" Then
                    strStatusText += IIf(strStatusText <> "", ", ", "") + "Net0"
                End If
                'Brutto 0
                If Mid(strBitLog, 6, 1) <> "0" Then
                    strStatusText += IIf(strStatusText <> "", ", ", "") + "Brut0"
                End If
                'Diff
                If Mid(strBitLog, 7, 1) <> "0" Then
                    strStatusText += IIf(strStatusText <> "", ", ", "") + "Diff"
                End If

                If Val(strBitLog) = 0 Then
                    strStatusText += " ok"
                End If

                'BitLog und Text schreiben
                subrow("strStatusUBBitLog") = strBitLog
                subrow("strStatusUBText") = strStatusText

                strStatusOverAll = strStatusOverAll Or strBitLog
                strStatusText = String.Empty
                'Application.DoEvents()

            Next

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
            MessageBox.Show(ex.Message, "Fehler Kredi-Subbuchungen " + lngKredID.ToString)
            Err.Clear()

        Finally
            selsubrow = Nothing
            strSteuer = Nothing

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

    Friend Function FcChCeckKredOP(ByRef strOPNbr As String, ByVal strKredRGNbr As String) As Int16

        Dim strKredOPNbr As String

        '0=ok, 1=OP erstellt oder falsch 9=Problem

        'OP - Nr. Testen
        Try

            If Not strOPNbr Is Nothing And strOPNbr <> "" Then

                strKredOPNbr = CStr(Convert.ToString(Array.FindAll(strOPNbr.ToArray, Function(c As Char) Char.IsNumber(c))))

                If strOPNbr <> strKredOPNbr Then
                    strOPNbr = strKredOPNbr
                    Return 0
                Else
                    Return 0
                End If

            Else
                If strKredRGNbr <> "" Then
                    strOPNbr = CStr(Convert.ToString(Array.FindAll(strKredRGNbr.ToArray, Function(c As Char) Char.IsNumber(c))))
                    Return 0
                Else

                    Return 9
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Err.Clear()
            Return 9

        End Try


    End Function

    Friend Function FcCheckKrediOPDouble(strKreditor As String,
                                                strOPNr As String,
                                                strKredCurrency As String,
                                                strKredTyp As String) As Int16

        'Return 0=ok, 1=Beleg existiert, 9=Problem

        Dim intBelegReturn As Int32

        Try
            'Bei Kreditoren zählt externe RG-Nummer als Test
            intBelegReturn = objKrBuha.doesBelegExistExtern(strKreditor,
                                                            strKredCurrency,
                                                            strOPNr,
                                                            strKredTyp,
                                                            "")
            If intBelegReturn = 0 Then
                Return 0
            Else
                Return 1
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Kreditor - Doppelcheck auf Kreditor " + strKreditor + ", OP " + strOPNr)
            Return 9

        End Try

    End Function

    Friend Function FcIsAllKrediRebilled(ByVal objdbKrediSub As DataTable,
                                                ByVal intRGNummer As Int32) As Int16
        'Returns 0=mind. nicht 1 Rebill, 1=alle Rebill, 9=Problem

        Dim drKrediSub() As DataRow

        Try

            'Zuerst betroffene Buchungen selektieren
            drKrediSub = objdbKrediSub.Select("lngKredID=" + intRGNummer.ToString + " AND booRebilling=false")

            If drKrediSub.Length > 0 Then
                Return 0
            Else
                Return 1
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem PGV - Buchung Kreditoren", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 9

        Finally

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

    Friend Function FcCheckPayType(ByRef intPayType As Int16,
                                          ByVal strReferenz As String,
                                          ByVal strKrediBank As String) As Int16

        '0=ok, 1=IBAN Nr. aber nicht IBAN-Typ, 6=ESR-Nr aber keine Bank oder ungültige, 4=keine Referenz, 5=keine korrekte QR-IBAN 2=QR-ESR, 6=ESR Bank-Referenz nicht korrekt, 7=IBAN ist QR-IBAN, 9=Problem

        Try

            If Len(strReferenz) > 0 Then
                'Wurde eine IBAN - Nr. übergeben aber Typ ist nicht IBAN
                If Len(strReferenz) >= 21 Then ' And intPayType <> 9 Then
                    ''Sind die ersten 2 Positionen nicht numerisch?
                    'If Strings.Asc(Left(strReferenz, 1)) < 48 And Strings.Asc(Left(strReferenz, 1)) > 57 Then '1 Zeichen nicht numerisch
                    '    If Strings.Asc(Mid(strReferenz, 2, 1)) < 48 And Strings.Asc(Mid(strReferenz, 2, 1)) > 57 Then '2 Zeichen nicht numerisch
                    '        intPayType = 9
                    '        Return 1
                    '    End If
                    'End If
                    If Main.FcAreFirst2Chars(strReferenz) = 0 And intPayType <> 9 And Mid(strReferenz, 5, 1) <> "3" Then 'Falscher PayType bei IBAN-Nr.
                        intPayType = 9
                        Return 1
                    End If
                    'QR-ESR?
                    'Bank - Referenz IBAN?
                    If Main.FcAreFirst2Chars(strReferenz) = 0 Then 'IBAN - Referenz
                        'If Main.FcAreFirst2Chars(strKrediBank) = 0 Then
                        'intPayType = 9
                        'Return 0
                        'Else
                        'normale IBAN
                        'Check ob nicht QR-IBAN als Zahl-IBAN erfasst
                        If Mid(strReferenz, 5, 1) = "3" And Strings.Left(strReferenz, 2) = "CH" Then
                            intPayType = 9
                            Return 7
                        Else
                            intPayType = 9
                            Return 0
                        End If
                        'End If
                    Else 'QR-ESR ?
                        If Main.FcAreFirst2Chars(IIf(strKrediBank = "", "00", strKrediBank)) = 0 Then 'IBAN als Bank
                            'QR-IBAN?
                            If Mid(strKrediBank, 5, 1) = "3" Then
                                intPayType = 10
                                Return 2
                            Else
                                'keine QR-IBAN-ESR-Ref
                                'intPayType = 3
                                Return 5
                            End If
                        Else

                            If Len(strKrediBank) <> 9 Then 'ESR aber keine gültige Bank
                                'ESR, falsch deklariert
                                If intPayType <> 3 Then
                                    intPayType = 3
                                End If
                                Return 6
                            Else
                                'Debug.Print("Checksum " + Strings.Left(strKrediBank, 8) + " " + Strings.Right(strKrediBank, 1) + ", " + Main.FcModulo10(Strings.Left(strKrediBank, 8)).ToString)
                                If Main.FcModulo10(Strings.Left(strKrediBank, 8)).ToString <> Strings.Right(strKrediBank, 1) Then
                                    Return 6
                                Else
                                    Return 0 'Bank ok
                                End If

                            End If
                        End If
                    End If
                ElseIf intPayType = 0 Then
                    Return 9
                End If
                'If Len(strKrediBank) <> 9 Then 'ESR aber keine gültige Bank
                '    Return 3
                'Else
                '    Return 0 'Bank ok
                'End If

                'Else
            Else
                If intPayType = 9 And Len(strReferenz) = 0 Then
                    intPayType = 3 'Nicht IBAN
                    Return 4
                    'ElseIf intPayType = 0 Then
                    '    Return 9
                End If

            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Fc CheckPayType")
            Return 9

        Finally

        End Try

    End Function

    Friend Function FcIsPrivateKreditorCreatable(ByVal lngKrediNbr As Long,
                                                ByRef intPayType As Int16,
                                                ByVal strIBANFromInv As String,
                                                ByVal intintBank As Int16,
                                                ByVal strKrediBank As String,
                                                ByVal strcmbBuha As String,
                                                ByVal intAccounting As Int16) As Int16

        'Return: 0=creatable und erstellt, 3=Kreditor konnte nicht erstellt werden, 4=Betrieb nicht gefunden, 5=Nicht geprüft, 6=Aufwandskonto nicht existent, 9=Nicht hinterlegt

        Dim intCreatable As Int16
        Dim objdtKreditor As New DataTable
        Dim objdtKredZB As New DataTable
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
        Dim intKredZB As Int16
        Dim objdsKreditor As New DataSet
        Dim objDAKreditor As New MySqlDataAdapter
        Dim objdbconnKred As New MySqlConnection
        Dim objsqlConnKred As New MySqlCommand
        Dim intAufwandsKonto As Int32
        Dim booReadAufwandsKono As Boolean
        Dim objdtSachB As New DataTable("dtbliSachB")
        Dim strSachB As String
        Dim objdbconnZHDB02 As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))
        Dim objsqlcommandZHDB02 As New MySqlCommand


        Try

            'Angaben einlesen
            objdbconnKred.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(Main.FcReadFromSettingsII("Buchh_PKKrediTableConnection", intAccounting))
            If objdbconnKred.State = ConnectionState.Closed Then
                objdbconnKred.Open()
            End If
            If objdbconnZHDB02.State = ConnectionState.Closed Then
                objdbconnZHDB02.Open()
            End If
            objsqlConnKred.Connection = objdbconnKred
            objsqlConnKred.CommandText = "SELECT Lastname, " +
                                                "Firstname, " +
                                                "Street, " +
                                                "ZipCode, " +
                                                "City, " +
                                                "KrediGegenKonto, " +
                                                "'Privatperson' AS Gruppe, " +
                                                "IF(country IS NULL, 'CH', country) AS country, " +
                                                "Phone, " +
                                                "Fax, " +
                                                "Email, " +
                                                "IF(Language IS NULL, 'DE', Language) AS Language, " +
                                                "BankName, " +
                                                "BankZipCode, " +
                                                "BankCountry, " +
                                                "IBAN, " +
                                                "BankBIC, " +
                                                "BankName, " +
                                                "BankZipCode, " +
                                                "BankBIC, " +
                                                "PCKto, " +
                                                "IF(Currency IS NULL, 'CHF', Currency) AS Currency, " +
                                                "BankIntern, " +
                                                "KrediZKonditionID, " +
                                                "KrediAufwandskonto, " +
                                                "ReviewedOn " +
                                           "FROM t_customer WHERE PKNr=" + lngKrediNbr.ToString

            'objsqlcommandZHDB02.Connection = objdbconnZHDB02
            objdtKreditor.Load(objsqlConnKred.ExecuteReader)

            'Gefunden?
            If objdtKreditor.Rows.Count > 0 Then
                'Debug.Print("Gefunden, kann erstellt werden")

                If IsDBNull(objdtKreditor.Rows(0).Item("ReviewedOn")) Then
                    'PK wurde nicht geprüft

                    Return 5

                Else

                    'Prüfen, ob Aufwandskonto definiert ist
                    intReturnValue = FcCheckKonto(objdtKreditor.Rows(0).Item("KrediAufwandskonto"),
                                                       0,
                                                       0,
                                                       True)
                    If intReturnValue <> 0 Then
                        booReadAufwandsKono = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettingsII("Buchh_KrediTakeAufwKto", intAccounting)))
                        If booReadAufwandsKono Then
                            'Zu nehmendes Aufwandskonto einlesen
                            intAufwandsKonto = FcReadFromSettingsII("Buchh_KrediAufwKto", intAccounting)
                            objdtKreditor.Rows(0).Item("Rep_Kred_Aufwandskonto") = intAufwandsKonto
                            'Prüfen ob dieses Konto existiert
                            intReturnValue = FcCheckKonto(objdtKreditor.Rows(0).Item("KrediAufwandskonto"),
                                                       0,
                                                       0,
                                                       True)
                            If intReturnValue <> 0 Then
                                Return 6
                                'Sonst weiter 
                            End If

                        Else
                            Return 6
                        End If

                    End If

                    'Sachbearbeiter setzen
                    'Default setzen
                    objsqlcommandZHDB02.Connection = objdbconnZHDB02
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
                                                             objdtKreditor.Rows(0).Item("BankIntern"),
                                                             intintBank)

                    'Zahlungsbedingung suchen
                    intReturnValue = FcGetKZkondFromCust(lngKrediNbr,
                                                         intKredZB,
                                                         intAccounting)

                    ''objdtKreditor.Clear()
                    ''Es muss der Weg über ein Dataset genommen werden da sosnt constraint-Meldungen kommen
                    'objsqlConnKred.CommandText = "SELECT Tab_Repbetriebe.PKNr, 
                    '                                      t_sage_zahlungskondition.SageID " +
                    '                              "FROM Tab_Repbetriebe INNER JOIN t_sage_zahlungskondition ON Tab_Repbetriebe.Rep_Kred_ZKonditionID = t_sage_zahlungskondition.ID " +
                    '                              "WHERE Tab_Repbetriebe.PKNr=" + lngKrediNbr.ToString
                    'objDAKreditor.SelectCommand = objsqlConnKred
                    'objdsKreditor.EnforceConstraints = False
                    'objDAKreditor.Fill(objdsKreditor)

                    ''objdsKreditor.Load(objsqlcommandZHDB02.ExecuteReader)
                    ''objdtKreditor.Load(objsqlcommandZHDB02.ExecuteReader)
                    'If Not IsDBNull(objdsKreditor.Tables(0).Rows(0).Item("SageID")) Then
                    '    intKredZB = objdsKreditor.Tables(0).Rows(0).Item("SageID")
                    'Else
                    '    intKredZB = 1
                End If

                'Land von Text auf Auto-Kennzeichen ändern
                strLand = objdtKreditor.Rows(0).Item("country")
                'Select Case IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Land")), "Schweiz", objdtKreditor.Rows(0).Item("Rep_Land"))
                '    Case "Schweiz"
                '        strLand = "CH"
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
                Select Case Strings.UCase(IIf(IsDBNull(objdtKreditor.Rows(0).Item("Language")), "D", objdtKreditor.Rows(0).Item("Language")))
                    Case "D", "DE", ""
                        intLangauage = 2055
                    Case "F", "FR"
                        intLangauage = 4108
                    Case "I", "IT"
                        intLangauage = 2064
                    Case Else
                        intLangauage = 2057 'Englisch
                End Select

                'Variablen zuweisen für die Erstellung des Kreditors
                'IBAN von RG übernehmen sonst von Default holen
                If strIBANFromInv = "" Then
                    strIBANNr = IIf(IsDBNull(objdtKreditor.Rows(0).Item("IBAN")), "", objdtKreditor.Rows(0).Item("IBAN"))
                Else
                    strIBANNr = strIBANFromInv
                End If
                strBankName = IIf(IsDBNull(objdtKreditor.Rows(0).Item("BankName")), "", objdtKreditor.Rows(0).Item("BankName"))
                strBankAddress1 = ""
                strBankPLZ = IIf(IsDBNull(objdtKreditor.Rows(0).Item("BankZipCode")), "", objdtKreditor.Rows(0).Item("BankZipCode"))
                strBankOrt = ""
                strBankAddress2 = strBankPLZ + " " + strBankOrt
                strBankBIC = IIf(IsDBNull(objdtKreditor.Rows(0).Item("BankBIC")), "", objdtKreditor.Rows(0).Item("BankBIC"))
                strBankClearing = IIf(IsDBNull(objdtKreditor.Rows(0).Item("PCKto")), "", objdtKreditor.Rows(0).Item("PCKto"))

                If intPayType = 9 Or Len(strIBANNr) = 21 Then 'IBAN

                    If intPayType <> 9 Then 'Type nicht IBAN angegeben aber IBAN - Nr. erfasst
                        intPayType = 9
                    End If
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

                'QR-IBAN
                If intPayType = 10 And Len(strKrediBank) >= 21 Then
                    strIBANNr = strKrediBank
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

                intCreatable = FcCreateKreditor(lngKrediNbr,
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("LastName")), "", objdtKreditor.Rows(0).Item("LastName")), '+ " " + IIf(IsDBNull(objdtKreditor.Rows(0).Item("FirstName")), "", objdtKreditor.Rows(0).Item("FirstName")),
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Street")), "", objdtKreditor.Rows(0).Item("Street")),
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("ZipCode")), "", objdtKreditor.Rows(0).Item("ZipCode")),
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("City")), "", objdtKreditor.Rows(0).Item("City")),
                                          objdtKreditor.Rows(0).Item("KrediGegenKonto"),
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Gruppe")), "", objdtKreditor.Rows(0).Item("Gruppe")),
                                          "",
                                          "",
                                          strLand,
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Phone")), "", objdtKreditor.Rows(0).Item("Phone")),
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Fax")), "", objdtKreditor.Rows(0).Item("Fax")),
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Email")), "", objdtKreditor.Rows(0).Item("Email")),
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
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Currency")), "CHF", objdtKreditor.Rows(0).Item("Currency")),
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("KrediAufwandskonto")), 4200, objdtKreditor.Rows(0).Item("KrediAufwandskonto")),
                                          intKredZB,
                                          intintBank,
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("FirstName")), "", objdtKreditor.Rows(0).Item("FirstName")))

                If intCreatable = 0 Then
                    'MySQL
                    'strSQL = "INSERT INTO Tbl_RTFAutomail (RGNbr, MailCreateDate, MailCreateWho, MailTo, MailSender, MailTitle, MAilMsg, MailSent) VALUES (" +
                    '                                     lngKrediNbr.ToString + ", Date('" + Format(Today(), "yyyy-MM-dd").ToString + "'), 'Sage200Imp', " +
                    '                                     "'rene.hager@mssag.ch', 'Sage200@mssag.ch', 'Kreditor " +
                    '                                     lngKrediNbr.ToString + " wurde erstell im Mandant " + strcmbBuha + "', 'Bitte kontrollieren und Daten erg&auml;nzen.', false)"
                    ' objlocMySQLRGConn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(strMDBName)
                    'objlocMySQLRGConn.Open()
                    'objlocMySQLRGcmd.Connection = objlocMySQLRGConn
                    'objsqlcommandZHDB02.CommandText = strSQL
                    'intAffected = objsqlcommandZHDB02.ExecuteNonQuery()

                    intCreatable = FcWriteDatetoPrivate(lngKrediNbr,
                                                             intAccounting,
                                                             1)


                    Return 0

                End If

            Else

                Return 4

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Fehler Erstellung Kreditor " + lngKrediNbr.ToString + ", IBAN " + strIBANNr + " Bank " + strKrediBank)
            Return 9

        Finally
            objdbconnZHDB02.Close()
            objdbconnZHDB02 = Nothing
            objsqlcommandZHDB02 = Nothing

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

    Friend Function FcGetKZkondFromCust(lngKrediiNbr As Long,
                                              ByRef intDZkond As Int16,
                                              intAccounting As Int16) As Int16

        'Returns 0=ok, 1=Repbetrieb nicht gefunden, 9=Problem; intDZKond wird abgefüllt

        Dim objdbconnZHDB02 As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))
        Dim objsqlcommandZHDB02 As New MySqlCommand
        Dim intDZKondDefault As Int16

        Dim objdsDebitor As New DataSet
        Dim objDADebitor As New MySqlDataAdapter
        Dim objdtDZKond As New DataTable("tbllocDZKond")

        Try

            If objdbconnZHDB02.State = ConnectionState.Closed Then
                objdbconnZHDB02.Open()
            End If
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
                                              "AND t_sage_zahlungskondition.IsKredi = true"

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
                                              "AND t_sage_zahlungskondition.IsKredi = true"
                objdtDZKond.Load(objsqlcommandZHDB02.ExecuteReader)
                intDZKondDefault = objdtDZKond.Rows(0).Item("SageID")

            End If

            'Zahlungsbedingung suchen
            'Es muss der Weg über ein Dataset genommen werden da sosnt constraint-Meldungen kommen
            objsqlcommandZHDB02.CommandText = "Select t_customer.PKNr, t_sage_zahlungskondition.SageID " +
                                                  "FROM t_customer INNER JOIN t_sage_zahlungskondition On t_customer.DebiZKonditionID = t_sage_zahlungskondition.ID " +
                                                  "WHERE t_customer.PKNr=" + lngKrediiNbr.ToString
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
            MessageBox.Show(ex.Message, "Kreditor - Z-Bedingung - von Cust lesen", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            intDZkond = intDZKondDefault
            Return 9

        Finally
            objdbconnZHDB02.Close()
            objdbconnZHDB02 = Nothing
            objsqlcommandZHDB02 = Nothing
            objdsDebitor = Nothing
            objDADebitor = Nothing
            objdtDZKond = Nothing
            'Application.DoEvents()

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

    Friend Function FcCreateKreditor(intKreditorNewNbr As Int32,
                                           strKredName As String,
                                           strKredStreet As String,
                                           strKredPLZ As String,
                                           strKredOrt As String,
                                           intKredSammelKto As Int32,
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
                                           intAufwandsKonto As Int16,
                                           intKredZB As Int16,
                                           intintBank As Int16,
                                           strFirstName As String) As Int16

        Dim strKredCountry As String = strLand
        Dim strKredCurrency As String = strCurrency
        Dim strKredSprachCode As String = intLangauage.ToString
        Dim strKredSperren As String = "N"
        'Dim intKredErlKto As Integer = 2000
        Dim intKredVorErfKto As Int32
        'Dim intKredAufwandKto As Int32 = 4200
        'Dim shrKredZahlK As Short = 1
        Dim intKredToleranzNbr As Integer = 1
        Dim intKredMahnGroup As Integer = 1
        Dim strKredWerbung As String = "N"
        Dim strText As String = String.Empty
        Dim strTelefon1 As String
        Dim strTelefax As String

        'Kreditor erstellen

        Try

            strText = IIf(strGruppe = "", "", "Gruppe: " + strGruppe) + IIf(strVertretung = "" Or "0", "", strText + vbCrLf + "Vertretung: " + strVertretung)
            strTelefon1 = IIf(strTel = "" Or strTel = "0", "", strTel)
            strTelefax = IIf(strFax = "" Or strFax = "0", "", strFax)
            'Vorerfassung
            If strCurrency = "CHF" Then
                intKredVorErfKto = 2040
            Else
                intKredVorErfKto = 2041
            End If


            Call objKrBuha.SetCommonInfo2(intKreditorNewNbr,
                                         strKredName,
                                         strFirstName,
                                         "",
                                         strKredStreet,
                                         "",
                                         "",
                                         strKredCountry,
                                         strKredPLZ,
                                         strKredOrt,
                                         strTelefon1,
                                         "",
                                         strTelefax,
                                         strMail,
                                         "",
                                         strKredCurrency,
                                         "",
                                         "",
                                         strAnsprechpartner,
                                         strKredSprachCode,
                                         strText)

            Call objKrBuha.SetExtendedInfo7(strKredSperren,
                                           strKreditLimite,
                                           strMwStNr,
                                           intKredSammelKto.ToString,
                                           intKredVorErfKto.ToString,
                                           intAufwandsKonto.ToString,
                                           "",
                                           "",
                                           "",
                                           intKredZB.ToString,
                                           "",
                                           strKredWerbung)

            If intPayDefault = 9 Then 'IBAN
                If Len(strZVIBAN) > 15 Then

                    If Strings.Mid(strZVIBAN, 5, 1) <> "3" Or Strings.Left(strZVIBAN, 2) <> "CH" Then

                        Call objKrBuha.SetZahlungsverbindung("B",
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
                    Else
                        'Typ ist 10 (=QR)
                        Call objKrBuha.SetZahlungsverbindung("Q",
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
            End If

            If intPayDefault = 10 And Strings.Len(strZVIBAN) > 0 Then 'QR - IBAN

                Call objKrBuha.SetZahlungsverbindung("Q",
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

            Call objKrBuha.WriteKreditor3(intintBank.ToString, 0)

            Return 0
            'intDebAdrLaufN = DbBhg.GetAdressLaufnr()
            'intDebBankLaufNr = DbBhg.GetZahlungsverbLaufnr()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem beim Anlegen Kreditor " + intKreditorNewNbr.ToString + ", " + strKredName)
            Err.Clear()
            Return 1

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

    Friend Function FcIsKreditorCreatable(ByVal lngKrediNbr As Long,
                                                ByVal strcmbBuha As String,
                                                ByRef intPayType As Int16,
                                                ByVal strIBANFromInv As String,
                                                ByVal intintBank As Int16,
                                                ByVal strKrediBank As String,
                                                ByVal intAccounting As Int16) As Int16

        'Return: 0=creatable und erstellt, 3=Kreditor konnte nicht erstellt werden, 4=Betrieb nicht gefunden, 5=Nicht geprüft, 6=Aufwandskonto nicht existent, 9=Nicht hinterlegt

        Dim intCreatable As Int16
        Dim objdtKreditor As New DataTable
        Dim objdtKredZB As New DataTable
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
        Dim intKredZB As Int16
        Dim objdsKreditor As New DataSet
        Dim objDAKreditor As New MySqlDataAdapter
        Dim objdbconnKred As New MySqlConnection
        Dim objsqlConnKred As New MySqlCommand
        Dim intAufwandsKonto As Int32
        Dim booReadAufwandsKono As Boolean
        Dim objdbconnZHDB02 As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))
        Dim objsqlcommandZHDB02 As New MySqlCommand

        Try

            'Angaben einlesen
            objdbconnKred.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(Main.FcReadFromSettingsII("Buchh_PKKrediTableConnection", intAccounting))
            If objdbconnKred.State = ConnectionState.Closed Then
                objdbconnKred.Open()
            End If
            If objdbconnZHDB02.State = ConnectionState.Closed Then
                objdbconnZHDB02.Open()
            End If
            objsqlConnKred.Connection = objdbconnKred
            objsqlConnKred.CommandText = "SELECT Rep_Firma, 
                                                      Rep_Strasse, 
                                                      Rep_PLZ, 
                                                      Rep_Ort, 
                                                      Rep_KredGegenKonto, 
                                                      Rep_Gruppe, 
                                                      Rep_Vertretung, 
                                                      Rep_Ansprechpartner, 
                                                      Rep_Land, 
                                                      Rep_Tel1, 
                                                      Rep_Fax, 
                                                      Rep_Mail, " +
                                                     "Rep_Language, 
                                                      Rep_Kredi_MWSTNr, 
                                                      Rep_Kreditlimite, 
                                                      Rep_Kred_Pay_Def, 
                                                      Rep_Kred_Bank_Name, 
                                                      Rep_Kred_Bank_PLZ, 
                                                      Rep_Kred_Bank_Ort, 
                                                      Rep_Kred_IBAN, 
                                                      Rep_Kred_Bank_BIC, " +
                                                     "Rep_Kred_Currency, 
                                                      Rep_Kred_PCKto, 
                                                      Rep_Kred_Aufwandskonto,
                                                      ReviewedOn 
                                                FROM Tab_Repbetriebe WHERE PKNr=" + lngKrediNbr.ToString

            'objsqlcommandZHDB02.Connection = objdbconnZHDB02
            objdtKreditor.Load(objsqlConnKred.ExecuteReader)

            'Gefunden?
            If objdtKreditor.Rows.Count > 0 Then
                'Debug.Print("Gefunden, kann erstellt werden")

                If IsDBNull(objdtKreditor.Rows(0).Item("ReviewedOn")) Then
                    'PK wurde nicht geprüft

                    Return 5

                Else

                    'Prüfen, ob Aufwandskonto definiert ist
                    intReturnValue = FcCheckKonto(objdtKreditor.Rows(0).Item("Rep_Kred_Aufwandskonto"),
                                                       0,
                                                       0,
                                                       True)
                    If intReturnValue <> 0 Then
                        booReadAufwandsKono = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettingsII("Buchh_KrediTakeAufwKto", intAccounting)))
                        If booReadAufwandsKono Then
                            'Zu nehmendes Aufwandskonto einlesen
                            intAufwandsKonto = FcReadFromSettingsII("Buchh_KrediAufwKto", intAccounting)
                            objdtKreditor.Rows(0).Item("Rep_Kred_Aufwandskonto") = intAufwandsKonto
                            'Prüfen ob dieses Konto existiert
                            intReturnValue = FcCheckKonto(objdtKreditor.Rows(0).Item("Rep_Kred_Aufwandskonto"),
                                                       0,
                                                       0,
                                                       True)
                            If intReturnValue <> 0 Then
                                Return 6
                                'Sonst weiter 
                            End If

                        Else
                            Return 6
                        End If

                    End If

                    'Zahlungsbedingung suchen
                    'objdtKreditor.Clear()
                    'Es muss der Weg über ein Dataset genommen werden da sosnt constraint-Meldungen kommen
                    objsqlConnKred.CommandText = "SELECT Tab_Repbetriebe.PKNr, 
                                                          t_sage_zahlungskondition.SageID " +
                                                  "FROM Tab_Repbetriebe INNER JOIN t_sage_zahlungskondition ON Tab_Repbetriebe.Rep_Kred_ZKonditionID = t_sage_zahlungskondition.ID " +
                                                  "WHERE Tab_Repbetriebe.PKNr=" + lngKrediNbr.ToString
                    objDAKreditor.SelectCommand = objsqlConnKred
                    objdsKreditor.EnforceConstraints = False
                    objDAKreditor.Fill(objdsKreditor)

                    'objdsKreditor.Load(objsqlcommandZHDB02.ExecuteReader)
                    'objdtKreditor.Load(objsqlcommandZHDB02.ExecuteReader)
                    If Not IsDBNull(objdsKreditor.Tables(0).Rows(0).Item("SageID")) Then
                        intKredZB = objdsKreditor.Tables(0).Rows(0).Item("SageID")
                    Else
                        intKredZB = 1
                    End If

                    'Land von Text auf Auto-Kennzeichen ändern
                    Select Case IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Land")), "Schweiz", objdtKreditor.Rows(0).Item("Rep_Land"))
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
                        Case Else
                            strLand = "CH"
                    End Select

                    'Sprache zuweisen von 1-Stelligem String nach Sage 200 Regionen
                    Select Case Strings.UCase(IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Language")), "D", objdtKreditor.Rows(0).Item("Rep_Language")))
                        Case "D", "DE", ""
                            intLangauage = 2055
                        Case "F", "FR"
                            intLangauage = 4108
                        Case "I", "IT"
                            intLangauage = 2064
                        Case Else
                            intLangauage = 2057 'Englisch
                    End Select

                    'Variablen zuweisen für die Erstellung des Kreditors
                    'IBAN von RG übernehmen sonst von Default holen
                    If strIBANFromInv = "" Then
                        strIBANNr = IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Kred_IBAN")), "", objdtKreditor.Rows(0).Item("Rep_Kred_IBAN"))
                    Else
                        strIBANNr = strIBANFromInv
                    End If
                    strBankName = IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Kred_Bank_Name")), "", objdtKreditor.Rows(0).Item("Rep_Kred_Bank_Name"))
                    strBankAddress1 = ""
                    strBankPLZ = IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Kred_Bank_PLZ")), "", objdtKreditor.Rows(0).Item("Rep_Kred_Bank_PLZ"))
                    strBankOrt = IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Kred_Bank_Ort")), "", objdtKreditor.Rows(0).Item("Rep_Kred_Bank_Ort"))
                    strBankAddress2 = strBankPLZ + " " + strBankOrt
                    strBankBIC = IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Kred_Bank_BIC")), "", objdtKreditor.Rows(0).Item("Rep_Kred_Bank_BIC"))
                    strBankClearing = IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Kred_PCKto")), "", objdtKreditor.Rows(0).Item("Rep_Kred_PCKto"))

                    If intPayType = 9 Or Len(strIBANNr) = 21 Then 'IBAN

                        If intPayType <> 9 Then 'Type nicht IBAN angegeben aber IBAN - Nr. erfasst
                            intPayType = 9
                        End If
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

                    'QR-IBAN
                    If intPayType = 10 And Len(strKrediBank) >= 21 Then
                        strIBANNr = strKrediBank
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

                    intCreatable = FcCreateKreditor(lngKrediNbr,
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Firma")), "", objdtKreditor.Rows(0).Item("Rep_Firma")),
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Strasse")), "", objdtKreditor.Rows(0).Item("Rep_Strasse")),
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_PLZ")), "", objdtKreditor.Rows(0).Item("Rep_PLZ")),
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Ort")), "", objdtKreditor.Rows(0).Item("Rep_Ort")),
                                          objdtKreditor.Rows(0).Item("Rep_KredGegenKonto"),
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Gruppe")), "", objdtKreditor.Rows(0).Item("Rep_Gruppe")),
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Vertretung")), "", objdtKreditor.Rows(0).Item("Rep_Vertretung")),
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Ansprechpartner")), "", objdtKreditor.Rows(0).Item("Rep_Ansprechpartner")),
                                          strLand,
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Tel1")), "", objdtKreditor.Rows(0).Item("Rep_Tel1")),
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Fax")), "", objdtKreditor.Rows(0).Item("Rep_Fax")),
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Mail")), "", objdtKreditor.Rows(0).Item("Rep_Mail")),
                                          intLangauage,
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Kredi_MWStNr")), "", objdtKreditor.Rows(0).Item("Rep_Kredi_MWStNr")),
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Kreditlimite")), "", objdtKreditor.Rows(0).Item("Rep_Kreditlimite")),
                                          intPayType,
                                          strBankName,
                                          strBankPLZ,
                                          strBankOrt,
                                          strIBANNr,
                                          strBankBIC,
                                          strBankClearing,
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Kred_Currency")), "CHF", objdtKreditor.Rows(0).Item("Rep_Kred_Currency")),
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Kred_Aufwandskonto")), 4200, objdtKreditor.Rows(0).Item("Rep_Kred_Aufwandskonto")),
                                          intKredZB,
                                          intintBank,
                                          "")

                    If intCreatable = 0 Then
                        'MySQL
                        'strSQL = "INSERT INTO Tbl_RTFAutomail (RGNbr, MailCreateDate, MailCreateWho, MailTo, MailSender, MailTitle, MAilMsg, MailSent) VALUES (" +
                        '                                     lngKrediNbr.ToString + ", Date('" + Format(Today(), "yyyy-MM-dd").ToString + "'), 'Sage200Imp', " +
                        '                                     "'rene.hager@mssag.ch', 'Sage200@mssag.ch', 'Kreditor " +
                        '                                     lngKrediNbr.ToString + " wurde erstell im Mandant " + strcmbBuha + "', 'Bitte kontrollieren und Daten erg&auml;nzen.', false)"
                        ' objlocMySQLRGConn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(strMDBName)
                        'objlocMySQLRGConn.Open()
                        'objlocMySQLRGcmd.Connection = objlocMySQLRGConn
                        'objsqlcommandZHDB02.CommandText = strSQL
                        'intAffected = objsqlcommandZHDB02.ExecuteNonQuery()



                        Return 0
                    Else
                        Return 3

                    End If

                End If

            Else
                Return 4
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Fehler Erstellung Kreditor " + lngKrediNbr.ToString + ", IBAN " + strIBANNr + " Bank " + strKrediBank)
            Err.Clear()
            Return 9

        Finally
            objdbconnZHDB02.Close()
            objdbconnZHDB02 = Nothing
            objsqlcommandZHDB02 = Nothing

        End Try

    End Function

    Friend Function FcReadKreditorName(ByRef strKreditorName As String,
                                       intKrediNbr As Int32,
                                       strCurrency As String) As Int16

        'Dim strKreditorName As String
        Dim strKreditorAr() As String

        Try

            If strCurrency = "" Then

                strKreditorName = objKrBuha.ReadKreditor3(intKrediNbr * -1, strCurrency)

            Else

                strKreditorName = objKrBuha.ReadKreditor3(intKrediNbr, strCurrency)
                'strKreditorName = objKrBhg.ReadKreditor3(1, "CHF")
                'Call objKrBhg.ReadKrediStamm2()
                'Do Until strKreditorName = "EOF"
                '    strKreditorName = objKrBhg.GetKStammZeile3()
                'Loop
            End If

            strKreditorAr = Split(strKreditorName, "{>}")
            strKreditorName = strKreditorAr(0)

            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message, "kreditor-Name " + Err.Number.ToString)
            Err.Clear()
            Return 1

        Finally
            strKreditorAr = Nothing

        End Try

    End Function

    Friend Function FcCheckKreditBank(intKreditor As Int32,
                                         intPayType As Int16,
                                         strIBAN As String,
                                         strBank As String,
                                         strKredCur As String,
                                         ByRef intEBank As Int32) As Int16

        'Falls Typetype 9 (IBAN) ist, dann Zahlungsverbindungen prüfen

        Dim strZahlVerbindungLine As String = String.Empty
        Dim strZahlVerbindung() As String
        Dim booBankExists As Boolean = False
        Dim intReturnValue As Int16
        Dim strBankName As String = String.Empty
        Dim strBankAddress1 As String = String.Empty
        Dim strBankAddress2 As String = String.Empty
        Dim strBankCountry As String = String.Empty
        Dim strBankBIC As String = String.Empty
        Dim strBankClearing As String = String.Empty
        Dim strBankOrt As String = String.Empty
        Dim strBankPLZ As String = String.Empty

        Try

            If intPayType = 9 Or intPayType = 10 Then 'IBAN oder QR-

                Call objKrBuha.ReadZahlungsverb(intKreditor * -1)

                Do Until strZahlVerbindungLine = "EOF"

                    strZahlVerbindungLine = objKrBuha.GetZahlungsverbZeile()
                    'Debug.Print("Line " + strZahlVerbindungLine)
                    strZahlVerbindung = Split(strZahlVerbindungLine, "{>}")
                    If strZahlVerbindungLine <> "EOF" Then
                        If strZahlVerbindung(3) = "K" Then
                            'Debug.Print("BankV " + strZahlVerbindung(4) + ", " + strIBAN)
                            If strZahlVerbindung(4) = strIBAN Then
                                booBankExists = True
                                'Debug.Print("Gefunden " + strZahlVerbindungLine)
                                intEBank = strZahlVerbindung(0)
                            End If

                        End If
                    End If
                Loop

                If Not booBankExists Then
                    'MessageBox.Show("Bankverbindung muss erstellt werden " + strIBAN)

                    intReturnValue = FcGetIBANDetails(strIBAN,
                                                      strBankName,
                                                      strBankAddress1,
                                                      strBankAddress2,
                                                      strBankBIC,
                                                      strBankCountry,
                                                      strBankClearing)

                    If intReturnValue = 0 Then 'Angaben vollständig und kein Problem
                        'Kombinierte PLZ / Ort Feld trennen
                        strBankPLZ = Strings.Left(strBankAddress2, InStr(strBankAddress2, " "))
                        strBankOrt = Trim(Strings.Right(strBankAddress2, Len(strBankAddress2) - InStr(strBankAddress2, " ")))

                        'Evtl Typ falsch gesetzt?
                        If Strings.Mid(strIBAN, 5, 1) <> "3" Or Strings.Left(strIBAN, 2) <> "CH" Then
                            'IBAN
                            Call objKrBuha.WriteBank2(intKreditor,
                                                 strKredCur,
                                                 "B",
                                                 strIBAN,
                                                 strBankName,
                                                 "",
                                                 "",
                                                 strBankPLZ,
                                                 strBankOrt,
                                                 strBankCountry,
                                                 strBankClearing,
                                                 "J",
                                                 strBankBIC,
                                                 "",
                                                 "0",
                                                 "",
                                                 strIBAN)
                        Else

                            Stop
                            Call objKrBuha.WriteBank2(intKreditor,
                                                     strKredCur,
                                                     "Q",
                                                     strIBAN,
                                                     strBankName,
                                                     "",
                                                     "",
                                                     strBankPLZ,
                                                     strBankOrt,
                                                     strBankCountry,
                                                     strBankClearing,
                                                     "J",
                                                     strBankBIC,
                                                     "",
                                                     "",
                                                     "",
                                                     strIBAN)

                        End If
                        Return 0

                    End If

                End If

            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Fehler Check-Kredi-Bank " + intKreditor.ToString + ", " + strIBAN)
            Return 9

        End Try

    End Function

    Friend Function FcSQLParse(strSQLToParse As String,
                                      strRGNbr As String,
                                      objdtBookings As DataTable,
                                      strDebiCredit As String) As String

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

    Public Shared Function FcCheckKrediExistance(ByRef intBelegNbr As Int32,
                                                 ByVal strTyp As String,
                                                 ByVal intTeqNr As Int32,
                                                 ByVal intTeqNrLY As Int32,
                                                 ByVal intTeqNrPLY As Int32) As Int16

        '0=ok, 1=Beleg existierte schon, 9=Problem

        'Prinzip: in Tabelle kredibuchung suchen da API - Funktion nur in spezifischen Kreditor sucht

        Dim intReturnvalue As Int32
        Dim intStatus As Int16
        Dim tblKrediBeleg As New DataTable
        Dim intEntryBelNbr As Int32 = intBelegNbr
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
                objdbMSSQLCmd.CommandText = "SELECT lfnbrk FROM kredibuchung WHERE teqnbr IN(" + intTeqNr.ToString + ")" +
                                                                        " AND typ='" + strTyp + "'" +
                                                                        " AND belnbrint=" + intBelegNbr.ToString

                tblKrediBeleg.Rows.Clear()
                tblKrediBeleg.Load(objdbMSSQLCmd.ExecuteReader)
                If tblKrediBeleg.Rows.Count > 0 Then
                    intReturnvalue = tblKrediBeleg.Rows(0).Item("lfnbrk")
                    'objKrBhg.IncrBelNbr = "J"
                    'intBelegNbr = objKrBhg.GetNextBelNbr(strTyp)
                    'Hat Hochzählen geklappt
                    'If intBelegNbr <= intEntryBelNbr Then
                    intBelegNbr += 1
                    'End If
                Else
                    intReturnvalue = 0
                End If
            Loop

            Return intStatus


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Kreditor - BelegExistenzprüfung Problem " + intBelegNbr.ToString)
            Err.Clear()
            Return 9

        Finally
            objdbMSSQLConn.Close()
            objdbMSSQLCmd = Nothing
            objdbMSSQLConn = Nothing
            tblKrediBeleg = Nothing

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

    Friend Function FcGetSteuerFeld2(ByRef strSteuerFeld As String,
                                           lngKto As Long,
                                           strDebiSubText As String,
                                           dblBrutto As Double,
                                           strMwStKey As String,
                                           dblMwSt As Double,
                                           datValuta As Date) As Int16

        'Setzt Steuer-Feld mit Valuzta-Datum

        Try

            If dblMwSt <> 0 Then

                strSteuerFeld = objfiBuha.GetSteuerfeld(lngKto.ToString,
                                                      strDebiSubText,
                                                      dblBrutto.ToString,
                                                      strMwStKey,
                                                      dblMwSt.ToString,
                                                      Format(datValuta, "yyyyMMdd"))

            Else

                strSteuerFeld = objFiBebu.GetSteuerfeld(lngKto.ToString,
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

    Friend Function FcPGVKTreatment(ByVal tblKrediB As DataTable,
                                                ByVal intKRGNbr As Int32,
                                                ByVal intKBelegNr As Int32,
                                                ByVal strCur As String,
                                                ByVal datValuta As Date,
                                                ByVal strIType As String,
                                                ByVal datPGVStart As Date,
                                                ByVal datPGVEnd As Date,
                                                ByVal intITotal As Int16,
                                                ByVal intITY As Int16,
                                                ByVal intINY As Int16,
                                                ByVal intAcctTY As Int16,
                                                ByVal intAcctNY As Int16,
                                                ByVal strPeriode As String,
                                                ByVal objdbcon As MySqlConnection,
                                                ByVal objsqlcon As SqlConnection,
                                                ByVal objsqlcmd As SqlCommand,
                                                ByVal intAccounting As Int16,
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
        Dim drKrediSub() As DataRow
        Dim strBebuEintragSoll As String
        Dim strBebuEintragHaben As String
        Dim strPeriodenInfoA() As String
        Dim strPeriodenInfo As String
        Dim intReturnValue As Int32
        Dim strActualYear As String
        Dim datPGVEndSave As Date
        Dim datValutaSave As Date

        Try

            'Jahr retten
            strActualYear = strYear
            'Valuta saven
            datValutaSave = datValuta
            'Zuerst betroffene Buchungen selektieren
            drKrediSub = tblKrediB.Select("lngKredID=" + intKRGNbr.ToString)

            'Durch die Buchungen steppen
            For Each drKSubrow As DataRow In drKrediSub
                'Auflösung
                '=========

                'If strPGVType = "RV" Then
                '    'Damit die Periodenbuchung auf den ersten gebucht wird.
                '    datPGVStart = "2023-01-01"
                'End If
                datValuta = datValutaSave

                'Evtl. Aufteilen auf 2 Jahre
                For intYearLooper As Int16 = 0 To Year(DateAdd(DateInterval.Month, intITotal - 1, datPGVStart)) - Year(datValuta)

                    If intYearLooper = 0 And intITotal > 1 Then '2022 Then
                        dblNettoBetrag = drKSubrow("dblNetto") / intITotal * intITY
                        intSollKonto = intAcctTY
                    Else
                        dblNettoBetrag = drKSubrow("dblNetto") / intITotal * intINY
                        intSollKonto = intAcctNY
                    End If

                    If dblNettoBetrag <> 0 Then 'Falls in einem Jahr nichts zu buchen ist

                        strBelegDatum = Format(datValuta, "yyyyMMdd").ToString

                        'If intITotal = 1 Then
                        '    strDebiTextSoll = drKSubrow("strKredSubText") + ", TP Auflösung"
                        'Else
                        strDebiTextSoll = drKSubrow("strKredSubText") + ", PGV Auflösung"
                        'End If

                        strSteuerFeldSoll = "STEUERFREI"

                        intHabenKonto = drKSubrow("lngKto")

                        'If intITotal = 1 Then
                        '    strDebiTextHaben = drKSubrow("strKredSubText") + ", TP Auflösung"
                        '    If strPGVType = "VR" Then
                        '        'Valuta - Datum auf 01.01.22 legen, Achtung provisorisch
                        '        strValutaDatum = "20220101"
                        '        strBelegDatum = "20220101"
                        '    Else
                        '        strValutaDatum = Format(datValuta, "yyyyMMdd").ToString
                        '    End If
                        'Else
                        strDebiTextHaben = drKSubrow("strKredSubText") + ", PGV Auflösung"
                        strValutaDatum = Format(datValuta, "yyyyMMdd").ToString
                        'End If

                        strSteuerFeldHaben = "STEUERFREI"

                        'Falls nicht CHF dann umrechnen und auf CHF setzen
                        If strCur <> "CHF" Then
                            dblKursD = FcGetKurs(strCur, strValutaDatum)
                            strDebiCurrency = "CHF"
                        Else
                            dblKursD = 1.0#
                            strDebiCurrency = strCur
                        End If
                        dblKursH = dblKursD

                        'KORE
                        If drKSubrow("lngKST") > 0 Then

                            If drKSubrow("intSollHaben") = 1 Then 'Haben
                                strBebuEintragSoll = drKSubrow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
                                strBebuEintragHaben = Nothing
                            Else
                                strBebuEintragSoll = Nothing
                                strBebuEintragHaben = drKSubrow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
                            End If
                        Else
                            strBebuEintragSoll = Nothing
                            strBebuEintragHaben = Nothing

                        End If

                        If Year(datValuta) = 2023 And Year(datValuta) <> Val(strYear) Then 'Achtung provisorisch
                            'Zuerst Info-Table löschen
                            objdtInfo.Clear()
                            'Application.DoEvents()
                            'Im 2021 anmelden
                            intReturnValue = FcLoginSage2(objdbcon,
                                                          objsqlcon,
                                                          objsqlcmd,
                                                          objFinanz,
                                                          objfiBuha,
                                                          objdbBuha,
                                                          objdbPIFb,
                                                          objFiBebu,
                                                          objKrBuha,
                                                          intAccounting,
                                                          objdtInfo,
                                                          "2023",
                                                          strYear,
                                                          intTeqNbr,
                                                          intTeqNbrLY,
                                                          intTeqNbrPLY,
                                                          datPeriodFrom,
                                                          datPeriodTo,
                                                          strPeriodStatus)
                            'Application.DoEvents()

                        ElseIf Year(datValuta) = 2024 And Year(datValuta) <> Val(strYear) Then
                            'Zuerst Info-Table löschen
                            objdtInfo.Clear()
                            'Application.DoEvents()
                            'Im 2022 anmelden
                            intReturnValue = FcLoginSage2(objdbcon,
                                                          objsqlcon,
                                                          objsqlcmd,
                                                          objFinanz,
                                                          objfiBuha,
                                                          objdbBuha,
                                                          objdbPIFb,
                                                          objFiBebu,
                                                          objKrBuha,
                                                          intAccounting,
                                                          objdtInfo,
                                                          "2024",
                                                          strYear,
                                                          intTeqNbr,
                                                          intTeqNbrLY,
                                                          intTeqNbrPLY,
                                                          datPeriodFrom,
                                                          datPeriodTo,
                                                          strPeriodStatus)
                            'Application.DoEvents()

                        End If

                        'Buchen
                        Call objfiBuha.WriteBuchung(0,
                               intKBelegNr,
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

                'If strPGVType = "VR" Then
                '    'Falls VR dann PGVEnd zurück
                '    datPGVEnd = datPGVEndSave
                'End If

                'Falls FY dann 1312 auf 1311
                'Gab es eine Neutralisierung fürs FJ?
                If intINY > 0 And intITotal > 1 Then
                    'Was ist die aktuelle angemeldete Periode ?
                    strPeriodenInfo = objFinanz.GetPeriListe(0)
                    strPeriodenInfoA = Split(strPeriodenInfo, "{>}")

                    'Ist aktuell angemeldete Periode = FJ
                    If Year(datPGVEnd) <> Val(Strings.Left(strPeriodenInfo, 4)) Then
                        'Zuerst Info-Table löschen
                        objdtInfo.Clear()
                        'Application.DoEvents()
                        'Login ins FJ
                        intReturnValue = FcLoginSage2(objdbcon,
                                                          objsqlcon,
                                                          objsqlcmd,
                                                          objFinanz,
                                                          objfiBuha,
                                                          objdbBuha,
                                                          objdbPIFb,
                                                          objFiBebu,
                                                          objKrBuha,
                                                          intAccounting,
                                                          objdtInfo,
                                                          Year(datPGVEnd).ToString,
                                                          strYear,
                                                          intTeqNbr,
                                                          intTeqNbrLY,
                                                          intTeqNbrPLY,
                                                          datPeriodFrom,
                                                          datPeriodTo,
                                                          strPeriodStatus)

                        'Application.DoEvents()

                        '2311 -> 2312
                        datValuta = "2024-01-01" 'Achtung provisorisch
                        strValutaDatum = Format(datValuta, "yyyyMMdd").ToString
                        strBelegDatum = strValutaDatum
                        intSollKonto = intAcctTY
                        intHabenKonto = intAcctNY
                        strDebiTextSoll = drKSubrow("strKredSubText") + ", PGV AJ / FJ"
                        strDebiTextHaben = drKSubrow("strKredSubText") + ", PGV AJ / FJ"
                        strBebuEintragSoll = Nothing
                        strBebuEintragHaben = Nothing

                        'Buchen
                        Call objfiBuha.WriteBuchung(0,
                               intKBelegNr,
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

                'Falls nicht CHF dann umrechnen und auf CHF setzen
                If strCur <> "CHF" Then
                    dblKursD = FcGetKurs(strCur, strValutaDatum)
                    strDebiCurrency = "CHF"
                Else
                    dblKursD = 1.0#
                    strDebiCurrency = strCur
                End If
                dblKursH = dblKursD

                'Einzelene Monate buchen
                For intMonthLooper As Int16 = 0 To intITotal - 1
                    datValuta = DateAdd(DateInterval.Month, intMonthLooper, datPGVStart)
                    strValutaDatum = Format(datValuta, "yyyyMMdd").ToString
                    strBelegDatum = strValutaDatum
                    intSollKonto = drKSubrow("lngKto")
                    'If intITotal = 1 Then
                    '    strDebiTextSoll = drKSubrow("strKredSubText") + ", TP"
                    'Else
                    strDebiTextSoll = drKSubrow("strKredSubText") + ", PGV M " + (intMonthLooper + 1).ToString + "/ " + intITotal.ToString
                    'End If

                    dblNettoBetrag = drKSubrow("dblNetto") / intITotal
                    If intITotal = 1 Then
                        intHabenKonto = intAcctNY
                    Else
                        intHabenKonto = intAcctTY
                    End If

                    strDebiTextHaben = strDebiTextSoll

                    If drKSubrow("intSollHaben") = 1 Then 'Haben
                        strBebuEintragSoll = Nothing
                        strBebuEintragHaben = drKSubrow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
                        strSteuerFeldHaben = "STEUERFREI"
                    Else
                        strBebuEintragSoll = drKSubrow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
                        strSteuerFeldSoll = "STEUERFREI"
                        strBebuEintragHaben = Nothing
                    End If

                    If Year(datValuta) = 2023 And Year(datValuta) <> Val(strYear) Then 'Achtung provisorisch
                        'Zuerst Info-Table löschen
                        objdtInfo.Clear()
                        'Application.DoEvents()
                        'Im 2021 anmelden
                        intReturnValue = FcLoginSage2(objdbcon,
                                                          objsqlcon,
                                                          objsqlcmd,
                                                          objFinanz,
                                                          objfiBuha,
                                                          objdbBuha,
                                                          objdbPIFb,
                                                          objFiBebu,
                                                          objKrBuha,
                                                          intAccounting,
                                                          objdtInfo,
                                                          "2023",
                                                          strYear,
                                                          intTeqNbr,
                                                          intTeqNbrLY,
                                                          intTeqNbrPLY,
                                                          datPeriodFrom,
                                                          datPeriodTo,
                                                          strPeriodStatus)
                        'Application.DoEvents()

                    ElseIf Year(datValuta) = 2024 And Year(datValuta) <> Val(strYear) Then
                        'Zuerst Info-Table löschen
                        objdtInfo.Clear()
                        'Application.DoEvents()
                        'Im 2022 anmelden
                        intReturnValue = FcLoginSage2(objdbcon,
                                                          objsqlcon,
                                                          objsqlcmd,
                                                          objFinanz,
                                                          objfiBuha,
                                                          objdbBuha,
                                                          objdbPIFb,
                                                          objFiBebu,
                                                          objKrBuha,
                                                          intAccounting,
                                                          objdtInfo,
                                                          "2024",
                                                          strYear,
                                                          intTeqNbr,
                                                          intTeqNbrLY,
                                                          intTeqNbrPLY,
                                                          datPeriodFrom,
                                                          datPeriodTo,
                                                          strPeriodStatus)
                        'Application.DoEvents()

                    End If

                    'Buchen
                    Call objfiBuha.WriteBuchung(0,
                               intKBelegNr,
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
            If strYear <> strActualYear Then
                'Zuerst Info-Table löschen
                objdtInfo.Clear()
                'Application.DoEvents()
                'Im Aufrufjahr anmelden
                intReturnValue = FcLoginSage2(objdbcon,
                                                  objsqlcon,
                                                  objsqlcmd,
                                                  objFinanz,
                                                  objfiBuha,
                                                  objdbBuha,
                                                  objdbPIFb,
                                                  objFiBebu,
                                                  objKrBuha,
                                                  intAccounting,
                                                  objdtInfo,
                                                  strActualYear,
                                                  strYear,
                                                  intTeqNbr,
                                                  intTeqNbrLY,
                                                  intTeqNbrPLY,
                                                  datPeriodFrom,
                                                  datPeriodTo,
                                                  strPeriodStatus)
                'Application.DoEvents()
            End If
            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem PGV - Buchung Kreditoren", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 9

        Finally
            drKrediSub = Nothing
            strPeriodenInfoA = Nothing

        End Try

    End Function

    Friend Function FcLoginSage2(ByRef objdbconn As MySqlConnection,
                                       ByRef objsqlConn As SqlClient.SqlConnection,
                                       ByRef objsqlCom As SqlClient.SqlCommand,
                                       ByRef objFinanz As SBSXASLib.AXFinanz,
                                       ByRef objfiBuha As SBSXASLib.AXiFBhg,
                                       ByRef objdbBuha As SBSXASLib.AXiDbBhg,
                                       ByRef objdbPIFb As SBSXASLib.AXiPlFin,
                                       ByRef objFiBebu As SBSXASLib.AXiBeBu,
                                       ByRef objkrBuha As SBSXASLib.AXiKrBhg,
                                       ByVal intAccounting As Int16,
                                       ByRef objdtInfo As DataTable,
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
        Dim strPeriodenInfo As String
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
            Call objFinanz.ConnectSBSdb(System.Configuration.ConfigurationManager.AppSettings("OwnSageServer"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageDB"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageID"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSagePsw"), "")

            objdbconn.Open()
            strMandant = FcReadFromSettingsII("Buchh200_Name",
                                            intAccounting)
            objdbconn.Close()
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
            objdtInfo.Rows.Add("GeschäftsJ", strPeriode(3) + "-" + strPeriode(4) + ", teq: " + strPeriode(8).ToString + ", " + intTeqNbrLY.ToString + ", " + intTeqNbrPLY.ToString)
            objdtInfo.Rows.Add("Buchungen/ Status", strPeriode(5) + "-" + strPeriode(6) + "/ " + strPeriode(2))
            strYear = Strings.Left(strPeriode(4), 4)

            FcReturns = FcReadPeriodenDef(objsqlConn,
                                      objsqlCom,
                                      strPeriode(8),
                                      objdtInfo,
                                      strYear)

            'Perioden-Definition vom Tool einlesen
            'In einer ersten Phase nur erster DS einlesen
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

            'Finanz Buha öffnen
            If Not IsNothing(objfiBuha) Then
                objfiBuha = Nothing
            End If
            objfiBuha = New SBSXASLib.AXiFBhg
            objfiBuha = objFinanz.GetFibuObj()
            'Debitor öffnen
            If Not IsNothing(objdbBuha) Then
                objdbBuha = Nothing
            End If
            objdbBuha = New SBSXASLib.AXiDbBhg
            objdbBuha = objFinanz.GetDebiObj()
            If Not IsNothing(objdbPIFb) Then
                objdbPIFb = Nothing
            End If
            objdbPIFb = New SBSXASLib.AXiPlFin
            objdbPIFb = objfiBuha.GetCheckObj()
            If Not IsNothing(objFiBebu) Then
                objFiBebu = Nothing
            End If
            objFiBebu = New SBSXASLib.AXiBeBu
            objFiBebu = objFinanz.GetBeBuObj()
            'Kreditor
            If Not IsNothing(objkrBuha) Then
                objkrBuha = Nothing
            End If
            objkrBuha = New SBSXASLib.AXiKrBhg
            objkrBuha = objFinanz.GetKrediObj

            'Application.DoEvents()

        Catch ex As Exception
            MsgBox("OpenMandant:" + vbCrLf + "Error" + vbCrLf + "Error # " + Str(Err.Number) + " was generated by " + Err.Source + vbCrLf + Err.Description + " Fehlernummer" & Str(Err.Number And 65535))
            Err.Clear()
            End

        Finally
            objdtPeriodeLY = Nothing
            dtPeriods = Nothing

        End Try

    End Function

    Friend Function FcReadPeriodenDef(ByRef objSQLConnection As SqlClient.SqlConnection,
                                             ByRef objSQLCommand As SqlClient.SqlCommand,
                                             ByVal intPeriodenNr As Int32,
                                             ByRef objdtInfo As DataTable,
                                             ByVal strYear As String) As Int16

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

    Friend Function FcPGVKTreatmentYC(ByVal tblKrediB As DataTable,
                                                ByVal intKRGNbr As Int32,
                                                ByVal intKBelegNr As Int32,
                                                ByVal strCur As String,
                                                ByVal datValuta As Date,
                                                ByVal strIType As String,
                                                ByVal datPGVStart As Date,
                                                ByVal datPGVEnd As Date,
                                                ByVal intITotal As Int16,
                                                ByVal intITY As Int16,
                                                ByVal intINY As Int16,
                                                ByVal intAcctTY As Int16,
                                                ByVal intAcctNY As Int16,
                                                ByVal strPeriode As String,
                                                ByVal objdbcon As MySqlConnection,
                                                ByVal objsqlcon As SqlConnection,
                                                ByVal objsqlcmd As SqlCommand,
                                                ByVal intAccounting As Int16,
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
        Dim drKrediSub() As DataRow
        Dim strBebuEintragSoll As String
        Dim strBebuEintragHaben As String
        Dim strPeriodenInfoA() As String
        Dim strPeriodenInfo As String
        Dim intReturnValue As Int32
        Dim strActualYear As String
        Dim datPGVEndSave As Date
        Dim datValutaSave As Date


        Try

            'Jahr retten
            strActualYear = strYear
            'Valuta saven
            datValutaSave = datValuta
            'Zuerst betroffene Buchungen selektieren
            drKrediSub = tblKrediB.Select("lngKredID=" + intKRGNbr.ToString)

            'Durch die Buchungen steppen
            For Each drKSubrow As DataRow In drKrediSub
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
                        datPGVStart = "2024-01-01"
                        datValuta = datValutaSave
                        intITY = 1
                        intINY = 0
                        intAcctTY = 2312
                    End If

                End If

                'Evtl. Aufteilen auf 2 Jahre
                For intYearLooper As Int16 = Year(datValuta) To Year(datPGVEnd)

                    If intYearLooper = 2023 Then
                        dblNettoBetrag = drKSubrow("dblNetto") / intITotal * intITY
                        intSollKonto = intAcctTY
                    Else
                        dblNettoBetrag = drKSubrow("dblNetto") / intITotal * intINY
                        intSollKonto = intAcctNY
                    End If

                    If dblNettoBetrag <> 0 Then 'Falls in einem Jahr nichts zu buchen ist

                        strBelegDatum = Format(datValuta, "yyyyMMdd").ToString

                        If intITotal = 1 Then
                            If Year(datValuta) = 2023 Then
                                strDebiTextSoll = drKSubrow("strKredSubText") + ", TP"
                            Else
                                strDebiTextSoll = drKSubrow("strKredSubText") + ", TP Auflösung"
                            End If
                        Else
                            strDebiTextSoll = drKSubrow("strKredSubText") + ", PGV Auflösung"
                        End If

                        strSteuerFeldSoll = "STEUERFREI"

                        intHabenKonto = drKSubrow("lngKto")

                        If intITotal = 1 Then
                            strDebiTextHaben = strDebiTextSoll
                            If strPGVType = "VR" Then
                                'Valuta - Datum auf 01.01.24 legen, Achtung provisorisch
                                strValutaDatum = "20240101"
                                strBelegDatum = "20240101"
                            Else
                                'strValutaDatum = Format(datValuta, "yyyyMMdd").ToString
                                strValutaDatum = Format(datValuta, "yyyyMMdd").ToString
                                strBelegDatum = Format(datValuta, "yyyyMMdd").ToString
                            End If
                        Else
                            strDebiTextHaben = drKSubrow("strKredSubText") + ", PGV Auflösung"
                            strValutaDatum = Format(datValuta, "yyyyMMdd").ToString
                        End If

                        strSteuerFeldHaben = "STEUERFREI"

                        'Falls nicht CHF dann umrechnen und auf CHF setzen
                        If strCur <> "CHF" Then
                            dblKursD = FcGetKurs(strCur, strValutaDatum)
                            strDebiCurrency = "CHF"
                        Else
                            dblKursD = 1.0#
                            strDebiCurrency = strCur
                        End If
                        dblKursH = dblKursD

                        'KORE
                        If drKSubrow("lngKST") > 0 Then

                            If drKSubrow("intSollHaben") = 1 Then 'Haben
                                strBebuEintragSoll = Nothing
                                strBebuEintragHaben = drKSubrow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
                            Else
                                strBebuEintragSoll = Nothing
                                strBebuEintragHaben = drKSubrow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
                            End If
                        Else
                            strBebuEintragSoll = Nothing
                            strBebuEintragHaben = Nothing

                        End If

                        If Year(datValuta) = 2023 And Year(datValuta) <> Val(strYear) Then 'Achtung provisorisch
                            'Zuerst Info-Table löschen
                            objdtInfo.Clear()
                            'Application.DoEvents()
                            'Im 2021 anmelden
                            intReturnValue = FcLoginSage2(objdbcon,
                                                          objsqlcon,
                                                          objsqlcmd,
                                                          objFinanz,
                                                          objfiBuha,
                                                          objdbBuha,
                                                          objdbPIFb,
                                                          objFiBebu,
                                                          objKrBuha,
                                                          intAccounting,
                                                          objdtInfo,
                                                          "2023",
                                                          strYear,
                                                          intTeqNbr,
                                                          intTeqNbrLY,
                                                          intTeqNbrPLY,
                                                          datPeriodFrom,
                                                          datPeriodTo,
                                                          strPeriodStatus)
                            'Application.DoEvents()

                        ElseIf Year(datValuta) = 2024 And Year(datValuta) <> Val(strYear) Then
                            'Zuerst Info-Table löschen
                            objdtInfo.Clear()
                            'Application.DoEvents()
                            'Im 2022 anmelden
                            intReturnValue = FcLoginSage2(objdbcon,
                                                          objsqlcon,
                                                          objsqlcmd,
                                                          objFinanz,
                                                          objfiBuha,
                                                          objdbBuha,
                                                          objdbPIFb,
                                                          objFiBebu,
                                                          objKrBuha,
                                                          intAccounting,
                                                          objdtInfo,
                                                          "2024",
                                                          strYear,
                                                          intTeqNbr,
                                                          intTeqNbrLY,
                                                          intTeqNbrPLY,
                                                          datPeriodFrom,
                                                          datPeriodTo,
                                                          strPeriodStatus)
                            'Application.DoEvents()

                        End If

                        'doppelte Beleg-Nummern zulassen in HB
                        objfiBuha.CheckDoubleIntBelNbr = "N"

                        'Buchen
                        Call objfiBuha.WriteBuchung(0,
                               intKBelegNr,
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

                'Falls FY dann 1312 auf 1311
                'Gab es eine Neutralisierung fürs FJ?
                If intINY > 0 And intITotal > 1 Then
                    'Was ist die aktuelle angemeldete Periode ?
                    strPeriodenInfo = objFinanz.GetPeriListe(0)
                    strPeriodenInfoA = Split(strPeriodenInfo, "{>}")

                    'Ist aktuell angemeldete Periode = FJ
                    If Year(datPGVEnd) <> Val(Strings.Left(strPeriodenInfo, 4)) Then
                        'Zuerst Info-Table löschen
                        objdtInfo.Clear()
                        Application.DoEvents()
                        'Login ins FJ
                        intReturnValue = FcLoginSage2(objdbcon,
                                                          objsqlcon,
                                                          objsqlcmd,
                                                          objFinanz,
                                                          objfiBuha,
                                                          objdbBuha,
                                                          objdbPIFb,
                                                          objFiBebu,
                                                          objKrBuha,
                                                          intAccounting,
                                                          objdtInfo,
                                                          Year(datPGVEnd).ToString,
                                                          strYear,
                                                          intTeqNbr,
                                                          intTeqNbrLY,
                                                          intTeqNbrPLY,
                                                          datPeriodFrom,
                                                          datPeriodTo,
                                                          strPeriodStatus)

                        'Application.DoEvents()

                        '2311 -> 2312
                        datValuta = "2024-01-01" 'Achtung provisorisch
                        strValutaDatum = Format(datValuta, "yyyyMMdd").ToString
                        strBelegDatum = strValutaDatum
                        intSollKonto = intAcctTY
                        intHabenKonto = intAcctNY
                        strDebiTextSoll = drKSubrow("strKredSubText") + ", PGV AJ / FJ"
                        strDebiTextHaben = drKSubrow("strKredSubText") + ", PGV AJ / FJ"
                        strBebuEintragSoll = Nothing
                        strBebuEintragHaben = Nothing

                        'Buchen
                        Call objfiBuha.WriteBuchung(0,
                               intKBelegNr,
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
                    intSollKonto = drKSubrow("lngKto")
                    If intITotal = 1 Then
                        If Year(datValuta) = 2023 Then
                            strDebiTextSoll = drKSubrow("strKredSubText") + ", TP"
                        Else
                            strDebiTextSoll = drKSubrow("strKredSubText") + ", TP Auflösung"
                        End If

                    Else
                        strDebiTextSoll = drKSubrow("strKredSubText") + ", PGV M " + (intMonthLooper + 1).ToString + "/ " + intITotal.ToString
                    End If

                    dblNettoBetrag = drKSubrow("dblNetto") / intITotal
                    If intITotal = 1 Then
                        intHabenKonto = intAcctNY
                    Else
                        intHabenKonto = intAcctTY
                    End If

                    strDebiTextHaben = strDebiTextSoll

                    If drKSubrow("intSollHaben") = 1 Then 'Haben
                        strBebuEintragSoll = drKSubrow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
                        strBebuEintragHaben = Nothing
                    Else
                        strBebuEintragSoll = drKSubrow("lngKST").ToString + "{<}" + strDebiTextSoll + "{<}" + "CALCULATE" + "{>}"
                        strBebuEintragHaben = Nothing
                    End If

                    If Year(datValuta) = 2023 And Year(datValuta) <> Val(strYear) Then 'Achtung provisorisch
                        'Zuerst Info-Table löschen
                        objdtInfo.Clear()
                        'Application.DoEvents()
                        'Im 2021 anmelden
                        intReturnValue = FcLoginSage2(objdbcon,
                                                          objsqlcon,
                                                          objsqlcmd,
                                                          objFinanz,
                                                          objfiBuha,
                                                          objdbBuha,
                                                          objdbPIFb,
                                                          objFiBebu,
                                                          objKrBuha,
                                                          intAccounting,
                                                          objdtInfo,
                                                          "2023",
                                                          strYear,
                                                          intTeqNbr,
                                                          intTeqNbrLY,
                                                          intTeqNbrPLY,
                                                          datPeriodFrom,
                                                          datPeriodTo,
                                                          strPeriodStatus)
                        'Application.DoEvents()

                    ElseIf Year(datValuta) = 2024 And Year(datValuta) <> Val(strYear) Then
                        'Zuerst Info-Table löschen
                        objdtInfo.Clear()
                        'Application.DoEvents()
                        'Im 2022 anmelden
                        intReturnValue = FcLoginSage2(objdbcon,
                                                          objsqlcon,
                                                          objsqlcmd,
                                                          objFinanz,
                                                          objfiBuha,
                                                          objdbBuha,
                                                          objdbPIFb,
                                                          objFiBebu,
                                                          objKrBuha,
                                                          intAccounting,
                                                          objdtInfo,
                                                          "2024",
                                                          strYear,
                                                          intTeqNbr,
                                                          intTeqNbrLY,
                                                          intTeqNbrPLY,
                                                          datPeriodFrom,
                                                          datPeriodTo,
                                                          strPeriodStatus)
                        'Application.DoEvents()

                    End If

                    'doppelte Beleg-Nummern zulassen in HB
                    objfiBuha.CheckDoubleIntBelNbr = "N"

                    'Buchen
                    Call objfiBuha.WriteBuchung(0,
                               intKBelegNr,
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
            If strYear <> strActualYear Then
                'Zuerst Info-Table löschen
                objdtInfo.Clear()
                'Application.DoEvents()
                'Im Aufrufjahr anmelden
                intReturnValue = FcLoginSage2(objdbcon,
                                                  objsqlcon,
                                                  objsqlcmd,
                                                  objFinanz,
                                                  objfiBuha,
                                                  objdbBuha,
                                                  objdbPIFb,
                                                  objFiBebu,
                                                  objKrBuha,
                                                  intAccounting,
                                                  objdtInfo,
                                                  strActualYear,
                                                  strYear,
                                                  intTeqNbr,
                                                  intTeqNbrLY,
                                                  intTeqNbrPLY,
                                                  datPeriodFrom,
                                                  datPeriodTo,
                                                  strPeriodStatus)
                'Application.DoEvents()
            End If
            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem PGV - Buchung Kreditoren", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 9

        Finally
            drKrediSub = Nothing
            strPeriodenInfoA = Nothing

        End Try

    End Function

    Friend Function FcWriteToKrediRGTable(ByVal intMandant As Int32,
                                                 ByVal strKredID As String,
                                                 ByVal datDate As Date,
                                                 ByVal intBelegNr As Int32) As Int16

        'Returns 0=ok, 1=Problem

        Dim strSQL As String
        Dim intAffected As Int16
        Dim objdbAccessConn As New OleDb.OleDbConnection
        Dim objlocOLEdbcmd As New OleDb.OleDbCommand
        Dim objlocOracmd As New OracleCommand
        Dim objlocMySQLRGConn As New MySqlConnection
        Dim objlocMySQLRGcmd As New MySqlCommand
        Dim strNameKRGTable As String
        Dim strBelegNrName As String
        Dim strKRGNbrFieldName As String
        Dim strKRGTableType As String
        Dim strKRGNbrFieldType As String
        Dim strMDBName As String

        'objMySQLConn.Open()

        strMDBName = FcReadFromSettingsII("Buchh_KRGTableMDB", intMandant)
        strKRGTableType = FcReadFromSettingsII("Buchh_KRGTableType", intMandant)
        strNameKRGTable = FcReadFromSettingsII("Buchh_TableKred", intMandant)
        strBelegNrName = FcReadFromSettingsII("Buchh_TableKRGBelegNrName", intMandant)
        strKRGNbrFieldName = FcReadFromSettingsII("Buchh_TableKRGNbrFieldName", intMandant)
        strKRGNbrFieldType = FcReadFromSettingsII("Buchh_TableKRGNbrFieldType", intMandant)
        'strSQL = "UPDATE " + strNameRGTable + " SET gebucht=true, gebuchtDatum=#" + Format(datDate, "yyyy-MM-dd").ToString + "#, " + strBelegNrName + "=" + intBelegNr.ToString + " WHERE " + strRGNbrFieldName + "=" + strRGNbr

        Try

            If strKRGTableType = "A" Then
                'Access
                Call FcInitAccessConnecation(objdbAccessConn, strMDBName)
                'objdbAccessConn.Open()
                strSQL = "UPDATE " + strNameKRGTable + " SET Kredigebucht=true, KredigebuchtDatum=#" + Format(datDate, "yyyy-MM-dd").ToString + "#, " + strBelegNrName + "='" + intBelegNr.ToString + "' WHERE " + strKRGNbrFieldName + "=" + IIf(strKRGNbrFieldType = "T", "'", "") + strKredID + IIf(strKRGNbrFieldType = "T", "'", "")
                objlocOLEdbcmd.CommandText = strSQL
                objlocOLEdbcmd.Connection = objdbAccessConn
                objlocOLEdbcmd.Connection.Open()
                intAffected = objlocOLEdbcmd.ExecuteNonQuery()
                objlocOLEdbcmd.Connection.Close()

            ElseIf strKRGTableType = "M" Then
                'MySQL
                'Bei IG andere Feldnamen
                If intMandant = 25 Then
                    strSQL = "UPDATE " + strNameKRGTable + " SET IGKBooked=true, IGKBDate=DATE('" + Format(datDate, "yyyy-MM-dd").ToString + "'), " + strBelegNrName + "='" + intBelegNr.ToString + "' WHERE " + strKRGNbrFieldName + "=" + IIf(strKRGNbrFieldType = "T", "'", "") + strKredID + IIf(strKRGNbrFieldType = "T", "'", "")
                Else
                    strSQL = "UPDATE " + strNameKRGTable + " SET Kredigebucht=true, KredigebuchtDatum=DATE('" + Format(datDate, "yyyy-MM-dd").ToString + "'), " + strBelegNrName + "='" + intBelegNr.ToString + "' WHERE " + strKRGNbrFieldName + "=" + IIf(strKRGNbrFieldType = "T", "'", "") + strKredID + IIf(strKRGNbrFieldType = "T", "'", "")
                End If

                objlocMySQLRGConn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(strMDBName)
                objlocMySQLRGConn.Open()
                objlocMySQLRGcmd.Connection = objlocMySQLRGConn
                objlocMySQLRGcmd.CommandText = strSQL
                intAffected = objlocMySQLRGcmd.ExecuteNonQuery()
                If intAffected = 0 Then
                    Return 9
                End If

            End If

            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return 1

        Finally
            'If objdbAccessConn.State = ConnectionState.Open Then
            ' objdbAccessConn.Close()
            'End If

            If objlocMySQLRGConn.State = ConnectionState.Open Then
                objlocMySQLRGConn.Close()
            End If

            'If objMySQLConn.State = ConnectionState.Open Then
            '    objMySQLConn.Close()
            'End If
            objdbAccessConn = Nothing
            objlocOLEdbcmd = Nothing
            objlocMySQLRGConn = Nothing
            objlocMySQLRGcmd = Nothing


        End Try

    End Function


End Class