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



    Public Sub InitDB()

        Dim strIdentityName As String

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

        Catch ex As Exception


        End Try

    End Sub

    Private Sub frmKredDisp_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        FELD_SEP = "{<}"
        REC_SEP = "{>}"
        KSTKTR_SEP = "{-}"

        FELD_SEP_OUT = "{>}"
        REC_SEP_OUT = "{<}"

        Call InitDB()

    End Sub


    Friend Function FcKrediDisplay(intMandant As Int32,
                                   LstMandant As ListBox,
                                   LstBPerioden As ListBox) As Int16

        Dim objdbConn As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString"))
        Dim objdbtaskcmd As New MySqlCommand
        Dim objdbMSSQLConn As New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("SQLConnectionString"))
        Dim objdbSQLcommand As New SqlCommand

        Dim intFcReturns As Int16
        Dim strPeriode As String
        Dim strYearCh As String
        Dim BgWCheckKrediLocArgs As New BgWCheckDebitArgs
        Dim objdbtasks As New DataTable

        'Dim objFinanz As New SBSXASLib.AXFinanz
        'Dim objfiBuha As New SBSXASLib.AXiFBhg
        'Dim objdbBuha As New SBSXASLib.AXiDbBhg
        'Dim objdbPIFb As New SBSXASLib.AXiPlFin
        'Dim objFiBebu As New SBSXASLib.AXiBeBu
        'Dim objKrBuha As New SBSXASLib.AXiKrBhg


        Try

            Me.Cursor = Cursors.WaitCursor

            'Zuerst in tblImportTasks setzen
            objdbtaskcmd.Connection = objdbConn
            objdbtaskcmd.Connection.Open()
            objdbtaskcmd.CommandText = "SELECT * FROM tblimporttasks WHERE IdentityName='" + frmImportMain.LblIdentity.Text + "' AND Type='C'"
            objdbtasks.Load(objdbtaskcmd.ExecuteReader())
            If objdbtasks.Rows.Count > 0 Then
                'update
                objdbtaskcmd.CommandText = "UPDATE tblimporttasks SET Mandant=" + Convert.ToString(LstMandant.SelectedIndex) + ", Periode=" + Convert.ToString(LstBPerioden.SelectedIndex) + " WHERE IdentityName='" + frmImportMain.LblIdentity.Text + "' AND Type='C'"
            Else
                'insert
                objdbtaskcmd.CommandText = "INSERT INTO tblimporttasks (IdentityName, Type, Mandant, Periode) VALUES ('" + frmImportMain.LblIdentity.Text + "', 'C', " + Convert.ToString(LstMandant.SelectedIndex) + ", " + Convert.ToString(LstBPerioden.SelectedIndex) + ")"
            End If
            objdbtaskcmd.ExecuteNonQuery()
            objdbtaskcmd.Connection.Close()

            'DGVs
            dgvBookings.DataSource = Nothing
            dgvBookingSub.DataSource = Nothing

            Me.butImport.Enabled = False

            'Zuerst evtl. vorhandene DS löschen in Tabelle
            MySQLdaKreditoren.DeleteCommand.Connection.Open()
            MySQLdaKreditoren.DeleteCommand.ExecuteNonQuery()
            MySQLdaKreditoren.DeleteCommand.Connection.Close()

            MySQLdaKreditorenSub.DeleteCommand.Connection.Open()
            MySQLdaKreditorenSub.DeleteCommand.ExecuteNonQuery()
            MySQLdaKreditorenSub.DeleteCommand.Connection.Close()

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

            dgvInfo.DataSource = dsKreditoren.Tables("tblKreditorenInfo")

            'Datums-Tabelle erstellen
            dsKreditoren.Tables.Add("tblDebitorenDates")
            Dim col7 As DataColumn = New DataColumn("intYear")
            col7.DataType = System.Type.GetType("System.Int16")
            col7.Caption = "Year"
            dsKreditoren.Tables("tblDebitorenDates").Columns.Add(col7)
            Dim col3 As DataColumn = New DataColumn("strDatType")
            col3.DataType = System.Type.GetType("System.String")
            col3.MaxLength = 50
            col3.Caption = "Datum-Typ"
            dsKreditoren.Tables("tblDebitorenDates").Columns.Add(col3)
            Dim col4 As DataColumn = New DataColumn("datFrom")
            col4.DataType = System.Type.GetType("System.DateTime")
            col4.Caption = "Von"
            dsKreditoren.Tables("tblDebitorenDates").Columns.Add(col4)
            Dim col5 As DataColumn = New DataColumn("datTo")
            col5.DataType = System.Type.GetType("System.DateTime")
            col5.Caption = "Bis"
            dsKreditoren.Tables("tblDebitorenDates").Columns.Add(col5)
            Dim col6 As DataColumn = New DataColumn("strStatus")
            col6.DataType = System.Type.GetType("System.String")
            col6.Caption = "S"
            dsKreditoren.Tables("tblDebitorenDates").Columns.Add(col6)
            dgvDates.DataSource = dsKreditoren.Tables("tblDebitorenDates")

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
                                  dsKreditoren.Tables("tblKreditorenInfo"),
                                  dsKreditoren.Tables("tblDebitorenDates"),
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
                                       dsKreditoren.Tables("tblDebitorenDates"),
                                       strPeriode)
                Else
                    'Periode ezreugen und auf N stellen
                    strYearCh = Convert.ToString(Val(strYear) - 1)
                    dsKreditoren.Tables("tblDebitorenDates").Rows.Add(strYearCh, "GJ n/a", DateSerial(Convert.ToUInt16(strYearCh), 1, 1), DateSerial(Convert.ToUInt16(strYearCh), 12, 31), "N")
                End If

                'Gibt es ein Folgehahr?
                If LstBPerioden.SelectedIndex + 1 < LstBPerioden.Items.Count Then
                    strPeriode = LstBPerioden.Items(LstBPerioden.SelectedIndex + 1)
                    'Peeriodendef holen
                    Call Main.FcLoginSage4(intMandant,
                                       dsKreditoren.Tables("tblDebitorenDates"),
                                       strPeriode)
                Else
                    'Periode ezreugen und auf N stellen
                    strYearCh = Convert.ToString(Val(strYear) + 1)
                    dsKreditoren.Tables("tblDebitorenDates").Rows.Add(strYearCh, "GJ n/a", DateSerial(Convert.ToUInt16(strYearCh), 1, 1), DateSerial(Convert.ToUInt16(strYearCh), 12, 31), "N")
                End If

            ElseIf LstBPerioden.Items.Count = 1 Then 'es gibt genau 1 Jahr
                'gewähltes Jahr checken
                Call Main.FcLoginSage4(intMandant,
                                       dsKreditoren.Tables("tblDebitorenDates"),
                                       strPeriode)
                'VJ erzeugen
                strYearCh = Convert.ToString(Val(strYear) - 1)
                dsKreditoren.Tables("tblDebitorenDates").Rows.Add(strYearCh, "GJ n/a", DateSerial(Convert.ToUInt16(strYearCh), 1, 1), DateSerial(Convert.ToUInt16(strYearCh), 12, 31), "N")

                'FJ erzeugen
                strYearCh = Convert.ToString(Val(strYear) + 1)
                dsKreditoren.Tables("tblDebitorenDates").Rows.Add(strYearCh, "GJ n/a", DateSerial(Convert.ToUInt16(strYearCh), 1, 1), DateSerial(Convert.ToUInt16(strYearCh), 12, 31), "N")

            End If


            'Dim clImp As New ClassImport
            'clImp.FcKreditFill(intMandant)
            'clImp = Nothing

            BgWLoadKredi.RunWorkerAsync(intMandant)

            Do While BgWLoadKredi.IsBusy
                Application.DoEvents()
            Loop


            'Tabellentyp darstellen
            Me.lblDB.Text = Main.FcReadFromSettingsII("Buchh_KRGTableType", intMandant)


            MySQLdaKreditoren.Fill(dsKreditoren, "tblKrediHeadsFromUser")
            MySQLdaKreditorenSub.Fill(dsKreditoren, "tblKrediSubsFromUser")


            'Application.DoEvents()

            'Dim clCheck As New ClassCheck
            'clCheck.FcCheckKredit(intMandant,
            '                  dsKreditoren,
            '                  Finanz,
            '                  FBhg,
            '                  KrBhg,
            '                  BeBu,
            '                  dsKreditoren.Tables("tblKreditorenInfo"),
            '                  dsKreditoren.Tables("tblDebitorenDates"),
            '                  frmImportMain.lstBoxMandant.Text,
            '                  strYear,
            '                  strPeriode,
            '                  datPeriodFrom,
            '                  datPeriodTo,
            '                  strPeriodStatus,
            '                  frmImportMain.chkValutaCorrect.Checked,
            '                  frmImportMain.dtpValutaCorrect.Value)

            'clCheck = Nothing

            BgWCheckKrediLocArgs.intMandant = intMandant
            BgWCheckKrediLocArgs.strMandant = frmImportMain.lstBoxMandant.GetItemText(frmImportMain.lstBoxMandant.SelectedItem)
            BgWCheckKrediLocArgs.intTeqNbr = intTeqNbr
            BgWCheckKrediLocArgs.intTeqNbrLY = intTeqNbrLY
            BgWCheckKrediLocArgs.intTeqNbrPLY = intTeqNbrPLY
            BgWCheckKrediLocArgs.strYear = strYear
            BgWCheckKrediLocArgs.strPeriode = LstBPerioden.GetItemText(LstBPerioden.SelectedItem)
            BgWCheckKrediLocArgs.booValutaCor = frmImportMain.chkValutaCorrect.Checked
            BgWCheckKrediLocArgs.datValutaCor = frmImportMain.dtpValutaCorrect.Value

            BgWCheckKredi.RunWorkerAsync(BgWCheckKrediLocArgs)

            Do While BgWCheckKredi.IsBusy
                Application.DoEvents()
            Loop

            Debug.Print("Vor Refresh DGV")

            'Grid neu aufbauen
            dgvBookings.DataSource = dsKreditoren.Tables("tblKrediHeadsFromUser")
            dgvBookingSub.DataSource = dsKreditoren.Tables("tblKrediSubsFromUser")

            intFcReturns = FcInitdgvInfo(dgvInfo)
            intFcReturns = FcInitdgvKreditoren(dgvBookings)
            intFcReturns = FcInitdgvKrediSub(dgvBookingSub)
            intFcReturns = FcInitdgvDate(dgvDates)


            'Anzahl schreiben
            txtNumber.Text = Me.dsKreditoren.Tables("tblKrediHeadsFromUser").Rows.Count.ToString

            Me.Cursor = Cursors.Default

            Me.butImport.Enabled = True
            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Generelles Problem Kredi-Check" + Err.Number.ToString)
            Err.Clear()
            Return 1

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
            objdbtaskcmd = Nothing
            objdbtasks = Nothing

            BgWCheckKrediLocArgs = Nothing

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


            Me.Cursor = Cursors.WaitCursor
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

        Finally
            'Neu aufbauen
            'butKreditoren_Click(butDebitoren, EventArgs.Empty)

            Me.Cursor = Cursors.Default
            'Me.butImportK.Enabled = True
            'Me.Close()

        End Try


    End Sub

    Private Sub dsKreditoren_MergeFailed(sender As Object, e As MergeFailedEventArgs) Handles dsKreditoren.MergeFailed

        MessageBox.Show("dsKreditoren_MergeFailed")

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

            Debug.Print("BW Read Start " + Convert.ToString(intAccounting))
            objmysqlcomdwritehead.Connection = objdbConnZHDB02
            objmysqlcomdwritesub.Connection = objdbConnZHDB02

            'Für den Save der Records
            strIdentityName = System.Security.Principal.WindowsIdentity.GetCurrent().Name
            strIdentityName = Strings.Replace(strIdentityName, "\", "/")

            strMDBName = Main.FcReadFromSettingsII("Buchh_KRGTableMDB",
                                              intAccounting)

            strSQL = Main.FcReadFromSettingsII("Buchh_SQLHeadKred",
                                          intAccounting)

            strKRGTableType = Main.FcReadFromSettingsII("Buchh_KRGTableType",
                                                   intAccounting)

            objdslocKredihead.EnforceConstraints = False


            Debug.Print("BW Read Before Read Head " + Convert.ToString(intAccounting))

            If strKRGTableType = "A" Then

                'Access
                Call Main.FcInitAccessConnecation(objdbAccessConn,
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
            objdslocKredihead.AcceptChanges()

            objdslocKredisub.EnforceConstraints = False


            strSQLToParse = Main.FcReadFromSettingsII("Buchh_SQLDetailKred",
                                                    intAccounting)

            intFcReturns = Main.FcInitInsCmdKHeads(objmysqlcomdwritehead)

            Debug.Print("BW Read Write Heads " + Convert.ToString(intAccounting))

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
                strSQLSub = MainKreditor.FcSQLParseKredi(strSQLToParse,
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

                Debug.Print("BW Read Write Subs")
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
        Dim objFinanz As New SBSXASLib.AXFinanz
        Dim objfiBuha As New SBSXASLib.AXiFBhg
        Dim objKrBuha As New SBSXASLib.AXiKrBhg
        Dim objFiBebu As New SBSXASLib.AXiBeBu

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

        Try

            Debug.Print("Start Kredi Check " + Convert.ToString(BgWCheckKrediArgsInProc.intMandant))
            'Finanz-Obj init
            'Login
            Call objFinanz.ConnectSBSdb(System.Configuration.ConfigurationManager.AppSettings("OwnSageServer"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageDB"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageID"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSagePsw"), "")

            strMandant = Main.FcReadFromSettingsII("Buchh200_Name",
                                            BgWCheckKrediArgsInProc.intMandant)

            booAccOk = objFinanz.CheckMandant(strMandant)
            'Open Mandantg
            objFinanz.OpenMandant(strMandant, BgWCheckKrediArgsInProc.strPeriode)

            objfiBuha = objFinanz.GetFibuObj()
            objKrBuha = objFinanz.GetKrediObj()
            objFiBebu = objFinanz.GetBeBuObj()

            'Variablen einesen
            booAutoCorrect = Convert.ToBoolean(Convert.ToInt16(Main.FcReadFromSettingsII("Buchh_HeadKAutoCorrect", BgWCheckKrediArgsInProc.intMandant)))
            booCpyKSTToSub = Convert.ToBoolean(Convert.ToInt16(Main.FcReadFromSettingsII("Buchh_KKSTHeadToSub", BgWCheckKrediArgsInProc.intMandant)))
            booPKPrivate = IIf(Main.FcReadFromSettingsII("Buchh_PKKrediTable", BgWCheckKrediArgsInProc.intMandant) = "t_customer", True, False)
            booDiffHeadText = IIf(Main.FcReadFromSettingsII("Buchh_KTextSpecial", BgWCheckKrediArgsInProc.intMandant) = "0", False, True)
            booLeaveSubText = IIf(Main.FcReadFromSettingsII("Buchh_KSubLeaveText", BgWCheckKrediArgsInProc.intMandant) = "0", False, True)
            booDiffSubText = IIf(Main.FcReadFromSettingsII("Buchh_KSubTextSpecial", BgWCheckKrediArgsInProc.intMandant) = "0", False, True)

            For Each row As DataRow In dsKreditoren.Tables("tblKrediHeadsFromUser").Rows

                'If row("lngKredID") = "117383" Then Stop
                'Runden
                row("dblKredNetto") = Decimal.Round(row("dblKredNetto"), 2, MidpointRounding.AwayFromZero)
                row("dblKredMwSt") = Decimal.Round(row("dblKredMwst"), 2, MidpointRounding.AwayFromZero)
                row("dblKredBrutto") = Decimal.Round(row("dblKredBrutto"), 2, MidpointRounding.AwayFromZero)

                'Status-String erstellen
                'Kreditor 01
                intReturnValue = MainKreditor.FcGetRefKrediNr(IIf(IsDBNull(row("lngKredNbr")), 0, row("lngKredNbr")),
                                                            BgWCheckKrediArgsInProc.intMandant,
                                                            intKreditorNew)

                If intKreditorNew <> 0 Then
                    intReturnValue = MainKreditor.FcCheckKreditor(intKreditorNew,
                                                                  row("intBuchungsart"),
                                                                  objKrBuha)
                Else
                    intReturnValue = 2
                End If
                strBitLog = Trim(intReturnValue.ToString)

                'Kto 02
                'intReturnValue = FcCheckKonto(row("lngKredKtoNbr"), objfiBuha, row("dblKredMwSt"), 0)
                intReturnValue = 0
                strBitLog += Trim(intReturnValue.ToString)

                'Currency 03
                intReturnValue = Main.FcCheckCurrency(row("strKredCur"), objfiBuha)
                strBitLog += Trim(intReturnValue.ToString)

                'Sub 04
                intReturnValue = Main.FcCheckKrediSubBookings2(row("lngKredID"),
                                                         dsKreditoren.Tables("tblKrediSubsFromUser"),
                                                         intSubNumber,
                                                         dblSubBrutto,
                                                         dblSubNetto,
                                                         dblSubMwSt,
                                                         IIf(IsDBNull(row("datKredValDatum")), #1789-09-17#, row("datKredValDatum")),
                                                         objfiBuha,
                                                         objFiBebu,
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

                            dsKreditoren.Tables("tblKrediSubsFromUser").AcceptChanges()

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
                intReturnValue = Main.FcCheckBelegHead(row("intBuchungsart"),
                                                  IIf(IsDBNull(row("dblKredBrutto")), 0, row("dblKredBrutto")),
                                                  IIf(IsDBNull(row("dblKredNetto")), 0, row("dblKredNetto")),
                                                  IIf(IsDBNull(row("dblKredMwSt")), 0, row("dblKredMwSt")),
                                                  dblRDiffBrutto)
                strBitLog += Trim(intReturnValue.ToString)

                'OP - Nummer prüfen 08
                'intReturnValue = FcCreateDebRef(objdbconn, intAccounting, row("strDebiBank"), row("strDebRGNbr"), row("intBuchungsart"), strDebiReferenz)
                strCleanOPNbr = IIf(IsDBNull(row("strOPNr")), "", row("strOPNr"))
                intReturnValue = MainKreditor.FcChCeckKredOP(strCleanOPNbr, IIf(IsDBNull(row("strKredRGNbr")), "", row("strKredRGNbr")))
                row("strOPNr") = strCleanOPNbr
                strBitLog += Trim(intReturnValue.ToString)
                'OP - Verdopplung 09
                If row("dblKredBrutto") < 0 Then
                    strKredTyp = "G"
                Else
                    strKredTyp = "R"
                End If
                intReturnValue = MainKreditor.FcCheckKrediOPDouble(objKrBuha,
                                                                   intKreditorNew,
                                                                   row("strKredRGNbr"),
                                                                   row("strKredCur"),
                                                                   strKredTyp)
                strBitLog += Trim(intReturnValue.ToString)

                'PGV => Prüfung vor Valuta-Datum da Valuta-Datum verändert wird. PGV soll nur möglich sein wenn rebilled
                If Not IsDBNull(row("datPGVFrom")) And MainKreditor.FcIsAllKrediRebilled(dsKreditoren.Tables("tblKrediSubsFromUser"), row("lngKredID")) = 0 Then
                    row("booPGV") = True
                ElseIf Not IsDBNull(row("datPGVFrom")) And MainKreditor.FcIsAllKrediRebilled(dsKreditoren.Tables("tblKrediSubsFromUser"), row("lngKredID")) = 1 Then
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
                intReturnValue = Main.FcCheckDate2(IIf(IsDBNull(row("datKredValDatum")), row("datKredRGDatum"), row("datKredValDatum")),
                                              strYear,
                                              dsKreditoren.Tables("tblDebitorenDates"),
                                              False)

                'Bei PGV checken ob PGV-Startdatum in blockierter Periode
                If row("booPGV") And intReturnValue = 0 Then
                    'Ist TP ?
                    If intMonthsAJ + intMonthsNJ = 1 Then
                        'Ist Differenz Jahre grösser 1?
                        If Math.Abs(Convert.ToInt16(strYear) - Year(row("datPGVTo"))) > 1 Then
                            intReturnValue = 4
                        Else
                            intReturnValue = Main.FcCheckDate2(row("datPGVTo"),
                                                      strYear,
                                                      dsKreditoren.Tables("tblDebitorenDates"),
                                                      True)
                        End If
                    Else
                        'mehrere Monate PGV
                        For intMonthCounter = 0 To intPGVMonths - 1
                            'Ist Differenz Jahre grösser 1?
                            If Math.Abs(Convert.ToInt16(strYear) - Year(row("datPGVFrom"))) > 1 Then
                                intReturnValue = 4
                            Else
                                intReturnValue = Main.FcCheckDate2(DateAndTime.DateAdd(DateInterval.Month, intMonthCounter, row("datPGVFrom")),
                                                          strYear,
                                                          dsKreditoren.Tables("tblDebitorenDates"),
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
                intReturnValue = Main.FcCheckDate2(IIf(IsDBNull(row("datKredRGDatum")), #1789-09-17#, row("datKredRGDatum")),
                                              strYear,
                                              dsKreditoren.Tables("tblDebitorenDates"),
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
                            If Strings.Right(row("strKredRef"), 1) <> Main.FcModulo10(Strings.Left(row("strKredRef"), Len(row("strKredRef")) - 1)) Then
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
                intReturnValue = Main.FcCheckDebiIntBank(BgWCheckKrediArgsInProc.intMandant,
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
                intReturnValue = Main.FcCheckPayType(intPayType,
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
                            intReturnValue = MainKreditor.FcIsPrivateKreditorCreatable(intKreditorNew,
                                                                                        objKrBuha,
                                                                                        objfiBuha,
                                                                                        IIf(IsDBNull(row("intPayType")), 3, row("intPayType")),
                                                                                        IIf(IsDBNull(row("strKredRef")), "", row("strKredRef")),
                                                                                        intintBank,
                                                                                        IIf(IsDBNull(row("strKrediBank")), "", row("strKrediBank")),
                                                                                        BgWCheckKrediArgsInProc.strMandant,
                                                                                        BgWCheckKrediArgsInProc.intMandant)
                        Else
                            intReturnValue = MainKreditor.FcIsKreditorCreatable(intKreditorNew,
                                                                            objKrBuha,
                                                                            objfiBuha,
                                                                            BgWCheckKrediArgsInProc.strMandant,
                                                                            IIf(IsDBNull(row("intPayType")), 9, row("intPayType")),
                                                                            IIf(IsDBNull(row("strKredRef")), "", row("strKredRef")),
                                                                            intintBank,
                                                                            IIf(IsDBNull(row("strKrediBank")), "", row("strKrediBank")),
                                                                            BgWCheckKrediArgsInProc.intMandant)

                        End If
                        If intReturnValue = 0 Then
                            strStatus += " erstellt"
                            intReturnValue = MainKreditor.FcReadKreditorName(objKrBuha,
                                                                                strKreditorNew,
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
                    intReturnValue = MainKreditor.FcReadKreditorName(objKrBuha,
                                                                        strKreditorNew,
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
                        intReturnValue = MainKreditor.FcCheckKreditBank(objKrBuha,
                                                       intKreditorNew,
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
                    row("strKredKtoBez") = MainDebitor.FcReadDebitorKName(objfiBuha,
                                                                          row("lngKredKtoNbr"))
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
                    strKrediHeadText = MainDebitor.FcSQLParse(Main.FcReadFromSettingsII("Buchh_KTextSpecialText",
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
                        strKrediSubText = MainDebitor.FcSQLParse(Main.FcReadFromSettingsII("Buchh_KSubTextSpecialText",
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
            objFinanz = Nothing
            objfiBuha = Nothing
            objKrBuha = Nothing
            objFiBebu = Nothing
            selsubrow = Nothing
            BgWCheckKrediArgsInProc = Nothing

            System.GC.Collect()

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

        Me.dsKreditoren = Nothing
        Me.Dispose()
        Application.Restart()

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

        Dim objFinanz As New SBSXASLib.AXFinanz
        Dim objfiBuha As New SBSXASLib.AXiFBhg
        Dim objdbBuha As New SBSXASLib.AXiDbBhg
        Dim objdbPIFb As New SBSXASLib.AXiPlFin
        Dim objFiBebu As New SBSXASLib.AXiBeBu
        Dim objKrBuha As New SBSXASLib.AXiKrBhg

        Try

            objdbSQLcommand.Connection = objdbMSSQLConn

            'Start in Sync schreiben
            intReturnValue = WFDBClass.FcWriteStartToSync(objdbConnZHDB02,
                                                          BgWImportKrediArgsInProc.intMandant,
                                                          2,
                                                          dsKreditoren.Tables("tblKrediHeadsFromUser").Rows.Count)

            'Finanz-Obj init
            'Login
            Call objFinanz.ConnectSBSdb(System.Configuration.ConfigurationManager.AppSettings("OwnSageServer"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageDB"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageID"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSagePsw"), "")

            strMandant = Main.FcReadFromSettingsII("Buchh200_Name",
                                            BgWImportKrediArgsInProc.intMandant)

            booAccOk = objFinanz.CheckMandant(strMandant)
            'Open Mandant
            objFinanz.OpenMandant(strMandant, strPeriode)
            objfiBuha = objFinanz.GetFibuObj()
            objdbBuha = objFinanz.GetDebiObj()
            objdbPIFb = objfiBuha.GetCheckObj()
            objFiBebu = objFinanz.GetBeBuObj()
            objKrBuha = objFinanz.GetKrediObj()

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

                            intReturnValue = MainKreditor.FcCheckKrediExistance(objdbMSSQLConn,
                                                                                objdbSQLcommand,
                                                                                intKredBelegsNummer,
                                                                                "G",
                                                                                BgWImportKrediArgsInProc.intTeqNbr,
                                                                                BgWImportKrediArgsInProc.intTeqNbrLY,
                                                                                BgWImportKrediArgsInProc.intTeqNbrPLY,
                                                                                objKrBuha)

                        Else
                            strBuchType = "R"
                            'strZahlSperren = "N"
                            'Belegsnummer abholen
                            objKrBuha.IncrBelNbr = "J"
                            intKredBelegsNummer = objKrBuha.GetNextBelNbr("R")
                            'Muss auf Nicht hochzählen gesetzt werden da Sage 200 nicht merkt, dass Beleg-Nr. schon vergeben worden sind. => In den Einstellungen muss von Zeit zu Zeit der Zähler geändert werden
                            objKrBuha.IncrBelNbr = "N"

                            intReturnValue = MainKreditor.FcCheckKrediExistance(objdbMSSQLConn,
                                                                                objdbSQLcommand,
                                                                                intKredBelegsNummer,
                                                                                "R",
                                                                                BgWImportKrediArgsInProc.intTeqNbr,
                                                                                BgWImportKrediArgsInProc.intTeqNbrLY,
                                                                                BgWImportKrediArgsInProc.intTeqNbrPLY,
                                                                                objKrBuha)

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
                            dblKurs = Main.FcGetKurs(strCurrency,
                                                     strValutaDatum,
                                                     objfiBuha)
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
                                intReturnValue = Main.FcGetSteuerFeld2(objfiBuha,
                                                                      strSteuerFeld,
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
                            dblKurs = Main.FcGetKurs(strCurrency,
                                                     strValutaDatum,
                                                     objfiBuha)
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
                                        intReturnValue = Main.FcGetSteuerFeld(objfiBuha,
                                                                                 strSteuerFeldSoll,
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
                                        intReturnValue = Main.FcGetSteuerFeld(objfiBuha,
                                                                                  strSteuerFeldHaben,
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

                                intReturnValue = MainKreditor.FcPGVKTreatment(objfiBuha,
                                                                       objFinanz,
                                                                       objdbBuha,
                                                                       objdbPIFb,
                                                                       objFiBebu,
                                                                       objKrBuha,
                                                                       dsKreditoren.Tables("tblKrediSubsFromUser"),
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
                                intReturnValue = MainKreditor.FcPGVKTreatmentYC(objfiBuha,
                                                                       objFinanz,
                                                                       objdbBuha,
                                                                       objdbPIFb,
                                                                       objFiBebu,
                                                                       objKrBuha,
                                                                       dsKreditoren.Tables("tblKrediSubsFromUser"),
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
                        strKRGReferTo = Main.FcReadFromSettingsII("Buchh_TableKRGReferTo", BgWImportKrediArgsInProc.intMandant)
                        'If objdbConn.State = ConnectionState.Open Then
                        '    objdbConn.Close()
                        'End If
                        'Status in File RG-Tabelle schreiben
                        Debug.Print("Booking before Writing to RG-Table " + intKredBelegsNummer.ToString)
                        intReturnValue = MainKreditor.FcWriteToKrediRGTable(BgWImportKrediArgsInProc.intMandant,
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
            objKrBuha = Nothing
            objfiBuha = Nothing
            objdbPIFb = Nothing
            objdbBuha = Nothing
            objFiBebu = Nothing
            objFinanz = Nothing

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
End Class