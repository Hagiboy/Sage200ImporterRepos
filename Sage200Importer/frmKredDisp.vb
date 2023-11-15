Option Strict Off
Option Explicit On

Imports MySql.Data.MySqlClient
Imports System.Data.OracleClient
Imports System.Data.SqlClient
Imports CLClassSage200.WFSage200Import
Imports System.IO
Imports Google.Protobuf.WellKnownTypes
Imports Org.BouncyCastle.Crypto.Prng

Public Class frmKredDisp

    Dim Finanz As SBSXASLib.AXFinanz
    Dim FBhg As SBSXASLib.AXiFBhg
    Dim DbBhg As SBSXASLib.AXiDbBhg
    Dim KrBhg As SBSXASLib.AXiKrBhg
    Dim BsExt As SBSXASLib.AXiBSExt
    Dim Adr As SBSXASLib.AXiAdr
    Dim BeBu As SBSXASLib.AXiBeBu
    Dim PIFin As SBSXASLib.AXiPlFin

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

    Dim objdbConn As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString"))
    Dim objdbMSSQLConn As New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("SQLConnectionString"))
    Dim objdbSQLcommand As New SqlCommand
    Dim objdbAccessConn As New OleDb.OleDbConnection
    Dim objOracleConn As New OracleConnection("Data Source=(DESCRIPTION=" _
                    + "(ADDRESS=(PROTOCOL=TCP)(HOST=10.0.0.29)(PORT=1521))" _
                    + "(CONNECT_DATA=(SERVICE_NAME=CISNEW)));" _
                    + "User Id=cis;Password=sugus;")



    Public Sub InitDB()

        Dim strIdentityName As String

        Try

            strIdentityName = System.Security.Principal.WindowsIdentity.GetCurrent().Name
            strIdentityName = Strings.Replace(strIdentityName, "\", "/")

            'Dim daDebitorenHead As New MySqlDataAdapter()
            mysqlconn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString")
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


    Friend Function FcKrediDisplay(intMandant As Int16,
                                  strPeriode As String) As Int16

        Dim intFcReturns As Int16


        Me.Cursor = Cursors.WaitCursor

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

        Call Main.FcLoginSage2(objdbConn,
                                  objdbMSSQLConn,
                                  objdbSQLcommand,
                                  Finanz,
                                  FBhg,
                                  DbBhg,
                                  PIFin,
                                  BeBu,
                                  KrBhg,
                                  intMandant,
                                  dsKreditoren.Tables("tblKreditorenInfo"),
                                  strPeriode,
                                  strYear,
                                  intTeqNbr,
                                  intTeqNbrLY,
                                  intTeqNbrPLY,
                                  datPeriodFrom,
                                  datPeriodTo,
                                  strPeriodStatus)

        Dim clImp As New ClassImport
        clImp.FcKreditFill(intMandant)
        clImp = Nothing

        'Grid neu aufbauen
        MySQLdaKreditoren.Fill(dsKreditoren, "tblKrediHeadsFromUser")
        MySQLdaKreditorenSub.Fill(dsKreditoren, "tblKrediSubsFromUser")

        dgvBookings.DataSource = dsKreditoren.Tables("tblKrediHeadsFromUser")
        dgvBookingSub.DataSource = dsKreditoren.Tables("tblKrediSubsFromUser")

        intFcReturns = FcInitdgvInfo(dgvInfo)
        intFcReturns = FcInitdgvKreditoren(dgvBookings)
        intFcReturns = FcInitdgvKrediSub(dgvBookingSub)

        Application.DoEvents()

        Dim clCheck As New ClassCheck
        clCheck.FcCheckKredit(intMandant,
                              dsKreditoren,
                              Finanz,
                              FBhg,
                              KrBhg,
                              PIFin,
                              dsKreditoren.Tables("tblKreditorenInfo"),
                              frmImportMain.cmbBuha.Text,
                              strYear,
                              strPeriode,
                              datPeriodFrom,
                              datPeriodTo,
                              strPeriodStatus,
                              frmImportMain.chkValutaCorrect.Checked,
                              frmImportMain.dtpValutaCorrect.Value)

        clCheck = Nothing

        'Anzahl schreiben
        txtNumber.Text = Me.dsKreditoren.Tables("tblKrediHeadsFromUser").Rows.Count.ToString
        Me.Cursor = Cursors.Default

        Me.butImport.Enabled = True



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
            dgvBookingSub.Columns("intSollHaben").Width = 30
            dgvBookingSub.Columns("intSollHaben").DisplayIndex = 1
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
End Class