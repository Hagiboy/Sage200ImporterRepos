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

        'Tabellentyp darstellen
        Me.lblDB.Text = Main.FcReadFromSettingsII("Buchh_KRGTableType", intMandant)

        'Grid neu aufbauen
        MySQLdaKreditoren.Fill(dsKreditoren, "tblKrediHeadsFromUser")
        MySQLdaKreditorenSub.Fill(dsKreditoren, "tblKrediSubsFromUser")

        dgvBookings.DataSource = dsKreditoren.Tables("tblKrediHeadsFromUser")
        dgvBookingSub.DataSource = dsKreditoren.Tables("tblKrediSubsFromUser")

        intFcReturns = FcInitdgvInfo(dgvInfo)
        intFcReturns = FcInitdgvKreditoren(dgvBookings)
        intFcReturns = FcInitdgvKrediSub(dgvBookingSub)

        'Application.DoEvents()

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

    Private Sub dgvBookings_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvBookings.CellContentClick

        Try

            If e.RowIndex >= 0 Then

                dgvBookingSub.DataSource = dsKreditoren.Tables("tblKrediHeadsFromUser").Select("lngKredID=" + dgvBookings.Rows(e.RowIndex).Cells("lngKredID").Value.ToString).CopyToDataTable

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

        Dim intReturnValue As Int32
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
        Dim strSachBID As String = String.Empty
        Dim strVerkID As String = String.Empty
        Dim strMahnerlaubnis As String
        Dim sngAktuelleMahnstufe As Single
        Dim dblBetrag As Double
        Dim dblKurs As Double
        Dim strExtBelegNbr As String = String.Empty
        Dim strSkonto As String = String.Empty
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
        Dim strSteuerFeldSoll As String = String.Empty
        Dim strSteuerFeldHaben As String = String.Empty
        Dim strBeBuEintragSoll As String = String.Empty
        Dim strBeBuEintragHaben As String = String.Empty
        Dim strKrediTextSoll As String = String.Empty
        Dim strKrediTextHaben As String = String.Empty
        Dim dblKursSoll As Double = 0
        Dim dblKursHaben As Double = 0

        Dim selKrediSub() As DataRow
        Dim strSteuerInfo() As String
        Dim strDebiLine As String
        Dim strDebitor() As String

        Dim booBookingok As Boolean

        'Sammelbeleg
        Dim intCommonKonto As Int32
        Dim strKRGReferTo As String

        'Dim intTeqNbr As Int32

        Try


            Me.Cursor = Cursors.WaitCursor
            'Application.DoEvents()
            'Button disablen damit er nicht noch einmal geklickt werden kann.
            Me.butImport.Enabled = False

            'Start in Sync schreiben
            intReturnValue = WFDBClass.FcWriteStartToSync(objdbConn,
                                                          frmImportMain.cmbBuha.SelectedValue,
                                                          2,
                                                          dsKreditoren.Tables("tblKrediHeadsFromUser").Rows.Count)

            'intTeqNbr = Conversion.Val(Strings.Right(objdtInfo.Rows(1).Item(1), 3))

            'Kopfbuchung
            For Each row As DataRow In dsKreditoren.Tables("tblKrediHeadsFromUser").Rows

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
                        KrBhg.CheckDoubleIntBelNbr = "N"

                        'Eindeutigkeit externer Beleg-Nummer setzen
                        KrBhg.CheckDoubleExtBelNbr = "J"

                        'If IsDBNull(row("strOPNr")) Or row("StrOPNr") = "" Then
                        'strExtBelegNbr = row("strOPNr")

                        'Zuerst Beleg-Nummerieungung aktivieren
                        'KrBhg.IncrBelNbr = "N"
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
                            KrBhg.IncrBelNbr = "J"
                            intKredBelegsNummer = KrBhg.GetNextBelNbr("G")
                            KrBhg.IncrBelNbr = "N"

                            intReturnValue = MainKreditor.FcCheckKrediExistance(objdbMSSQLConn,
                                                                                objdbSQLcommand,
                                                                                intKredBelegsNummer,
                                                                                "G",
                                                                                intTeqNbr,
                                                                                intTeqNbrLY,
                                                                                intTeqNbrPLY,
                                                                                KrBhg)

                        Else
                            strBuchType = "R"
                            'strZahlSperren = "N"
                            'Belegsnummer abholen
                            KrBhg.IncrBelNbr = "J"
                            intKredBelegsNummer = KrBhg.GetNextBelNbr("R")
                            'Muss auf Nicht hochzählen gesetzt werden da Sage 200 nicht merkt, dass Beleg-Nr. schon vergeben worden sind. => In den Einstellungen muss von Zeit zu Zeit der Zähler geändert werden
                            KrBhg.IncrBelNbr = "N"

                            intReturnValue = MainKreditor.FcCheckKrediExistance(objdbMSSQLConn,
                                                                                objdbSQLcommand,
                                                                                intKredBelegsNummer,
                                                                                "R",
                                                                                intTeqNbr,
                                                                                intTeqNbrLY,
                                                                                intTeqNbrPLY,
                                                                                KrBhg)

                        End If

                        strValutaDatum = Format(row("datKredValDatum"), "yyyyMMdd").ToString
                        strBelegDatum = Format(row("datKredRGDatum"), "yyyyMMdd").ToString
                        strVerfallDatum = String.Empty
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
                            strTeilnehmer = String.Empty
                        End If
                        'End If
                        strMahnerlaubnis = String.Empty 'Format(row("datDebRGDatum"), "yyyyMMdd").ToString
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

                        selKrediSub = dsKreditoren.Tables("tblKrediSubsFromUser").Select("lngKredID=" + row("lngKredID").ToString)

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


                            strSteuerFeld = String.Empty
                            strBeBuEintrag = String.Empty

                            'Status Sub schreiben
                            'Application.DoEvents()

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
                        intKredBelegsNummer = FBhg.GetNextBelNbr()

                        'Prüfen, ob wirklich frei
                        intReturnValue = 10
                        Do Until intReturnValue = 0
                            intReturnValue = FBhg.doesBelegExist(intKredBelegsNummer,
                                                                 "NOT_SET",
                                                                 "NOT_SET",
                                                                 Strings.Left(frmImportMain.cmbPerioden.SelectedItem, 4) + "0101",
                                                                 Strings.Left(frmImportMain.cmbPerioden.SelectedItem, 4) + "1231")
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
                            dblKurs = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg)
                        Else
                            dblKurs = 1.0#
                        End If

                        selKrediSub = dsKreditoren.Tables("tblKrediSubsFromUser").Select("lngKredID=" + row("lngKredID").ToString)

                        If selKrediSub.Length = 2 Then

                            For Each SubRow As DataRow In selKrediSub

                                If SubRow("intSollHaben") = 0 Then 'Soll

                                    intSollKonto = SubRow("lngKto")
                                    dblKursSoll = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg, intSollKonto)
                                    'strSteuerInfo = Split(FBhg.GetKontoInfo(intSollKonto.ToString), "{>}")
                                    'Debug.Print("Konto-Info Soll: " + strSteuerInfo(26))
                                    dblSollBetrag = SubRow("dblNetto")
                                    strKrediTextSoll = SubRow("strKredSubText")
                                    If SubRow("dblMwSt") > 0 Then
                                        strSteuerFeldSoll = Main.FcGetSteuerFeld(FBhg,
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
                                    dblKursHaben = Main.FcGetKurs(strCurrency, strValutaDatum, FBhg, intHabenKonto)
                                    'strSteuerInfo = Split(FBhg.GetKontoInfo(intSollKonto.ToString), "{>}")
                                    'Debug.Print("Konto-Info Haben: " + strSteuerInfo(26))
                                    dblHabenBetrag = SubRow("dblNetto") * -1
                                    'dblHabenBetrag = dblSollBetrag
                                    strKrediTextHaben = SubRow("strKredSubText")
                                    If SubRow("dblMwSt") > 0 Then
                                        strSteuerFeldHaben = Main.FcGetSteuerFeld(FBhg,
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
                        If row("booPGV") Then
                            'Bei PGV Buchungen
                            If IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")) = "" Or
                               (IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")) = "RV" And row("intPGVMthsAY") + row("intPGVMthsNY") > 1) Then

                                intReturnValue = MainKreditor.FcPGVKTreatment(FBhg,
                                                                       Finanz,
                                                                       DbBhg,
                                                                       PIFin,
                                                                       BeBu,
                                                                       KrBhg,
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
                                                                       frmImportMain.cmbPerioden.SelectedItem,
                                                                       objdbConn,
                                                                       objdbMSSQLConn,
                                                                       objdbSQLcommand,
                                                                       frmImportMain.cmbBuha.SelectedValue,
                                                                       dsKreditoren.Tables("tblKreditorenInfo"),
                                                                       strYear,
                                                                       intTeqNbr,
                                                                       intTeqNbrLY,
                                                                       intTeqNbrPLY,
                                                                       IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")),
                                                                       datPeriodFrom,
                                                                       datPeriodTo,
                                                                       strPeriodStatus)

                            Else

                                'TP
                                intReturnValue = MainKreditor.FcPGVKTreatmentYC(FBhg,
                                                                       Finanz,
                                                                       DbBhg,
                                                                       PIFin,
                                                                       BeBu,
                                                                       KrBhg,
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
                                                                       frmImportMain.cmbPerioden.SelectedItem,
                                                                       objdbConn,
                                                                       objdbMSSQLConn,
                                                                       objdbSQLcommand,
                                                                       frmImportMain.cmbBuha.SelectedValue,
                                                                       dsKreditoren.Tables("tblKreditorenInfo"),
                                                                       strYear,
                                                                       intTeqNbr,
                                                                       intTeqNbrLY,
                                                                       intTeqNbrPLY,
                                                                       IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")),
                                                                       datPeriodFrom,
                                                                       datPeriodTo,
                                                                       strPeriodStatus)


                            End If


                        End If

                        'Status Head schreiben
                        row("strKredBookStatus") = row("strKredStatusBitLog")
                        row("booBooked") = True
                        row("datBooked") = Now()
                        row("lngBelegNr") = intKredBelegsNummer
                        'strKRGReferTo = "lngKredID"
                        'strKRGReferTo = "strKredRGNbr"
                        If objdbConn.State = ConnectionState.Closed Then
                            objdbConn.Open()
                        End If
                        strKRGReferTo = Main.FcReadFromSettings(objdbConn, "Buchh_TableKRGReferTo", frmImportMain.cmbBuha.SelectedValue)
                        If objdbConn.State = ConnectionState.Open Then
                            objdbConn.Close()
                        End If
                        'Status in File RG-Tabelle schreiben
                        intReturnValue = MainKreditor.FcWriteToKrediRGTable(frmImportMain.cmbBuha.SelectedValue,
                                                                        row(strKRGReferTo),
                                                                        row("datBooked"),
                                                                        row("lngBelegNr"),
                                                                        objdbAccessConn,
                                                                        objOracleConn,
                                                                        objdbConn)
                        If intReturnValue <> 0 Then
                            'Throw an exception
                            MessageBox.Show("Achtung, Beleg-Nummer: " + row("lngBelegNr").ToString + " konnte nicht In die RG-Tabelle geschrieben werden auf RG-ID: " + row("lngKredID").ToString + ".", "RG-Table Update nicht möglich", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        End If

                    End If

                End If

            Next

            'In sync-Tabelle schreiben
            intReturnValue = WFDBClass.FcWriteEndToSync(objdbConn,
                                                        frmImportMain.cmbBuha.SelectedValue,
                                                        2,
                                                        0,
                                                        IIf(booBookingok, "ok", "Probleme"))


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " bei Buchung-Beleg " + strExtKredBelegsNummer)

        Finally
            'Neu aufbauen
            'butKreditoren_Click(butDebitoren, EventArgs.Empty)

            Me.Cursor = Cursors.Default
            'Me.butImportK.Enabled = True
            Me.Close()

        End Try


    End Sub
End Class