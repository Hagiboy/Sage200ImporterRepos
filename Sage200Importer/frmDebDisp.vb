Option Strict Off
Option Explicit On

Imports MySql.Data.MySqlClient
Imports System.Data.OracleClient
Imports System.Data.SqlClient
Imports CLClassSage200.WFSage200Import
Imports System.IO

Public Class frmDebDisp

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

    Friend Function FcDebiDisplay(intMandant As Int16,
                                  strPeriode As String) As Int16

        Dim intFcReturns As Int16


        Me.Cursor = Cursors.WaitCursor

        'intMode = 0

        Me.butImport.Enabled = False

        'DGV Debitoren
        'dgvBookings.DataSource = Nothing
        'dgvBookingSub.DataSource = Nothing

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
                                  dsDebitoren.Tables("tblDebitorenInfo"),
                                  strPeriode,
                                  strYear,
                                  intTeqNbr,
                                  intTeqNbrLY,
                                  intTeqNbrPLY,
                                  datPeriodFrom,
                                  datPeriodTo,
                                  strPeriodStatus)

        Dim clImp As New ClassImport
        clImp.FcDebitFill(intMandant)
        clImp = Nothing

        'Tabellentyp darstellen
        Me.lblDB.Text = Main.FcReadFromSettingsII("Buchh_RGTableType", intMandant)

        'Grid neu aufbauen
        MySQLdaDebitoren.Fill(dsDebitoren, "tblDebiHeadsFromUser")
        MySQLdaDebitorenSub.Fill(dsDebitoren, "tblDebiSubsFromUser")

        dgvBookings.DataSource = dsDebitoren.Tables("tblDebiHeadsFromUser")
        dgvBookingSub.DataSource = dsDebitoren.Tables("tblDebiSubsFromUser")

        intFcReturns = FcInitdgvInfo(dgvInfo)
        intFcReturns = FcInitdgvBookings(dgvBookings)
        intFcReturns = FcInitdgvDebiSub(dgvBookingSub)

        'Application.DoEvents()

        Dim clCheck As New ClassCheck
        clCheck.FcClCheckDebit(intMandant,
                                   dsDebitoren,
                                   Finanz,
                                   FBhg,
                                   DbBhg,
                                   PIFin,
                                   BeBu,
                                   dsDebitoren.Tables("tblDebitorenInfo"),
                                   frmImportMain.cmbBuha.Text,
                                   intTeqNbr,
                                   intTeqNbrLY,
                                   intTeqNbrPLY,
                                   strYear,
                                   strPeriode,
                                   datPeriodFrom,
                                   datPeriodTo,
                                   strPeriodStatus,
                                   frmImportMain.chkValutaCorrect.Checked,
                                   frmImportMain.dtpValutaCorrect.Value)
        clCheck = Nothing


        'Anzahl schreiben
        txtNumber.Text = Me.dsDebitoren.Tables("tblDebiHeadsFromUser").Rows.Count.ToString
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

        Dim intReturnValue As Int32
        Dim intDebBelegsNummer As Int32

        Dim intDebitorNbr As Int32
        Dim strBuchType As String
        Dim strBelegDatum As String
        Dim strValutaDatum As String
        Dim strVerfallDatum As String
        Dim strReferenz As String
        Dim intKondition As Int32
        Dim strSachBID As String = String.Empty
        Dim strVerkID As String = String.Empty
        Dim strMahnerlaubnis As String
        Dim sngAktuelleMahnstufe As Single
        Dim dblBetrag As Double
        Dim dblKurs As Double
        Dim strExtBelegNbr As String = String.Empty
        Dim strSkonto As String = String.Empty
        Dim strCurrency As String
        Dim strDebiText As String

        Dim intGegenKonto As Int32
        Dim strFibuText As String
        Dim dblNettoBetrag As Double
        Dim dblBebuBetrag As Double
        Dim strBeBuEintrag As String = String.Empty
        Dim strSteuerFeld As String

        Dim intSollKonto As Int32
        Dim intHabenKonto As Int32
        Dim dblSollBetrag As Double
        Dim dblHabenBetrag As Double
        Dim strSteuerFeldSoll As String = String.Empty
        Dim strSteuerFeldHaben As String = String.Empty
        Dim strBeBuEintragSoll As String = String.Empty
        Dim strBeBuEintragHaben As String = String.Empty
        Dim strDebiTextSoll As String = String.Empty
        Dim strDebiTextHaben As String = String.Empty
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
        Dim booErfOPExt As Boolean

        Dim intLaufNbr As Int32
        Dim strBeleg As String
        Dim strBelegArr() As String
        Dim dblSplitPayed As Double
        Dim strErrMessage As String


        Try


            Me.Cursor = Cursors.WaitCursor
            'Butteon desaktivieren
            Me.butImport.Enabled = False

            'Start in Sync schreiben
            intReturnValue = WFDBClass.FcWriteStartToSync(objdbConn,
                                                          frmImportMain.cmbBuha.SelectedValue,
                                                          1,
                                                          dsDebitoren.Tables("tblDebiHeadsFromUser").Rows.Count)

            'Setting soll erfasste OP als externe Beleg-Nr. genommen werden und lngDebIdentNbr als Beleg-Nr.
            objdbConn.Open()
            booErfOPExt = Convert.ToBoolean(Convert.ToInt16(Main.FcReadFromSettings(objdbConn, "Buchh_ErfOPExt", frmImportMain.cmbBuha.SelectedValue)))
            objdbConn.Close()

            'Kopfbuchung
            For Each row In Me.dsDebitoren.Tables("tblDebiHeadsFromUser").Rows

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
                            'Falls booGSToInv (Gutschrift zu Rechnung) dann OP-Nummer vorgeben, sonst hochzählen lassen
                            If row("booCrToInv") Then
                                'Beleg-Nummerierung desaktivieren
                                DbBhg.IncrBelNbr = "N"
                                'Eingelesene OP-Nummer (=Verknüpfte OP-Nr.) = interne Beleg-Nummer
                                intDebBelegsNummer = Main.FcCleanRGNrStrict(row("strOPNr"))
                                strExtBelegNbr = row("strDebRGNbr")
                            Else
                                'Zuerst Beleg-Nummerieungung aktivieren
                                DbBhg.IncrBelNbr = "J"
                                'Belegsnummer abholen
                                intDebBelegsNummer = DbBhg.GetNextBelNbr("G")
                                'Prüfen ob wirklich frei und falls nicht hochzählen
                                intReturnValue = MainDebitor.FcCheckDebiExistance(objdbMSSQLConn,
                                                                                  objdbSQLcommand,
                                                                                  intDebBelegsNummer,
                                                                                  "G",
                                                                                  intTeqNbr,
                                                                                  intTeqNbrLY,
                                                                                  intTeqNbrPLY)


                                'intReturnValue = 10
                                'Do Until intReturnValue = 0

                                '    intReturnValue = DbBhg.doesBelegExist(row("lngDebNbr").ToString,
                                '                                      row("strDebCur"),
                                '                                      intDebBelegsNummer.ToString,
                                '                                      "NOT_SET",
                                '                                      "G",
                                '                                      "NOT_SET")
                                '    If intReturnValue <> 0 Then
                                '        intDebBelegsNummer += 1
                                '    End If
                                'Loop
                                strExtBelegNbr = row("strOPNr")
                            End If

                            'Beträge drehen
                            row("dblDebBrutto") = row("dblDebBrutto") * -1
                            row("dblDebMwSt") = row("dblDebMwSt") * -1
                            row("dblDebNetto") = row("dblDebNetto") * -1

                            strBuchType = "G"

                        Else

                            If String.IsNullOrEmpty(row("strOPNr")) Then
                                'strExtBelegNbr = row("strOPNr")

                                'Zuerst Beleg-Nummerieungung aktivieren
                                DbBhg.IncrBelNbr = "J"
                                'Belegsnummer abholen
                                intDebBelegsNummer = DbBhg.GetNextBelNbr("R")
                                intReturnValue = MainDebitor.FcCheckDebiExistance(objdbMSSQLConn,
                                                                                  objdbSQLcommand,
                                                                                  intDebBelegsNummer,
                                                                                  "R",
                                                                                  intTeqNbr,
                                                                                  intTeqNbrLY,
                                                                                  intTeqNbrPLY)
                            Else
                                If Strings.Len(Main.FcCleanRGNrStrict(row("strOPNr"))) > 9 Then
                                    'Zahl zu gross
                                    DbBhg.IncrBelNbr = "J"
                                    'Belegsnummer abholen
                                    intDebBelegsNummer = DbBhg.GetNextBelNbr("R")
                                    intReturnValue = MainDebitor.FcCheckDebiExistance(objdbMSSQLConn,
                                                                                      objdbSQLcommand,
                                                                                      intDebBelegsNummer,
                                                                                      "R",
                                                                                      intTeqNbr,
                                                                                      intTeqNbrLY,
                                                                                      intTeqNbrPLY)
                                    strExtBelegNbr = row("strOPNr")
                                Else
                                    'Beleg-Nummerierung abschalten
                                    DbBhg.IncrBelNbr = "N"
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
                        strDebiLine = DbBhg.ReadDebitor3(row("lngDebNbr") * -1, "")
                        strDebitor = Split(strDebiLine, "{>}")
                        strSachBID = strDebitor(30)
                        'strExtBelegNbr = row("strDebRGNbr")
                        intDebitorNbr = row("lngDebNbr")
                        strValutaDatum = Format(row("datDebValDatum"), "yyyyMMdd").ToString
                        strBelegDatum = Format(row("datDebRGDatum"), "yyyyMMdd").ToString
                        If IsDBNull(row("datDebDue")) Then
                            strVerfallDatum = String.Empty
                        Else
                            strVerfallDatum = Format(row("datDebDue"), "yyyyMMdd").ToString
                        End If
                        strReferenz = row("strDebReferenz")
                        strMahnerlaubnis = String.Empty 'Format(row("datDebRGDatum"), "yyyyMMdd").ToString
                        dblBetrag = row("dblDebBrutto")
                        'Bei SplittBill 2ter Rechnung Text anfügen
                        If row("booLinked") Then
                            strDebiText = row("strDebText") + ", FRG "
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
                        'Zahl-Kondition
                        intKondition = IIf(IsDBNull(row("intZKond")), 1, row("intZKond"))

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

                        selDebiSub = dsDebitoren.Tables("tblDebiSubsFromUser").Select("strRGNr='" + row("strDebRGNbr") + "'")
                        strRGNbr = row("strDebRGNbr")

                        For Each SubRow As DataRow In selDebiSub

                            'Bei zweiter Splitt-Bill Rechung hier eingreifen
                            'Gegenkonto auf 1092, MwStKey auf 'null' setzen, KST = 0
                            'If row("booLinked") Then
                            '    If row("booLinkedPayed") Then
                            '        intGegenKonto = 2331
                            '    Else
                            '        intGegenKonto = 1092
                            '    End If
                            '    SubRow("dblNetto") = SubRow("dblBrutto")
                            '    SubRow("strMwStKey") = "null"
                            '    SubRow("lngKST") = 0
                            'Else
                            intGegenKonto = SubRow("lngKto")
                            'End If
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
                                strBeBuEintrag = Nothing
                            End If
                            If Not IsDBNull(SubRow("strMwStKey")) And
                                    SubRow("strMwStKey") <> "null" And
                                    SubRow("lngKto") <> 6906 Then 'And SubRow("strMwStKey") <> "25" Then
                                If strBuchType = "R" Then
                                    strSteuerFeld = Main.FcGetSteuerFeld(FBhg,
                                                                         SubRow("lngKto"),
                                                                         SubRow("strDebSubText"),
                                                                         SubRow("dblBrutto") * -1,
                                                                         SubRow("strMwStKey"),
                                                                         SubRow("dblMwSt") * -1)     '"25{<}DEBI D Bhg Export MwSt{<}0{>}"
                                Else
                                    strSteuerFeld = Main.FcGetSteuerFeld(FBhg,
                                                                         SubRow("lngKto"),
                                                                         SubRow("strDebSubText"),
                                                                         SubRow("dblBrutto"),
                                                                         SubRow("strMwStKey"),
                                                                         SubRow("dblMwSt"))     '"25{<}DEBI D Bhg Export MwSt{<}0{>}"
                                End If
                            Else
                                strSteuerFeld = "STEUERFREI"
                            End If
                            'strSteuerInfo = Split(FBhg.GetKontoInfo(intGegenKonto.ToString), "{>}")
                            'Debug.Print("Konto-Info: " + strSteuerInfo(26))

                            Try

                                booBooingok = True
                                Call DbBhg.SetVerteilung(intGegenKonto.ToString,
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

                            strSteuerFeld = Nothing
                            strBeBuEintrag = Nothing

                            'Status Sub schreiben
                            'Application.DoEvents()

                        Next

                        Try

                            booBooingok = True
                            Call DbBhg.WriteBuchung()

                            'Bei SplittBill 2ter Rechnung TZahlung auf LinkedRG machen
                            'Prinzip: Beleg einlesen anhand und Betrag ausrechnen => Summe Beleg - diesen Beleg
                            If row("booLinked") And Mid(row("strDebStatusBitLog"), 13, 1) = "0" Then 'Nur wenn Beleg in gleicher Buha
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
                                        'Ausbuchen?, wohin mit dem Betrag?
                                    Else

                                        'Welcher Betrag wurde schon bezahlt?
                                        dblSplitPayed = dblBetrag

                                        'Teilzahlung buchen
                                        Call DbBhg.SetZahlung(1944,
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
                                        'Application.DoEvents()

                                        Call DbBhg.WriteTeilzahlung4(intLaufNbr.ToString,
                                                                 row("lngDebIdentNbr").ToString + ", TZ " + row("strDebRGNbr").ToString,
                                                                 "NOT_SET",
                                                                 ,
                                                                 "NOT_SET",
                                                                 "NOT_SET",
                                                                 "Default",
                                                                 "Default")
                                        'Application.DoEvents()

                                    End If

                                End If

                            End If

                        Catch ex As Exception
                            'MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Belegerstellung " + intDebBelegsNummer.ToString + ", RG " + strRGNbr)
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
                                                                     String.Concat(Microsoft.VisualBasic.Left(frmImportMain.cmbPerioden.SelectedItem, 4) - 1, "0101"),
                                                                     String.Concat(Microsoft.VisualBasic.Left(frmImportMain.cmbPerioden.SelectedItem, 4), "1231"))
                                If intReturnValue <> 0 Then
                                    intDebBelegsNummer += 1
                                End If
                            Loop
                            'Debug.Print("Belegnummer taken:  " + intDebBelegsNummer.ToString)
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
                                'Application.DoEvents()

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

                                'Application.DoEvents()

                            Catch ex As Exception
                                MessageBox.Show(ex.Message, "Problem " + (Err.Number And 65535).ToString + " Belegerstellung " + intDebBelegsNummer.ToString + ", RG " + strRGNbr, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
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

                                    'Application.DoEvents()

                                End If

                            Next

                            'Sammelbeleg schreiben
                            Try

                                booBooingok = True
                                Call FBhg.WriteSammelBhgT()

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

                                intReturnValue = MainDebitor.FcPGVDTreatment(FBhg,
                                                                   Finanz,
                                                                   DbBhg,
                                                                   PIFin,
                                                                   BeBu,
                                                                   KrBhg,
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
                                                                   frmImportMain.cmbPerioden.SelectedItem,
                                                                   objdbConn,
                                                                   objdbMSSQLConn,
                                                                   objdbSQLcommand,
                                                                   frmImportMain.cmbBuha.SelectedValue,
                                                                   dsDebitoren.Tables("tblDebitorenInfo"),
                                                                   strYear,
                                                                   intTeqNbr,
                                                                   intTeqNbrLY,
                                                                   intTeqNbrPLY,
                                                                   IIf(IsDBNull(row("strPGVType")), "", row("strPGVType")),
                                                                   datPeriodFrom,
                                                                   datPeriodTo,
                                                                   strPeriodStatus)


                            Else
                                intReturnValue = MainDebitor.FcPGVDTreatmentYC(FBhg,
                                                                   Finanz,
                                                                   DbBhg,
                                                                   PIFin,
                                                                   BeBu,
                                                                   KrBhg,
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
                                                                   frmImportMain.cmbPerioden.SelectedItem,
                                                                   objdbConn,
                                                                   objdbMSSQLConn,
                                                                   objdbSQLcommand,
                                                                   frmImportMain.cmbBuha.SelectedValue,
                                                                   dsDebitoren.Tables("tblDebitorenInfo"),
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
                        row("strDebBookStatus") = row("strDebStatusBitLog")
                        row("booBooked") = True
                        row("datBooked") = Now()
                        row("lngBelegNr") = intDebBelegsNummer
                        'Application.DoEvents()

                        'Status in File RG-Tabelle schreiben
                        intReturnValue = MainDebitor.FcWriteToRGTable(frmImportMain.cmbBuha.SelectedValue,
                                                                      row("strDebRGNbr"),
                                                                      row("datBooked"),
                                                                      row("lngBelegNr"),
                                                                      objdbAccessConn,
                                                                      objOracleConn,
                                                                      objdbConn,
                                                                      row("booDatChanged"),
                                                                      row("datDebRGDatum"),
                                                                      row("datDebValDatum"))
                        If intReturnValue <> 0 Then
                            'Throw an exception
                        End If

                        'Evtl. Query nach Buchung ausführen
                        Call MainDebitor.FcExecuteAfterDebit(frmImportMain.cmbBuha.SelectedValue, objdbConn)
                    End If

                End If

            Next
            'Status in t_sage_syncstatus schreiben
            'intReturnValue = MainDebitor.FcWriteEndToSync(objdbConn,
            '                                              cmbBuha.SelectedValue,
            '                                              1,
            '                                              Date.Now,
            '                                              0,
            '                                              IIf(booBooingok, "ok", "Probleme"))

            intReturnValue = WFDBClass.FcWriteEndToSync(objdbConn,
                                                        frmImportMain.cmbBuha.SelectedValue,
                                                        1,
                                                        0,
                                                        IIf(booBooingok, "ok", "Probleme"))




        Catch ex As Exception
            MessageBox.Show(ex.Message, "Generelles Problem " + (Err.Number And 65535).ToString + " Belegerstellung " + intDebBelegsNummer.ToString + ", RG " + strRGNbr, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        Finally

            'Application.DoEvents()
            'Grid neu aufbauen, Daten von Mandant einlesen
            'Call butDebitoren.PerformClick()

            Me.Cursor = Cursors.Default
            'Me.butImport.Enabled = False
            Me.Close()
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

End Class