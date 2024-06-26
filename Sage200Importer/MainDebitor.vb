﻿Option Strict Off
Option Explicit On

Imports MySql.Data.MySqlClient
Imports System.Data.OracleClient
Imports System.Data.SqlClient
Imports System.Net
Imports System.IO
Imports System.Xml
Imports System.Data.OleDb


Public Class MainDebitor

    Public Shared Function FcFillDebit(ByVal intAccounting As Int16,
                                       ByVal objdbaccessconn As OleDb.OleDbConnection,
                                       ByVal objdbmysqlcon As MySqlConnection) As Integer

        Dim strSQL As String
        Dim strSQLSub As String
        Dim strRGTableType As String
        Dim objRGMySQLConn As New MySqlConnection
        Dim objlocMySQLcmd As New MySqlCommand
        Dim objdbAccessConnLoc As New OleDb.OleDbConnection
        Dim objlocOLEdbcmdLoc As New OleDb.OleDbCommand
        Dim strConnection As String
        Dim objdtlocDebiSub As New DataTable
        Dim objmysqlcomdwritesub As New MySqlCommand
        Dim objmysqlconZHDB02 As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))
        Dim strmysqlSaveSub As String
        Dim strIdentityName As String
        Dim objdtLocDebiHead As New DataTable
        Dim objmysqlcomdwritehead As New MySqlCommand
        Dim strmysqlSaveHead As String

        'Dim objDTDebiHead As New DataTable
        'Dim objdrSub As DataRow
        Dim intFcReturns As Int16
        Dim strMDBName As String
        Dim strSQLToParse As String


        Try

            objmysqlcomdwritesub.Connection = objmysqlconZHDB02
            objmysqlcomdwritehead.Connection = objmysqlconZHDB02

            strMDBName = Main.FcReadFromSettingsII("Buchh_RGTableMDB",
                                                 intAccounting)

            'Head Debitoren löschen
            'objdtHead.Clear()
            'objdtHead.Constraints.Clear()

            'objdtSub.Clear()
            'objdtSub.Constraints.Clear()

            strSQL = Main.FcReadFromSettingsII("Buchh_SQLHead",
                                             intAccounting)
            strRGTableType = Main.FcReadFromSettingsII("Buchh_RGTableType",
                                                     intAccounting)

            'objlocMySQLcmd.CommandText = strSQL
            If strRGTableType = "A" Then
                'Access

                'Call Main.FcInitAccessConnecation(objdbaccessconn,
                '                                  strMDBName)

                'objdbAccessConnLoc = objdbaccessconn
                'objdbAccessConnLoc.Open()
                objdbaccessconn.Open()
                objlocOLEdbcmdLoc.CommandText = strSQL
                'objlocOLEdbcmdLoc.Connection = objdbAccessConnLoc
                objlocOLEdbcmdLoc.Connection = objdbaccessconn
                'objdtHead.Load(objlocOLEdbcmdLoc.ExecuteReader)
                objdtLocDebiHead.Load(objlocOLEdbcmdLoc.ExecuteReader)
                'objdbAccessConnLoc.Close()
                objdbaccessconn.Close()
            ElseIf strRGTableType = "M" Then
                strConnection = System.Configuration.ConfigurationManager.AppSettings(strMDBName)
                objRGMySQLConn = objdbmysqlcon.Clone()
                objRGMySQLConn.ConnectionString = strConnection
                'frmImportMain.mysqlcongen.Open()
                'frmImportMain.mysqlcongen.ChangeDatabase("AHZ")
                'objRGMySQLConn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(strMDBName)
                objlocMySQLcmd.Connection = objRGMySQLConn
                objlocMySQLcmd.CommandText = strSQL
                'frmImportMain.mysqlcmdgen.CommandText = strSQL
                objRGMySQLConn.Open()
                'frmImportMain.mysqlcongen.Open()
                'objdtHead.Load(objlocMySQLcmd.ExecuteReader)
                objdtLocDebiHead.Load(objlocMySQLcmd.ExecuteReader)
                objRGMySQLConn.Close()
                'frmImportMain.mysqlcongen.Close()
            End If
            'objlocMySQLcmd.Connection = objdbconn
            'objDTDebiHead.Load(objlocMySQLcmd.ExecuteReader)
            'Durch die Records steppen und Sub-Tabelle füllen
            strSQLToParse = Main.FcReadFromSettingsII("Buchh_SQLDetail",
                                                    intAccounting)
            intFcReturns = FcInitInsCmdDHeads(objmysqlcomdwritehead)
            For Each row As DataRow In objdtLocDebiHead.Rows

                objmysqlcomdwritehead.Connection.Open()
                objmysqlcomdwritehead.Parameters("@IdentityName").Value = strIdentityName
                objmysqlcomdwritehead.Parameters("@ProcessID").Value = Process.GetCurrentProcess().Id
                objmysqlcomdwritehead.Parameters("@intBuchhaltung").Value = intAccounting
                objmysqlcomdwritehead.Parameters("@strDebRGNbr").Value = row("strDebRGNbr")
                objmysqlcomdwritehead.Parameters("@intBuchungsart").Value = row("intBuchungsart")
                objmysqlcomdwritehead.Parameters("@intRGArt").Value = row("intRGArt")
                objmysqlcomdwritehead.Parameters("@strRGArt").Value = row("strRGArt")
                objmysqlcomdwritehead.Parameters("@booLinked").Value = False
                objmysqlcomdwritehead.Parameters("@booLinkedPayed").Value = False
                objmysqlcomdwritehead.Parameters("@strOPNr").Value = row("strOPNr")
                objmysqlcomdwritehead.Parameters("@lngDebNbr").Value = row("lngDebNbr")
                'objmysqlcomdwritehead.Parameters("@strDebBez").Value = ""
                objmysqlcomdwritehead.Parameters("@lngDebKtoNbr").Value = row("lngDebKtoNbr")
                'objmysqlcomdwritehead.Parameters("@strDebKtoBez").Value = ""
                objmysqlcomdwritehead.Parameters("@dblDebNetto").Value = row("dblDebNetto")
                objmysqlcomdwritehead.Parameters("@dblDebMwSt").Value = row("dblDebMwSt")
                objmysqlcomdwritehead.Parameters("@dblDebBrutto").Value = row("dblDebBrutto")
                objmysqlcomdwritehead.Parameters("@lngDebIdentNbr").Value = row("lngDebIdentNbr")
                objmysqlcomdwritehead.Parameters("@strDebText").Value = row("strDebText")
                objmysqlcomdwritehead.Parameters("@strDebreferenz").Value = row("strDebReferenz")
                objmysqlcomdwritehead.Parameters("@datDebRGDatum").Value = row("datDebRGDatum")
                objmysqlcomdwritehead.Parameters("@datDebValDatum").Value = row("datDebValDatum")
                objmysqlcomdwritehead.Parameters("@strDebStatusBitLog").Value = ""
                objmysqlcomdwritehead.Parameters("@strDebStatusText").Value = ""
                objmysqlcomdwritehead.Parameters("@strDebBookStatus").Value = ""
                objmysqlcomdwritehead.ExecuteNonQuery()
                objmysqlcomdwritehead.Connection.Close()
                'If row("intBuchungsart") = 1 Then
                '    objdrSub = objdtSub.NewRow()
                '    objdrSub("strRGNr") = row("strDebRGNbr")
                '    objdrSub("intSollHaben") = 2
                '    objdrSub("lngKto") = row("lngDebKtoNbr")
                '    objdrSub("dblBrutto") = row("dblDebBrutto")
                '    objdrSub("dblNetto") = row("dblDebNetto")
                '    objdrSub("dblMwSt") = row("dblDebMwSt")
                '    objdrSub("strDebSubText") = row("Betrifft").ToString + " " + row("betrifft1").ToString
                '    objdtSub.Rows.Add(objdrSub)
                'End If
                'If row("strDebRGNbr") = "" Then Stop
                strSQLSub = MainDebitor.FcSQLParse(strSQLToParse,
                                                   row("strDebRGNbr"),
                                                   objdtLocDebiHead,
                                                   "D")
                If strRGTableType = "A" Then
                    'objdbAccessConnLoc.Open()
                    objdbaccessconn.Open()
                    objlocOLEdbcmdLoc.CommandText = strSQLSub
                    objdtlocDebiSub.Load(objlocOLEdbcmdLoc.ExecuteReader)
                    'objdbAccessConnLoc.Close()
                    objdbaccessconn.Close()
                ElseIf strRGTableType = "M" Then
                    objlocMySQLcmd.CommandText = strSQLSub
                    'frmImportMain.mysqlcmdgen.CommandText = strSQLSub
                    objRGMySQLConn.Open()
                    'frmImportMain.mysqlcongen.Open()
                    objdtlocDebiSub.Load(objlocMySQLcmd.ExecuteReader)
                    'objdtSub.Load(objlocMySQLcmd.ExecuteReader)
                    objRGMySQLConn.Close()
                    'frmImportMain.mysqlcongen.Close()
                End If

                'Application.DoEvents()

            Next
            For Each drsub As DataRow In objdtlocDebiSub.Rows
                '    objdtSub.ImportRow(drsub)
                strmysqlSaveSub = "INSERT INTO tbldebitorensub (IdentityName, ProcessID, strRGNr, intSollHaben, lngKto, lngKST, dblNetto, dblMwSt, dblBrutto) "
                strmysqlSaveSub += " VALUES('"
                strmysqlSaveSub += strIdentityName + "', " + Process.GetCurrentProcess().Id.ToString + ", '" + drsub("strRGNr") + "', " + drsub("intSollHaben").ToString + ", " + drsub("lngKto").ToString + ", " + drsub("lngKST").ToString + ", " + drsub("dblNetto").ToString + ", " + drsub("dblMwSt").ToString + ", " + drsub("dblBrutto").ToString + ")"
                objmysqlcomdwritesub.CommandText = strmysqlSaveSub
                objmysqlcomdwritesub.Connection.Open()
                objmysqlcomdwritesub.ExecuteNonQuery()
                objmysqlcomdwritesub.Connection.Close()

            Next
            'Tabellen runden
            'intFcReturns = FcRoundInTable(objdtHead, "dblDebNetto", 2)
            'intFcReturns = FcRoundInTable(objdtHead, "dblDebBrutto", 2)
            'intFcReturns = FcRoundInTable(objdtHead, "dblDebMwSt", 2)
            'intFcReturns = FcRoundInTable(objdtSub, "dblNetto", 2)
            'intFcReturns = FcRoundInTable(objdtSub, "dblMwSt", 2)
            'intFcReturns = FcRoundInTable(objdtSub, "dblBrutto", 2)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Debitoren-Daten-Lesen Problem", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        Finally

            If objdbAccessConnLoc.State = ConnectionState.Open Then
                objdbAccessConnLoc.Close()
            End If
            If objRGMySQLConn.State = ConnectionState.Open Then
                objRGMySQLConn.Close()
            End If
            If frmImportMain.mysqlcongen.State = ConnectionState.Open Then
                frmImportMain.mysqlcongen.Close()
            End If

            objRGMySQLConn = Nothing
            objdbAccessConnLoc = Nothing
            objlocMySQLcmd = Nothing
            objlocOLEdbcmdLoc = Nothing

        End Try

    End Function


    Public Shared Function FcWriteDebiHeads(ByVal tblHeads As DataTable,
                                            ByVal intBuha As Int16) As Int16

        Dim objmysqlconZHDB02 As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))
        Dim strIdentityName As String
        Dim objmysqlcomdwritehead As New MySqlCommand
        Dim intFcReturns As Int16

        Try

            strIdentityName = System.Security.Principal.WindowsIdentity.GetCurrent().Name
            strIdentityName = Strings.Replace(strIdentityName, "\", "/")
            objmysqlcomdwritehead.Connection = objmysqlconZHDB02

            intFcReturns = FcInitInsCmdDHeads(objmysqlcomdwritehead)

            For Each row As DataRow In tblHeads.Rows

                objmysqlcomdwritehead.Connection.Open()
                objmysqlcomdwritehead.Parameters("@IdentityName").Value = strIdentityName
                objmysqlcomdwritehead.Parameters("@ProcessID").Value = Process.GetCurrentProcess().Id
                objmysqlcomdwritehead.Parameters("@intBuchhaltung").Value = intBuha
                objmysqlcomdwritehead.Parameters("@strDebRGNbr").Value = row("strDebRGNbr")
                objmysqlcomdwritehead.Parameters("@booLinked").Value = False
                objmysqlcomdwritehead.Parameters("@booLinkedPayed").Value = False
                objmysqlcomdwritehead.Parameters("@lngDebNbr").Value = row("lngDebNbr")
                objmysqlcomdwritehead.Parameters("@strDebBez").Value = ""
                objmysqlcomdwritehead.Parameters("@lngDebKtoNbr").Value = row("lngDebKtoNbr")
                objmysqlcomdwritehead.Parameters("@strDebKtoBez").Value = ""
                objmysqlcomdwritehead.Parameters("@datDebRGDatum").Value = row("datDebRGDatum")
                objmysqlcomdwritehead.Parameters("@datDebValDatum").Value = row("datDebValDatum")
                objmysqlcomdwritehead.Parameters("@strDebStatusBitLog").Value = ""
                objmysqlcomdwritehead.Parameters("@strDebStatusText").Value = ""
                objmysqlcomdwritehead.Parameters("@strDebBookStatus").Value = ""
                objmysqlcomdwritehead.ExecuteNonQuery()
                objmysqlcomdwritehead.Connection.Close()

            Next
            Return 0


        Catch ex As Exception
            MessageBox.Show("Error bei Schreiben DebiHeads to Table")
            Return 1

        End Try


    End Function


    Public Shared Function FcInitInsCmdDHeads(ByRef mysqlinscmd As MySqlCommand) As Int16

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
            inscmdFields += ", booLinked"
            inscmdValues += ", @booLinked"
            inscmdFields += ", booLinkedPayed"
            inscmdValues += ", @booLinkedPayed"
            inscmdFields += ", booGS"
            inscmdValues += ", @booGS"
            inscmdFields += ", strOPNr"
            inscmdValues += ", @strOPNr"
            inscmdFields += ", lngDebNbr"
            inscmdValues += ", @lngDebNbr"
            'inscmdFields += ", strDebBez"
            'inscmdValues += ", @strDebBez"
            inscmdFields += ", lngDebKtoNbr"
            inscmdValues += ", @lngDebKtoNbr"
            'inscmdFields += ", strDebKtoBez"
            'inscmdValues += ", @strDebKtoBez"
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
            inscmdFields += ", strDebStatusBitLog"
            inscmdValues += ", @strDebStatusBitLog"
            inscmdFields += ", strDebStatusText"
            inscmdValues += ", @strDebStatusText"
            inscmdFields += ", strDebBookStatus"
            inscmdValues += ", @strDebBookStatus"

            'strIdentityName = System.Security.Principal.WindowsIdentity.GetCurrent().Name
            'strIdentityName = Strings.Replace(strIdentityName, "\", "/")

            'Dim daDebitorenHead As New MySqlDataAdapter()
            'mysqlconn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString")

            'Ins cmd DebiHead
            mysqlinscmd.CommandText = "INSERT INTO tbldebitorenjhead (" + inscmdFields + ") VALUES (" + inscmdValues + ")"
            mysqlinscmd.Parameters.Add("@IdentityName", MySqlDbType.String).SourceColumn = "IdentityName"
            mysqlinscmd.Parameters.Add("@ProcessID", MySqlDbType.Int16).SourceColumn = "ProcessID"
            mysqlinscmd.Parameters.Add("@intBuchhaltung", MySqlDbType.Int16).SourceColumn = "intBuchhaltung"
            mysqlinscmd.Parameters.Add("@strDebRGNbr", MySqlDbType.String).SourceColumn = "strDebRGNbr"
            mysqlinscmd.Parameters.Add("@intBuchungsart", MySqlDbType.Int16).SourceColumn = "intBuchungsart"
            mysqlinscmd.Parameters.Add("@intRGArt", MySqlDbType.Int16).SourceColumn = "intRGArt"
            mysqlinscmd.Parameters.Add("@strRGArt", MySqlDbType.String).SourceColumn = "strRGArt"
            mysqlinscmd.Parameters.Add("@booLinked", MySqlDbType.Byte).SourceColumn = "booLinked"
            mysqlinscmd.Parameters.Add("@booLinkedPayed", MySqlDbType.Byte).SourceColumn = "booLinkedPayed"
            mysqlinscmd.Parameters.Add("@booGS", MySqlDbType.Byte).SourceColumn = "booGS"
            mysqlinscmd.Parameters.Add("@strOPNr", MySqlDbType.String).SourceColumn = "strOPNr"
            mysqlinscmd.Parameters.Add("@lngDebNbr", MySqlDbType.Int32).SourceColumn = "lngDebNbr"
            'mysqlinscmd.Parameters.Add("@strDebBez", MySqlDbType.String).SourceColumn = "strDebBez"
            mysqlinscmd.Parameters.Add("@lngDebKtoNbr", MySqlDbType.Int32).SourceColumn = "lngDebKtoNbr"
            'mysqlinscmd.Parameters.Add("@strDebKtoBez", MySqlDbType.String).SourceColumn = "strDebKtoBez"
            mysqlinscmd.Parameters.Add("@dblDebNetto", MySqlDbType.Decimal).SourceColumn = "dblDebNetto"
            mysqlinscmd.Parameters.Add("@dblDebMwst", MySqlDbType.Decimal).SourceColumn = "dblDebMwSt"
            mysqlinscmd.Parameters.Add("@dblDebBrutto", MySqlDbType.Decimal).SourceColumn = "dblDebBrutto"
            mysqlinscmd.Parameters.Add("@strDebText", MySqlDbType.String).SourceColumn = "strDebText"
            mysqlinscmd.Parameters.Add("@lngDebIdentNbr", MySqlDbType.Int32).SourceColumn = "lngDebIdentNbr"
            mysqlinscmd.Parameters.Add("@strDebReferenz", MySqlDbType.String).SourceColumn = "strDebReferenz"
            mysqlinscmd.Parameters.Add("@datDebRGDatum", MySqlDbType.Date).SourceColumn = "datDebRGDatum"
            mysqlinscmd.Parameters.Add("@datDebValDatum", MySqlDbType.Date).SourceColumn = "datDebValDatum"
            mysqlinscmd.Parameters.Add("@strDebStatusBitLog", MySqlDbType.String).SourceColumn = "strDebStatusBitLog"
            mysqlinscmd.Parameters.Add("@strDebStatusText", MySqlDbType.String).SourceColumn = "strDebStatusText"
            mysqlinscmd.Parameters.Add("@strDebBookStatus", MySqlDbType.String).SourceColumn = "strDebBookStatus"
            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem HeadCommand Init", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 1

        End Try

    End Function

    Public Shared Function FcReadDebitorKName(ByRef objfiBuha As SBSXASLib.AXiFBhg,
                                              ByVal lngDebKtoNbr As Long) As String

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


        Finally
            'Application.DoEvents()

        End Try

    End Function

    Public Shared Function FcReadDebitorName(ByRef objDbBhg As SBSXASLib.AXiDbBhg,
                                             ByVal intDebiNbr As Int32,
                                             ByVal strCurrency As String) As String

        Dim strDebitorName As String
        Dim strDebitorAr() As String

        Try

            If strCurrency = "" Then

                strDebitorName = objDbBhg.ReadDebitor3(intDebiNbr * -1, strCurrency)

            Else

                strDebitorName = objDbBhg.ReadDebitor3(intDebiNbr, strCurrency)

            End If

            strDebitorAr = Split(strDebitorName, "{>}")

            Return strDebitorAr(0)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Debitoren-Daten-Lesen Problem", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        Finally
            'Application.DoEvents()

        End Try


    End Function

    Friend Shared Function FcGetRefDebiNr(lngDebiNbr As Int32,
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
        Dim strMDBName As String = Main.FcReadFromSettingsII("Buchh_PKTableConnection",
                                                           intAccounting)
        'Dim objOrcommand As New OracleClient.OracleCommand
        Dim strSQL As String
        Dim intFunctionReturns As Int16

        Try

            strTableName = Main.FcReadFromSettingsII("Buchh_PKTable",
                                                   intAccounting)
            strTableType = Main.FcReadFromSettingsII("Buchh_PKTableType",
                                                   intAccounting)
            strDebFieldName = Main.FcReadFromSettingsII("Buchh_PKField",
                                                      intAccounting)
            strDebNewField = Main.FcReadFromSettingsII("Buchh_PKNewField",
                                                     intAccounting)

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
                    objdbConnDeb.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(Main.FcReadFromSettingsII("Buchh_PKTableConnection", intAccounting))
                    objdbConnDeb.Open()
                    'objsqlCommDeb.CommandText = "SELECT " + strDebFieldName + ", " + strDebNewField + ", " + strCompFieldName + ", " + strStreetFieldName + ", " + strZIPFieldName + ", " + strTownFieldName + ", " + strSageName + ", " + strDebiAccField +
                    '                            " FROM " + strTableName + " WHERE " + strDebFieldName + "=" + lngDebiNbr.ToString
                    objsqlCommDeb.CommandText = strSQL
                    objsqlCommDeb.Connection = objdbConnDeb
                    objdtDebitor.Load(objsqlCommDeb.ExecuteReader)
                    objdbConnDeb.Close()

                ElseIf strTableType = "A" Then 'Access
                    'Access
                    Call Main.FcInitAccessConnecation(objdbAccessConn, strMDBName)
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
                            intPKNewField = Main.FcGetPKNewFromRep(IIf(IsDBNull(objdtDebitor.Rows(0).Item("ID")), 0, objdtDebitor.Rows(0).Item("ID")),
                                                           "P")
                        Else
                            'D.h. Neue PK-Nr. wird nie von anderer Tabelle gelesen als t_customer oder Repbetriebe, bei einem <> t_customer muss de Rep_Betiebnr mitgegeben werden
                            intPKNewField = Main.FcGetPKNewFromRep(IIf(IsDBNull(objdtDebitor.Rows(0).Item(strDebNewField)), 0, objdtDebitor.Rows(0).Item(strDebNewField)),
                                                           "R")

                            'Stop
                        End If

                        If intPKNewField = 0 Then
                            'PK wurde nicht vergeben => Eine neue erzeugen und in der Tabelle Rep_Betriebe 
                            If strTableName = "t_customer" Then
                                intFunctionReturns = Main.FcNextPrivatePKNr(objdtDebitor.Rows(0).Item("ID"),
                                                                            intDebiNew)
                                If intFunctionReturns = 0 And intDebiNew > 0 Then 'Vergabe hat geklappt
                                    intFunctionReturns = Main.FcWriteNewPrivateDebToRepbetrieb(objdtDebitor.Rows(0).Item("ID"),
                                                                                               intDebiNew)
                                    If intFunctionReturns = 0 Then 'Schreiben hat geklappt
                                        Return 1
                                    End If
                                End If
                            Else
                                intFunctionReturns = Main.FcNextPKNr(objdtDebitor.Rows(0).Item(strDebNewField),
                                                                     intDebiNew,
                                                                     intAccounting,
                                                                     "D")
                                If intFunctionReturns = 0 And intDebiNew > 0 Then 'Vergabe hat geklappt
                                    intFunctionReturns = Main.FcWriteNewDebToRepbetrieb(objdtDebitor.Rows(0).Item(strDebNewField),
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
                            intFunctionReturns = Main.FcNextPKNr(lngDebiNbr,
                                                                 intDebiNew,
                                                                 intAccounting,
                                                                 "D")
                            If intFunctionReturns = 0 And intDebiNew > 0 Then 'Vergabe hat geklappt
                                intFunctionReturns = Main.FcWriteNewDebToRepbetrieb(lngDebiNbr,
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

    Public Shared Function FcGetDZkondFromCust(ByVal lngDebiNbr As Long,
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

    Public Shared Function FcGetDZkondFromRep(ByVal lngDebiNbr As Long,
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

    Public Shared Function FcGetDZKondSageID(ByVal intDZkond As Int16,
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


    Public Shared Function FcIsDebitorCreatable(ByVal lngDebiNbr As Long,
                                                ByRef objDbBhg As SBSXASLib.AXiDbBhg,
                                                ByVal strcmbBuha As String,
                                                ByVal intAccounting As Int16) As Int16

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

        Try

            'Angaben einlesen
            objdbConnDeb.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(Main.FcReadFromSettingsII("Buch_TabRepConnection", intAccounting))

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
                    intReturnValue = Main.FcCheckDebiIntBank(intAccounting,
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
                        intReturnValue = Main.FcGetIBANDetails(strIBANNr,
                                                          strBankName,
                                                          strBankAddress1,
                                                          strBankAddress2,
                                                          strBankBIC,
                                                          strBankCountry,
                                                          strBankClearing)

                        'Kombinierte PLZ / Ort Feld trennen
                        strBankPLZ = Left(strBankAddress2, InStr(strBankAddress2, " "))
                        strBankOrt = Trim(Right(strBankAddress2, Len(strBankAddress2) - InStr(strBankAddress2, " ")))
                    End If

                    intCreatable = FcCreateDebitor(objDbBhg,
                                              lngDebiNbr,
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

    Public Shared Function FcCreateDebitor(ByRef objDbBhg As SBSXASLib.AXiDbBhg,
                                       ByVal intDebitorNewNbr As Int32,
                                       ByVal strSuchtext As String,
                                       ByVal strDebName As String,
                                       ByVal strDebStreet As String,
                                       ByVal strDebPLZ As String,
                                       ByVal strDebOrt As String,
                                       ByVal intDebSammelKto As Int32,
                                       ByVal strGruppe As String,
                                       ByVal strVertretung As String,
                                       ByVal strAnsprechpartner As String,
                                       ByVal strLand As String,
                                       ByVal strTel As String,
                                       ByVal strFax As String,
                                       ByVal strMail As String,
                                       ByVal intLangauage As Int32,
                                       ByVal strMwStNr As String,
                                       ByVal strKreditLimite As String,
                                       ByVal intPayDefault As Int16,
                                       ByVal strZVBankName As String,
                                       ByVal strZVBankPLZ As String,
                                       ByVal strZVBankOrt As String,
                                       ByVal strZVIBAN As String,
                                       ByVal strZVBIC As String,
                                       ByVal strZVClearing As String,
                                       ByVal strCurrency As String,
                                       ByVal intDebErlKto As Int16,
                                       ByVal intDebZB As Int16,
                                       ByVal strSachB As String,
                                       ByVal intintBank As Int16,
                                       ByVal strFirtName As String) As Int16

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

            Call objDbBhg.SetCommonInfo2(intDebitorNewNbr,
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

            Call objDbBhg.SetExtendedInfo8(strDebSperren,
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
                Call objDbBhg.SetIndividInfoText(1,
                                                 strSuchtext)
            End If

            If intPayDefault = 9 Then 'IBAN
                If Len(strZVIBAN) > 15 Then
                    Call objDbBhg.SetZahlungsverbindung("B",
                                                        strZVIBAN,
                                                        strZVBankName,
                                                        "",
                                                        "",
                                                        strZVBankPLZ.ToString,
                                                        strZVBankOrt,
                                                        Left(strZVIBAN, 2),
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
            Call objDbBhg.WriteDebitor3(0, intintBank.ToString)

            'Mail über Erstellung absetzen


            Return 0
            'intDebAdrLaufN = DbBhg.GetAdressLaufnr()
            'intDebBankLaufNr = DbBhg.GetZahlungsverbLaufnr()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Debitor - Erstellung " + intDebitorNewNbr.ToString + ", " + strDebName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

            Return 1

        End Try

    End Function

    Public Shared Function FcCheckDebitor(ByVal lngDebitor As Long,
                                          ByVal intBuchungsart As Integer,
                                          ByRef objdbBuha As SBSXASLib.AXiDbBhg) As Integer

        Dim strReturn As String

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

    End Function

    Public Shared Function FcCheckDZKond(ByVal strMandant As String,
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

    Public Shared Function FcWriteToRGTable(ByVal intMandant As Int32,
                                            ByVal strRGNbr As String,
                                            ByVal datDate As Date,
                                            ByVal intBelegNr As Int32,
                                            ByRef objdbAccessConn As OleDb.OleDbConnection,
                                            ByRef objOracleConn As OracleConnection,
                                            ByRef objMySQLConn As MySqlConnection,
                                            ByVal booDatChanged As Boolean,
                                            ByVal datDebRGDatum As Date,
                                            ByVal datDebValDatum As Date) As Int16

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

        strMDBName = Main.FcReadFromSettingsII("Buchh_RGTableMDB", intMandant)
        strRGTableType = Main.FcReadFromSettingsII("Buchh_RGTableType", intMandant)
        strNameRGTable = Main.FcReadFromSettingsII("Buchh_TableDeb", intMandant)
        strBelegNrName = Main.FcReadFromSettingsII("Buchh_TableRGBelegNrName", intMandant)
        strRGNbrFieldName = Main.FcReadFromSettingsII("Buchh_TableRGNbrFieldName", intMandant)
        strDebRGFieldName = Main.FcReadFromSettingsII("Buchh_DRGDateField", intMandant)
        strDebValFieldName = Main.FcReadFromSettingsII("Buchh_DValDateField", intMandant)

        Try

            If strRGTableType = "A" Then
                'Access
                Call Main.FcInitAccessConnecation(objdbAccessConn, strMDBName)

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

    Public Shared Function FcExecuteBeforeDebit(ByVal intMandant As Integer,
                                                ByRef objMySQLConn As MySqlConnection) As Int16

        Dim strSQL As String
        Dim strBeforeDebiRunType As String
        Dim strMDBName As String
        Dim objlocMySQLRGConn As New MySqlConnection
        Dim objlocMySQLRGcmd As New MySqlCommand
        Dim intAffected As Int16


        Try

            objMySQLConn.Open()
            strSQL = Main.FcReadFromSettings(objMySQLConn, "Buchh_SQLbeforeDebiRun", intMandant)
            strBeforeDebiRunType = Main.FcReadFromSettings(objMySQLConn, "Buchh_SQLbeforeDebiType", intMandant)
            strMDBName = Main.FcReadFromSettings(objMySQLConn, "Buchh_SQLbeforeDebiMDB", intMandant)

            If strSQL <> "" Then

                If strBeforeDebiRunType = "A" Then
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

                ElseIf strBeforeDebiRunType = "M" Then
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
            MessageBox.Show(ex.Message, "Vor Debitor - Ausführung")
            Return 1

        Finally
            If objMySQLConn.State = ConnectionState.Open Then
                objMySQLConn.Close()
            End If

            If objlocMySQLRGConn.State = ConnectionState.Open Then
                objlocMySQLRGConn.Close()
            End If

        End Try

    End Function

    Public Shared Function FcExecuteAfterDebit(ByVal intMandant As Integer,
                                               ByRef objMySQLConn As MySqlConnection) As Int16

        Dim strSQL As String
        Dim strAfterDebiRunType As String
        Dim strMDBName As String
        Dim objlocMySQLRGConn As New MySqlConnection
        Dim objlocMySQLRGcmd As New MySqlCommand
        Dim intAffected As Int16


        Try

            objMySQLConn.Open()
            strSQL = Main.FcReadFromSettings(objMySQLConn, "Buchh_SQLafterDebiRun", intMandant)
            strAfterDebiRunType = Main.FcReadFromSettings(objMySQLConn, "Buchh_SQLafterDebiType", intMandant)
            strMDBName = Main.FcReadFromSettings(objMySQLConn, "Buchh_SQLafterDebiMDB", intMandant)

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
            If objMySQLConn.State = ConnectionState.Open Then
                objMySQLConn.Close()
            End If

            If objlocMySQLRGConn.State = ConnectionState.Open Then
                objlocMySQLRGConn.Close()
            End If

        End Try

    End Function

    Public Shared Function FcCheckDebiIntBank(ByVal intAccounting As Integer,
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

    Public Shared Function FcSQLParse(ByVal strSQLToParse As String,
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
                            strField = FcGetKundenzeichen(RowBooking(0).Item("lngDebIdentNbr"))
                        Case Else
                            strField = "unknown field"
                    End Select
                    strSQLToParse = Left(strSQLToParse, intPipePositionBegin - 1) & strField & Right(strSQLToParse, Len(strSQLToParse) - intPipePositionEnd)
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

    Public Shared Function FcGetKundenzeichen(ByVal lngJournalNr As Int32) As String
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
            objdtJournalKZ.Dispose()
            objdtJournalKZ = Nothing

            objdbConnCIS.Close()
            objdbConnCIS.Dispose()
            objdbConnCIS = Nothing

            objdbCmdCIS.Dispose()
            objdbCmdCIS = Nothing


        End Try

    End Function

    Public Shared Function FcIsPrivateDebitorCreatable(ByVal lngDebiNbr As Long,
                                             ByRef objDbBhg As SBSXASLib.AXiDbBhg,
                                             ByVal strcmbBuha As String,
                                             ByVal intAccounting As Int16) As Int16

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
                    intReturnValue = Main.FcCheckDebiIntBank(intAccounting,
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
                        intReturnValue = Main.FcGetIBANDetails(strIBANNr,
                                                      strBankName,
                                                      strBankAddress1,
                                                      strBankAddress2,
                                                      strBankBIC,
                                                      strBankCountry,
                                                      strBankClearing)

                        'Kombinierte PLZ / Ort Feld trennen
                        strBankPLZ = Left(strBankAddress2, InStr(strBankAddress2, " "))
                        strBankOrt = Trim(Right(strBankAddress2, Len(strBankAddress2) - InStr(strBankAddress2, " ")))
                    End If

                    'Currency - Check
                    If objdtDebitor.Rows(0).Item("DebiGegenKonto") = 1105 And lngDebiNbr >= 40000 Then
                        strCurrency = "EUR"
                    Else
                        strCurrency = "CHF"
                    End If

                    intCreatable = FcCreateDebitor(objDbBhg,
                                              lngDebiNbr,
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

                        intCreatable = MainDebitor.FcWriteDatetoPrivate(lngDebiNbr,
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

    Public Shared Function FcWriteDatetoPrivate(ByVal intNewPKNr As Int32,
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

    Public Shared Function FcCheckLinkedRG(ByRef objDbBhg As SBSXASLib.AXiDbBhg,
                                           ByVal intNewDebiNbr As Int32,
                                           ByVal strDebiCur As String,
                                           ByVal intBelegNbr As Int32) As Int16

        'Returns 0=ok, 1=Beleg nicht existent, 2=Beleg existiert, ist aber bezahlt, 9=Problem

        Dim intLaufNbr As Int32
        Dim strBeleg As String
        Dim strBelegArr() As String

        Try

            intLaufNbr = objDbBhg.doesBelegExist2(intNewDebiNbr.ToString,
                                                  strDebiCur,
                                                  intBelegNbr.ToString,
                                                  "NOT_SET",
                                                  "R",
                                                  "NOT_SET",
                                                  "NOT_SET",
                                                  "NOT_SET")

            If intLaufNbr > 0 Then
                'Prüfung ob Beleg bezahlt
                strBeleg = objDbBhg.GetBeleg(intNewDebiNbr.ToString,
                                             intLaufNbr.ToString)

                strBelegArr = Split(strBeleg, "{>}")
                If strBelegArr(4) = "B" Then
                    Return 2
                Else
                    Return 0
                End If

            Else
                Return 1
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Fehler Prüfen Splitt-Bill Bel " + intBelegNbr.ToString)
            Return 9

        Finally

        End Try

    End Function

    Public Shared Function FcGetDebitorFromLinkedRG(ByVal lngRGNbr As Int32,
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
                strTableName = Main.FcReadFromSettingsII("Buchh_TableDeb", intAccounting)
                strTableType = Main.FcReadFromSettingsII("Buchh_RGTableType", intAccounting)
                strDebFieldName = "RGNr"
                'strDebNewField = Main.FcReadFromSettings(objdbconn, "Buchh_PKNewField", intAccounting)
                'strDebNewFieldType = Main.FcReadFromSettings(objdbconn, "Buchh_PKNewFType", intAccounting)
                'strCompFieldName = Main.FcReadFromSettings(objdbconn, "Buchh_PKCompany", intAccounting)
                'strStreetFieldName = Main.FcReadFromSettings(objdbconn, "Buchh_PKStreet", intAccounting)
                'strZIPFieldName = Main.FcReadFromSettings(objdbconn, "Buchh_PKZIP", intAccounting)
                'strTownFieldName = Main.FcReadFromSettings(objdbconn, "Buchh_PKTown", intAccounting)
                'strSageName = Main.FcReadFromSettings(objdbconn, "Buchh_PKSageName", intAccounting)
                'strDebiAccField = Main.FcReadFromSettings(objdbconn, "Buchh_DPKAccount", intAccounting)

                strSQL = "SELECT PKNr " + 'strDebFieldName + ", " + strDebNewField + ", " + strCompFieldName + ", " + strStreetFieldName + ", " + strZIPFieldName + ", " + strTownFieldName + ", " + strSageName + ", " + strDebiAccField +
                     "FROM " + strTableName + " WHERE " + strDebFieldName + "=" + lngRGNbr.ToString

                If strTableName <> "" And strDebFieldName <> "" Then

                    If strTableType = "O" Then 'Oracle
                        'objOrdbconn.Open()
                        'objOrcommand.CommandText = "SELECT " + strDebFieldName + ", " + strDebNewField + ", " + strCompFieldName + ", " + strStreetFieldName + ", " + strZIPFieldName + ", " + strTownFieldName + ", " + strSageName + ", " + strDebiAccField +
                        '                            " FROM " + strTableName + " WHERE " + strDebFieldName + "=" + lngDebiNbr.ToString
                        objOrcommand.CommandText = strSQL
                        objdtDebitor.Load(objOrcommand.ExecuteReader)
                        'Ist DebiNrNew Linked oder Direkt
                        'If strDebNewFieldType = "D" Then

                        'objOrdbconn.Close()
                    ElseIf strTableType = "M" Then 'MySQL
                        intDebiNew = 0
                        'MySQL - Tabelle einlesen
                        objdbConnDeb.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(Main.FcReadFromSettingsII("Buchh_RGTableMDB", intAccounting))
                        objdbConnDeb.Open()
                        'objsqlCommDeb.CommandText = "SELECT " + strDebFieldName + ", " + strDebNewField + ", " + strCompFieldName + ", " + strStreetFieldName + ", " + strZIPFieldName + ", " + strTownFieldName + ", " + strSageName + ", " + strDebiAccField +
                        '                            " FROM " + strTableName + " WHERE " + strDebFieldName + "=" + lngDebiNbr.ToString
                        objsqlCommDeb.CommandText = strSQL
                        objsqlCommDeb.Connection = objdbConnDeb
                        objdtDebitor.Load(objsqlCommDeb.ExecuteReader)
                        objdbConnDeb.Close()

                        'ElseIf strTableType = "A" Then 'Access
                        '    'Access
                        '    Call Main.FcInitAccessConnecation(objdbAccessConn, strMDBName)
                        '    objlocOLEdbcmd.CommandText = strSQL
                        '    objdbAccessConn.Open()
                        '    objlocOLEdbcmd.Connection = objdbAccessConn
                        '    objdtDebitor.Load(objlocOLEdbcmd.ExecuteReader)
                        '    objdbAccessConn.Close()

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

    Public Shared Function FcWriteEndToSync(ByRef objdbcon As MySqlConnection,
                                     ByVal intMandant As Int32,
                                     ByVal intProzess As Int16,
                                     ByVal datLastRun As Date,
                                     ByVal intlastDuration As Int32,
                                     ByVal strLastResult As String) As Int16

        Dim objdbCmd As New MySqlCommand
        Dim strSQL As String
        Dim intAffected As Int16

        Try
            If objdbcon.State = ConnectionState.Closed Then
                objdbcon.Open()
            End If
            objdbCmd.Connection = objdbcon
            strSQL = "UPDATE t_sage_syncstatus SET LastRun='" + Format(datLastRun, "yyyy-MM-dd HH:mm:ss") + "', " +
                                                  "LastDuration=" + intlastDuration.ToString + ", " +
                                                  "LastResult='" + strLastResult + "' " +
                                     "WHERE MandantID=" + intMandant.ToString +
                                     " AND ProcessID=" + intProzess.ToString

            objdbCmd.CommandText = strSQL
            intAffected = objdbCmd.ExecuteNonQuery()
            If intAffected <> 1 Then
                Return 2
            Else
                Return 0
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Fehler Status in Sync-Tabelle schreiben")
            Return 9

        Finally
            If objdbcon.State = ConnectionState.Open Then
                objdbcon.Close()
            End If

        End Try

    End Function

    Public Shared Function FcPGVDTreatmentYC(ByRef objFBhg As SBSXASLib.AXiFBhg,
                                                ByRef objFinanz As SBSXASLib.AXFinanz,
                                                ByRef objDbBhg As SBSXASLib.AXiDbBhg,
                                                ByRef objPiFin As SBSXASLib.AXiPlFin,
                                                ByRef objBebu As SBSXASLib.AXiBeBu,
                                                ByRef objKrBhg As SBSXASLib.AXiKrBhg,
                                                ByVal tblDebiB As DataTable,
                                                ByVal strDRGNbr As String,
                                                ByVal intDBelegNr As Int32,
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
        Dim drDebiSub() As DataRow
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
                        datPGVStart = "2024-01-01"
                        datValuta = datValutaSave
                        intITY = 1
                        intINY = 0
                        intAcctTY = 1312
                    End If
                End If

                'Evtl. Aufteilen auf 2 Jahre
                For intYearLooper As Int16 = Year(datValuta) To Year(datPGVEnd)

                    If intYearLooper = 2023 Then
                        dblNettoBetrag = drDSubrow("dblNetto") * -1 / intITotal * intITY
                        intHabenKonto = intAcctTY
                    Else
                        dblNettoBetrag = drDSubrow("dblNetto") * -1 / intITotal * intINY
                        intHabenKonto = intAcctNY
                    End If

                    If dblNettoBetrag <> 0 Then 'Falls in einem Jahr nichts zu buchen ist

                        strBelegDatum = Format(datValuta, "yyyyMMdd").ToString

                        If intITotal = 1 Then
                            If Year(datValuta) = 2023 Then
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
                                strValutaDatum = "20240101"
                                strBelegDatum = "20240101"
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
                            dblKursD = Main.FcGetKurs(strCur, strValutaDatum, objFBhg)
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

                        If Year(datValuta) = 2023 And Year(datValuta) <> Val(strYear) Then 'Achtung provisorisch
                            'Zuerst Info-Table löschen
                            objdtInfo.Clear()
                            'Application.DoEvents()
                            'Im 2022 anmelden
                            intReturnValue = Main.FcLoginSage2(objdbcon,
                                                          objsqlcon,
                                                          objsqlcmd,
                                                          objFinanz,
                                                          objFBhg,
                                                          objDbBhg,
                                                          objPiFin,
                                                          objBebu,
                                                          objKrBhg,
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
                            ''Application.DoEvents()

                        ElseIf Year(datValuta) = 2024 And Year(datValuta) <> Val(strYear) Then
                            'Zuerst Info-Table löschen
                            objdtInfo.Clear()
                            'Application.DoEvents()
                            'Im 2023 anmelden
                            intReturnValue = Main.FcLoginSage2(objdbcon,
                                                          objsqlcon,
                                                          objsqlcmd,
                                                          objFinanz,
                                                          objFBhg,
                                                          objDbBhg,
                                                          objPiFin,
                                                          objBebu,
                                                          objKrBhg,
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
                        objFBhg.CheckDoubleIntBelNbr = "N"

                        'Buchen
                        Call objFBhg.WriteBuchung(0,
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
                    strPeriodenInfo = objFinanz.GetPeriListe(0)
                    strPeriodenInfoA = Split(strPeriodenInfo, "{>}")

                    'Ist aktuell angemeldete Periode = FJ
                    If Year(datPGVEnd) <> Val(Strings.Left(strPeriodenInfo, 4)) Then
                        'Zuerst Info-Table löschen
                        objdtInfo.Clear()
                        'Application.DoEvents()
                        'Login ins FJ
                        intReturnValue = Main.FcLoginSage2(objdbcon,
                                                          objsqlcon,
                                                          objsqlcmd,
                                                          objFinanz,
                                                          objFBhg,
                                                          objDbBhg,
                                                          objPiFin,
                                                          objBebu,
                                                          objKrBhg,
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
                        objFBhg.CheckDoubleIntBelNbr = "N"

                        'Buchen
                        Call objFBhg.WriteBuchung(0,
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
                        If Year(datValuta) = 2023 Then
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

                    If Year(datValuta) = 2023 And Year(datValuta) <> Val(strYear) Then 'Achtung provisorisch
                        'Zuerst Info-Table löschen
                        objdtInfo.Clear()
                        'Application.DoEvents()
                        'Im 2022 anmelden
                        intReturnValue = Main.FcLoginSage2(objdbcon,
                                                          objsqlcon,
                                                          objsqlcmd,
                                                          objFinanz,
                                                          objFBhg,
                                                          objDbBhg,
                                                          objPiFin,
                                                          objBebu,
                                                          objKrBhg,
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
                        'Im 2023 anmelden
                        intReturnValue = Main.FcLoginSage2(objdbcon,
                                                          objsqlcon,
                                                          objsqlcmd,
                                                          objFinanz,
                                                          objFBhg,
                                                          objDbBhg,
                                                          objPiFin,
                                                          objBebu,
                                                          objKrBhg,
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
                    objFBhg.CheckDoubleIntBelNbr = "N"

                    'Buchen
                    Call objFBhg.WriteBuchung(0,
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
            If strYear <> strActualYear Then
                'Zuerst Info-Table löschen
                objdtInfo.Clear()
                'Application.DoEvents()
                'Im Aufrufjahr anmelden
                intReturnValue = Main.FcLoginSage2(objdbcon,
                                                  objsqlcon,
                                                  objsqlcmd,
                                                  objFinanz,
                                                  objFBhg,
                                                  objDbBhg,
                                                  objPiFin,
                                                  objBebu,
                                                  objKrBhg,
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
            MessageBox.Show(ex.Message, "Problem PGV - Buchung Debitoren", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 9

        Finally

        End Try

    End Function

    Public Shared Function FcPGVDTreatment(ByRef objFBhg As SBSXASLib.AXiFBhg,
                                                ByRef objFinanz As SBSXASLib.AXFinanz,
                                                ByRef objDbBhg As SBSXASLib.AXiDbBhg,
                                                ByRef objPiFin As SBSXASLib.AXiPlFin,
                                                ByRef objBebu As SBSXASLib.AXiBeBu,
                                                ByRef objKrBhg As SBSXASLib.AXiKrBhg,
                                                ByVal tblDebiB As DataTable,
                                                ByVal strDRGNbr As String,
                                                ByVal intDBelegNr As Int32,
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
        Dim drDebiSub() As DataRow
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
            'Zuerst betroffene Buchungen selektieren
            drDebiSub = tblDebiB.Select("strRGNr='" + strDRGNbr + "' AND dblNetto<>0")

            'Durch die Buchungen steppen
            For Each drDSubrow As DataRow In drDebiSub
                'Auflösung
                '=========

                datValuta = datValutaSave
                If strPGVType = "RV" Then
                    datPGVStart = "2024-01-01"
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
                            dblKursD = Main.FcGetKurs(strCur, strValutaDatum, objFBhg)
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
                        Call objFBhg.WriteBuchung(0,
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
                    strPeriodenInfo = objFinanz.GetPeriListe(0)
                    strPeriodenInfoA = Split(strPeriodenInfo, "{>}")

                    'Ist aktuell angemeldete Periode = FJ
                    If Year(datPGVEnd) <> Val(Strings.Left(strPeriodenInfo, 4)) Then
                        'Zuerst Info-Table löschen
                        objdtInfo.Clear()
                        'Application.DoEvents()
                        'Login ins FJ
                        intReturnValue = Main.FcLoginSage2(objdbcon,
                                                      objsqlcon,
                                                      objsqlcmd,
                                                      objFinanz,
                                                      objFBhg,
                                                      objDbBhg,
                                                      objPiFin,
                                                      objBebu,
                                                      objKrBhg,
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
                        intHabenKonto = intAcctTY
                        intSollKonto = intAcctNY
                        strDebiTextHaben = drDSubrow("strDebSubText") + ", PGV AJ / FJ"
                        strDebiTextSoll = drDSubrow("strDebSubText") + ", PGV AJ / FJ"
                        strBebuEintragHaben = Nothing
                        strBebuEintragSoll = Nothing

                        'Buchen
                        Call objFBhg.WriteBuchung(0,
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

                    If Year(datValuta) = 2023 And Year(datValuta) <> Val(strYear) Then 'Achtung provisorisch
                        'Zuerst Info-Table löschen
                        objdtInfo.Clear()
                        'Application.DoEvents()
                        'Im 2023 anmelden
                        intReturnValue = Main.FcLoginSage2(objdbcon,
                                                      objsqlcon,
                                                      objsqlcmd,
                                                      objFinanz,
                                                      objFBhg,
                                                      objDbBhg,
                                                      objPiFin,
                                                      objBebu,
                                                      objKrBhg,
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
                        'Im 2023 anmelden
                        intReturnValue = Main.FcLoginSage2(objdbcon,
                                                      objsqlcon,
                                                      objsqlcmd,
                                                      objFinanz,
                                                      objFBhg,
                                                      objDbBhg,
                                                      objPiFin,
                                                      objBebu,
                                                      objKrBhg,
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
                    Call objFBhg.WriteBuchung(0,
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
            If strYear <> strActualYear Then
                'Zuerst Info-Table löschen
                objdtInfo.Clear()
                'Application.DoEvents()
                'Im Aufrufjahr anmelden
                intReturnValue = Main.FcLoginSage2(objdbcon,
                                              objsqlcon,
                                              objsqlcmd,
                                              objFinanz,
                                              objFBhg,
                                              objDbBhg,
                                              objPiFin,
                                              objBebu,
                                              objKrBhg,
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
            MessageBox.Show(ex.Message, "Problem PGV - Buchung Debitoren", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 9

        Finally

        End Try

    End Function

    Public Shared Function FcCheckDebiExistance(ByRef intBelegNbr As Int32,
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

    Public Shared Function FcWriteDebHeadToDB(ByVal objdtDebitorenHeadRead As DataTable,
                                              ByVal intBuha As Int16) As Int16

        'Returns: 0=ok, 9=Problem

        'Tabelle speichern.
        Dim strSQLFields As String
        Dim strSQLValues As String
        Dim strActualValue As String
        Dim intdbAffected As Int16
        Dim objdbConn As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString"))
        Dim objdbcommand As New MySqlCommand

        Try

            For Each drDebitorenHead As DataRow In objdtDebitorenHeadRead.Rows
                For Each clDebitorenHead As DataColumn In drDebitorenHead.Table.Columns
                    'Feldnamen zusammenstellen
                    strSQLFields = strSQLFields + IIf(Len(strSQLFields) > 0, ", ", "(") + clDebitorenHead.ColumnName
                    'Je nach Feld andere Werte setzen
                    Select Case clDebitorenHead.ColumnName
                        Case "intBuchhaltung"
                            strActualValue = intBuha.ToString
                            'Case "booDebBook"
                            '    strActualValue = "false"
                        Case Else
                            'Wert je nach Typ mit oder ohne ' setzen
                            Select Case clDebitorenHead.DataType.Name
                                Case "String"
                                    strActualValue = drDebitorenHead.Item(clDebitorenHead.ColumnName).ToString
                                    'Apostrophs entfernen
                                    strActualValue = Replace(strActualValue, "'", "`")
                                    strActualValue = "'" + strActualValue + "'"
                                Case "DateTime"
                                    If IsDBNull(drDebitorenHead.Item(clDebitorenHead.ColumnName)) Then
                                        strActualValue = "null"
                                    Else
                                        strActualValue = "'" + DateAndTime.Year(drDebitorenHead.Item(clDebitorenHead.ColumnName)).ToString + "-" + DateAndTime.Month(drDebitorenHead.Item(clDebitorenHead.ColumnName)).ToString + "-" + DateAndTime.Day(drDebitorenHead.Item(clDebitorenHead.ColumnName)).ToString + "'"
                                    End If
                                Case "Boolean"
                                    If IsDBNull(drDebitorenHead.Item(clDebitorenHead.ColumnName)) Then
                                        strActualValue = "false"
                                    ElseIf drDebitorenHead.Item(clDebitorenHead.ColumnName) = 0 Then
                                        strActualValue = "false"
                                    Else
                                        strActualValue = "true"
                                    End If
                                Case Else
                                    If IsDBNull(drDebitorenHead.Item(clDebitorenHead.ColumnName)) Then
                                        strActualValue = "null"
                                    Else
                                        strActualValue = drDebitorenHead.Item(clDebitorenHead.ColumnName).ToString
                                    End If
                            End Select
                    End Select
                    strSQLValues = strSQLValues + IIf(Len(strSQLValues) > 0, ", ", "(") + strActualValue
                Next
                'Jetzt noch Identity und Process - ID hinzufügen
                strSQLFields = strSQLFields + ", IdentityName, ProcessID)"
                strSQLValues = strSQLValues + ", '" + System.Security.Principal.WindowsIdentity.GetCurrent().Name + "', " + Process.GetCurrentProcess().Id.ToString + ")"
                objdbConn.Open()
                objdbcommand.Connection = objdbConn
                objdbcommand.CommandText = "INSERT INTO tbldebitorenjhead " + strSQLFields + " VALUES " + strSQLValues
                intdbAffected = objdbcommand.ExecuteNonQuery()
                objdbConn.Close()
                strSQLFields = ""
                strSQLValues = ""
                strActualValue = ""
            Next
            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Debitor - HeadTabelleTODB ")
            Return 9

        Finally
            objdbcommand = Nothing
            objdbConn = Nothing

        End Try


    End Function

    Public Shared Function FcWriteDebSubToDB(ByVal objdtDebitorenHeadRead As DataTable) As Int16

        'Returns: 0=ok, 9=Problem

        'Tabelle speichern.
        Dim strSQLFields As String
        Dim strSQLValues As String
        Dim strActualValue As String
        Dim intdbAffected As Int16
        Dim objdbConn As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString"))
        Dim objdbcommand As New MySqlCommand

        Try

            For Each drDebitorenHead As DataRow In objdtDebitorenHeadRead.Rows
                For Each clDebitorenHead As DataColumn In drDebitorenHead.Table.Columns
                    'Feldnamen zusammenstellen
                    strSQLFields = strSQLFields + IIf(Len(strSQLFields) > 0, ", ", "(") + clDebitorenHead.ColumnName
                    'Je nach Feld andere Werte setzen

                    'Wert je nach Typ mit oder ohne ' setzen
                    Select Case clDebitorenHead.DataType.Name
                        Case "String"
                            strActualValue = "'" + drDebitorenHead.Item(clDebitorenHead.ColumnName).ToString + "'"
                        Case "DateTime"
                            If IsDBNull(drDebitorenHead.Item(clDebitorenHead.ColumnName)) Then
                                strActualValue = "null"
                            Else
                                strActualValue = "'" + DateAndTime.Year(drDebitorenHead.Item(clDebitorenHead.ColumnName)).ToString + "-" + DateAndTime.Month(drDebitorenHead.Item(clDebitorenHead.ColumnName)).ToString + "-" + DateAndTime.Day(drDebitorenHead.Item(clDebitorenHead.ColumnName)).ToString + "'"
                            End If
                        Case "Boolean"
                            If IsDBNull(drDebitorenHead.Item(clDebitorenHead.ColumnName)) Then
                                strActualValue = "false"
                            ElseIf drDebitorenHead.Item(clDebitorenHead.ColumnName) = 0 Then
                                strActualValue = "false"
                            Else
                                strActualValue = "true"
                            End If
                        Case Else
                            If IsDBNull(drDebitorenHead.Item(clDebitorenHead.ColumnName)) Then
                                strActualValue = "null"
                            Else
                                strActualValue = drDebitorenHead.Item(clDebitorenHead.ColumnName).ToString
                            End If
                    End Select

                    strSQLValues = strSQLValues + IIf(Len(strSQLValues) > 0, ", ", "(") + strActualValue
                Next
                'Jetzt noch Identity und Process - ID hinzufügen
                strSQLFields = strSQLFields + ", IdentityName, ProcessID)"
                strSQLValues = strSQLValues + ", '" + System.Security.Principal.WindowsIdentity.GetCurrent().Name + "', " + Process.GetCurrentProcess().Id.ToString + ")"
                objdbConn.Open()
                objdbcommand.Connection = objdbConn
                objdbcommand.CommandText = "INSERT INTO tbldebitorensub " + strSQLFields + " VALUES " + strSQLValues
                intdbAffected = objdbcommand.ExecuteNonQuery()
                objdbConn.Close()
                strSQLFields = ""
                strSQLValues = ""
                strActualValue = ""
            Next
            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Debitor - SubTabelleTODB ")
            Return 9

        Finally
            objdbcommand = Nothing
            objdbConn = Nothing

        End Try


    End Function


End Class
