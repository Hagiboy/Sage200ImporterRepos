﻿Option Strict Off
Option Explicit On

Imports MySql.Data.MySqlClient
Imports System.Data.OracleClient
Imports System.Net
Imports System.IO
Imports System.Xml


Public Class MainDebitor

    Public Shared Function FcFillDebit(ByVal intAccounting As Integer,
                                       ByRef objdtHead As DataTable,
                                       ByRef objdtSub As DataTable,
                                       ByRef objdbconn As MySqlConnection,
                                       ByRef objdbAccessConn As OleDb.OleDbConnection,
                                       ByRef objOracleCon As OracleConnection,
                                       ByRef objOracleCmd As OracleCommand) As Integer

        Dim strSQL As String
        Dim strSQLSub As String
        Dim strRGTableType As String
        Dim objRGMySQLConn As New MySqlConnection
        Dim objlocMySQLcmd As New MySqlCommand
        Dim objlocOLEdbcmd As New OleDb.OleDbCommand

        Dim objDTDebiHead As New DataTable
        'Dim objdrSub As DataRow
        'Dim intFcReturns As Int16
        Dim strMDBName As String

        objdbconn.Open()

        strMDBName = Main.FcReadFromSettings(objdbconn, "Buchh_RGTableMDB", intAccounting)

        'Head Debitoren löschen
        objdtHead.Clear()
        strSQL = Main.FcReadFromSettings(objdbconn, "Buchh_SQLHead", intAccounting)
        strRGTableType = Main.FcReadFromSettings(objdbconn, "Buchh_RGTableType", intAccounting)

        Try

            'objlocMySQLcmd.CommandText = strSQL
            If strRGTableType = "A" Then
                'Access
                Call Main.FcInitAccessConnecation(objdbAccessConn, strMDBName)

                objlocOLEdbcmd.CommandText = strSQL
                objdbAccessConn.Open()
                objlocOLEdbcmd.Connection = objdbAccessConn
                objdtHead.Load(objlocOLEdbcmd.ExecuteReader)
            ElseIf strRGTableType = "M" Then
                objRGMySQLConn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(strMDBName)
                objlocMySQLcmd.Connection = objRGMySQLConn
                objlocMySQLcmd.CommandText = strSQL
                objRGMySQLConn.Open()
                objdtHead.Load(objlocMySQLcmd.ExecuteReader)
            End If
            'objlocMySQLcmd.Connection = objdbconn
            'objDTDebiHead.Load(objlocMySQLcmd.ExecuteReader)
            'Durch die Records steppen und Sub-Tabelle füllen
            For Each row In objdtHead.Rows
                'Debug.Print(strSQLSub)
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
                strSQLSub = MainDebitor.FcSQLParse(Main.FcReadFromSettings(objdbconn,
                                                                           "Buchh_SQLDetail",
                                                                           intAccounting),
                                                   row("strDebRGNbr"),
                                                   objdtHead,
                                                   objOracleCon,
                                                   objOracleCmd,
                                                   "D")
                If strRGTableType = "A" Then
                    objlocOLEdbcmd.CommandText = strSQLSub
                    objdtSub.Load(objlocOLEdbcmd.ExecuteReader)
                ElseIf strRGTableType = "M" Then
                    objlocMySQLcmd.CommandText = strSQLSub
                    objdtSub.Load(objlocMySQLcmd.ExecuteReader)
                End If
            Next
            'Tabellen runden
            'intFcReturns = FcRoundInTable(objdtHead, "dblDebNetto", 2)
            'intFcReturns = FcRoundInTable(objdtHead, "dblDebBrutto", 2)
            'intFcReturns = FcRoundInTable(objdtHead, "dblDebMwSt", 2)
            'intFcReturns = FcRoundInTable(objdtSub, "dblNetto", 2)
            'intFcReturns = FcRoundInTable(objdtSub, "dblMwSt", 2)
            'intFcReturns = FcRoundInTable(objdtSub, "dblBrutto", 2)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Debitoren-Daten-Lesen Problem")

        Finally

            If objdbAccessConn.State = ConnectionState.Open Then
                objdbAccessConn.Close()
            End If
            If objRGMySQLConn.State = ConnectionState.Open Then
                objRGMySQLConn.Close()
            End If
            objdbconn.Close()

        End Try

    End Function

    Public Shared Function FcReadDebitorKName(ByRef objfiBuha As SBSXASLib.AXiFBhg, ByVal lngDebKtoNbr As Long) As String

        Dim strDebitorKName As String
        Dim strDebitorKAr() As String

        strDebitorKName = objfiBuha.GetKontoInfo(lngDebKtoNbr)

        strDebitorKAr = Split(strDebitorKName, "{>}")

        If strDebitorKName <> "EOF" Then
            Return strDebitorKAr(8)
        Else
            Return "EOF"
        End If

    End Function

    Public Shared Function FcReadDebitorName(ByRef objDbBhg As SBSXASLib.AXiDbBhg, ByVal intDebiNbr As Int32, ByVal strCurrency As String) As String

        Dim strDebitorName As String
        Dim strDebitorAr() As String

        If strCurrency = "" Then

            strDebitorName = objDbBhg.ReadDebitor3(intDebiNbr * -1, strCurrency)

        Else

            strDebitorName = objDbBhg.ReadDebitor3(intDebiNbr, strCurrency)

        End If

        strDebitorAr = Split(strDebitorName, "{>}")

        Return strDebitorAr(0)

    End Function

    Public Shared Function FcGetRefDebiNr(ByRef objdbconn As MySqlConnection,
                                          ByRef objdbconnZHDB02 As MySqlConnection,
                                          ByRef objsqlcommand As MySqlCommand,
                                          ByRef objsqlcommandZHDB02 As MySqlCommand,
                                          ByRef objOrdbconn As OracleClient.OracleConnection,
                                          ByRef objOrcommand As OracleClient.OracleCommand,
                                          ByRef objdbAccessConn As OleDb.OleDbConnection,
                                          ByVal lngDebiNbr As Int32,
                                          ByVal intAccounting As Int32,
                                          ByRef intDebiNew As Int32) As Int16

        'Return 0=ok, 1=Neue Debi genereiert und gesetzt, 2=Rep_Ref nicht definiert, 3=Nicht in Tab_Repbetriebe, 4=keine Angaben in Tab_Repbetriebe

        Dim strTableName, strTableType, strDebFieldName, strDebNewField, strDebNewFieldType, strCompFieldName, strStreetFieldName, strZIPFieldName, strTownFieldName, strSageName, strDebiAccField As String
        'Dim intCreatable As Int16
        Dim objdtDebitor As New DataTable
        Dim intPKNewField As Int32
        Dim objdbConnDeb As New MySqlConnection
        Dim objsqlCommDeb As New MySqlCommand

        Dim objlocOLEdbcmd As New OleDb.OleDbCommand
        Dim strMDBName As String = Main.FcReadFromSettings(objdbconn, "Buchh_PKTableConnection", intAccounting)
        Dim strSQL As String
        Dim intFunctionReturns As Int16

        strTableName = Main.FcReadFromSettings(objdbconn, "Buchh_PKTable", intAccounting)
        strTableType = Main.FcReadFromSettings(objdbconn, "Buchh_PKTableType", intAccounting)
        strDebFieldName = Main.FcReadFromSettings(objdbconn, "Buchh_PKField", intAccounting)
        strDebNewField = Main.FcReadFromSettings(objdbconn, "Buchh_PKNewField", intAccounting)
        strDebNewFieldType = Main.FcReadFromSettings(objdbconn, "Buchh_PKNewFType", intAccounting)
        strCompFieldName = Main.FcReadFromSettings(objdbconn, "Buchh_PKCompany", intAccounting)
        strStreetFieldName = Main.FcReadFromSettings(objdbconn, "Buchh_PKStreet", intAccounting)
        strZIPFieldName = Main.FcReadFromSettings(objdbconn, "Buchh_PKZIP", intAccounting)
        strTownFieldName = Main.FcReadFromSettings(objdbconn, "Buchh_PKTown", intAccounting)
        strSageName = Main.FcReadFromSettings(objdbconn, "Buchh_PKSageName", intAccounting)
        strDebiAccField = Main.FcReadFromSettings(objdbconn, "Buchh_DPKAccount", intAccounting)

        strSQL = "SELECT * " + 'strDebFieldName + ", " + strDebNewField + ", " + strCompFieldName + ", " + strStreetFieldName + ", " + strZIPFieldName + ", " + strTownFieldName + ", " + strSageName + ", " + strDebiAccField +
                 " FROM " + strTableName + " WHERE " + strDebFieldName + "=" + lngDebiNbr.ToString

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
                objdbConnDeb.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(Main.FcReadFromSettings(objdbconn, "Buchh_PKTableConnection", intAccounting))
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
                        intPKNewField = Main.FcGetPKNewFromRep(objdbconnZHDB02,
                                                           objsqlcommandZHDB02,
                                                           IIf(IsDBNull(objdtDebitor.Rows(0).Item("ID")), 0, objdtDebitor.Rows(0).Item("ID")),
                                                           "P")
                    Else
                        intPKNewField = Main.FcGetPKNewFromRep(objdbconnZHDB02,
                                                           objsqlcommandZHDB02,
                                                           IIf(IsDBNull(objdtDebitor.Rows(0).Item(strDebNewField)), 0, objdtDebitor.Rows(0).Item(strDebNewField)),
                                                           "R")
                    End If

                    If intPKNewField = 0 Then
                        'PK wurde nicht vergeben => Eine neue erzeugen und in der Tabelle Rep_Betriebe 
                        If strTableName = "t_customer" Then
                            intFunctionReturns = Main.FcNextPrivatePKNr(objdbconn, objdtDebitor.Rows(0).Item("ID"), intDebiNew)
                            If intFunctionReturns = 0 And intDebiNew > 0 Then 'Vergabe hat geklappt
                                intFunctionReturns = Main.FcWriteNewPrivateDebToRepbetrieb(objdbconn, objdtDebitor.Rows(0).Item("ID"), intDebiNew)
                                If intFunctionReturns = 0 Then 'Schreiben hat geklappt
                                    Return 1
                                End If
                            End If
                        Else
                            intFunctionReturns = Main.FcNextPKNr(objdbconnZHDB02, objdtDebitor.Rows(0).Item(strDebNewField), intDebiNew)
                            If intFunctionReturns = 0 And intDebiNew > 0 Then 'Vergabe hat geklappt
                                intFunctionReturns = Main.FcWriteNewDebToRepbetrieb(objdbconnZHDB02, objdtDebitor.Rows(0).Item(strDebNewField), intDebiNew)
                                If intFunctionReturns = 0 Then 'Schreiben hat geklappt
                                    Return 1
                                End If
                            End If
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
                        intFunctionReturns = Main.FcNextPKNr(objdbconnZHDB02, lngDebiNbr, intDebiNew)
                        If intFunctionReturns = 0 And intDebiNew > 0 Then 'Vergabe hat geklappt
                            intFunctionReturns = Main.FcWriteNewDebToRepbetrieb(objdbconnZHDB02, lngDebiNbr, intDebiNew)
                            If intFunctionReturns = 0 Then 'Schreiben hat geklappt
                                Return 1
                            End If
                        End If
                    End If
                    Return 0
                End If
            End If
            'Else
            'intDebiNew = 0
            'Return 4
        End If

        'End If

        Return intPKNewField

    End Function

    Public Shared Function FcIsDebitorCreatable(ByRef objdbconn As MySqlConnection,
                                                ByRef objdbconnZHDB02 As MySqlConnection,
                                                ByRef objsqlcommandZHDB02 As MySqlCommand,
                                                ByVal lngDebiNbr As Long,
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
        Dim strBankName As String = ""
        Dim strBankAddress1 As String = ""
        Dim strBankAddress2 As String = ""
        Dim strBankPLZ As String = ""
        Dim strBankOrt As String = ""
        Dim strBankBIC As String = ""
        Dim strBankCountry As String = ""
        Dim strBankClearing As String = ""
        Dim intReturnValue As Int16
        Dim intDebZB As Int16
        Dim objdsDebitor As New DataSet
        Dim objDADebitor As New MySqlDataAdapter
        Dim objdtSachB As New DataTable("tbliSachB")
        Dim strSachB As String
        Dim intPayType As Int16
        Dim intintBank As Int16

        Try

            'Angaben einlesen
            If objdbconnZHDB02.State = ConnectionState.Closed Then
                objdbconnZHDB02.Open()
            End If
            objsqlcommandZHDB02.Connection = objdbconnZHDB02
            objsqlcommandZHDB02.CommandText = "SELECT Rep_Nr, Rep_Firma, Rep_Strasse, Rep_PLZ, Rep_Ort, Rep_DebiKonto, Rep_Gruppe, Rep_Vertretung, Rep_Ansprechpartner, IF(Rep_Land IS NULL, 'Schweiz', Rep_Land) AS Rep_Land, Rep_Tel1, Rep_Fax, Rep_Mail, " +
                                                "IF(Rep_Language IS NULL, 'D', Rep_Language) AS Rep_Language, Rep_Kredi_MWSTNr, Rep_Kreditlimite, Rep_Kred_Pay_Def, Rep_Kred_Bank_Name, Rep_Kred_Bank_PLZ, Rep_Kred_Bank_Ort, Rep_Kred_IBAN, Rep_Kred_Bank_BIC, " +
                                                "IF(Rep_Kred_Currency IS NULL, 'CHF', Rep_Kred_Currency) AS Rep_Kred_Currency, Rep_Kred_PCKto, Rep_DebiErloesKonto, Rep_Kred_BankIntern FROM Tab_Repbetriebe WHERE PKNr=" + lngDebiNbr.ToString
            objdtDebitor.Load(objsqlcommandZHDB02.ExecuteReader)

            'Gefunden?
            If objdtDebitor.Rows.Count > 0 Then
                'Debug.Print("Gefunden, kann erstellt werden")

                'Sachbearbeiter suchen
                'Ist Ausnahme definiert?
                objsqlcommandZHDB02.CommandText = "SELECT CustomerID FROM t_rep_sagesachbearbeiter WHERE Rep_Nr=" + objdtDebitor.Rows(0).Item("Rep_Nr").ToString + " AND Buchh_Nr=" + intAccounting.ToString
                objdtSachB.Load(objsqlcommandZHDB02.ExecuteReader)
                If objdtSachB.Rows.Count > 0 Then 'Ausnahme definiert auf Rep-Betrieb
                    strSachB = Trim(objdtSachB.Rows(0).Item("CustomerID").ToString)
                Else
                    'Default setzen
                    objsqlcommandZHDB02.CommandText = "SELECT CustomerID FROM t_rep_sagesachbearbeiter WHERE Rep_Nr=2535 AND Buchh_Nr=" + intAccounting.ToString
                    objdtSachB.Load(objsqlcommandZHDB02.ExecuteReader)
                    If objdtSachB.Rows.Count > 0 Then 'Default ist definiert
                        strSachB = Trim(objdtSachB.Rows(0).Item("CustomerID").ToString)
                    Else
                        strSachB = ""
                        MessageBox.Show("Kein Sachbearbeiter - Default gesetzt für Buha " + strcmbBuha, "Debitorenerstellung")
                    End If
                End If

                'interne Bank
                intReturnValue = Main.FcCheckDebiIntBank(objdbconn,
                                                         intAccounting,
                                                         objdtDebitor.Rows(0).Item("Rep_Kred_BankIntern"),
                                                         intintBank)

                'Zahlungsbedingung suchen
                'objdtKreditor.Clear()
                'Es muss der Weg über ein Dataset genommen werden da sosnt constraint-Meldungen kommen
                objsqlcommandZHDB02.CommandText = "Select Tab_Repbetriebe.PKNr, t_sage_zahlungskondition.SageID " +
                                                  "FROM Tab_Repbetriebe INNER JOIN t_sage_zahlungskondition On Tab_Repbetriebe.Rep_DebiZKonditionID = t_sage_zahlungskondition.ID " +
                                                  "WHERE Tab_Repbetriebe.PKNr=" + lngDebiNbr.ToString
                objDADebitor.SelectCommand = objsqlcommandZHDB02
                objdsDebitor.EnforceConstraints = False
                objDADebitor.Fill(objdsDebitor)

                'objdsKreditor.Load(objsqlcommandZHDB02.ExecuteReader)
                'objdtKreditor.Load(objsqlcommandZHDB02.ExecuteReader)
                If Not IsDBNull(objdsDebitor.Tables(0).Rows(0).Item("SageID")) Then
                    intDebZB = objdsDebitor.Tables(0).Rows(0).Item("SageID")
                Else
                    intDebZB = 1
                End If

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
                    Case Else
                        strLand = "CH"
                End Select

                'Sprache zuweisen von 1-Stelligem String nach Sage 200 Regionen
                Select Case IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Language")), "D", objdtDebitor.Rows(0).Item("Rep_Language"))
                    Case "D", ""
                        intLangauage = 2055
                    Case "F"
                        intLangauage = 4108
                    Case "I"
                        intLangauage = 2064
                    Case Else
                        intLangauage = 2057 'Englisch
                End Select

                'Variablen zuweisen für die Erstellung des Debitors
                strIBANNr = IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Kred_IBAN")), "", objdtDebitor.Rows(0).Item("Rep_Kred_IBAN"))
                strBankName = IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Kred_Bank_Name")), "", objdtDebitor.Rows(0).Item("Rep_Kred_Bank_Name"))
                strBankAddress1 = ""
                strBankPLZ = IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Kred_Bank_PLZ")), "", objdtDebitor.Rows(0).Item("Rep_Kred_Bank_PLZ"))
                strBankOrt = IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Kred_Bank_Ort")), "", objdtDebitor.Rows(0).Item("Rep_Kred_Bank_Ort"))
                strBankAddress2 = strBankPLZ + " " + strBankOrt
                strBankBIC = IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Kred_Bank_BIC")), "", objdtDebitor.Rows(0).Item("Rep_Kred_Bank_BIC"))
                strBankClearing = IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Kred_PCKto")), "", objdtDebitor.Rows(0).Item("Rep_Kred_PCKto"))

                If Len(strIBANNr) = 21 Then 'IBAN
                    'If intPayType <> 9 Then 'Type nicht IBAN angegeben aber IBAN - Nr. erfasst
                    intPayType = 9
                    'End If
                    intReturnValue = Main.FcGetIBANDetails(objdbconn,
                                                      strIBANNr,
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
                                          IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Kred_Currency")), "CHF", objdtDebitor.Rows(0).Item("Rep_Kred_Currency")),
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
            Else
                Return 4

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Debitor - Erstellbar - Abklärung")
            Return 9

        Finally
            objdbconnZHDB02.Close()

        End Try

    End Function

    Public Shared Function FcCreateDebitor(ByRef objDbBhg As SBSXASLib.AXiDbBhg,
                                       ByVal intDebitorNewNbr As Int32,
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
        Dim strText As String = ""
        Dim strTelefon1 As String
        Dim strTelefax As String

        strText = IIf(strGruppe = "", "", "Gruppe: " + strGruppe) + IIf(strVertretung = "" Or "0", "", strText + vbCrLf + "Vertretung: " + strVertretung)
        strTelefon1 = IIf(strTel = "" Or strTel = "0", "", strTel)
        strTelefax = IIf(strFax = "" Or strFax = "0", "", strFax)

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
            MessageBox.Show(ex.Message, "Debitor - Erstellung " + intDebitorNewNbr.ToString + ", " + strDebName)

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

    Public Shared Function FcWriteToRGTable(ByVal intMandant As Int32, ByVal strRGNbr As String, ByVal datDate As Date, ByVal intBelegNr As Int32, ByRef objdbAccessConn As OleDb.OleDbConnection, ByRef objOracleConn As OracleConnection, ByRef objMySQLConn As MySqlConnection) As Int16

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


        objMySQLConn.Open()

        strMDBName = Main.FcReadFromSettings(objMySQLConn, "Buchh_RGTableMDB", intMandant)
        strRGTableType = Main.FcReadFromSettings(objMySQLConn, "Buchh_RGTableType", intMandant)
        strNameRGTable = Main.FcReadFromSettings(objMySQLConn, "Buchh_TableDeb", intMandant)
        strBelegNrName = Main.FcReadFromSettings(objMySQLConn, "Buchh_TableRGBelegNrName", intMandant)
        strRGNbrFieldName = Main.FcReadFromSettings(objMySQLConn, "Buchh_TableRGNbrFieldName", intMandant)
        'strSQL = "UPDATE " + strNameRGTable + " SET gebucht=true, gebuchtDatum=#" + Format(datDate, "yyyy-MM-dd").ToString + "#, " + strBelegNrName + "=" + intBelegNr.ToString + " WHERE " + strRGNbrFieldName + "=" + strRGNbr

        Try

            If strRGTableType = "A" Then
                'Access
                Call Main.FcInitAccessConnecation(objdbAccessConn, strMDBName)

                strSQL = "UPDATE " + strNameRGTable + " SET gebucht=true, gebuchtDatum=#" + Format(datDate, "yyyy-MM-dd").ToString + "#, " + strBelegNrName + "=" + intBelegNr.ToString + " WHERE " + strRGNbrFieldName + "=" + strRGNbr
                objdbAccessConn.Open()
                objlocOLEdbcmd.CommandText = strSQL
                objlocOLEdbcmd.Connection = objdbAccessConn
                intAffected = objlocOLEdbcmd.ExecuteNonQuery()

            ElseIf strRGTableType = "M" Then
                'MySQL
                strSQL = "UPDATE " + strNameRGTable + " SET gebucht=true, gebuchtDatum=DATE('" + Format(datDate, "yyyy-MM-dd").ToString + "'), " + strBelegNrName + "=" + intBelegNr.ToString + " WHERE " + strRGNbrFieldName + "=" + strRGNbr
                objlocMySQLRGConn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(strMDBName)
                objlocMySQLRGConn.Open()
                objlocMySQLRGcmd.Connection = objlocMySQLRGConn
                objlocMySQLRGcmd.CommandText = strSQL
                intAffected = objlocMySQLRGcmd.ExecuteNonQuery()


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

    Public Shared Function FcExecuteBeforeDebit(ByVal intMandant As Integer, ByRef objMySQLConn As MySqlConnection) As Int16

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

    Public Shared Function FcExecuteAfterDebit(ByVal intMandant As Integer, ByRef objMySQLConn As MySqlConnection) As Int16

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

    Public Shared Function FcCheckDebiIntBank(ByRef objdbconn As MySqlConnection, ByVal intAccounting As Integer, ByVal striBankS50 As String, ByRef intIBankS200 As String) As Int16

        '0=ok, 1=Sage50 iBank nicht gefunden, 2=Kein Standard gesetzt, 3=Nichts angegeben, auf Standard gesetzt, 9=Problem

        Dim objdbcommand As New MySqlCommand
        Dim objdtiBank As New DataTable

        Try
            'wurde i Bank definiert?
            If striBankS50 <> "" Then
                'Sage 50 - Bank suchen
                objdbcommand.Connection = objdbconn
                If objdbconn.State = ConnectionState.Closed Then
                    objdbconn.Open()
                End If
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
                objdbcommand.CommandText = "SELECT intSage200 FROM tblaccoutningbank WHERE booStandard=true AND intAccountingID=" + intAccounting.ToString
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
            If objdbconn.State = ConnectionState.Open Then
                'objdbconn.Close()
            End If

        End Try

    End Function

    Public Shared Function FcSQLParse(ByVal strSQLToParse As String,
                                      ByVal strRGNbr As String,
                                      ByVal objdtBookings As DataTable,
                                      ByRef objOracleConn As OracleClient.OracleConnection,
                                      ByRef objOracleCommand As OracleClient.OracleCommand,
                                      ByVal strDebiCredit As String) As String

        'Funktion setzt in eingelesenem SQL wieder Variablen ein
        Dim intPipePositionBegin, intPipePositionEnd As Integer
        Dim strWork, strField As String
        Dim RowBooking() As DataRow

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
                        strField = FcGetKundenzeichen(RowBooking(0).Item("lngDebIdentNbr"), objOracleConn, objOracleCommand)
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

        Return strSQLToParse


    End Function

    Public Shared Function FcGetKundenzeichen(ByVal lngJournalNr As Int32, ByRef objOracleCon As OracleConnection, ByRef objOracleCmd As OracleCommand) As String

        Dim objdtJournalKZ As New DataTable

        objOracleCmd.CommandText = "SELECT KUNDENZEICHEN FROM TAB_JOURNALSTAMM WHERE JORNALNR=" + lngJournalNr.ToString
        objdtJournalKZ.Load(objOracleCmd.ExecuteReader)

        Return objdtJournalKZ.Rows(0).Item(0)

    End Function

    Public Shared Function FcIsPrivateDebitorCreatable(ByRef objdbconn As MySqlConnection,
                                             ByRef objdbconnZHDB02 As MySqlConnection,
                                             ByRef objsqlcommandZHDB02 As MySqlCommand,
                                             ByVal lngDebiNbr As Long,
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
        Dim strBankName As String = ""
        Dim strBankAddress1 As String = ""
        Dim strBankAddress2 As String = ""
        Dim strBankPLZ As String = ""
        Dim strBankOrt As String = ""
        Dim strBankBIC As String = ""
        Dim strBankCountry As String = ""
        Dim strBankClearing As String = ""
        Dim intReturnValue As Int16
        Dim intDebZB As Int16
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
            objsqlcommandZHDB02.CommandText = "SELECT Lastname, Firstname, Street, ZipCode, City, DebiKonto, 'Privatperson' AS Gruppe, IF(Country IS NULL, 'CH', country) AS country, Phone, Fax, Email, " +
                                                "IF(Language IS NULL, 'DE',Language) AS Language, BankName, BankZipCode, BankCountry, IBAN, BankBIC, " +
                                                "IF(Currency IS NULL, 'CHF', Currency) AS Currency, DebiKonto, '3200' AS ErloesKonto, BankIntern, DebiZKonditionID FROM t_customer WHERE PKNr=" + lngDebiNbr.ToString
            objdtDebitor.Load(objsqlcommandZHDB02.ExecuteReader)

            'Gefunden?
            If objdtDebitor.Rows.Count > 0 Then
                'Debug.Print("Gefunden, kann erstellt werden")

                'Sachbearbeiter suchen
                'Default setzen
                objsqlcommandZHDB02.CommandText = "SELECT CustomerID FROM t_rep_sagesachbearbeiter WHERE Rep_Nr=2535 AND Buchh_Nr=" + intAccounting.ToString
                objdtSachB.Load(objsqlcommandZHDB02.ExecuteReader)
                If objdtSachB.Rows.Count > 0 Then 'Default ist definiert
                    strSachB = Trim(objdtSachB.Rows(0).Item("CustomerID").ToString)
                Else
                    strSachB = ""
                    MessageBox.Show("Kein Sachbearbeiter - Default gesetzt für Buha " + strcmbBuha, "Debitorenerstellung")
                End If

                'interne Bank
                intReturnValue = Main.FcCheckDebiIntBank(objdbconn,
                                                         intAccounting,
                                                         objdtDebitor.Rows(0).Item("BankIntern"),
                                                         intintBank)


                'Zahlungsbedingung suchen
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
                If IIf(IsDBNull(objdtDebitor.Rows(0).Item("DebiZKonditionID")), 0, objdtDebitor.Rows(0).Item("DebiZKonditionID")) <> 0 Then
                    intDebZB = objdtDebitor.Rows(0).Item("DebiZKonditionID")
                Else
                    intDebZB = 1
                End If

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
                Select Case IIf(IsDBNull(objdtDebitor.Rows(0).Item("Language")), "DE", objdtDebitor.Rows(0).Item("Language"))
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
                strBankAddress1 = ""
                strBankPLZ = IIf(IsDBNull(objdtDebitor.Rows(0).Item("BankZipCode")), "", objdtDebitor.Rows(0).Item("BankZipCode"))
                strBankOrt = ""
                strBankAddress2 = strBankPLZ + " " + strBankOrt
                strBankBIC = IIf(IsDBNull(objdtDebitor.Rows(0).Item("BankBIC")), "", objdtDebitor.Rows(0).Item("BankBIC"))
                strBankClearing = ""

                If Len(strIBANNr) = 21 Then 'IBAN
                    'If intPayType <> 9 Then 'Type nicht IBAN angegeben aber IBAN - Nr. erfasst
                    intPayType = 9
                    'End If
                    intReturnValue = Main.FcGetIBANDetails(objdbconn,
                                                      strIBANNr,
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
                If objdtDebitor.Rows(0).Item("DebiKonto") = 1105 And lngDebiNbr >= 40000 Then
                    strCurrency = "EUR"
                Else
                    strCurrency = "CHF"
                End If

                intCreatable = FcCreateDebitor(objDbBhg,
                                              lngDebiNbr,
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("LastName")), "", objdtDebitor.Rows(0).Item("LastName")),
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("Street")), "", objdtDebitor.Rows(0).Item("Street")),
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("ZipCode")), "", objdtDebitor.Rows(0).Item("ZipCode")),
                                              IIf(IsDBNull(objdtDebitor.Rows(0).Item("City")), "", objdtDebitor.Rows(0).Item("City")),
                                              objdtDebitor.Rows(0).Item("DebiKonto"),
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

                    intCreatable = MainDebitor.FcWriteDatetoPrivate(objdbconn,
                                                             lngDebiNbr,
                                                             intAccounting,
                                                             0)


                End If


                Return 0
            Else
                Return 4

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Debitor - Erstellbar - Abklärung")
            Return 9

        Finally
            objdbconnZHDB02.Close()

        End Try

    End Function

    Public Shared Function FcWriteDatetoPrivate(ByRef objdbcon As MySqlConnection,
                                                   ByVal intNewPKNr As Int32,
                                                   ByVal intAccounting As Int16,
                                                   ByVal intDebitKredit As Int16) As Int16

        '0=ok, 1=PKNr nicht existent, 2=DS konnte nicht erstellt werden, 9=Problem

        Dim objdbCmd As New MySqlCommand
        Dim intAffected As Int16
        Dim strSQL As String
        Dim intRepNr As Int32
        Dim objdtPrivate As New DataTable
        Dim strDebiCreatedField As String

        Try

            If intDebitKredit = 0 Then
                strDebiCreatedField = "DebiCreatedPKOn"
            Else
                strDebiCreatedField = "CrediCreatedPKON"
            End If

            'Zuerst CustomerID suchen
            If objdbcon.State = ConnectionState.Closed Then
                objdbcon.Open()
            End If
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
            If objdbcon.State() = ConnectionState.Open Then
                objdbcon.Close()
            End If
            objdtPrivate.Dispose()
            objdbCmd.Dispose()

        End Try

    End Function

    Public Shared Function FcCheckLinkedRG(ByRef objDbBhg As SBSXASLib.AXiDbBhg,
                                           ByVal intNewDebiNbr As Int32,
                                           ByVal strDebiCur As String,
                                           ByVal intBelegNbr As Int32) As Int16

        'Returns 0=ok, 1=Beleg nicht existent, 9=Problem

        Dim intLaufNbr As Int32

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
                Return 0
            Else
                Return 1
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Fehler Prüfen Splitt-Bill Bel " + intBelegNbr.ToString)
            Return 9

        Finally

        End Try

    End Function

    Public Shared Function FcGetDebitorFromLinkedRG(ByRef objdbconn As MySqlConnection,
                                                    ByRef objdbconnZHDB02 As MySqlConnection,
                                                    ByRef objsqlcommand As MySqlCommand,
                                                    ByRef objsqlcommandZHDB02 As MySqlCommand,
                                                    ByRef objOrdbconn As OracleClient.OracleConnection,
                                                    ByRef objOrcommand As OracleClient.OracleCommand,
                                                    ByRef objdbAccessConn As OleDb.OleDbConnection,
                                                    ByVal lngRGNbr As Int32,
                                                    ByVal intAccounting As Int32,
                                                    ByRef intDebiNew As Int32) As Int16

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

        'Dim objlocOLEdbcmd As New OleDb.OleDbCommand
        'Dim strMDBName As String = Main.FcReadFromSettings(objdbconn, "Buchh_PKTableConnection", intAccounting)
        Dim strSQL As String
        'Dim intFunctionReturns As Int16

        Try

            strTableName = Main.FcReadFromSettings(objdbconn, "Buchh_TableDeb", intAccounting)
            strTableType = Main.FcReadFromSettings(objdbconn, "Buchh_RGTableType", intAccounting)
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
                    objdbConnDeb.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(Main.FcReadFromSettings(objdbconn, "Buchh_RGTableMDB", intAccounting))
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

            'Return intPKNewField

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Fehler Prüfen Splitt-Bill")
            Return 9

        Finally

        End Try

    End Function


End Class
