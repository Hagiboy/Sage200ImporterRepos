Option Strict Off
Option Explicit On

Imports MySql.Data.MySqlClient
Imports System.Data.OracleClient
Imports System.Data.SqlClient
Imports System.Net
Imports System.IO
Imports System.Xml
Imports Org.BouncyCastle.Crypto.Prng

Public Class MainKreditor


    Public Shared Function FcReadKreditorName(ByRef objKrBhg As SBSXASLib.AXiKrBhg, ByVal intKrediNbr As Int32, ByVal strCurrency As String) As String

        Dim strKreditorName As String
        Dim strKreditorAr() As String

        Try

            If strCurrency = "" Then

                strKreditorName = objKrBhg.ReadKreditor3(intKrediNbr * -1, strCurrency)

            Else

                strKreditorName = objKrBhg.ReadKreditor3(intKrediNbr, strCurrency)
                'strKreditorName = objKrBhg.ReadKreditor3(1, "CHF")
                'Call objKrBhg.ReadKrediStamm2()
                'Do Until strKreditorName = "EOF"
                '    strKreditorName = objKrBhg.GetKStammZeile3()
                'Loop
            End If

            strKreditorAr = Split(strKreditorName, "{>}")

            Return strKreditorAr(0)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "kreditor-Name " + Err.Number.ToString)

        End Try

    End Function

    Public Shared Function FcGetRefKrediNr(ByVal lngKrediNbr As Int32,
                                          ByVal intAccounting As Int32,
                                          ByRef intKrediNew As Int32) As Int16

        'Return 0=ok, 1=noch nicht implementiert, 2=Rep_Ref nicht definiert, 3=Nicht in Tab_Repbetriebe, 4=keine Angaben in Tab_Repbetriebe, 9=Problem

        Dim strTableName, strTableType, strKredFieldName, strKredNewField, strKredNewFieldType, strCompFieldName, strStreetFieldName, strZIPFieldName, strTownFieldName, strSageName, strKredAccField As String
        'Dim intCreatable As Int16
        Dim objdtKreditor As New DataTable
        Dim intPKNewField As Int32
        Dim objdbConnKred As New MySqlConnection
        Dim objsqlCommKred As New MySqlCommand

        Dim objdbAccessConn As OleDb.OleDbConnection
        Dim objlocOLEdbcmd As New OleDb.OleDbCommand
        Dim strMDBName As String = Main.FcReadFromSettingsII("Buchh_PKKrediTableConnection", intAccounting)
        Dim objOrcommand As OracleClient.OracleCommand
        Dim strSQL As String
        Dim intFunctionReturns As Int16

        Try

            strTableName = Main.FcReadFromSettingsII("Buchh_PKKrediTable", intAccounting)
            strTableType = Main.FcReadFromSettingsII("Buchh_PKKrediTableType", intAccounting)
            strKredFieldName = Main.FcReadFromSettingsII("Buchh_PKKrediField", intAccounting)
            strKredNewField = Main.FcReadFromSettingsII("Buchh_PKKrediNewField", intAccounting)
            strKredNewFieldType = Main.FcReadFromSettingsII("Buchh_PKKrediNewFType", intAccounting)
            'strCompFieldName = Main.FcReadFromSettingsII("Buchh_PKKrediCompany", intAccounting)
            'strStreetFieldName = Main.FcReadFromSettingsII("Buchh_PKKrediStreet", intAccounting)
            'strZIPFieldName = Main.FcReadFromSettingsII("Buchh_PKKrediZIP", intAccounting)
            'strTownFieldName = Main.FcReadFromSettingsII("Buchh_PKKrediTown", intAccounting)
            'strSageName = Main.FcReadFromSettingsII("Buchh_PKKrediSageName", intAccounting)
            'strKredAccField = Main.FcReadFromSettingsII("Buchh_PKKrediAccount", intAccounting)

            strSQL = "SELECT * " + 'strKredFieldName + ", " + strKredNewField + ", " + strCompFieldName + ", " + strStreetFieldName + ", " + strZIPFieldName + ", " + strTownFieldName + ", " + strSageName + ", " + strKredAccField +
                 " FROM " + strTableName + " WHERE " + strKredFieldName + "=" + lngKrediNbr.ToString

            If strTableName <> "" And strKredFieldName <> "" Then

                If strTableType = "O" Then 'Oracle
                    'objOrdbconn.Open()
                    objOrcommand.CommandText = strSQL
                    objdtKreditor.Load(objOrcommand.ExecuteReader)
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
                    Call Main.FcInitAccessConnecation(objdbAccessConn, strMDBName)
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
                            intPKNewField = Main.FcGetPKNewFromRep(IIf(IsDBNull(objdtKreditor.Rows(0).Item("ID")), 0, objdtKreditor.Rows(0).Item("ID")),
                                                                       "P")
                        Else
                            intPKNewField = Main.FcGetPKNewFromRep(objdtKreditor.Rows(0).Item(strKredNewField),
                                                                        "R") 'Rep_Nr
                            Stop
                        End If

                        If intPKNewField = 0 Then
                            'PK wurde nicht vergeben => Eine neue erzeugen und in der Tabelle Rep_Betriebe 
                            If strTableName = "t_customer" Then
                                intFunctionReturns = Main.FcNextPrivatePKNr(objdtKreditor.Rows(0).Item("ID"),
                                                                            intKrediNew)
                                If intFunctionReturns = 0 And intKrediNew > 0 Then 'Vergabe hat geklappt
                                    intFunctionReturns = Main.FcWriteNewPrivateDebToRepbetrieb(objdtKreditor.Rows(0).Item("ID"),
                                                                                                   intKrediNew)
                                    If intFunctionReturns = 0 Then 'Schreiben hat geklappt
                                        Return 1
                                    End If
                                End If
                            Else
                                intFunctionReturns = Main.FcNextPKNr(objdtKreditor.Rows(0).Item(strKredNewField),
                                                                         intKrediNew,
                                                                         intAccounting,
                                                                         "C")
                                If intFunctionReturns = 0 And intKrediNew > 0 Then 'Vergabe hat geklappt
                                    intFunctionReturns = Main.FcWriteNewDebToRepbetrieb(objdtKreditor.Rows(0).Item("Rep_Nr"),
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
                            intFunctionReturns = Main.FcNextPKNr(objdtKreditor.Rows(0).Item("Rep_Nr"),
                                                                    intKrediNew,
                                                                    intAccounting,
                                                                    "C")
                            If intFunctionReturns = 0 And intKrediNew > 0 Then 'Vergabe hat geklappt
                                intFunctionReturns = Main.FcWriteNewDebToRepbetrieb(objdtKreditor.Rows(0).Item("Rep_Nr"),
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
            objdtKreditor.Constraints.Clear()
            objdtKreditor.Rows.Clear()
            objdtKreditor.Columns.Clear()
            objdtKreditor.Dispose()
            objdtKreditor = Nothing

        End Try


    End Function

    Public Shared Function FcIsPrivateKreditorCreatable(ByVal lngKrediNbr As Long,
                                                ByRef objKrBhg As SBSXASLib.AXiKrBhg,
                                                ByRef objFiBhg As SBSXASLib.AXiFBhg,
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
                    intReturnValue = Main.FcCheckKonto(objdtKreditor.Rows(0).Item("KrediAufwandskonto"),
                                                       objFiBhg,
                                                       0,
                                                       0,
                                                       True)
                    If intReturnValue <> 0 Then
                        booReadAufwandsKono = Convert.ToBoolean(Convert.ToInt16(Main.FcReadFromSettingsII("Buchh_KrediTakeAufwKto", intAccounting)))
                        If booReadAufwandsKono Then
                            'Zu nehmendes Aufwandskonto einlesen
                            intAufwandsKonto = Main.FcReadFromSettingsII("Buchh_KrediAufwKto", intAccounting)
                            objdtKreditor.Rows(0).Item("Rep_Kred_Aufwandskonto") = intAufwandsKonto
                            'Prüfen ob dieses Konto existiert
                            intReturnValue = Main.FcCheckKonto(objdtKreditor.Rows(0).Item("KrediAufwandskonto"),
                                                       objFiBhg,
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
                    intReturnValue = Main.FcCheckDebiIntBank(intAccounting,
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

                'QR-IBAN
                If intPayType = 10 And Len(strKrediBank) >= 21 Then
                    strIBANNr = strKrediBank
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

                intCreatable = FcCreateKreditor(objKrBhg,
                                          lngKrediNbr,
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

                    intCreatable = MainDebitor.FcWriteDatetoPrivate(lngKrediNbr,
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

        End Try

    End Function


    Public Shared Function FcIsKreditorCreatable(ByVal lngKrediNbr As Long,
                                                ByRef objKrBhg As SBSXASLib.AXiKrBhg,
                                                ByRef objFiBhg As SBSXASLib.AXiFBhg,
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
                    intReturnValue = Main.FcCheckKonto(objdtKreditor.Rows(0).Item("Rep_Kred_Aufwandskonto"),
                                                       objFiBhg,
                                                       0,
                                                       0,
                                                       True)
                    If intReturnValue <> 0 Then
                        booReadAufwandsKono = Convert.ToBoolean(Convert.ToInt16(Main.FcReadFromSettingsII("Buchh_KrediTakeAufwKto", intAccounting)))
                        If booReadAufwandsKono Then
                            'Zu nehmendes Aufwandskonto einlesen
                            intAufwandsKonto = Main.FcReadFromSettingsII("Buchh_KrediAufwKto", intAccounting)
                            objdtKreditor.Rows(0).Item("Rep_Kred_Aufwandskonto") = intAufwandsKonto
                            'Prüfen ob dieses Konto existiert
                            intReturnValue = Main.FcCheckKonto(objdtKreditor.Rows(0).Item("Rep_Kred_Aufwandskonto"),
                                                       objFiBhg,
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

                    'QR-IBAN
                    If intPayType = 10 And Len(strKrediBank) >= 21 Then
                        strIBANNr = strKrediBank
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

                    intCreatable = FcCreateKreditor(objKrBhg,
                                          lngKrediNbr,
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
            Return 9

        Finally
            objdbconnZHDB02.Close()

        End Try

    End Function

    Public Shared Function FcCreateKreditor(ByRef objKrBhg As SBSXASLib.AXiKrBhg,
                                           ByVal intKreditorNewNbr As Int32,
                                           ByVal strKredName As String,
                                           ByVal strKredStreet As String,
                                           ByVal strKredPLZ As String,
                                           ByVal strKredOrt As String,
                                           ByVal intKredSammelKto As Int32,
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
                                           ByVal intAufwandsKonto As Int16,
                                           ByVal intKredZB As Int16,
                                           ByVal intintBank As Int16,
                                           ByVal strFirstName As String) As Int16

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


            Call objKrBhg.SetCommonInfo2(intKreditorNewNbr,
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

            Call objKrBhg.SetExtendedInfo7(strKredSperren,
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

                        Call objKrBhg.SetZahlungsverbindung("B",
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
                    Else
                        'Typ ist 10 (=QR)
                        Call objKrBhg.SetZahlungsverbindung("Q",
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
            End If

            If intPayDefault = 10 Then 'QR - IBAN

                Call objKrBhg.SetZahlungsverbindung("Q",
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

            Call objKrBhg.WriteKreditor3(intintBank.ToString, 0)

            Return 0
            'intDebAdrLaufN = DbBhg.GetAdressLaufnr()
            'intDebBankLaufNr = DbBhg.GetZahlungsverbLaufnr()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem beim Anlegen Kreditor " + intKreditorNewNbr.ToString + ", " + strKredName)

            Return 1

        End Try

    End Function

    Public Shared Function FcWriteToKrediRGTable(ByVal intMandant As Int32,
                                                 ByVal strKredID As String,
                                                 ByVal datDate As Date,
                                                 ByVal intBelegNr As Int32,
                                                 ByRef objdbAccessConn As OleDb.OleDbConnection,
                                                 ByRef objOracleConn As OracleConnection,
                                                 ByRef objMySQLConn As MySqlConnection) As Int16

        'Returns 0=ok, 1=Problem

        Dim strSQL As String
        Dim intAffected As Int16
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

        strMDBName = Main.FcReadFromSettingsII("Buchh_KRGTableMDB", intMandant)
        strKRGTableType = Main.FcReadFromSettingsII("Buchh_KRGTableType", intMandant)
        strNameKRGTable = Main.FcReadFromSettingsII("Buchh_TableKred", intMandant)
        strBelegNrName = Main.FcReadFromSettingsII("Buchh_TableKRGBelegNrName", intMandant)
        strKRGNbrFieldName = Main.FcReadFromSettingsII("Buchh_TableKRGNbrFieldName", intMandant)
        strKRGNbrFieldType = Main.FcReadFromSettingsII("Buchh_TableKRGNbrFieldType", intMandant)
        'strSQL = "UPDATE " + strNameRGTable + " SET gebucht=true, gebuchtDatum=#" + Format(datDate, "yyyy-MM-dd").ToString + "#, " + strBelegNrName + "=" + intBelegNr.ToString + " WHERE " + strRGNbrFieldName + "=" + strRGNbr

        Try

            If strKRGTableType = "A" Then
                'Access
                Call Main.FcInitAccessConnecation(objdbAccessConn, strMDBName)
                strSQL = "UPDATE " + strNameKRGTable + " SET Kredigebucht=true, KredigebuchtDatum=#" + Format(datDate, "yyyy-MM-dd").ToString + "#, " + strBelegNrName + "='" + intBelegNr.ToString + "' WHERE " + strKRGNbrFieldName + "=" + IIf(strKRGNbrFieldType = "T", "'", "") + strKredID + IIf(strKRGNbrFieldType = "T", "'", "")

                objdbAccessConn.Open()
                objlocOLEdbcmd.CommandText = strSQL
                objlocOLEdbcmd.Connection = objdbAccessConn
                intAffected = objlocOLEdbcmd.ExecuteNonQuery()

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
            If objdbAccessConn.State = ConnectionState.Open Then
                objdbAccessConn.Close()
            End If

            If objlocMySQLRGConn.State = ConnectionState.Open Then
                objlocMySQLRGConn.Close()
            End If

            'If objMySQLConn.State = ConnectionState.Open Then
            '    objMySQLConn.Close()
            'End If

        End Try

    End Function

    Public Shared Function FcFillKredit(ByVal intAccounting As Integer,
                                       ByRef objdtHead As DataTable,
                                       ByRef objdtSub As DataTable,
                                       ByRef objdbconn As MySqlConnection,
                                       ByRef objdbAccessConn As OleDb.OleDbConnection) As Integer

        '0=ok, 1=Keine Defintion, 9=Problem

        Dim strSQL As String
        Dim strSQLSub As String
        Dim strKRGTableType As String
        Dim objRGMySQLConn As New MySqlConnection
        Dim objlocMySQLcmd As New MySqlCommand
        Dim objlocOLEdbcmd As New OleDb.OleDbCommand

        Dim objDTDebiHead As New DataTable
        Dim strMDBName As String
        Dim objdrSub As DataRow
        Dim intFcReturns As Int16

        objdbconn.Open()

        strMDBName = Main.FcReadFromSettings(objdbconn, "Buchh_KRGTableMDB", intAccounting)

        'Head Debitzoren löschen
        objdtHead.Clear()
        objdtHead.Constraints.Clear()
        strSQL = Main.FcReadFromSettings(objdbconn, "Buchh_SQLHeadKred", intAccounting)
        strKRGTableType = Main.FcReadFromSettings(objdbconn, "Buchh_KRGTableType", intAccounting)

        Try

            If IsDBNull(strSQL) Or strSQL = "" Then
                Return 1
            Else
                'objlocMySQLcmd.CommandText = strSQL
                If strKRGTableType = "A" Then
                    'Access
                    Call Main.FcInitAccessConnecation(objdbAccessConn, strMDBName)

                    objlocOLEdbcmd.CommandText = strSQL
                    objdbAccessConn.Open()
                    objlocOLEdbcmd.Connection = objdbAccessConn
                    objdtHead.Load(objlocOLEdbcmd.ExecuteReader)
                ElseIf strKRGTableType = "M" Then
                    'MySQL
                    objRGMySQLConn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(strMDBName)
                    objlocMySQLcmd.Connection = objRGMySQLConn
                    objlocMySQLcmd.CommandText = strSQL
                    objRGMySQLConn.Open()
                    objdtHead.Load(objlocMySQLcmd.ExecuteReader)
                End If
                'objlocMySQLcmd.Connection = objdbconn
                'objDTDebiHead.Load(objlocMySQLcmd.ExecuteReader)
                'Durch die Records steppen und Sub-Tabelle füllen

                For Each row As DataRow In objdtHead.Rows
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
                    strSQLSub = FcSQLParseKredi(Main.FcReadFromSettings(objdbconn, "Buchh_SQLDetailKred", intAccounting), row("lngKredID"), objdtHead)
                    If strKRGTableType = "A" Then
                        objlocOLEdbcmd.CommandText = strSQLSub
                        objdtSub.Load(objlocOLEdbcmd.ExecuteReader)
                    ElseIf strKRGTableType = "M" Then
                        objlocMySQLcmd.CommandText = strSQLSub
                        objdtSub.Load(objlocMySQLcmd.ExecuteReader)
                        'Debug.Print("Ok, " + strSQLSub)
                    End If
                Next
                Return 0
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return 9

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

    Public Shared Function FcCheckKreditor(ByVal lngKreditor As Long, ByVal intBuchungsart As Integer, ByRef objKrBuha As SBSXASLib.AXiKrBhg) As Integer

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

        End Try

    End Function

    Public Shared Function FcChCeckKredOP(ByRef strOPNbr As String, ByVal strKredRGNbr As String) As Int16

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

            Return 9

        End Try


    End Function

    Public Shared Function FcGetKZkondFromCust(ByVal lngKrediiNbr As Long,
                                              ByRef intDZkond As Int16,
                                              ByVal intAccounting As Int16) As Int16

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
            objdsDebitor.Dispose()
            objDADebitor.Dispose()
            'Application.DoEvents()

        End Try


    End Function


    Public Shared Function FcCheckKreditBank(ByVal objKrBhg As SBSXASLib.AXiKrBhg,
                                         ByVal intKreditor As Int32,
                                         ByVal intPayType As Int16,
                                         ByVal strIBAN As String,
                                         ByVal strBank As String,
                                         ByVal strKredCur As String,
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

                Call objKrBhg.ReadZahlungsverb(intKreditor * -1)

                Do Until strZahlVerbindungLine = "EOF"

                    strZahlVerbindungLine = objKrBhg.GetZahlungsverbZeile()
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

                    intReturnValue = Main.FcGetIBANDetails(strIBAN,
                                                      strBankName,
                                                      strBankAddress1,
                                                      strBankAddress2,
                                                      strBankBIC,
                                                      strBankCountry,
                                                      strBankClearing)

                    If intReturnValue = 0 Then 'Angaben vollständig und kein Problem
                        'Kombinierte PLZ / Ort Feld trennen
                        strBankPLZ = Left(strBankAddress2, InStr(strBankAddress2, " "))
                        strBankOrt = Trim(Right(strBankAddress2, Len(strBankAddress2) - InStr(strBankAddress2, " ")))

                        'Evtl Typ falsch gesetzt?
                        If Strings.Mid(strIBAN, 5, 1) <> "3" Or Strings.Left(strIBAN, 2) <> "CH" Then
                            'IBAN
                            Call objKrBhg.WriteBank2(intKreditor,
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
                            Call objKrBhg.WriteBank2(intKreditor,
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

    Public Shared Function FcSQLParseKredi(ByVal strSQLToParse As String,
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
                    strSQLToParse = Left(strSQLToParse, intPipePositionBegin - 1) + IIf(strFieldType = "T", "'", "") + strField + IIf(strFieldType = "T", "'", "") + Right(strSQLToParse, Len(strSQLToParse) - intPipePositionEnd)
                    'Neuer Anfang suchen für evtl. weitere |
                    intPipePositionBegin = InStr(strSQLToParse, "|")
                    'intPipePositionBegin = InStr(intPipePositionEnd + 1, strSQLToParse, "|")
                    intPipePositionEnd = InStr(intPipePositionBegin + 1, strSQLToParse, "|")
                Loop
            End If

            Return strSQLToParse

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Fehler Parsing " + Err.Number.ToString)

        End Try




    End Function

    Public Shared Function FcCheckKrediOPDouble(ByRef objKrBuha As SBSXASLib.AXiKrBhg,
                                                ByVal strKreditor As String,
                                                ByVal strOPNr As String,
                                                ByVal strKredCurrency As String,
                                                ByVal strKredTyp As String) As Int16

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

    Public Shared Function FcCheckKrediExistance(ByRef objdbMSSQLConn As SqlConnection,
                                                 ByRef objdbMSSQLCmd As SqlCommand,
                                                 ByRef intBelegNbr As Int32,
                                                 ByVal strTyp As String,
                                                 ByVal intTeqNr As Int32,
                                                 ByVal intTeqNrLY As Int32,
                                                 ByVal intTeqNrPLY As Int32,
                                                 ByRef objKrBhg As SBSXASLib.AXiKrBhg) As Int16

        '0=ok, 1=Beleg existierte schon, 9=Problem

        'Prinzip: in Tabelle kredibuchung suchen da API - Funktion nur in spezifischen Kreditor sucht

        Dim intReturnvalue As Int32
        Dim intStatus As Int16
        Dim tblKrediBeleg As New DataTable
        Dim intEntryBelNbr As Int32 = intBelegNbr

        Try

            'Prüfung
            intReturnvalue = 10
            intStatus = 0

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
            Return 9

        Finally
            objdbMSSQLConn.Close()
            tblKrediBeleg.Constraints.Clear()
            tblKrediBeleg.Dispose()
            tblKrediBeleg = Nothing

        End Try


    End Function

    Public Shared Function FcPGVKTreatmentYC(ByRef objFBhg As SBSXASLib.AXiFBhg,
                                                ByRef objFinanz As SBSXASLib.AXFinanz,
                                                ByRef objDbBhg As SBSXASLib.AXiDbBhg,
                                                ByRef objPiFin As SBSXASLib.AXiPlFin,
                                                ByRef objBebu As SBSXASLib.AXiBeBu,
                                                ByRef objKrBhg As SBSXASLib.AXiKrBhg,
                                                ByVal tblKrediB As DataTable,
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
                            dblKursD = Main.FcGetKurs(strCur, strValutaDatum, objFBhg)
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
                        intSollKonto = intAcctTY
                        intHabenKonto = intAcctNY
                        strDebiTextSoll = drKSubrow("strKredSubText") + ", PGV AJ / FJ"
                        strDebiTextHaben = drKSubrow("strKredSubText") + ", PGV AJ / FJ"
                        strBebuEintragSoll = Nothing
                        strBebuEintragHaben = Nothing

                        'Buchen
                        Call objFBhg.WriteBuchung(0,
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
            MessageBox.Show(ex.Message, "Problem PGV - Buchung Kreditoren", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 9

        Finally

        End Try

    End Function


    Public Shared Function FcPGVKTreatment(ByRef objFBhg As SBSXASLib.AXiFBhg,
                                                ByRef objFinanz As SBSXASLib.AXFinanz,
                                                ByRef objDbBhg As SBSXASLib.AXiDbBhg,
                                                ByRef objPiFin As SBSXASLib.AXiPlFin,
                                                ByRef objBebu As SBSXASLib.AXiBeBu,
                                                ByRef objKrBhg As SBSXASLib.AXiKrBhg,
                                                ByVal tblKrediB As DataTable,
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
                            dblKursD = Main.FcGetKurs(strCur, strValutaDatum, objFBhg)
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
                        intSollKonto = intAcctTY
                        intHabenKonto = intAcctNY
                        strDebiTextSoll = drKSubrow("strKredSubText") + ", PGV AJ / FJ"
                        strDebiTextHaben = drKSubrow("strKredSubText") + ", PGV AJ / FJ"
                        strBebuEintragSoll = Nothing
                        strBebuEintragHaben = Nothing

                        'Buchen
                        Call objFBhg.WriteBuchung(0,
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
                    dblKursD = Main.FcGetKurs(strCur, strValutaDatum, objFBhg)
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
            MessageBox.Show(ex.Message, "Problem PGV - Buchung Kreditoren", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 9

        Finally

        End Try

    End Function

    Public Shared Function FcIsAllKrediRebilled(ByVal objdbKrediSub As DataTable,
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

End Class