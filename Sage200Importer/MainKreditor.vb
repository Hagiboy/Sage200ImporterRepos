Option Strict Off
Option Explicit On

Imports MySql.Data.MySqlClient
Imports System.Data.OracleClient
Imports System.Net
Imports System.IO
Imports System.Xml

Public Class MainKreditor


    Public Shared Function FcReadKreditorName(ByRef objKrBhg As SBSXASLib.AXiKrBhg, ByVal intKrediNbr As Int32, ByVal strCurrency As String) As String

        Dim strKreditorName As String
        Dim strKreditorAr() As String

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

    End Function

    Public Shared Function FcGetRefKrediNr(ByRef objdbconn As MySqlConnection,
                                          ByRef objdbconnZHDB02 As MySqlConnection,
                                          ByRef objsqlcommand As MySqlCommand,
                                          ByRef objsqlcommandZHDB02 As MySqlCommand,
                                          ByRef objOrdbconn As OracleClient.OracleConnection,
                                          ByRef objOrcommand As OracleClient.OracleCommand,
                                          ByRef objdbAccessConn As OleDb.OleDbConnection,
                                          ByVal lngKrediNbr As Int32,
                                          ByVal intAccounting As Int32,
                                          ByRef intKrediNew As Int32) As Int16

        'Return 0=ok, 1=noch nicht implementiert, 2=Rep_Ref nicht definiert, 3=Nicht in Tab_Repbetriebe, 4=keine Angaben in Tab_Repbetriebe

        Dim strTableName, strTableType, strKredFieldName, strKredNewField, strKredNewFieldType, strCompFieldName, strStreetFieldName, strZIPFieldName, strTownFieldName, strSageName, strKredAccField As String
        'Dim intCreatable As Int16
        Dim objdtKreditor As New DataTable
        Dim intPKNewField As Int32
        Dim objdbConnKred As New MySqlConnection
        Dim objsqlCommKred As New MySqlCommand

        Dim objlocOLEdbcmd As New OleDb.OleDbCommand
        Dim strMDBName As String = Main.FcReadFromSettings(objdbconn, "Buchh_PKKrediTableConnection", intAccounting)
        Dim strSQL As String
        Dim intFunctinReturns As Int16

        strTableName = Main.FcReadFromSettings(objdbconn, "Buchh_PKKrediTable", intAccounting)
        strTableType = Main.FcReadFromSettings(objdbconn, "Buchh_PKKrediTableType", intAccounting)
        strKredFieldName = Main.FcReadFromSettings(objdbconn, "Buchh_PKKrediField", intAccounting)
        strKredNewField = Main.FcReadFromSettings(objdbconn, "Buchh_PKKrediNewField", intAccounting)
        strKredNewFieldType = Main.FcReadFromSettings(objdbconn, "Buchh_PKKrediNewFType", intAccounting)
        strCompFieldName = Main.FcReadFromSettings(objdbconn, "Buchh_PKKrediCompany", intAccounting)
        strStreetFieldName = Main.FcReadFromSettings(objdbconn, "Buchh_PKKrediStreet", intAccounting)
        strZIPFieldName = Main.FcReadFromSettings(objdbconn, "Buchh_PKKrediZIP", intAccounting)
        strTownFieldName = Main.FcReadFromSettings(objdbconn, "Buchh_PKKrediTown", intAccounting)
        strSageName = Main.FcReadFromSettings(objdbconn, "Buchh_PKKrediSageName", intAccounting)
        strKredAccField = Main.FcReadFromSettings(objdbconn, "Buchh_PKKrediAccount", intAccounting)

        strSQL = "SELECT " + strKredFieldName + ", " + strKredNewField + ", " + strCompFieldName + ", " + strStreetFieldName + ", " + strZIPFieldName + ", " + strTownFieldName + ", " + strSageName + ", " + strKredAccField +
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
                objdbConnKred.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(Main.FcReadFromSettings(objdbconn, "Buchh_PKKrediTableConnection", intAccounting))
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
                If IsDBNull(objdtKreditor.Rows(0).Item(strKredNewField)) And strTableName <> "Tab_Repbetriebe" Then
                    intKrediNew = 0
                    Return 2
                Else

                    If strTableName <> "Tab_Repbetriebe" Then
                        'intPKNewField = objdtKreditor.Rows(0).Item(strKredNewField)
                        intPKNewField = Main.FcGetPKNewFromRep(objdbconnZHDB02, objsqlcommandZHDB02, objdtKreditor.Rows(0).Item(strKredNewField)) 'Rep_Nr
                        If intPKNewField = 0 Then
                            'PK wurde nicht vergeben => Eine neue erzeugen und in der Tabelle Rep_Betriebe 
                            intFunctinReturns = Main.FcNextPKNr(objdbconnZHDB02, objdtKreditor.Rows(0).Item(strKredNewField), intKrediNew)
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
                            intFunctinReturns = Main.FcNextPKNr(objdbconnZHDB02, objdtKreditor.Rows(0).Item(strKredNewField), intKrediNew)
                        End If
                        Return 0
                    End If
                End If
            Else
                intKrediNew = 0
                Return 4
            End If

        End If

        Return intPKNewField

    End Function

    Public Shared Function FcIsKreditorCreatable(ByRef objdbconn As MySqlConnection,
                                                ByRef objdbconnZHDB02 As MySqlConnection,
                                                ByRef objsqlcommandZHDB02 As MySqlCommand,
                                                ByVal lngKrediNbr As Long,
                                                ByRef objKrBhg As SBSXASLib.AXiKrBhg,
                                                ByVal strcmbBuha As String,
                                                ByRef intPayType As Int16,
                                                ByVal strIBANFromInv As String) As Int16

        'Return: 0=creatable und erstellt, 3=Kreditor konnte nicht erstellt werden, 4=Betrieb nicht gefunden, 9=Nicht hinterlegt

        Dim intCreatable As Int16
        Dim objdtKreditor As New DataTable
        Dim objdtKredZB As New DataTable
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
        Dim intKredZB As Int16
        Dim objdsKreditor As New DataSet
        Dim objDAKreditor As New MySqlDataAdapter

        Try

            'Angaben einlesen
            objdbconnZHDB02.Open()
            objsqlcommandZHDB02.CommandText = "SELECT Rep_Firma, Rep_Strasse, Rep_PLZ, Rep_Ort, Rep_KredGegenKonto, Rep_Gruppe, Rep_Vertretung, Rep_Ansprechpartner, Rep_Land, Rep_Tel1, Rep_Fax, Rep_Mail, " +
                                                "Rep_Language, Rep_Kredi_MWSTNr, Rep_Kreditlimite, Rep_Kred_Pay_Def, Rep_Kred_Bank_Name, Rep_Kred_Bank_PLZ, Rep_Kred_Bank_Ort, Rep_Kred_IBAN, Rep_Kred_Bank_BIC, " +
                                                "Rep_Kred_Currency, Rep_Kred_PCKto, Rep_Kred_Aufwandskonto FROM Tab_Repbetriebe WHERE PKNr=" + lngKrediNbr.ToString
            objsqlcommandZHDB02.Connection = objdbconnZHDB02
            objdtKreditor.Load(objsqlcommandZHDB02.ExecuteReader)

            'Gefunden?
            If objdtKreditor.Rows.Count > 0 Then
                'Debug.Print("Gefunden, kann erstellt werden")

                'Zahlungsbedingung suchen
                'objdtKreditor.Clear()
                'Es muss der Weg über ein Dataset genommen werden da sosnt constraint-Meldungen kommen
                objsqlcommandZHDB02.CommandText = "SELECT Tab_Repbetriebe.PKNr, t_sage_zahlungskondition.SageID " +
                                                  "FROM Tab_Repbetriebe INNER JOIN t_sage_zahlungskondition ON Tab_Repbetriebe.Rep_Kred_ZKonditionID = t_sage_zahlungskondition.ID " +
                                                  "WHERE Tab_Repbetriebe.PKNr=" + lngKrediNbr.ToString
                objDAKreditor.SelectCommand = objsqlcommandZHDB02
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
                Select Case IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Language")), "D", objdtKreditor.Rows(0).Item("Rep_Language"))
                    Case "D", ""
                        intLangauage = 2055
                    Case "F"
                        intLangauage = 4108
                    Case "I"
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
                                          intKredZB)

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
                Else
                    Return 4
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
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
                                           ByVal intKredZB As Int16) As Int16

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
        Dim strText As String = ""
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
                                         "",
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
                End If
            End If
            Call objKrBhg.WriteKreditor3(0)

            Return 0
            'intDebAdrLaufN = DbBhg.GetAdressLaufnr()
            'intDebBankLaufNr = DbBhg.GetZahlungsverbLaufnr()

        Catch ex As Exception
            MessageBox.Show(ex.Message)

            Return 1

        End Try

    End Function

    Public Shared Function FcWriteToKrediRGTable(ByVal intMandant As Int32,
                                                 ByVal lngKredID As Int32,
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
        Dim strMDBName As String

        objMySQLConn.Open()

        strMDBName = Main.FcReadFromSettings(objMySQLConn, "Buchh_KRGTableMDB", intMandant)
        strKRGTableType = Main.FcReadFromSettings(objMySQLConn, "Buchh_KRGTableType", intMandant)
        strNameKRGTable = Main.FcReadFromSettings(objMySQLConn, "Buchh_TableKred", intMandant)
        strBelegNrName = Main.FcReadFromSettings(objMySQLConn, "Buchh_TableKRGBelegNrName", intMandant)
        strKRGNbrFieldName = Main.FcReadFromSettings(objMySQLConn, "Buchh_TableKRGNbrFieldName", intMandant)
        'strSQL = "UPDATE " + strNameRGTable + " SET gebucht=true, gebuchtDatum=#" + Format(datDate, "yyyy-MM-dd").ToString + "#, " + strBelegNrName + "=" + intBelegNr.ToString + " WHERE " + strRGNbrFieldName + "=" + strRGNbr

        Try

            If strKRGTableType = "A" Then
                'Access
                Call Main.FcInitAccessConnecation(objdbAccessConn, strMDBName)
                strSQL = "UPDATE " + strNameKRGTable + " SET Kredigebucht=true, KredigebuchtDatum=#" + Format(datDate, "yyyy-MM-dd").ToString + "#, " + strBelegNrName + "='" + intBelegNr.ToString + "' WHERE " + strKRGNbrFieldName + "=" + lngKredID.ToString

                objdbAccessConn.Open()
                objlocOLEdbcmd.CommandText = strSQL
                objlocOLEdbcmd.Connection = objdbAccessConn
                intAffected = objlocOLEdbcmd.ExecuteNonQuery()

            ElseIf strKRGTableType = "M" Then
                'MySQL
                'Bei IG andere Feldnamen
                If intMandant = 25 Then
                    strSQL = "UPDATE " + strNameKRGTable + " SET IGKBooked=true, IGKBDate=DATE('" + Format(datDate, "yyyy-MM-dd").ToString + "'), " + strBelegNrName + "='" + intBelegNr.ToString + "' WHERE " + strKRGNbrFieldName + "=" + lngKredID.ToString
                Else
                    strSQL = "UPDATE " + strNameKRGTable + " SET Kredigebucht=true, KredigebuchtDatum=DATE('" + Format(datDate, "yyyy-MM-dd").ToString + "'), " + strBelegNrName + "='" + intBelegNr.ToString + "' WHERE " + strKRGNbrFieldName + "=" + lngKredID.ToString
                End If

                objlocMySQLRGConn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(strMDBName)
                objlocMySQLRGConn.Open()
                objlocMySQLRGcmd.Connection = objlocMySQLRGConn
                objlocMySQLRGcmd.CommandText = strSQL
                intAffected = objlocMySQLRGcmd.ExecuteNonQuery()


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

            If objMySQLConn.State = ConnectionState.Open Then
                objMySQLConn.Close()
            End If

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

    Public Shared Function FcCheckKreditBank(ByRef objdbcon As MySqlConnection,
                                         ByRef objdbconnZHDB02 As MySqlConnection,
                                         ByVal objKrBhg As SBSXASLib.AXiKrBhg,
                                         ByVal intKreditor As Int32,
                                         ByVal intPayType As Int16,
                                         ByVal strIBAN As String,
                                         ByVal strBank As String,
                                         ByVal strKredCur As String) As Int16

        'Falls Typetype 9 (IBAN) ist, dann Zahlungsverbindungen prüfen

        Dim strZahlVerbindungLine As String = ""
        Dim strZahlVerbindung() As String
        Dim booBankExists As Boolean = False
        Dim intReturnValue As Int16
        Dim strBankName As String = ""
        Dim strBankAddress1 As String = ""
        Dim strBankAddress2 As String = ""
        Dim strBankCountry As String = ""
        Dim strBankBIC As String = ""
        Dim strBankClearing As String = ""
        Dim strBankOrt As String = ""
        Dim strBankPLZ As String = ""

        Try

            If intPayType = 9 Then 'IBAN

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
                            End If

                        End If
                    End If
                Loop

                If Not booBankExists Then
                    'MessageBox.Show("Bankverbindung muss erstellt werden " + strIBAN)

                    intReturnValue = Main.FcGetIBANDetails(objdbcon,
                                                      strIBAN,
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
                        Return 0
                    End If

                End If

            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return 9

        End Try

    End Function

    Public Shared Function FcSQLParseKredi(ByVal strSQLToParse As String, ByVal lngKredID As Int32, ByVal objdtKredi As DataTable) As String

        'Funktion setzt in eingelesenem SQL wieder Variablen ein
        Dim intPipePositionBegin, intPipePositionEnd As Integer
        Dim strWork, strField As String
        Dim RowKredi() As DataRow

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
                strSQLToParse = Left(strSQLToParse, intPipePositionBegin - 1) & strField & Right(strSQLToParse, Len(strSQLToParse) - intPipePositionEnd)
                'Neuer Anfang suchen für evtl. weitere |
                intPipePositionBegin = InStr(strSQLToParse, "|")
                'intPipePositionBegin = InStr(intPipePositionEnd + 1, strSQLToParse, "|")
                intPipePositionEnd = InStr(intPipePositionBegin + 1, strSQLToParse, "|")
            Loop
        End If

        Return strSQLToParse


    End Function

    Public Shared Function FcCheckKrediOPDouble(ByRef objKrBuha As SBSXASLib.AXiKrBhg, ByVal strKreditor As String, ByVal strOPNr As String) As Int16

        'Return 0=ok, 1=Beleg existiert, 9=Problem

        Dim intBelegReturn As Int16

        Try
            intBelegReturn = objKrBuha.doesBelegExist(strKreditor, "CHF", strOPNr, "", "", "")
            If intBelegReturn = 0 Then
                Return 0
            Else
                Return 1
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return 9

        End Try

    End Function



End Class
