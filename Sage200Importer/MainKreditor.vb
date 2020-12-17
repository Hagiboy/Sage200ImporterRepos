Option Strict Off
Option Explicit On

Imports MySql.Data.MySqlClient
Imports System.Data.OracleClient
Imports System.Net
Imports System.IO
Imports System.Xml

Public Class MainKreditor

    'Public Shared Function FcCheckKrediSubBookings(ByVal lngKredID As Int32,
    '                                          ByRef objDtKrediSub As DataTable,
    '                                          ByRef intSubNumber As Int16,
    '                                          ByRef dblSubBrutto As Double,
    '                                          ByRef dblSubNetto As Double,
    '                                          ByRef dblSubMwSt As Double,
    '                                          ByRef objdbconn As MySqlConnection,
    '                                          ByRef objFiBhg As SBSXASLib.AXiFBhg,
    '                                          ByRef objFiPI As SBSXASLib.AXiPlFin,
    '                                          ByVal intBuchungsArt As Int32,
    '                                          ByVal booAutoCorrect As Boolean) As Int16

    '    'Functin Returns 0=ok, 1=Problem sub, 2=OP Diff zu Kopf, 3=OP nicht 0, 9=keine Subs

    '    'BitLog in Sub
    '    '1: Konto
    '    '2: KST
    '    '3: MwST
    '    '4: Brutto, Netto + MwSt 0
    '    '5: Netto 0
    '    '6: Brutto 0
    '    '7: Brutto - MwsT <> Netto

    '    Dim intReturnValue As Int32
    '    Dim strBitLog As String
    '    Dim strStatusText As String
    '    Dim strStrStCodeSage200 As String = ""
    '    Dim strKstKtrSage200 As String = ""
    '    Dim selsubrow() As DataRow
    '    Dim strStatusOverAll As String = "0000000"
    '    Dim strSteuer() As String

    '    'Summen bilden und Angaben prüfen
    '    intSubNumber = 0
    '    dblSubNetto = 0
    '    dblSubMwSt = 0
    '    dblSubBrutto = 0

    '    selsubrow = objDtKrediSub.Select("lngKredID=" + lngKredID.ToString)

    '    For Each subrow As DataRow In selsubrow

    '        strBitLog = ""
    '        'Runden
    '        subrow("dblNetto") = IIf(IsDBNull(subrow("dblNetto")), 0, Decimal.Round(subrow("dblNetto"), 2, MidpointRounding.AwayFromZero))
    '        subrow("dblMwSt") = IIf(IsDBNull(subrow("dblMwst")), 0, Decimal.Round(subrow("dblMwst"), 2, MidpointRounding.AwayFromZero))
    '        subrow("dblBrutto") = IIf(IsDBNull(subrow("dblBrutto")), 0, Decimal.Round(subrow("dblBrutto"), 2, MidpointRounding.AwayFromZero))
    '        subrow("dblMwStSatz") = IIf(IsDBNull(subrow("dblMwStSatz")), 0, Decimal.Round(subrow("dblMwStSatz"), 1, MidpointRounding.AwayFromZero))

    '        'MwSt prüfen
    '        If Not IsDBNull(subrow("strMwStKey")) Then
    '            intReturnValue = Main.FcCheckMwStToCorrect(objdbconn, subrow("strMwStKey"), subrow("dblMwStSatz"), subrow("dblMwSt"))
    '            intReturnValue = Main.FcCheckMwSt(objdbconn, objFiBhg, subrow("strMwStKey"), subrow("dblMwStSatz"), strStrStCodeSage200)
    '            If intReturnValue = 0 Then
    '                subrow("strMwStKey") = strStrStCodeSage200
    '                'Check of korrekt berechnet
    '                strSteuer = Split(objFiBhg.GetSteuerfeld(subrow("lngKto").ToString, "Zum Rechnen", subrow("dblBrutto").ToString, strStrStCodeSage200), "{<}")
    '                If Val(strSteuer(2)) <> subrow("dblMwst") Then
    '                    'Im Fall von Auto-Korrekt anpassen
    '                    'Stop
    '                    If booAutoCorrect Then
    '                        subrow("dblMwst") = Val(strSteuer(2))
    '                        subrow("dblBrutto") = subrow("dblNetto") + subrow("dblMwSt")
    '                    Else
    '                        intReturnValue = 1
    '                    End If
    '                End If
    '            Else
    '                subrow("strMwStKey") = "n/a"
    '            End If
    '        Else
    '            subrow("strMwStKey") = "null"
    '            intReturnValue = 0

    '        End If
    '        strBitLog += Trim(intReturnValue.ToString)


    '        'If subrow("intSollHaben") <> 2 Then
    '        intSubNumber += 1
    '        If subrow("intSollHaben") = 0 Then
    '            dblSubNetto += IIf(IsDBNull(subrow("dblNetto")), 0, subrow("dblNetto")) * -1
    '            dblSubMwSt += IIf(IsDBNull(subrow("dblMwSt")), 0, subrow("dblMwSt")) * -1
    '            dblSubBrutto += IIf(IsDBNull(subrow("dblBrutto")), 0, subrow("dblBrutto")) * -1
    '        Else
    '            dblSubNetto += subrow("dblNetto")
    '            dblSubMwSt += subrow("dblMwSt")
    '            dblSubBrutto += subrow("dblBrutto")
    '        End If

    '        'Konto prüfen
    '        If Not IsDBNull(subrow("lngKto")) Then
    '            intReturnValue = Main.FcCheckKonto(subrow("lngKto"), objFiBhg, IIf(IsDBNull(subrow("dblMwSt")), 0, subrow("dblMwSt")), IIf(IsDBNull(subrow("lngKST")), 0, subrow("lngKST")))
    '            If intReturnValue = 0 Then
    '                subrow("strKtoBez") = Main.FcReadDebitorKName(objFiBhg, subrow("lngKto"))
    '            ElseIf intReturnValue = 2 Then
    '                subrow("strKtoBez") = Main.FcReadDebitorKName(objFiBhg, subrow("lngKto")) + " MwSt!"
    '            ElseIf intReturnValue = 3 Then
    '                subrow("strKtoBez") = Main.FcReadDebitorKName(objFiBhg, subrow("lngKto")) + " NoKST"
    '                'Falls keine KST definiert KST auf 0 setzen
    '                subrow("lngKST") = 0
    '                'Error zurück setzen
    '                intReturnValue = 0
    '            Else
    '                subrow("strKtoBez") = "n/a"

    '            End If
    '        Else
    '            subrow("strKtoBez") = "null"
    '            subrow("lngKto") = 0
    '            intReturnValue = 1

    '        End If
    '        strBitLog += Trim(intReturnValue.ToString)

    '        'Kst/Ktr prüfen
    '        If IIf(IsDBNull(subrow("lngKST")), 0, subrow("lngKST")) > 0 Then
    '            intReturnValue = Main.FcCheckKstKtr(subrow("lngKST"), objFiBhg, objFiPI, subrow("lngKto"), strKstKtrSage200)
    '            If intReturnValue = 0 Then
    '                subrow("strKstBez") = strKstKtrSage200
    '            ElseIf intReturnValue = 1 Then
    '                subrow("strKstBez") = "KoArt"

    '            Else
    '                subrow("strKstBez") = "n/a"

    '            End If
    '        Else
    '            subrow("strKstBez") = "null"
    '            intReturnValue = 0

    '        End If
    '        strBitLog += Trim(intReturnValue.ToString)

    '        ''MwSt prüfen
    '        'If Not IsDBNull(subrow("strMwStKey")) Then
    '        '    intReturnValue = FcCheckMwSt(objdbconn, objFiBhg, subrow("strMwStKey"), subrow("lngMwStSatz"), strStrStCodeSage200)
    '        '    If intReturnValue = 0 Then
    '        '        subrow("strMwStKey") = strStrStCodeSage200
    '        '        'Check of korrekt berechnet
    '        '        strSteuer = Split(objFiBhg.GetSteuerfeld(subrow("lngKto").ToString, "Zum Rechnen", subrow("dblBrutto").ToString, strStrStCodeSage200), "{<}")
    '        '        If Val(strSteuer(2)) <> subrow("dblMwst") Then
    '        '            'Im Fall von Auto-Korrekt anpassen
    '        '            Stop
    '        '        End If
    '        '    Else
    '        '        subrow("strMwStKey") = "n/a"

    '        '    End If
    '        'Else
    '        '    subrow("strMwStKey") = "null"
    '        '    intReturnValue = 0

    '        'End If
    '        'strBitLog += Trim(intReturnValue.ToString)

    '        'Brutto + MwSt + Netto = 0
    '        If IIf(IsDBNull(subrow("dblBrutto")), 0, subrow("dblBrutto")) = 0 And IIf(IsDBNull(subrow("dblMwSt")), 0, subrow("dblMwSt")) = 0 And IIf(IsDBNull(subrow("dblNetto")), 0, subrow("dblNetto")) Then
    '            strBitLog += "1"

    '        Else
    '            strBitLog += "0"
    '        End If

    '        'Netto = 0
    '        If IIf(IsDBNull(subrow("dblNetto")), 0, subrow("dblNetto")) = 0 Then
    '            strBitLog += "1"

    '        Else
    '            strBitLog += "0"
    '        End If

    '        'Brutto = 0
    '        If IIf(IsDBNull(subrow("dblBrutto")), 0, subrow("dblBrutto")) = 0 Then
    '            strBitLog += "1"

    '        Else
    '            strBitLog += "0"
    '        End If

    '        'Brutto - MwSt <> Netto
    '        If Math.Round(IIf(IsDBNull(subrow("dblBrutto")), 0, subrow("dblBrutto")) - IIf(IsDBNull(subrow("dblMwSt")), 0, subrow("dblMwSt")), 2, MidpointRounding.AwayFromZero) <> IIf(IsDBNull(subrow("dblNetto")), 0, subrow("dblNetto")) Then
    '            strBitLog += "1"

    '        Else
    '            strBitLog += "0"
    '        End If


    '        'Statustext zusammen setzten
    '        strStatusText = ""
    '        'MwSt
    '        If Left(strBitLog, 1) <> "0" Then
    '            strStatusText += IIf(strStatusText <> "", ", ", "") + "MwSt"
    '        End If
    '        'Konto
    '        If Mid(strBitLog, 2, 1) <> "0" Then
    '            If Left(strBitLog, 1) = "2" Then
    '                strStatusText = "Kto MwSt"
    '            ElseIf Mid(strBitLog, 2, 1) = "3" Then
    '                strStatusText = "Kto nKST"
    '            Else
    '                strStatusText = "Kto"
    '            End If
    '        End If
    '        'Kst/Ktr
    '        If Mid(strBitLog, 3, 1) <> "0" Then
    '            strStatusText += IIf(strStatusText <> "", ", ", "") + "KST"
    '        End If
    '        'Alles 0
    '        If Mid(strBitLog, 4, 1) <> "0" Then
    '            strStatusText += IIf(strStatusText <> "", ", ", "") + "All0"
    '        End If
    '        'Netto 0
    '        If Mid(strBitLog, 5, 1) <> "0" Then
    '            strStatusText += IIf(strStatusText <> "", ", ", "") + "Net0"
    '        End If
    '        'Brutto 0
    '        If Mid(strBitLog, 6, 1) <> "0" Then
    '            strStatusText += IIf(strStatusText <> "", ", ", "") + "Brut0"
    '        End If
    '        'Diff
    '        If Mid(strBitLog, 7, 1) <> "0" Then
    '            strStatusText += IIf(strStatusText <> "", ", ", "") + "Diff"
    '        End If

    '        If Val(strBitLog) = 0 Then
    '            strStatusText = "ok"
    '        End If

    '        'BitLog und Text schreiben
    '        subrow("strStatusUBBitLog") = strBitLog
    '        subrow("strStatusUBText") = strStatusText

    '        strStatusOverAll = strStatusOverAll Or strBitLog

    '    Next

    '    'Rückgabe der ganzen Funktion Sub-Prüfung
    '    If intSubNumber = 0 Then 'keine Subs
    '        Return 9
    '    Else
    '        If Val(strStatusOverAll) > 0 Then
    '            Return 1
    '        Else
    '            Return 0
    '            'If intBuchungsArt = 1 Then
    '            '    'OP - Buchung
    '            '    'If dblSubNetto <> 0 Or dblSubBrutto <> 0 Or dblSubMwSt <> 0 Then 'Diff
    '            '    'Return 2
    '            '    'Else
    '            '    Return 0
    '            '    'End If
    '            'Else
    '            '    'Belegsbuchung 'Nur Brutto 0 - Test
    '            '    If dblSubBrutto <> 0 Then
    '            '        Return 3
    '            '    Else
    '            '        Return 0
    '            '    End If
    '            'End If
    '        End If
    '    End If

    'End Function


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
        Dim strMDBName As String = Main.FcReadFromSettings(objdbconn, "Buchh_PKTableConnection", intAccounting)
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
                If IsDBNull(objdtKreditor.Rows(0).Item(strKredNewField)) Then
                    intKrediNew = 0
                    Return 2
                Else

                    If strTableName <> "Tab_Repbetriebe" Then
                        'intPKNewField = objdtKreditor.Rows(0).Item(strKredNewField)
                        intPKNewField = Main.FcGetPKNewFromRep(objdbconnZHDB02, objsqlcommandZHDB02, objdtKreditor.Rows(0).Item(strKredNewField)) 'Rep_Nr
                        If intPKNewField = 0 Then
                            'PK wurde nicht vergeben => Eine neue erzeugen und in der Tabelle Rep_Betriebe 
                            intFunctinReturns = Main.FcNextPKNr(objdbconnZHDB02, objdtKreditor.Rows(0).Item(strKredNewField))
                            intKrediNew = 0
                            Return 3
                        Else
                            intKrediNew = intPKNewField
                            Return 0
                        End If
                    Else 'Wenn Angaben nicht von anderer Tabelle kommen
                        'Prüfen ob Repbetrieb schon eine neue Nummer erhalten hat
                        If Not IsDBNull(objdtKreditor.Rows(0).Item(strKredNewField)) Then
                            intKrediNew = objdtKreditor.Rows(0).Item(strKredNewField)
                        Else
                            intFunctinReturns = Main.FcNextPKNr(objdbconnZHDB02, objdtKreditor.Rows(0).Item(strKredNewField))
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
                                                ByRef intPayType As Int16) As Int16

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
                        strLand = "NA"
                End Select

                'Sprache zuweisen von 1-Stelligem String nach Sage 200 Regionen
                Select Case IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Language")), "D", objdtKreditor.Rows(0).Item("Rep_Language"))
                    Case "D"
                        intLangauage = 2055
                    Case "F"
                        intLangauage = 4108
                    Case "I"
                        intLangauage = 2064
                    Case Else
                        intLangauage = 2057 'Englisch
                End Select

                'Variablen zuweisen für die Erstellung des Kreditors
                strIBANNr = IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Kred_IBAN")), "", objdtKreditor.Rows(0).Item("Rep_Kred_IBAN"))
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
                strSQL = "UPDATE " + strNameKRGTable + " SET Kredigebucht=true, KredigebuchtDatum=DATE('" + Format(datDate, "yyyy-MM-dd").ToString + "'), " + strBelegNrName + "='" + intBelegNr.ToString + "' WHERE " + strKRGNbrFieldName + "=" + lngKredID.ToString
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
                    strSQLSub = Main.FcSQLParseKredi(Main.FcReadFromSettings(objdbconn, "Buchh_SQLDetailKred", intAccounting), row("lngKredID"), objdtHead)
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

    'Public Shared Function FcCheckKredit(ByVal intAccounting As Integer,
    '                                    ByRef objdtKredits As DataTable,
    '                                    ByRef objdtKreditSubs As DataTable,
    '                                    ByRef objFinanz As SBSXASLib.AXFinanz,
    '                                    ByRef objfiBuha As SBSXASLib.AXiFBhg,
    '                                    ByRef objKrBuha As SBSXASLib.AXiKrBhg,
    '                                    ByRef objdbPIFb As SBSXASLib.AXiPlFin,
    '                                    ByRef objdbconn As MySqlConnection,
    '                                    ByRef objdbconnZHDB02 As MySqlConnection,
    '                                    ByRef objsqlcommand As MySqlCommand,
    '                                    ByRef objsqlcommandZHDB02 As MySqlCommand,
    '                                    ByRef objOrdbconn As OracleClient.OracleConnection,
    '                                    ByRef objOrcommand As OracleClient.OracleCommand,
    '                                    ByRef objdbAccessConn As OleDb.OleDbConnection,
    '                                    ByRef objdtInfo As DataTable,
    '                                    ByVal strcmbBuha As String) As Integer

    '    'DebiBitLog 1=PK, 2=Konto, 3=Währung, 4=interne Bank, 5=OP Kopf, 6=RG-Datum, 7=Valuta Datum, 8=Subs, 9=OP doppelt
    '    Dim strBitLog As String = ""
    '    Dim intReturnValue As Integer
    '    Dim strStatus As String = ""
    '    Dim intSubNumber As Int16
    '    Dim dblSubNetto As Double
    '    Dim dblSubMwSt As Double
    '    Dim dblSubBrutto As Double
    '    Dim booAutoCorrect As Boolean
    '    Dim selsubrow() As DataRow
    '    Dim strKrediReferenz As String
    '    Dim booDiffHeadText As Boolean
    '    Dim strKrediiHeadText As String
    '    Dim booDiffSubText As Boolean
    '    Dim strKrediSubText As String
    '    Dim intKreditorNew As Int32
    '    Dim strCleanOPNbr As String

    '    Try

    '        objdbconn.Open()
    '        objOrdbconn.Open()

    '        For Each row As DataRow In objdtKredits.Rows

    '            '
    '            If row("lngKredID") = "48535" Then Stop
    '            'Runden
    '            row("dblKredNetto") = Decimal.Round(row("dblKredNetto"), 2, MidpointRounding.AwayFromZero)
    '            row("dblKredMwSt") = Decimal.Round(row("dblKredMwst"), 2, MidpointRounding.AwayFromZero)
    '            row("dblKredBrutto") = Decimal.Round(row("dblKredBrutto"), 2, MidpointRounding.AwayFromZero)
    '            'Status-String erstellen
    '            'Kreditor 01
    '            intReturnValue = FcGetRefKrediNr(objdbconn,
    '                                             objdbconnZHDB02,
    '                                             objsqlcommand,
    '                                             objsqlcommandZHDB02,
    '                                             objOrdbconn,
    '                                             objOrcommand,
    '                                             objdbAccessConn,
    '                                             IIf(IsDBNull(row("lngKredNbr")), 0, row("lngKredNbr")),
    '                                             intAccounting,
    '                                             intKreditorNew)

    '            strBitLog += Trim(intReturnValue.ToString)
    '            If intKreditorNew <> 0 Then
    '                intReturnValue = FcCheckKreditor(intKreditorNew, row("intBuchungsart"), objKrBuha)
    '                'intReturnValue = FcCheckKreditBank(objKrBuha, intKreditorNew, row("intPayType"), row("strKredRef"), row("strKrediBank"), objdbconnZHDB02)
    '                'intReturnValue = 3
    '            Else
    '                intReturnValue = 2
    '            End If
    '            strBitLog = Trim(intReturnValue.ToString)

    '            'Kto 02
    '            intReturnValue = FcCheckKonto(row("lngKredKtoNbr"), objfiBuha, row("dblKredMwSt"), 0)
    '            strBitLog += Trim(intReturnValue.ToString)

    '            'Currency 03
    '            intReturnValue = FcCheckCurrency(row("strKredCur"), objfiBuha)
    '            strBitLog += Trim(intReturnValue.ToString)

    '            'Sub 04
    '            booAutoCorrect = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettings(objdbconn, "Buchh_HeadAutoCorrect", intAccounting)))
    '            'booAutoCorrect = False
    '            intReturnValue = FcCheckKrediSubBookings(row("lngKredID"), objdtKreditSubs, intSubNumber, dblSubBrutto, dblSubNetto, dblSubMwSt, objdbconn, objfiBuha, objdbPIFb, row("intBuchungsart"), booAutoCorrect)
    '            strBitLog += Trim(intReturnValue.ToString)

    '            'Autokorrektur 05
    '            'booAutoCorrect = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettings(objdbconn, "Buchh_HeadAutoCorrect", intAccounting)))
    '            'booAutoCorrect = False
    '            If booAutoCorrect Then
    '                'Git es etwas zu korrigieren?
    '                If IIf(IsDBNull(row("dblKredBrutto")), 0, row("dblKredBrutto")) <> dblSubBrutto Or
    '                    IIf(IsDBNull(row("dblKredNetto")), 0, row("dblKredNetto")) <> dblSubNetto Or
    '                    IIf(IsDBNull(row("dblKredMwSt")), 0, row("dblKredMwSt")) <> dblSubMwSt Then
    '                    row("dblKredBrutto") = Math.Round(dblSubBrutto * -1, 2, MidpointRounding.AwayFromZero)
    '                    row("dblKredNetto") = dblSubNetto * -1
    '                    row("dblKredMwSt") = dblSubMwSt * -1
    '                    ''In Sub korrigieren
    '                    'selsubrow = objdtDebitSubs.Select("strRGNr='" + row("strDebRGNbr") + "' AND intSollHaben=2")
    '                    'If selsubrow.Length = 1 Then
    '                    '    selsubrow(0).Item("dblBrutto") = dblSubBrutto * -1
    '                    '    selsubrow(0).Item("dblMwSt") = dblSubMwSt * -1
    '                    '    selsubrow(0).Item("dblNetto") = dblSubNetto * -1
    '                    'End If
    '                    strBitLog += "1"
    '                Else
    '                    strBitLog += "0"
    '                End If
    '            Else
    '                strBitLog += "0"
    '            End If

    '            'Diff Kopf - Sub? 06
    '            If row("intBuchungsart") = 1 Then 'OP
    '                If IIf(IsDBNull(row("dblKredBrutto")), 0, row("dblKredBrutto")) + dblSubBrutto <> 0 _
    '                    Or IIf(IsDBNull(row("dblKredMwSt")), 0, row("dblKredMwSt")) + dblSubMwSt <> 0 _
    '                    Or IIf(IsDBNull(row("dblKredNetto")), 0, row("dblKredNetto")) + dblSubNetto <> 0 Then
    '                    strBitLog += "1"
    '                Else
    '                    strBitLog += "0"
    '                End If
    '            Else
    '                'Test ob sub 0
    '                If dblSubBrutto <> 0 Then
    '                    strBitLog += "1"
    '                Else
    '                    strBitLog += "0"
    '                End If
    '            End If
    '            'OP Kopf balanced? 07
    '            intReturnValue = FcCheckBelegHead(row("intBuchungsart"), IIf(IsDBNull(row("dblKredBrutto")), 0, row("dblKredBrutto")), IIf(IsDBNull(row("dblKredNetto")), 0, row("dblKredNetto")), IIf(IsDBNull(row("dblKredMwSt")), 0, row("dblKredMwSt")))
    '            strBitLog += Trim(intReturnValue.ToString)
    '            'OP - Nummer prüfen 08
    '            'intReturnValue = FcCreateDebRef(objdbconn, intAccounting, row("strDebiBank"), row("strDebRGNbr"), row("intBuchungsart"), strDebiReferenz)
    '            strCleanOPNbr = IIf(IsDBNull(row("strOPNr")), "", row("strOPNr"))
    '            intReturnValue = FcChCeckKredOP(strCleanOPNbr, IIf(IsDBNull(row("strKredRGNbr")), "", row("strKredRGNbr")))
    '            row("strOPNr") = strCleanOPNbr
    '            strBitLog += Trim(intReturnValue.ToString)
    '            'OP - Verdopplung 09
    '            intReturnValue = FcCheckKrediOPDouble(objKrBuha, IIf(IsDBNull(row("lngKredNbr")), 0, row("lngKredNbr")), row("strKredRGNbr"))
    '            strBitLog += Trim(intReturnValue.ToString)
    '            'Valuta - Datum 10
    '            intReturnValue = FcChCeckDate(row("datKredValDatum"), objdtInfo)
    '            strBitLog += Trim(intReturnValue.ToString)
    '            'RG - Datum 11
    '            intReturnValue = FcChCeckDate(row("datKredRGDatum"), objdtInfo)
    '            strBitLog += Trim(intReturnValue.ToString)
    '            ''intReturnValue = fcCheckIntBank()


    '            'Status-String auswerten
    '            'Kreditor
    '            If Left(strBitLog, 1) <> "0" Then
    '                strStatus = "Kred"
    '                If Left(strBitLog, 1) <> "2" Then
    '                    intReturnValue = FcIsKreditorCreatable(objdbconn, objdbconnZHDB02, objsqlcommandZHDB02, intKreditorNew, objKrBuha, strcmbBuha, row("intPayType"))
    '                    If intReturnValue = 0 Then
    '                        strStatus += " erstellt"
    '                        row("strKredBez") = FcReadKreditorName(objKrBuha, intKreditorNew, row("strKredCur"))

    '                    Else
    '                        strStatus += " nicht erstellt."
    '                        row("strKredBez") = "n/a"
    '                    End If
    '                    row("lngKredNbr") = intKreditorNew
    '                Else
    '                    strStatus += " keine Ref"
    '                    row("strKredBez") = "n/a"
    '                End If
    '            Else
    '                row("strKredBez") = FcReadKreditorName(objKrBuha, intKreditorNew, row("strKredCur"))
    '                row("lngKredNbr") = intKreditorNew
    '                intReturnValue = MainKreditor.FcCheckKreditBank(objdbconn,
    '                                                   objdbconnZHDB02,
    '                                                   objKrBuha,
    '                                                   intKreditorNew,
    '                                                   row("intPayType"),
    '                                                   row("strKredRef"),
    '                                                   row("strKredRef"),
    '                                                   row("strKredCur"))
    '            End If
    '            'Konto
    '            If Mid(strBitLog, 2, 1) <> "0" Then
    '                If Mid(strBitLog, 2, 1) <> 2 Then
    '                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Kto"
    '                Else
    '                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Kto MwSt"
    '                End If
    '                row("strKredKtoBez") = "n/a"
    '            Else
    '                row("strKredKtoBez") = FcReadDebitorKName(objfiBuha, row("lngKredKtoNbr"))
    '            End If
    '            'Währung
    '            If Mid(strBitLog, 3, 1) <> "0" Then
    '                strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Cur"
    '            End If
    '            'Subbuchungen
    '            'Totale in Head schreiben
    '            row("intSubBookings") = intSubNumber.ToString
    '            row("dblSumSubBookings") = dblSubBrutto.ToString
    '            If Mid(strBitLog, 4, 1) <> "0" Then
    '                strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Sub"
    '            End If
    '            'Autokorretkur
    '            If Mid(strBitLog, 5, 1) <> "0" Then
    '                strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "AutoC"
    '            End If
    '            'Diff zu Subbuchungen
    '            If Mid(strBitLog, 6, 1) <> "0" Then
    '                strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "DiffS"
    '            End If
    '            'OP Kopf
    '            If Mid(strBitLog, 7, 1) <> "0" Then
    '                strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "BelK"
    '            End If
    '            'OP Nummer
    '            If Mid(strBitLog, 8, 1) <> "0" Then
    '                strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "OPNbr"
    '            End If
    '            'OP Doppelt
    '            If Mid(strBitLog, 9, 1) <> "0" Then
    '                strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "OPDbl"
    '                'Else
    '                '   row("strDebRef") = strDebiReferenz
    '            End If
    '            'Valuta Datum 
    '            If Mid(strBitLog, 10, 1) <> "0" Then
    '                strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "ValD"
    '                'Else
    '                '    row("strDebRef") = strDebiReferenz
    '            End If
    '            'RG Datum 
    '            If Mid(strBitLog, 11, 1) <> "0" Then
    '                strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "RgD"
    '                'Else
    '                '    row("strDebRef") = strDebiReferenz
    '            End If
    '            'OP - Nr.

    '            'Status schreiben
    '            If Val(strBitLog) = 0 Or Val(strBitLog) = 1000000 Then
    '                row("booKredBook") = True
    '                strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "ok"
    '            End If
    '            row("strKredStatusText") = strStatus
    '            row("strKredStatusBitLog") = strBitLog

    '            ''Wird ein anderer Text in der Head-Buchung gewünscht?
    '            'booDiffHeadText = IIf(FcReadFromSettings(objdbconn, "Buchh_TextSpecial", intAccounting) = "0", False, True)
    '            'If booDiffHeadText Then
    '            '    strDebiHeadText = FcSQLParse(FcReadFromSettings(objdbconn, "Buchh_TextSpecialText", intAccounting), row("strDebRGNbr"), objdtDebits)
    '            '    row("strDebText") = strDebiHeadText
    '            'End If

    '            ''Wird ein anderer Text in den Sub-Buchung gewünscht?
    '            'booDiffSubText = IIf(FcReadFromSettings(objdbconn, "Buchh_SubTextSpecial", intAccounting) = "0", False, True)
    '            'If booDiffSubText Then
    '            '    strDebiSubText = FcSQLParse(FcReadFromSettings(objdbconn, "Buchh_SubTextSpecialText", intAccounting), row("strDebRGNbr"), objdtDebits)
    '            'Else
    '            '    strDebiSubText = row("strDebText")
    '            'End If
    '            'selsubrow = objdtDebitSubs.Select("strRGNr='" + row("strDebRGNbr") + "'")
    '            'For Each subrow In selsubrow
    '            '    subrow("strDebSubText") = strDebiSubText
    '            'Next

    '            'Init
    '            strBitLog = ""
    '            strStatus = ""
    '            intSubNumber = 0
    '            dblSubBrutto = 0
    '            dblSubNetto = 0
    '            dblSubMwSt = 0
    '            intKreditorNew = 0

    '        Next

    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)

    '    Finally
    '        If objOrdbconn.State = ConnectionState.Open Then
    '            objOrdbconn.Close()
    '        End If
    '        If objdbconn.State = ConnectionState.Open Then
    '            objdbconn.Close()
    '        End If

    '    End Try


    'End Function


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
                                Debug.Print("Gefunden " + strZahlVerbindungLine)
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

End Class
