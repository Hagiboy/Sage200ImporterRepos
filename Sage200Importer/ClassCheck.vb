Option Strict Off
Option Explicit On

Imports MySql.Data.MySqlClient
Imports System.Data.OracleClient
Imports System.Data.SqlClient

Friend Class ClassCheck

    Friend Function FcClCheckDebit(intAccounting As Int32,
                                   ByRef objdtDebits As DataSet,
                                   ByRef objFinanz As SBSXASLib.AXFinanz,
                                   ByRef objfiBuha As SBSXASLib.AXiFBhg,
                                   ByRef objdbBuha As SBSXASLib.AXiDbBhg,
                                   ByRef objdbPIFb As SBSXASLib.AXiPlFin,
                                   ByRef objFiBebu As SBSXASLib.AXiBeBu,
                                   ByRef objdtInfo As DataTable,
                                   objdtDates As DataTable,
                                   strcmbBuha As String,
                                   intTeqNbr As Int16,
                                   intTeqNbrLY As Int16,
                                   intTeqNbrPLY As Int16,
                                   strYear As String,
                                   booValutaCorrect As Boolean,
                                   datValutaCorrect As Date) As Int16

        'DebiBitLog 1=PK, 2=Konto, 3=Währung, 4=interne Bank, 5=OP Kopf, 6=RG-Datum, 7=Valuta Datum, 8=Subs, 9=OP doppelt
        Dim strBitLog As String = String.Empty
        Dim intReturnValue As Integer
        Dim strStatus As String = String.Empty
        Dim intSubNumber As Int16
        Dim dblSubNetto As Double
        Dim dblSubMwSt As Double
        Dim dblSubBrutto As Double
        Dim booAutoCorrect As Boolean
        Dim booSplittBill As Boolean
        Dim booCpyKSTToSub As Boolean
        Dim booGeneratePymentBooking As Boolean
        Dim selsubrow() As DataRow
        Dim strDebiReferenz As String = String.Empty
        Dim booDiffHeadText As Boolean
        Dim strDebiHeadText As String
        Dim booDiffSubText As Boolean
        Dim strDebiSubText As String
        Dim intDebitorNew As Int32
        Dim intiBankSage200 As Int16
        Dim dblRDiffNetto As Double
        Dim dblRDiffMwSt As Double
        Dim dblRDiffBrutto As Double
        Dim decDebiDiff As Decimal
        Dim intDZKond As Int16
        Dim intDZKondS200 As Int16

        Dim booPKPrivate As Boolean
        Dim booCashSollCorrect As Boolean
        Dim strRGNbr As String
        Dim intLinkedDebitor As Int32
        'Dim intTeqNbr As Int16
        Dim intSBGegenKonto As Int16
        Dim selSBrows() As DataRow

        'Dim datValutaPGV As Date
        Dim datValutaSave As Date
        Dim intPGVMonths As Int16
        Dim intMonthCounter As Int16
        Dim intMonthsAJ As Int16
        Dim intMonthsNJ As Int16
        Dim booDateChanged As Boolean
        Dim strMandant As String




        Try


            'Variablen einlesen
            booAutoCorrect = Convert.ToBoolean(Convert.ToInt16(Main.FcReadFromSettingsII("Buchh_HeadAutoCorrect", intAccounting)))
            booCpyKSTToSub = Convert.ToBoolean(Convert.ToInt16(Main.FcReadFromSettingsII("Buchh_KSTHeadToSub", intAccounting)))
            booSplittBill = Convert.ToBoolean(Convert.ToInt16(Main.FcReadFromSettingsII("Buchh_LinkedBookings", intAccounting)))
            booCashSollCorrect = Convert.ToBoolean(Convert.ToInt16(Main.FcReadFromSettingsII("Buchh_CashSollKontoKorr", intAccounting)))
            booGeneratePymentBooking = Convert.ToBoolean(Convert.ToInt16(Main.FcReadFromSettingsII("Buchh_GeneratePaymentBooking", intAccounting)))


            Debug.Print("Start Check " + Convert.ToString(intAccounting))
            For Each row As DataRow In objdtDebits.Tables("tblDebiHeadsFromUser").Rows

                'If row("strDebRGNbr") = "101261" Then Stop
                strRGNbr = row("strDebRGNbr") 'Für Error-Msg

                'Runden
                row("dblDebNetto") = Decimal.Round(row("dblDebNetto"), 4, MidpointRounding.AwayFromZero)
                row("dblDebMwSt") = Decimal.Round(row("dblDebMwst"), 2, MidpointRounding.AwayFromZero)
                row("dblDebBrutto") = Decimal.Round(row("dblDebBrutto"), 4, MidpointRounding.AwayFromZero)
                'RG-Create - Datum auf RG-Datum setzen falls nicht vorhanden
                If IsDBNull(row("datRGCreate")) Then
                    row("datRGCreate") = row("datDebRGDatum")
                End If
                'Credit To Debit (Gutschrift zu Rechnung) - Option auf false setzen falls nicht vorhanden
                If IsDBNull(row("booCrToInv")) Then
                    row("booCrToInv") = False
                End If
                'CreatePaymentBooking auf 0 setzen falls nicht vorhanden
                If IsDBNull(row("intKtoPayed")) Then
                    row("intKtoPayed") = 0
                End If

                'Status-String erstellen
                'Debitor 01
                intReturnValue = MainDebitor.FcGetRefDebiNr(IIf(IsDBNull(row("lngDebNbr")), 0, row("lngDebNbr")),
                                                intAccounting,
                                                intDebitorNew)

                'strBitLog += Trim(intReturnValue.ToString)
                If intReturnValue = 1 Then 'Neue Debi-Nr wurde angelegt
                    strStatus = "NDeb "
                End If
                If intDebitorNew <> 0 Or intReturnValue = 4 Then
                    intReturnValue = MainDebitor.FcCheckDebitor(intDebitorNew,
                                                                row("intBuchungsart"),
                                                                objdbBuha)
                Else
                    intReturnValue = 2
                End If
                strBitLog = Trim(intReturnValue.ToString)
                'Application.DoEvents()

                'Kto 02
                'intReturnValue = FcCheckKonto(row("lngDebKtoNbr"), objfiBuha, row("dblDebMwSt"), 0)
                intReturnValue = 0
                strBitLog += Trim(intReturnValue.ToString)

                'Currency 03
                intReturnValue = FcCheckCurrency(row("strDebCur"), objfiBuha)
                strBitLog += Trim(intReturnValue.ToString)
                'Application.DoEvents()

                'Sub 04
                If booSplittBill And IIf(IsDBNull(row("intRGArt")), 0, row("intRGArt")) = 10 Then
                    row("booLinked") = True

                Else
                    row("booLinked") = False
                End If
                intReturnValue = FcCheckSubBookings(row("strDebRGNbr"),
                                                    objdtDebits.Tables("tblDebiSubsFromUser"),
                                                    intSubNumber,
                                                    dblSubBrutto,
                                                    dblSubNetto,
                                                    dblSubMwSt,
                                                    objfiBuha,
                                                    objdbPIFb,
                                                    objFiBebu,
                                                    row("intBuchungsart"),
                                                    booAutoCorrect,
                                                    booCpyKSTToSub,
                                                    row("lngDebiKST"),
                                                    row("lngDebKtoNbr"),
                                                    booCashSollCorrect,
                                                    booSplittBill)

                objdtDebits.Tables("tblDebiSubsFromUser").AcceptChanges()

                strBitLog += Trim(intReturnValue.ToString)

                'Gibt es eine Bezahlbuchung zu erstellen? 
                'booGeneratePymentBooking = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettings(objdbconn, "Buchh_GeneratePaymentBooking", intAccounting)))
                If booGeneratePymentBooking And row("intBuchungsart") <> 1 And row("intKtoPayed") > 0 Then
                    'Bedingungen erfüllt
                    Dim drPaymentBuchung As DataRow = objdtDebits.Tables("tblDebiSubsFromUser").NewRow
                    'Felder zuweisen
                    drPaymentBuchung("strRGNr") = row("strDebRGNbr")
                    drPaymentBuchung("intSollHaben") = 0
                    drPaymentBuchung("lngKto") = row("intKtoPayed")
                    drPaymentBuchung("strKtoBez") = "Bezahlung"
                    drPaymentBuchung("lngKST") = 0
                    drPaymentBuchung("strKstBez") = "keine"
                    drPaymentBuchung("lngProj") = 0
                    drPaymentBuchung("strProjBez") = "null"
                    drPaymentBuchung("dblNetto") = row("dblDebBrutto")
                    drPaymentBuchung("dblMwSt") = 0
                    drPaymentBuchung("dblBrutto") = row("dblDebBrutto")
                    drPaymentBuchung("dblMwStSatz") = 0
                    drPaymentBuchung("strMwStKey") = "null"
                    drPaymentBuchung("strArtikel") = "Bezahlvorgang"
                    drPaymentBuchung("strDebSubText") = "Bezahlvorgang"
                    objdtDebits.Tables("tblDebiSubsFromUser").Rows.Add(drPaymentBuchung)
                    drPaymentBuchung = Nothing
                    'Summe der Sub-Buchungen anpassen
                    dblSubBrutto = Decimal.Round(dblSubBrutto + row("dblDebBrutto"), 2, MidpointRounding.AwayFromZero)

                    objdtDebits.Tables("tblDebiSubsFromUser").AcceptChanges()

                End If

                'Autokorrektur 05
                'Bei SplitBill - erste Rechnung evtl. Rückzahlung im Total nicht beachten
                If booSplittBill And row("intRGArt") = 1 And IIf(IsDBNull(row("lngLinkedRG")), 0, row("lngLinkedRG")) > 0 Then
                    row("dblDebBrutto") = Decimal.Round(dblSubBrutto, 2, MidpointRounding.AwayFromZero) * -1
                    row("dblDebNetto") = Decimal.Round(dblSubNetto, 2, MidpointRounding.AwayFromZero) * -1
                    row("dblDebMwSt") = Decimal.Round(dblSubMwSt, 2, MidpointRounding.AwayFromZero) * -1
                End If
                'booAutoCorrect = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettings(objdbconn, "Buchh_HeadAutoCorrect", intAccounting)))
                If booAutoCorrect And row("intBuchungsart") = 1 Then
                    decDebiDiff = 0
                    'Git es etwas zu korrigieren?
                    If IIf(IsDBNull(row("dblDebNetto")), 0, row("dblDebNetto")) <> dblSubNetto * -1 Or
                        IIf(IsDBNull(row("dblDebMwSt")), 0, row("dblDebMwSt")) <> dblSubMwSt * -1 Or
                        IIf(IsDBNull(row("dblDebBrutto")), 0, row("dblDebBrutto")) <> dblSubBrutto * -1 Then
                        If Math.Abs(row("dblDebBrutto") + dblSubBrutto) < 1 Then
                            decDebiDiff = Decimal.Round(Math.Abs(row("dblDebBrutto") + dblSubBrutto), 4, MidpointRounding.AwayFromZero)
                            row("dblDebBrutto") = Decimal.Round(dblSubBrutto, 4, MidpointRounding.AwayFromZero) * -1
                            row("dblDebNetto") = Decimal.Round(dblSubNetto, 4, MidpointRounding.AwayFromZero) * -1
                            row("dblDebMwSt") = Decimal.Round(dblSubMwSt, 4, MidpointRounding.AwayFromZero) * -1
                            strBitLog += "1"
                        Else
                            strBitLog += "3"
                        End If
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
                    If row("intBuchungsart") = 1 Then

                        dblRDiffBrutto = 0
                        If IIf(IsDBNull(row("dblDebMwSt")), 0, row("dblDebMwSt")) <> dblSubMwSt * -1 Then
                            row("dblDebMwSt") = dblSubMwSt * -1
                        End If
                        If IIf(IsDBNull(row("dblDebNetto")), 0, row("dblDebNetto")) <> dblSubNetto * -1 Then
                            row("dblDebNetto") = dblSubNetto * -1
                        End If

                        'Für evtl. Rundungsdifferenzen einen Datensatz in die Sub-Tabelle hinzufügen
                        If IIf(IsDBNull(row("dblDebBrutto")), 0, row("dblDebBrutto")) + dblSubBrutto <> 0 Then '0 _

                            dblRDiffBrutto = Decimal.Round(dblSubBrutto + row("dblDebBrutto"), 4, MidpointRounding.AwayFromZero)
                            dblRDiffMwSt = 0 'row("dblDebMwSt") - Decimal.Round(dblSubMwSt, 2, MidpointRounding.AwayFromZero)
                            dblRDiffNetto = 0 ' row("dblDebNetto") - Decimal.Round(dblSubNetto, 2, MidpointRounding.AwayFromZero)

                            'Zu sub-Table hinzifügen
                            Dim objdrDebiSub As DataRow = objdtDebits.Tables("tblDebiSubsFromUser").NewRow
                            objdrDebiSub("strRGNr") = row("strDebRGNbr")
                            objdrDebiSub("intSollHaben") = 1
                            objdrDebiSub("lngKto") = 6906
                            objdrDebiSub("strKtoBez") = "Rundungsdifferenzen"
                            objdrDebiSub("lngKST") = 40
                            objdrDebiSub("strKstBez") = "SystemKST"
                            objdrDebiSub("dblNetto") = dblRDiffNetto
                            objdrDebiSub("dblMwSt") = dblRDiffMwSt
                            objdrDebiSub("dblBrutto") = dblRDiffBrutto * -1
                            objdrDebiSub("dblMwStSatz") = 0
                            objdrDebiSub("strMwStKey") = "null"
                            objdrDebiSub("strArtikel") = "Rundungsdifferenz"
                            objdrDebiSub("strDebSubText") = "Eingefügt"
                            objdrDebiSub("strStatusUBBitLog") = "00000000"
                            If Math.Abs(dblRDiffBrutto) > 6 Then
                                objdrDebiSub("strStatusUBText") = "Rund > 6"
                            Else
                                objdrDebiSub("strStatusUBText") = "ok"
                            End If
                            objdtDebits.Tables("tblDebiSubsFromUser").Rows.Add(objdrDebiSub)
                            objdrDebiSub = Nothing

                            objdtDebits.Tables("tblDebiSubsFromUser").AcceptChanges()

                            'Summe der Sub-Buchungen anpassen
                            dblSubBrutto = Decimal.Round(dblSubBrutto - dblRDiffBrutto, 2, MidpointRounding.AwayFromZero)
                            If Math.Abs(dblRDiffBrutto) > 6 Then
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

                End If

                'Diff Kopf - Sub? 06
                If row("intBuchungsart") = 1 Then 'OP
                    If IIf(IsDBNull(row("dblDebBrutto")), 0, row("dblDebBrutto")) + dblSubBrutto <> 0 _
                        Or IIf(IsDBNull(row("dblDebMwSt")), 0, row("dblDebMwSt")) + dblSubMwSt <> 0 _
                        Or IIf(IsDBNull(row("dblDebNetto")), 0, row("dblDebNetto")) + dblSubNetto <> 0 Then
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
                                                  IIf(IsDBNull(row("dblDebBrutto")), 0, row("dblDebBrutto")),
                                                  IIf(IsDBNull(row("dblDebNetto")), 0, row("dblDebNetto")),
                                                  IIf(IsDBNull(row("dblDebMwSt")), 0, row("dblDebMwSt")),
                                                  dblRDiffBrutto)
                strBitLog += Trim(intReturnValue.ToString)
                'strBitLog += "0"
                'Referenz 08
                If IIf(IsDBNull(row("strDebReferenz")), "", row("strDebReferenz")) = "" And row("intBuchungsart") = 1 Then
                    intReturnValue = FcCreateDebRef(intAccounting,
                                                    row("strDebiBank"),
                                                    row("strDebRGNbr"),
                                                    row("strOPNr"),
                                                    row("intBuchungsart"),
                                                    strDebiReferenz,
                                                    row("intPayType"))
                    If Len(strDebiReferenz) > 0 Then
                        row("strDebReferenz") = strDebiReferenz
                    Else
                        intReturnValue = 1
                    End If
                Else
                    strDebiReferenz = IIf(IsDBNull(row("strDebReferenz")), "", row("strDebReferenz"))
                    intReturnValue = 0
                End If
                strBitLog += Trim(intReturnValue.ToString)

                'Status-String auswerten, vorziehen um neue PK - Nummer auszulesen
                booPKPrivate = IIf(Main.FcReadFromSettingsII("Buchh_PKTable", intAccounting) = "t_customer", True, False)
                'Debitor
                If Left(strBitLog, 1) <> "0" Then
                    strStatus += "Deb"
                    If Left(strBitLog, 1) <> "2" Then
                        If booPKPrivate = True Then
                            intReturnValue = MainDebitor.FcIsPrivateDebitorCreatable(intDebitorNew,
                                                                                     objdbBuha,
                                                                                     strcmbBuha,
                                                                                     intAccounting)
                        Else
                            intReturnValue = MainDebitor.FcIsDebitorCreatable(intDebitorNew,
                                                                              objdbBuha,
                                                                              strcmbBuha,
                                                                              intAccounting)
                        End If
                        If intReturnValue = 0 Then
                            strStatus += " erstellt"
                            row("strDebBez") = MainDebitor.FcReadDebitorName(objdbBuha,
                                                                         intDebitorNew,
                                                                         row("strDebCur"))

                        ElseIf intReturnValue = 5 Then
                            strStatus += " not approved "
                            row("strDebBez") = "nap"
                        Else
                            strStatus += " nicht erstellt."
                            row("strDebBez") = "n/a"
                        End If
                        row("lngDebNbr") = intDebitorNew
                    Else
                        strStatus += " keine Ref"
                        row("strDebBez") = "n/a"
                    End If
                Else
                    'If row("intBuchungsart") = 1 Then
                    row("strDebBez") = MainDebitor.FcReadDebitorName(objdbBuha,
                                                                     intDebitorNew,
                                                                     row("strDebCur"))
                    'Falls Währung auf Debitor nicht definiert, dann Currency setzen
                    If row("strDebBez") = "EOF" And row("intBuchungsart") = 1 Then
                        strBitLog = Left(strBitLog, 2) + "1" + Right(strBitLog, Len(strBitLog) - 3)
                    End If
                    'Else
                    'row("strDebBez") = "Nicht relevant"
                    'End If
                    row("lngDebNbr") = intDebitorNew
                End If

                'OP - Verdopplung 09
                intReturnValue = FcCheckOPDouble(objdbBuha,
                                                 row("lngDebNbr"),
                                                 row("strOPNr"),
                                                 IIf(row("dblDebBrutto") > 0, "R", "G"),
                                                 row("strDebCur"))
                strBitLog += Trim(intReturnValue.ToString)

                'PGV => Prüfung vor Valuta-Datum da Valuta-Datum verändert wird
                If Not IsDBNull(row("datPGVFrom")) Then
                    row("booPGV") = True
                End If

                'Bei Datum-Korrektur vorgängig Datum ersetzen um PGV-Buchungen zu verhindern
                If booValutaCorrect Then
                    If row("datDebRGDatum") < datValutaCorrect Then
                        row("datDebRGDatum") = datValutaCorrect.ToShortDateString
                        strStatus = "RgDCor"
                    End If
                    If row("datDebValDatum") < datValutaCorrect Then
                        row("datDebValDatum") = datValutaCorrect.ToShortDateString
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "ValDCor"
                    End If
                End If

                booDateChanged = False
                'Jahresübergreifend RG- / Valuta-Datum
                If Year(row("datDebRGDatum")) <> Year(row("datDebValDatum")) And Year(row("datDebValDatum")) >= 2023 Then
                    'Not IsDBNull(row("datPGVFrom")) Then
                    row("booPGV") = True
                    'datValutaPGV = row("datDebValDatum")
                    'Bei Valuta-Datum in einem anderen Jahr Valuta-Datum ändern
                    If Year(row("datDebRGDatum")) < Year(row("datDebValDatum")) Then
                        row("strPGVType") = "RV"
                    Else
                        row("strPGVType") = "VR"
                    End If
                    datValutaSave = row("datDebValDatum")

                    If IsDBNull(row("datPGVFrom")) Then
                        If row("strPGVType") = "VR" Then
                            row("datPGVFrom") = Year(datValutaSave).ToString + "-" + Month(datValutaSave).ToString + "-" + DateAndTime.Day(datValutaSave).ToString
                            row("datPGVTo") = Year(datValutaSave).ToString + "-" + Month(datValutaSave).ToString + "-" + DateAndTime.Day(datValutaSave).ToString
                            row("datDebValDatum") = "2024-01-01"
                            booDateChanged = True
                        ElseIf row("strPGVType") = "RV" Then
                            row("datPGVFrom") = Year(datValutaSave).ToString + "-" + Month(datValutaSave).ToString + "-" + DateAndTime.Day(datValutaSave).ToString
                            row("datPGVTo") = Year(datValutaSave).ToString + "-" + Month(datValutaSave).ToString + "-" + DateAndTime.Day(datValutaSave).ToString
                            row("datDebValDatum") = row("datDebRGDatum")
                        End If
                    Else
                        If row("strPGVType") = "RV" Then
                            row("datDebValDatum") = row("datDebRGDatum")
                            booDateChanged = True
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
                'intReturnValue = FcChCeckDate(IIf(IsDBNull(row("datDebValDatum")), #1789-09-17#, row("datDebValDatum")),
                '                              objdtInfo,
                '                              datPeriodFrom,
                '                              datPeriodTo,
                '                              strPeriodStatus,
                '                              True)
                intReturnValue = FcCheckDate2(IIf(IsDBNull(row("datDebValDatum")), #1789-09-17#, row("datDebValDatum")),
                                              strYear,
                                              objdtDates,
                                              False)

                ''Falls Problem versuchen mit Valuta-Datum-Anpassung
                'If intReturnValue <> 0 And booValutaCorrect Then
                '    row("datDebValDatum") = Format(datValutaCorrect, "Short Date")
                '    booDateChanged = True
                '    intReturnValue = FcChCeckDate(IIf(IsDBNull(row("datDebValDatum")), #1789-09-17#, row("datDebValDatum")),
                '                              objdtInfo,
                '                              datPeriodFrom,
                '                              datPeriodTo,
                '                              strPeriodStatus,
                '                              True)
                '    If intReturnValue = 0 Then
                '        'Korrektur hat funktioniert Wert auf 2 setzen
                '        intReturnValue = 2
                '    Else
                '        intReturnValue = 3
                '    End If

                'End If

                'Bei PGV checken ob PGV-Startdatum in blockierter Periode
                If row("booPGV") And intReturnValue = 0 Then
                    'Ist TA ?
                    If intMonthsAJ + intMonthsNJ = 1 Then
                        'Ist Differenz Jahre grösser 1?
                        If Math.Abs(Convert.ToInt16(strYear) - Year(row("datPGVTo"))) > 1 Then
                            intReturnValue = 4
                        Else
                            intReturnValue = FcCheckDate2(row("datPGVTo"),
                                                      strYear,
                                                      objdtDates,
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
                                                          objdtDates,
                                                          True)
                            End If
                            If intReturnValue <> 0 Then
                                Exit For
                            End If
                        Next
                    End If
                    'intReturnValue = FcCheckPGVDate(row("datPGVFrom"),
                    '                                intAccounting)
                    'If intReturnValue <> 0 Then
                    '    'Falls TA-Buchung in blockierter Periode probieren mit Valuta-Korrektur
                    '    If intPGVMonths = 1 And booValutaCorrect Then
                    '        row("datDebValDatum") = Format(datValutaCorrect, "Short Date")
                    '        booDateChanged = True
                    '        intReturnValue = FcChCeckDate(IIf(IsDBNull(row("datDebValDatum")), #1789-09-17#, row("datDebValDatum")),
                    '                          objdtInfo,
                    '                          datPeriodFrom,
                    '                          datPeriodTo,
                    '                          strPeriodStatus,
                    '                          True)
                    '        If intReturnValue = 0 Then
                    '            'PGV - Flag entfernen
                    '            row("booPGV") = False
                    '            intReturnValue = 5
                    '        Else
                    '            intReturnValue = 3
                    '        End If
                    '    Else
                    '        intReturnValue = 4
                    '    End If
                    'End If

                End If
                strBitLog += Trim(intReturnValue.ToString)

                'RG - Datum 11
                intReturnValue = FcCheckDate2(IIf(IsDBNull(row("datDebRGDatum")), #1789-09-17#, row("datDebRGDatum")),
                                              strYear,
                                              objdtDates,
                                              False)

                'intReturnValue = FcChCeckDate(IIf(IsDBNull(row("datDebRGDatum")), #1789-09-17#, row("datDebRGDatum")),
                '                              objdtInfo,
                '                              datPeriodFrom,
                '                              datPeriodTo,
                '                              strPeriodStatus,
                '                              True)

                ''Falls Problem versuchen mit Valuta-Datum-Anpassung
                'If intReturnValue <> 0 And booValutaCorrect Then
                '    row("datDebRGDatum") = datValutaCorrect
                '    booDateChanged = True
                '    intReturnValue = FcChCeckDate(IIf(IsDBNull(row("datDebRGDatum")), #1789-09-17#, row("datDebRGDatum")),
                '                              objdtInfo,
                '                              datPeriodFrom,
                '                              datPeriodTo,
                '                              strPeriodStatus,
                '                              True)
                '    If intReturnValue = 0 Then
                '        'Korrektur hat funktioniert Wert auf 2 setzen
                '        intReturnValue = 2
                '    Else
                '        intReturnValue = 3
                '    End If

                'End If
                strBitLog += Trim(intReturnValue.ToString)
                'Falls ein Datum geändert wurde dann Flag setzen
                'If booDateChanged Then
                '    row("booDatChanged") = True
                'End If

                'Interne Bank 12
                'If row.Table.Columns("intPayType") Is Nothing Then
                '    row("intPayType") = 3
                'End If
                If IsDBNull(row("intPayType")) Then
                    row("intPayType") = 9
                End If
                intReturnValue = MainDebitor.FcCheckDebiIntBank(intAccounting,
                                                                IIf(IsDBNull(row("strDebiBank")), "", row("strDebiBank")),
                                                                row("intPayType"),
                                                                intiBankSage200)
                strBitLog += Trim(intReturnValue.ToString)

                'Bei SplittBill: Existiert verlinkter Beleg? 13
                If row("booLinked") Then
                    'Zuerst Debitor von erstem Beleg suchen
                    intDebitorNew = MainDebitor.FcGetDebitorFromLinkedRG(IIf(IsDBNull(row("lngLinkedRG")), 0, row("lngLinkedRG")),
                                                                         intAccounting,
                                                                         intLinkedDebitor,
                                                                         intTeqNbr,
                                                                         intTeqNbrLY,
                                                                         intTeqNbrPLY)
                    row("lngLinkedDeb") = intLinkedDebitor

                    intReturnValue = MainDebitor.FcCheckLinkedRG(objdbBuha,
                                                                 intLinkedDebitor,
                                                                 row("strDebCur"),
                                                                 row("lngLinkedRG"))
                    'Falls erste Rechnung bezahlt, dann Flag setzen
                    If intReturnValue = 2 Then
                        row("booLinkedPayed") = True
                        intSBGegenKonto = 2331
                    Else
                        row("booLinkedPayed") = False
                        intSBGegenKonto = 1092
                    End If

                    'UB - Löschen und Buchung erstellen ohne MwSt und ohne KST da schon in RG 1 beinhaltet
                    selSBrows = objdtDebits.Tables("tblDebiSubsFromUser").Select("strRGNr='" + row("strDebRGNbr") + "'")

                    For Each SBsubrow As DataRow In selSBrows
                        SBsubrow.Delete()
                    Next

                    Dim drSBBuchung As DataRow = objdtDebits.Tables("tblDebiSubsFromUser").NewRow
                    'Felder zuweisen
                    drSBBuchung("strRGNr") = row("strDebRGNbr")
                    drSBBuchung("intSollHaben") = 1
                    drSBBuchung("lngKto") = intSBGegenKonto
                    drSBBuchung("strKtoBez") = "SB - Buchung"
                    drSBBuchung("lngKST") = 0
                    drSBBuchung("strKstBez") = "keine"
                    drSBBuchung("lngProj") = 0
                    drSBBuchung("strProjBez") = "null"
                    drSBBuchung("dblNetto") = row("dblDebBrutto") * -1
                    drSBBuchung("dblMwSt") = 0
                    drSBBuchung("dblBrutto") = row("dblDebBrutto") * -1
                    drSBBuchung("dblMwStSatz") = 0
                    drSBBuchung("strMwStKey") = "null"
                    drSBBuchung("strArtikel") = "SB - Buchung"
                    drSBBuchung("strDebSubText") = row("lngDebIdentNbr").ToString + ", FRG, " + row("lngLinkedRG").ToString
                    objdtDebits.Tables("tblDebiSubsFromUser").Rows.Add(drSBBuchung)
                    drSBBuchung = Nothing

                    objdtDebits.Tables("tblDebiSubsFromUser").AcceptChanges()

                Else
                    intReturnValue = 0
                End If
                strBitLog += Trim(intReturnValue.ToString)

                'Zahlungs-Kondition 14
                'Falls Zahlungskondition vorhanden von RG holen, sonst von Tab_Repbetrieben
                'intZKondT=1 = von Rep_Betrieben, 0=von t_payment...
                If IsDBNull(row("intZKondT")) Then
                    row("intZKondT") = 1
                End If
                If IsDBNull(row("intZKond")) Then
                    row("intZKond") = 0
                    intDZKond = 0
                Else
                    'ID in effektive Sage 200 umwandeln (=von Tabelle lesen)
                    intReturnValue = MainDebitor.FcGetDZKondSageID(row("intZKond"),
                                                                   intDZKondS200)
                    row("intZKond") = intDZKondS200
                End If
                If row("intZKondT") = 1 And row("intZKond") = 0 Then
                    'Fall kein Privatekunde
                    If booPKPrivate = False Then
                        'Daten aus den Tab_Repbetriebe holen
                        intReturnValue = MainDebitor.FcGetDZkondFromRep(row("lngDebNbr"),
                                                                    intDZKond,
                                                                    intAccounting)
                    Else
                        'Daten aus der t_customer holen
                        intReturnValue = MainDebitor.FcGetDZkondFromCust(row("lngDebNbr"),
                                                                         intDZKond,
                                                                         intAccounting)
                    End If
                    row("intZKond") = intDZKond
                End If
                'Prüfem ob Zahlungs-Kondition - ID existiert in Sage 200 bei Mandant
                strMandant = Main.FcReadFromSettingsII("Buchh200_Name",
                                                intAccounting)
                intReturnValue = MainDebitor.FcCheckDZKond(strMandant,
                                                           row("intZKond"))
                strBitLog += Trim(intReturnValue.ToString)


                'Status-String auswerten
                ''Debitor
                'If Left(strBitLog, 1) <> "0" Then
                '    strStatus = "Deb"
                '    If Left(strBitLog, 1) <> "2" Then
                '        intReturnValue = FcIsDebitorCreatable(objdbconnZHDB02, objsqlcommandZHDB02, intDebitorNew, objdbBuha)
                '        If intReturnValue = 0 Then
                '            strStatus += " erstellt"
                '        Else
                '            strStatus += " nicht erstellt."
                '        End If
                '        row("strDebBez") = FcReadDebitorName(objdbBuha, intDebitorNew, row("strDebCur"))
                '        row("lngDebNbr") = intDebitorNew
                '    Else
                '        strStatus += " keine Ref"
                '        row("strDebBez") = "n/a"
                '    End If
                'Else
                '    row("strDebBez") = FcReadDebitorName(objdbBuha, intDebitorNew, row("strDebCur"))
                '    row("lngDebNbr") = intDebitorNew
                'End If
                'Konto
                If Mid(strBitLog, 2, 1) <> "0" Then
                    If Mid(strBitLog, 2, 1) <> 2 Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Kto"
                    Else
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Kto MwSt"
                    End If
                    row("strDebKtoBez") = "n/a"
                Else
                    row("strDebKtoBez") = MainDebitor.FcReadDebitorKName(objfiBuha,
                                                                         row("lngDebKtoNbr"))
                End If
                'Währung
                If Mid(strBitLog, 3, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Cur"
                End If
                'Subbuchungen
                'Totale in Head schreiben
                row("intSubBookings") = intSubNumber.ToString
                row("dblSumSubBookings") = dblSubBrutto.ToString
                If Mid(strBitLog, 4, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Sub"
                End If
                'Autokorretkur
                If Mid(strBitLog, 5, 1) <> "0" Then
                    If Mid(strBitLog, 5, 1) = "1" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "AutoC " + decDebiDiff.ToString
                    ElseIf Mid(strBitLog, 5, 1) = "2" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Round"
                        'Wieder auf 1 setzen damit Beleg gebucht werden kann
                        Mid(strBitLog, 5, 1) = "1"
                    ElseIf Mid(strBitLog, 5, 1) = "3" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Rnd>5"
                    End If
                End If
                'Diff zu Subbuchungen
                If Mid(strBitLog, 6, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "DiffS"
                End If
                'OP Kopf
                If Mid(strBitLog, 7, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "BelK"
                End If
                'Referenz
                If Mid(strBitLog, 8, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Ref"
                    'Else
                    '    row("strDebRef") = strDebiReferenz
                End If
                'OP
                If Mid(strBitLog, 9, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "OPDbl"
                    'Else
                    '    row("strDebRef") = strDebiReferenz
                End If
                'Valuta Datum 
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
                'RG Datum 
                If Mid(strBitLog, 11, 1) <> "0" Then
                    If Mid(strBitLog, 11, 1) = "1" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "RgD"
                    ElseIf Mid(strBitLog, 11, 1) = "2" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "RgDBlck"
                        'Korrektur hat geklappt, Wert wieder auf 0 setzen
                        'strBitLog = Left(strBitLog, 10) + "0" + Right(strBitLog, Len(strBitLog) - 11)
                        'ElseIf Mid(strBitLog, 11, 1) = "3" Then
                        '    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "RgDCorNok"
                        'ElseIf Mid(strBitLog, 11, 1) = "4" Then
                        '    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "PGVDblckRD"
                    End If
                End If
                'interne Bank
                If Mid(strBitLog, 12, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "iBnk"
                Else
                    row("strDebiBank") = intiBankSage200
                End If
                'Splitt-Bill
                If Mid(strBitLog, 13, 1) <> "0" Then
                    If Mid(strBitLog, 13, 1) = "1" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "SplBNo1"
                    Else
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "SplBBez1"
                    End If

                End If
                'Zahlungs-Kondition
                If Mid(strBitLog, 14, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "ZKond"
                End If

                'PGV keine Ziffer
                If row("booPGV") Then
                    If row("intPGVMthsAY") + row("intPGVMthsNY") = 1 Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "TA " + row("strPGVType")
                    Else
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "PGV " + row("strPGVType")
                    End If
                End If

                'Status schreiben
                '5 Autokorrektur trotzdem ok, 10 Valuta 2 trotzdem ok, 11 RG 2 trotdem ok
                If Val(strBitLog) = 0 Or Val(strBitLog) = 1000002200 Or Val(strBitLog) = 2200 Or Val(strBitLog) = 1000000000 Then
                    row("booDebBook") = True
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "ok"
                End If
                row("strDebStatusText") = strStatus
                row("strDebStatusBitLog") = strBitLog

                'Wird ein anderer Text in der Head-Buchung gewünscht?
                booDiffHeadText = IIf(Main.FcReadFromSettingsII("Buchh_TextSpecial", intAccounting) = "0", False, True)
                If booDiffHeadText Then
                    strDebiHeadText = MainDebitor.FcSQLParse(Main.FcReadFromSettingsII("Buchh_TextSpecialText",
                                                                                intAccounting),
                                                             row("strDebRGNbr"),
                                                             objdtDebits.Tables("tblDebiHeadsFromUser"),
                                                             "D")
                    row("strDebText") = strDebiHeadText
                End If

                'Wird ein anderer Text in den Sub-Buchung gewünscht?
                booDiffSubText = IIf(Main.FcReadFromSettingsII("Buchh_SubTextSpecial", intAccounting) = "0", False, True)
                If booDiffSubText And Not row("booLinked") Then
                    strDebiSubText = MainDebitor.FcSQLParse(Main.FcReadFromSettingsII("Buchh_SubTextSpecialText",
                                                                               intAccounting),
                                                            row("strDebRGNbr"),
                                                            objdtDebits.Tables("tblDebiHeadsFromUser"),
                                                            "D")
                Else
                    strDebiSubText = row("strDebText")
                End If
                'Falls nicht SB - Linked dann Text in SB ersetzen
                If Not row("booLinked") Then
                    selsubrow = objdtDebits.Tables("tblDebiSubsFromUser").Select("strRGNr='" + row("strDebRGNbr") + "'")
                    For Each subrow In selsubrow
                        subrow("strDebSubText") = strDebiSubText
                    Next
                End If

                'Init
                strBitLog = String.Empty
                strStatus = String.Empty
                intSubNumber = 0
                dblSubBrutto = 0
                dblSubNetto = 0
                dblSubMwSt = 0
                intDZKond = 0

                'Application.DoEvents()
                objdtDebits.Tables("tblDebiHeadsFromUser").AcceptChanges()

            Next
            Return 0


        Catch ex As Exception
            MessageBox.Show(ex.Message + vbCrLf + "Auf RG " + strRGNbr, "Debitor Kopfdaten-Check", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 1

        Finally
            selsubrow = Nothing
            selSBrows = Nothing
            Debug.Print("End Check " + Convert.ToString(intAccounting))

        End Try



    End Function

    Friend Function FcCheckPGVDate(ByVal datPGVDateToCheck As Date,
                                          ByVal intAccounting As Int16)


        Dim tblPeriods As New DataTable
        Dim intYearToCheck As Int16
        Dim objsqlcommand As New MySqlCommand
        Dim objdbconnZHDB02 As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))

        Try

            'Zuerst testen ob überhaupt eine Definition fürs Jahr existiert
            intYearToCheck = Year(datPGVDateToCheck)

            objsqlcommand.CommandText = "SELECT * FROM t_sage_buchhaltungen_periods WHERE year=" + intYearToCheck.ToString + " AND refMandant=" + intAccounting.ToString
            objsqlcommand.Connection = objdbconnZHDB02

            objdbconnZHDB02.Open()

            tblPeriods.Load(objsqlcommand.ExecuteReader)

            If tblPeriods.Rows.Count > 0 Then
                'Hat es eine Periode definiert?
                For Each drPeriods As DataRow In tblPeriods.Rows
                    If datPGVDateToCheck >= drPeriods.Item("periodFrom") And datPGVDateToCheck <= drPeriods.Item("periodTo") Then
                        'Periodendefinition existiert, ist Periode offen?
                        If drPeriods.Item("status") = "O" Then
                            Return 0
                        Else
                            Return 1
                        End If
                    End If
                Next
            Else
                Return 0
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "PGV-Datumscheck")
            Return 9

        Finally
            objdbconnZHDB02.Close()
            objdbconnZHDB02 = Nothing
            objsqlcommand = Nothing
            tblPeriods = Nothing

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


    Friend Function FcChCeckDate(ByVal datDateToCheck As Date,
                                        ByRef objdtInfo As DataTable,
                                        ByVal datPeriodFrom As Date,
                                        ByVal datPeriodTo As Date,
                                        ByVal strPeriodStatus As String,
                                        ByVal booInclPeriods As Boolean) As Int16

        'Returns 0=ok, 1=nicht erlaubt, 9=Problem
        Dim datGJVon As Date = Convert.ToDateTime(Left(objdtInfo.Rows(2).Item(1), 4) + "-" + Mid(objdtInfo.Rows(2).Item(1), 5, 2) + "-" + Mid(objdtInfo.Rows(2).Item(1), 7, 2) + " 00:00:00")
        Dim datGJBis As Date = Convert.ToDateTime(Mid(objdtInfo.Rows(2).Item(1), 10, 4) + "-" + Mid(objdtInfo.Rows(2).Item(1), 14, 2) + "-" + Mid(objdtInfo.Rows(2).Item(1), 16, 2) + " 23:59:59")
        Dim booBuhaOpen As Boolean = IIf(Right(objdtInfo.Rows(2).Item(1), 1) = "O", True, False)
        Dim datPerVon As Date
        Dim datPerBis As Date
        Dim intActualLine As Int16
        Dim booPeriodeOpen As Boolean

        Try

            'Ist Datum in Geschäftsjahr - Def und ist Buchen erlaubt?
            If datDateToCheck >= datGJVon And datDateToCheck <= datGJBis Then
                If booBuhaOpen Then
                    If booInclPeriods Then
                        If objdtInfo.Rows.Count > 3 Then
                            intActualLine = 4
                            booPeriodeOpen = True
                            Do While intActualLine < objdtInfo.Rows.Count
                                'Wurden zusätzliche Perioden defniert und falls ja, ist der Status offen?
                                datPerVon = Convert.ToDateTime(Left(objdtInfo.Rows(intActualLine).Item(1), 10) + " 00:00:00")
                                datPerBis = Convert.ToDateTime(Mid(objdtInfo.Rows(intActualLine).Item(1), 23, 10) + " 23:59:59")
                                booBuhaOpen = IIf(Right(objdtInfo.Rows(intActualLine).Item(1), 1) = "O", True, False)
                                If datDateToCheck >= datPerVon And datDateToCheck <= datPerBis Then
                                    If booBuhaOpen And booPeriodeOpen Then
                                        'If datDateToCheck >= datPeriodFrom And datDateToCheck <= datPeriodTo And strPeriodStatus = "O" Then
                                        booPeriodeOpen = True
                                        'Else
                                        'booPeriodeOpen = False
                                        'End If
                                    Else
                                        booPeriodeOpen = False
                                    End If
                                End If
                                intActualLine += 1
                            Loop
                            If booPeriodeOpen Then
                                Return 0
                            Else
                                Return 1
                            End If
                        Else
                            Return 0
                        End If
                    Else
                        Return 0
                    End If
                Else
                    Return 1
                End If
            Else
                Return 1
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Datumscheck")
            Return 9

        Finally
            'Application.DoEvents()

        End Try

    End Function


    Friend Function FcCheckOPDouble(ByRef objdbBuha As SBSXASLib.AXiDbBhg,
                                           ByVal strDebitor As String,
                                           ByVal strOPNr As String,
                                           ByVal strType As String,
                                           ByVal strCurrency As String) As Int16

        'Return 0=ok, 1=Beleg existiert, 9=Problem

        Dim intBelegReturn As Int32

        Try
            intBelegReturn = objdbBuha.doesBelegExist(strDebitor,
                                                      strCurrency,
                                                      Main.FcCleanRGNrStrict(strOPNr),
                                                      "NOT_SET",
                                                      strType,
                                                      "")
            If intBelegReturn = 0 Then
                'Zusätzlich extern überprüfen
                intBelegReturn = objdbBuha.doesBelegExistExtern(strDebitor,
                                                                strCurrency,
                                                                strOPNr,
                                                                strType,
                                                                "")
                If intBelegReturn <> 0 Then
                    Return 1
                Else
                    Return 0
                End If
            Else
                Return 1
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Check doppelte OP - Nr.")
            Return 9

        Finally
            'Application.DoEvents()

        End Try

    End Function


    Friend Function FcCreateDebRef(ByVal intAccounting As Integer,
                                          ByVal strBank As String,
                                          ByVal strRGNr As String,
                                          ByRef strOPNr As String,
                                          ByVal intBuchungsArt As Integer,
                                          ByRef strReferenz As String,
                                          ByVal intPayType As Integer) As Integer

        'Return 0=ok oder nicht nötig, 1=keine Angaben hinterlegt, 2=Berechnung hat nicht geklappt

        Dim strTLNNr As String
        Dim strCleanedNr As String = String.Empty
        Dim strRefFrom As String
        Dim intLengthWTlNr As Int16

        Try

            If intBuchungsArt = 1 Then
                'Checken ob Referenz aus OP - Nr. oder aus Rechnung erstellt werden soll

                strRefFrom = Main.FcReadFromSettingsII("Buchh_ESRNrFrom", intAccounting)
                If strRefFrom = "" Then
                    strRefFrom = "R"
                End If

                Select Case strRefFrom
                    Case "R"
                        strCleanedNr = strRGNr
                        strOPNr = strRGNr
                    Case "O"
                        strCleanedNr = strOPNr

                End Select

                strTLNNr = FcReadBankSettings(intAccounting,
                                              intPayType,
                                              strBank)

                'Bei HW Mandant an TLNr anhängen
                If intAccounting = 29 Then
                    strTLNNr += Strings.Left(strCleanedNr, 3)
                End If

                strCleanedNr = FcCleanRGNrStrict(strCleanedNr)

                intLengthWTlNr = 26 - Len(strTLNNr)

                strReferenz = strTLNNr + StrDup(intLengthWTlNr - Len(strCleanedNr), "0") + strCleanedNr + Trim(CStr(FcModulo10(strTLNNr + StrDup(intLengthWTlNr - Len(strCleanedNr), "0") + strCleanedNr)))

                Return 0

            Else
                Return 0
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem Referenzerstellung")
            Return 1

        Finally


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

        End Try


    End Function


    Friend Function FcCleanRGNrStrict(ByVal strRGNrToClean As String) As String

        Dim intCounter As Int16
        Dim strCleanRGNr As String = String.Empty

        Try

            For intCounter = 1 To Len(strRGNrToClean)
                If Mid(strRGNrToClean, intCounter, 1) = "0" Or Val(Mid(strRGNrToClean, intCounter, 1)) > 0 Then
                    strCleanRGNr += Mid(strRGNrToClean, intCounter, 1)
                End If

            Next

            Return strCleanRGNr

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem CleanString")

        End Try


    End Function

    Friend Function FcReadBankSettings(ByVal intAccounting As Int16,
                                              ByVal intPayType As Int16,
                                              ByVal strBank As String) As String

        Dim objlocdtBank As New DataTable("tbllocBank")
        Dim objlocMySQLcmd As New MySqlCommand
        Dim objdbconn As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionStringZHDB02"))

        Try


            If intPayType = 10 Then
                objlocMySQLcmd.CommandText = "SELECT strBLZ FROM t_sage_tblaccountingbank WHERE intAccountingID=" + intAccounting.ToString + " AND QRTNNR='" + strBank + "'"
            Else
                objlocMySQLcmd.CommandText = "SELECT strBLZ FROM t_sage_tblaccountingbank WHERE intAccountingID=" + intAccounting.ToString + " AND strBank='" + strBank + "'"
            End If

            objdbconn.Open()
            objlocMySQLcmd.Connection = objdbconn
            objlocdtBank.Load(objlocMySQLcmd.ExecuteReader)

            If objlocdtBank.Rows.Count = 0 Then
                Return "0"
            Else
                Return objlocdtBank.Rows(0).Item(0).ToString
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Bankleitzahl suchen.")

        Finally
            objdbconn.Close()
            objdbconn = Nothing
            objlocdtBank = Nothing
            objlocMySQLcmd = Nothing

        End Try


    End Function

    Friend Function FcCheckBelegHead(ByVal intBuchungsArt As Int16,
                                            ByVal dblBrutto As Double,
                                            ByVal dblNetto As Double,
                                            ByVal dblMwSt As Double,
                                            ByVal dblRDiff As Double) As Int16

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

    Friend Function FcCheckCurrency(ByVal strCurrency As String,
                                    ByRef objfiBuha As SBSXASLib.AXiFBhg) As Int16

        Dim strReturn As String
        Dim booFoundCurrency As Boolean

        Try

            booFoundCurrency = False
            strReturn = String.Empty

            Call objfiBuha.ReadWhg()

            'If strCurrency = "EUR" Then Stop

            strReturn = objfiBuha.GetWhgZeile()
            Do While strReturn <> "EOF"
                If Left(strReturn, 3) = strCurrency Then
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

    Friend Function FcCheckSubBookings(ByVal strDebRgNbr As String,
                                              ByRef objDtDebiSub As DataTable,
                                              ByRef intSubNumber As Int16,
                                              ByRef dblSubBrutto As Double,
                                              ByRef dblSubNetto As Double,
                                              ByRef dblSubMwSt As Double,
                                              ByRef objFiBhg As SBSXASLib.AXiFBhg,
                                              ByRef objFiPI As SBSXASLib.AXiPlFin,
                                              ByRef objFiBebu As SBSXASLib.AXiBeBu,
                                              ByVal intBuchungsArt As Int32,
                                              ByVal booAutoCorrect As Boolean,
                                              ByVal booCpyKSTToSub As Boolean,
                                              ByVal strKST As String,
                                              ByRef lngDebKonto As Int32,
                                              ByVal booCashSollKorrekt As Boolean,
                                              ByVal booSplittBill As Boolean) As Int16

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
        Dim intSollKonto As Int32 = lngDebKonto
        Dim strProjDesc As String

        Try

            'Summen bilden und Angaben prüfen
            intSubNumber = 0
            dblSubNetto = 0
            dblSubMwSt = 0
            dblSubBrutto = 0

            selsubrow = objDtDebiSub.Select("strRGNr='" + strDebRgNbr + "'")

            For Each subrow As DataRow In selsubrow

                'If subrow("lngKto") = 3409 Then
                '    Stop
                'End If

                strBitLog = String.Empty

                'Runden
                If IsDBNull(subrow("dblNetto")) Then
                    subrow("dblNetto") = 0
                Else
                    subrow("dblNetto") = Decimal.Round(subrow("dblNetto"), 4, MidpointRounding.AwayFromZero)
                End If
                If IsDBNull(subrow("dblMwst")) Then
                    subrow("dblMwst") = 0
                Else
                    subrow("dblMwst") = Decimal.Round(subrow("dblMwst"), 2, MidpointRounding.AwayFromZero)
                End If
                If IsDBNull(subrow("dblBrutto")) Then
                    subrow("dblBrutto") = 0
                Else
                    subrow("dblBrutto") = Decimal.Round(subrow("dblBrutto"), 4, MidpointRounding.AwayFromZero)
                End If
                If IsDBNull(subrow("dblMwStSatz")) Then
                    subrow("dblMwStSatz") = 0
                Else
                    subrow("dblMwStSatz") = Decimal.Round(subrow("dblMwStSatz"), 1, MidpointRounding.AwayFromZero)
                End If

                'Falls KTRToSub dann kopieren
                If booCpyKSTToSub Then
                    subrow("lngKST") = strKST
                End If

                'Zuerst key auf 'ohne' setzen wenn MwSt-Satz = 0 und Mwst-Betrag = 0
                If subrow("dblMwStSatz") = 0 And subrow("dblMwst") = 0 And IIf(IsDBNull(subrow("strMwStKey")), "", subrow("strMwStKey")) <> "ohne" And
                    IIf(IsDBNull(subrow("strMwStKey")), "", subrow("strMwStKey")) <> "null" Then
                    'Stop
                    If IIf(IsDBNull(subrow("strMwStKey")), "", subrow("strMwStKey")) <> "AUSL0" And IIf(IsDBNull(subrow("strMwStKey")), "", subrow("strMwStKey")) <> "frei" Then
                        subrow("strMwStKey") = "ohne"
                    End If
                End If

                'Zuerst evtl. falsch gesetzte KTR oder Steuer - Sätze prüfen
                If (subrow("lngKto") >= 10000 Or subrow("lngKto") < 3000) Then 'Or subrow("strMwStKey") = "ohne" Then
                    Select Case subrow("lngKto")
                        Case 1120, 1121, 1200, 1500 To 1599, 1600 To 1699, 1700 To 1799, 2030
                            'Nur KST - Buchung resetten
                            subrow("lngKST") = 0
                        Case Else
                            subrow("strMwStKey") = Nothing
                            subrow("lngKST") = 0
                    End Select
                End If

                'MwSt prüfen 01
                If Not IsDBNull(subrow("strMwStKey")) And IIf(IsDBNull(subrow("strMwStKey")), "", subrow("strMwStKey")) <> "null" Then
                    intReturnValue = FcCheckMwSt(objFiBhg,
                                                 subrow("strMwStKey"),
                                                 IIf(IsDBNull(subrow("dblMwStSatz")), 0, subrow("dblMwStSatz")),
                                                 strStrStCodeSage200,
                                                 subrow("lngKto"))
                    If intReturnValue = 0 Then
                        subrow("strMwStKey") = strStrStCodeSage200
                        'Check ob korrekt berechnet
                        strSteuer = Split(objFiBhg.GetSteuerfeld2(subrow("lngKto").ToString,
                                                                  "Zum Rechnen", subrow("dblBrutto").ToString,
                                                                  strStrStCodeSage200), "{<}")
                        If Val(strSteuer(2)) <> IIf(IsDBNull(subrow("dblMwst")), 0, subrow("dblMwst")) Then
                            'Im Fall von Auto-Korrekt anpassen wenn Toleranz
                            'Falls MwSt-Betrag nur in 3 und 4 Stelle anders, dann erfassten Betrag nehmen.
                            If Math.Abs(Val(strSteuer(2)) - IIf(IsDBNull(subrow("dblMwst")), 0, subrow("dblMwst"))) >= 0.01 Then
                                strStatusText += "MwSt " + subrow("dblMwst").ToString
                                subrow("dblMwst") = Val(strSteuer(2))
                                strStatusText += " cor -> " + subrow("dblMwst").ToString + ", "
                            End If
                            subrow("dblBrutto") = Decimal.Round(subrow("dblNetto") + subrow("dblMwSt"), 4, MidpointRounding.AwayFromZero)
                            If Val(strSteuer(3)) <> 0 Then 'Wurde ein anderer Steuersatz gewählt?
                                subrow("dblMwStSatz") = Val(strSteuer(3))
                                subrow("strMwStKey") = strSteuer(0)
                            End If

                        End If
                    Else
                        subrow("strMwStKey") = "n/a"
                    End If
                Else
                    subrow("strMwStKey") = "null"
                    subrow("dblMwst") = 0
                    intReturnValue = 0

                End If
                strBitLog += Trim(intReturnValue.ToString)

                'If subrow("intSollHaben") <> 2 Then
                intSubNumber += 1
                If subrow("intSollHaben") = 0 Then
                    dblSubNetto += IIf(IsDBNull(subrow("dblNetto")), 0, subrow("dblNetto"))
                    dblSubMwSt += IIf(IsDBNull(subrow("dblMwSt")), 0, subrow("dblMwSt"))
                    dblSubBrutto += IIf(IsDBNull(subrow("dblBrutto")), 0, subrow("dblBrutto"))
                    If Strings.Left(subrow("lngKto").ToString, 1) = "1" Then
                        intSollKonto = subrow("lngKto") 'Für Sollkonto - Korretkur
                    End If
                Else
                    dblSubNetto -= IIf(IsDBNull(subrow("dblNetto")), 0, subrow("dblNetto"))
                    subrow("dblNetto") = subrow("dblNetto") * -1
                    dblSubMwSt -= IIf(IsDBNull(subrow("dblMwSt")), 0, subrow("dblMwSt"))
                    subrow("dblMwSt") = subrow("dblMwSt") * -1
                    dblSubBrutto -= IIf(IsDBNull(subrow("dblBrutto")), 0, subrow("dblBrutto"))
                    subrow("dblBrutto") = subrow("dblBrutto") * -1
                End If

                'Runden
                dblSubNetto = Decimal.Round(dblSubNetto, 4, MidpointRounding.AwayFromZero)
                dblSubMwSt = Decimal.Round(dblSubMwSt, 4, MidpointRounding.AwayFromZero)
                dblSubBrutto = Decimal.Round(dblSubBrutto, 4, MidpointRounding.AwayFromZero)

                'Konto prüfen 02
                If IIf(IsDBNull(subrow("lngKto")), 0, subrow("lngKto")) > 0 Then
                    'Falls KSt nicht gültig, entfernen
                    If CInt(Left(subrow("lngKto").ToString, 1)) < 3 Then
                        subrow("lngKST") = 0
                        subrow("strKtoBez") = "K<3KST ->"
                    End If
                    intReturnValue = FcCheckKonto(subrow("lngKto"),
                                                  objFiBhg,
                                                  IIf(IsDBNull(subrow("dblMwSt")), 0, subrow("dblMwSt")),
                                                  IIf(IsDBNull(subrow("lngKST")), 0, subrow("lngKST")),
                                                  False)
                    If intReturnValue = 0 Then
                        subrow("strKtoBez") += MainDebitor.FcReadDebitorKName(objFiBhg, subrow("lngKto"))
                    ElseIf intReturnValue = 2 Then
                        subrow("strKtoBez") += MainDebitor.FcReadDebitorKName(objFiBhg, subrow("lngKto")) + " MwSt!"
                    ElseIf intReturnValue = 3 Then
                        subrow("strKtoBez") += MainDebitor.FcReadDebitorKName(objFiBhg, subrow("lngKto")) + " NoKST"
                        'ElseIf intReturnValue = 4 Then
                        '    subrow("strKtoBez") = MainDebitor.FcReadDebitorKName(objFiBhg, subrow("lngKto")) + " K<3KST"
                        '    subrow("lngKST") = 0
                        '    intReturnValue = 0
                    ElseIf intReturnValue = 5 Then
                        subrow("strKtoBez") += MainDebitor.FcReadDebitorKName(objFiBhg, subrow("lngKto")) + " K<3MwSt"
                    Else
                        subrow("strKtoBez") = "n/a"

                    End If
                Else
                    subrow("strKtoBez") = "null"
                    subrow("lngKto") = 0
                    intReturnValue = 1

                End If
                strBitLog += Trim(intReturnValue.ToString)

                'Kst/Ktr prüfen 03
                If IIf(IsDBNull(subrow("lngKST")), 0, subrow("lngKST")) > 0 Then
                    intReturnValue = FcCheckKstKtr(IIf(IsDBNull(subrow("lngKST")), 0, subrow("lngKST")),
                                                   objFiBhg,
                                                   objFiBebu,
                                                   subrow("lngKto"),
                                                   strKstKtrSage200)
                    If intReturnValue = 0 Then
                        subrow("strKstBez") = strKstKtrSage200
                    ElseIf intReturnValue = 1 Then
                        subrow("strKstBez") = "KoArt"
                    ElseIf intReturnValue = 3 Then
                        subrow("strKstBez") = "NoKST"
                        subrow("lngKST") = 0
                        intReturnValue = 0 'Kein Fehler
                    Else
                        subrow("strKstBez") = "n/a"

                    End If
                Else
                    subrow("lngKST") = 0
                    subrow("strKstBez") = "null"
                    intReturnValue = 0

                End If
                strBitLog += Trim(intReturnValue.ToString)

                'Projekt prüfen 04
                If IIf(IsDBNull(subrow("lngProj")), 0, subrow("lngProj")) > 0 Then
                    intReturnValue = FcCheckProj(objFiBebu,
                                                 subrow("lngProj"),
                                                 strProjDesc)
                    If intReturnValue = 0 Then
                        subrow("strProjBez") = strProjDesc
                    ElseIf intReturnValue = 9 Then
                        subrow("strProjBez") = "n/a"
                    End If

                Else
                    subrow("lngProj") = 0
                    subrow("strProjBez") = "null"
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

                'Brutto + MwSt + Netto = 0; 05
                If IIf(IsDBNull(subrow("dblBrutto")), 0, subrow("dblBrutto")) = 0 And IIf(IsDBNull(subrow("dblMwSt")), 0, subrow("dblMwSt")) = 0 And IIf(IsDBNull(subrow("dblNetto")), 0, subrow("dblNetto")) Then
                    strBitLog += "1"

                Else
                    strBitLog += "0"
                End If

                'Netto = 0; 06
                If IIf(IsDBNull(subrow("dblNetto")), 0, subrow("dblNetto")) = 0 And Not booSplittBill Then
                    strBitLog += "1"

                Else
                    strBitLog += "0"
                End If

                'Brutto = 0; 07
                If IIf(IsDBNull(subrow("dblBrutto")), 0, subrow("dblBrutto")) = 0 Then
                    strBitLog += "1"

                Else
                    strBitLog += "0"
                End If

                'Brutto - MwSt <> Netto; 08
                If Decimal.Round(subrow("dblNetto") + subrow("dblMwSt"), 4, MidpointRounding.AwayFromZero) <> subrow("dblBrutto") Then
                    strBitLog += "1"

                Else
                    strBitLog += "0"
                End If


                'Statustext zusammen setzten
                'strStatusText = ""
                'MwSt
                If Left(strBitLog, 1) <> "0" Then
                    strStatusText += IIf(strStatusText <> "", ", ", "") + "MwSt"
                End If
                'Konto
                If Mid(strBitLog, 2, 1) <> "0" Then
                    If Mid(strBitLog, 2, 1) = "2" Then
                        strStatusText = "Kto MwSt"
                    ElseIf Mid(strBitLog, 2, 1) = "5" Then
                        strStatusText = "MwstK<3K"
                    ElseIf Mid(strBitLog, 2, 1) = "3" Then
                        strStatusText = "NoKST"
                    Else
                        strStatusText = "Kto"
                    End If
                End If
                'Kst/Ktr
                If Mid(strBitLog, 3, 1) <> "0" Then
                    strStatusText += IIf(strStatusText <> "", ", ", "") + "KST"
                End If
                'Projekt 
                If Mid(strBitLog, 4, 1) <> "0" Then
                    strStatusText += IIf(strStatusText <> "", ", ", "") + "Proj"
                End If
                'Alles 0
                If Mid(strBitLog, 5, 1) <> "0" Then
                    strStatusText += IIf(strStatusText <> "", ", ", "") + "All0"
                End If
                'Netto 0
                If Mid(strBitLog, 6, 1) <> "0" Then
                    strStatusText += IIf(strStatusText <> "", ", ", "") + "Nett0"
                End If
                'Brutto 0
                If Mid(strBitLog, 7, 1) <> "0" Then
                    strStatusText += IIf(strStatusText <> "", ", ", "") + "Brut0"
                End If
                'Diff 0
                If Mid(strBitLog, 8, 1) <> "0" Then
                    strStatusText += IIf(strStatusText <> "", ", ", "") + "Diff"
                End If

                If Val(strBitLog) = 0 Then
                    strStatusText += " ok"
                End If

                'BitLog und Text schreiben
                subrow("strStatusUBBitLog") = strBitLog
                subrow("strStatusUBText") = strStatusText
                strStatusText = String.Empty

                strStatusOverAll = strStatusOverAll Or strBitLog
                'Application.DoEvents()

            Next

            'Falls Soll-Konto-Korretkur gesetzt hier Konto ändern
            If booCashSollKorrekt And intBuchungsArt = 4 Then
                lngDebKonto = intSollKonto
            End If

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
            MessageBox.Show(ex.Message, "Fehler Debi-Subbuchungen", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        Finally
            selsubrow = Nothing
            strSteuer = Nothing
            System.GC.Collect()

        End Try

    End Function

    Friend Function FcCheckProj(ByRef objFiBebu As SBSXASLib.AXiBeBu,
                                       ByVal intProj As Int32,
                                       ByRef strProjDesc As String) As Int16

        'Returns 0=ok, 9=Problem

        Dim strLine As String
        Dim booFoundProject As Boolean
        Dim strProjectAr() As String

        Try

            booFoundProject = False
            strLine = String.Empty

            Call objFiBebu.ReadProjektTree(0)

            strLine = objFiBebu.GetProjektLine()
            Do While strLine <> "EOF"
                strProjectAr = Split(strLine, "{>}")
                Debug.Print("Aktuelle Line: " + strLine)
                'Projekt gefunden?
                If Val(strProjectAr(0)) = intProj Then
                    booFoundProject = True
                    strProjDesc = strProjectAr(1)
                End If
                strLine = objFiBebu.GetProjektLine()
            Loop

            If booFoundProject Then
                Return 0
            Else
                Return 9
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem Check-Project")

        End Try


    End Function


    Friend Function FcCheckKstKtr(ByVal lngKST As Long,
                                         ByRef objFiBhg As SBSXASLib.AXiFBhg,
                                         ByRef objBebu As SBSXASLib.AXiBeBu,
                                         ByVal lngKonto As Long,
                                         ByRef strKstKtrSage200 As String) As Int16

        'return 0=ok, 1=Kst existiert kene Kostenart, 2=Kst nicht defniert, 3=nicht auf Konto anwendbar 1000 - 2999

        Dim strReturn As String
        Dim strReturnAr() As String
        Dim booKstKAok As Boolean
        Dim strKst, strKA As String
        Dim strKAZeile As String

        booKstKAok = False

        Try
            'If CInt(Left(lngKonto.ToString, 1)) >= 3 Then
            strReturn = objFiBhg.GetKstKtrInfo(lngKST.ToString)
            If strReturn = "EOF" Then
                Return 2
            Else
                Call objBebu.ReadKaLnk(lngKST.ToString)
                Do Until strKAZeile = "EOF"
                    strKAZeile = objBebu.GetKaLnkLine
                    strReturnAr = Split(strKAZeile, "{>}")
                    strKA = strReturnAr(0)
                    If strKA = Convert.ToString(lngKonto) Then
                        booKstKAok = True
                    End If
                    'strKst = Convert.ToString(lngKST)
                    'strKA = Convert.ToString(lngKonto)
                    'Ist Kst auf Kostenbart definiert?
                Loop

                'booKstKAok = objFiPI.CheckKstKtr(strKst, strKA)

                If booKstKAok Then
                    Return 0
                Else
                    Return 1
                End If
            End If


        Catch ex As Exception
            Return 1

        End Try

    End Function


    Friend Function FcCheckKonto(ByVal lngKtoNbr As Long,
                                        ByRef objfiBuha As SBSXASLib.AXiFBhg,
                                        ByVal dblMwSt As Double,
                                        ByVal lngKST As Int32,
                                        ByVal booExistanceOnly As Boolean) As Integer

        'Returns 0=ok, 1=existiert nicht, 2=existiert aber keine KST erlaubt, 3=KST nicht auf Konto definiert, 4=KST auf Konto > 3

        Dim strReturn As String
        Dim strKontoInfo() As String

        Try

            'If lngKtoNbr = 1173 Then Stop

            strReturn = objfiBuha.GetKontoInfo(lngKtoNbr.ToString)
            If strReturn = "EOF" Then
                Return 1
            Else
                strKontoInfo = Split(objfiBuha.GetKontoInfo(lngKtoNbr.ToString), "{>}")
                If booExistanceOnly Then
                    Return 0
                End If
                'KST?

                If lngKST > 0 Then
                    If CInt(Left(lngKtoNbr.ToString, 1)) >= 3 Then

                        If strKontoInfo(22) = "" Then
                            Return 3
                        Else
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


    Friend Function FcCheckMwSt(ByRef objFiBhg As SBSXASLib.AXiFBhg,
                                       ByVal strStrCode As String,
                                       ByVal dblStrWert As Double,
                                       ByRef strStrCode200 As String,
                                       ByVal intKonto As Int32) As Integer

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

    Friend Function FcCheckKredit(intAccounting As Integer,
                                  ByRef objdtKredits As DataSet,
                                  ByRef objFinanz As SBSXASLib.AXFinanz,
                                  ByRef objfiBuha As SBSXASLib.AXiFBhg,
                                  ByRef objKrBuha As SBSXASLib.AXiKrBhg,
                                  ByRef objBebu As SBSXASLib.AXiBeBu,
                                  ByRef objdtInfo As DataTable,
                                  objdtDates As DataTable,
                                  strcmbBuha As String,
                                  strYear As String,
                                  strPeriode As String,
                                  datPeriodFrom As Date,
                                  datPeriodTo As Date,
                                  strPeriodStatus As String,
                                  booValutaCoorect As Boolean,
                                  datValutaCorrect As Date) As Integer

        'DebiBitLog 1=PK, 2=Konto, 3=Währung, 4=interne Bank, 5=OP Kopf, 6=RG-Datum, 7=Valuta Datum, 8=Subs, 9=OP doppelt
        Dim strBitLog As String = String.Empty
        Dim intReturnValue As Integer
        Dim strStatus As String = String.Empty
        Dim intSubNumber As Int16
        Dim dblSubNetto As Double
        Dim dblSubMwSt As Double
        Dim dblSubBrutto As Double
        Dim booAutoCorrect As Boolean
        Dim selsubrow() As DataRow
        Dim strKrediReferenz As String
        Dim booDiffHeadText As Boolean
        Dim strKrediHeadText As String
        Dim booDiffSubText As Boolean
        Dim booLeaveSubText As Boolean
        Dim strKrediSubText As String
        Dim intKreditorNew As Int32
        Dim strCleanOPNbr As String
        Dim intintBank As Int16
        Dim intPayType As Int16
        Dim booCpyKSTToSub As Boolean
        Dim strKredTyp As String
        Dim strIBANToPass As String
        Dim lngKrediID As Int32
        Dim dblRDiffBrutto As Double
        Dim dblRDiffMwSt As Double
        Dim dblRDiffNetto As Double
        Dim datValutaPGV As Date
        Dim intPGVMonths As Int16
        Dim intMonthCounter As Int16
        Dim intMonthsAJ As Int16
        Dim intMonthsNJ As Int16
        Dim datValutaSave As Date
        Dim booPKPrivate As Boolean

        Try

            'objdbconn.Open()
            'objOrdbconn.Open()

            booAutoCorrect = Convert.ToBoolean(Convert.ToInt16(Main.FcReadFromSettingsII("Buchh_HeadKAutoCorrect", intAccounting)))
            'booAutoCorrect = False
            booCpyKSTToSub = Convert.ToBoolean(Convert.ToInt16(Main.FcReadFromSettingsII("Buchh_KKSTHeadToSub", intAccounting)))

            For Each row As DataRow In objdtKredits.Tables("tblKrediHeadsFromUser").Rows


                'If row("lngKredID") = "117383" Then Stop
                'Runden
                row("dblKredNetto") = Decimal.Round(row("dblKredNetto"), 2, MidpointRounding.AwayFromZero)
                row("dblKredMwSt") = Decimal.Round(row("dblKredMwst"), 2, MidpointRounding.AwayFromZero)
                row("dblKredBrutto") = Decimal.Round(row("dblKredBrutto"), 2, MidpointRounding.AwayFromZero)
                'lngKrediID = row("lngKredID")

                'Status-String erstellen
                'Kreditor 01
                intReturnValue = MainKreditor.FcGetRefKrediNr(IIf(IsDBNull(row("lngKredNbr")), 0, row("lngKredNbr")),
                                                            intAccounting,
                                                            intKreditorNew)

                'strBitLog += Trim(intReturnValue.ToString)
                If intKreditorNew <> 0 Then
                    intReturnValue = MainKreditor.FcCheckKreditor(intKreditorNew,
                                                                  row("intBuchungsart"),
                                                                  objKrBuha)
                    'intReturnValue = FcCheckKreditBank(objKrBuha, intKreditorNew, row("intPayType"), row("strKredRef"), row("strKrediBank"), objdbconnZHDB02)
                    'intReturnValue = 3
                Else
                    intReturnValue = 2
                End If
                strBitLog = Trim(intReturnValue.ToString)

                'Kto 02
                'intReturnValue = FcCheckKonto(row("lngKredKtoNbr"), objfiBuha, row("dblKredMwSt"), 0)
                intReturnValue = 0
                strBitLog += Trim(intReturnValue.ToString)

                'Currency 03
                intReturnValue = FcCheckCurrency(row("strKredCur"), objfiBuha)
                strBitLog += Trim(intReturnValue.ToString)

                'Sub 04
                'booAutoCorrect = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettings(objdbconn, "Buchh_HeadKAutoCorrect", intAccounting)))
                ''booAutoCorrect = False
                'booCpyKSTToSub = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettings(objdbconn, "Buchh_KKSTHeadToSub", intAccounting)))
                intReturnValue = FcCheckKrediSubBookings(row("lngKredID"),
                                                         objdtKredits.Tables("tblKrediSubsFromUser"),
                                                         intSubNumber,
                                                         dblSubBrutto,
                                                         dblSubNetto,
                                                         dblSubMwSt,
                                                         objfiBuha,
                                                         objBebu,
                                                         row("intBuchungsart"),
                                                         booAutoCorrect,
                                                         booCpyKSTToSub,
                                                         row("lngKrediKST"),
                                                         row("intPayType"),
                                                         IIf(IsDBNull(row("strKrediBank")), "", row("strKrediBank")))

                strBitLog += Trim(intReturnValue.ToString)

                'Autokorrektur 05
                'booAutoCorrect = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettings(objdbconn, "Buchh_HeadAutoCorrect", intAccounting)))
                'booAutoCorrect = False
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
                            Dim objdrKrediSub As DataRow = objdtKredits.Tables("tblKrediSubsFromUser").NewRow
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
                            objdtKredits.Tables("tblKrediSubsFromUser").Rows.Add(objdrKrediSub)

                            objdtKredits.Tables("tblKrediSubsFromUser").AcceptChanges()

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

                'Application.DoEvents()

                'PGV => Prüfung vor Valuta-Datum da Valuta-Datum verändert wird. PGV soll nur möglich sein wenn rebilled
                If Not IsDBNull(row("datPGVFrom")) And MainKreditor.FcIsAllKrediRebilled(objdtKredits.Tables("tblKrediSubsFromUser"), row("lngKredID")) = 0 Then
                    row("booPGV") = True
                ElseIf Not IsDBNull(row("datPGVFrom")) And MainKreditor.FcIsAllKrediRebilled(objdtKredits.Tables("tblKrediSubsFromUser"), row("lngKredID")) = 1 Then
                    row("strPGVType") = "XX"
                End If

                'Bei Datum-Korrekur vorgängig Datum ersetzen um PGV-Buchung zu verhindern
                If booValutaCoorect Then
                    If row("datKredRGDatum") < datValutaCorrect Then
                        row("datKredRGDatum") = datValutaCorrect.ToShortDateString
                        strStatus = "RgDCor"
                    End If
                    If row("datKredValDatum") < datValutaCorrect Then
                        row("datKredValDatum") = datValutaCorrect.ToShortDateString
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "ValDCor"
                    End If
                End If

                'Jahresübergreifend RG- / Valuta-Datum
                If Year(row("datKredRGDatum")) <> Year(row("datKredValDatum")) And Year(row("datKredValDatum")) >= 2023 Then

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
                'intReturnValue = FcChCeckDate(IIf(IsDBNull(row("datKredValDatum")), row("datKredRGDatum"), row("datKredValDatum")),
                '                              objdtInfo,
                '                              datPeriodFrom,
                '                              datPeriodTo,
                '                              strPeriodStatus,
                '                              True)
                intReturnValue = FcCheckDate2(IIf(IsDBNull(row("datKredValDatum")), row("datKredRGDatum"), row("datKredValDatum")),
                                              strYear,
                                              objdtDates,
                                              False)

                ''Falls Problem versuchen mit Valuta-Datum-Anpassung
                'If intReturnValue <> 0 And booValutaCoorect Then
                '    row("datKredValDatum") = Format(datValutaCorrect, "Short Date")
                '    intReturnValue = FcChCeckDate(IIf(IsDBNull(row("datKredValDatum")), #1789-09-17#, row("datKredValDatum")),
                '                                  objdtInfo,
                '                                  datPeriodFrom,
                '                                  datPeriodTo,
                '                                  strPeriodStatus,
                '                                  True)
                '    If intReturnValue = 0 Then
                '        intReturnValue = 2
                '    Else
                '        intReturnValue = 3
                '    End If

                'End If

                'Bei PGV checken ob PGV-Startdatum in blockierter Periode
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
                                                      objdtDates,
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
                                                          objdtDates,
                                                          True)
                            End If
                            If intReturnValue <> 0 Then
                                Exit For
                            End If
                        Next
                    End If
                    'intReturnValue = FcCheckPGVDate(row("datPGVFrom"),
                    '                                intAccounting)
                    'If intReturnValue <> 0 Then
                    '    'Falls TA-Buchung in blockierter Periode probieren mit Valuta-Korrektur
                    '    If intPGVMonths = 1 And booValutaCorrect Then
                    '        row("datDebValDatum") = Format(datValutaCorrect, "Short Date")
                    '        booDateChanged = True
                    '        intReturnValue = FcChCeckDate(IIf(IsDBNull(row("datDebValDatum")), #1789-09-17#, row("datDebValDatum")),
                    '                          objdtInfo,
                    '                          datPeriodFrom,
                    '                          datPeriodTo,
                    '                          strPeriodStatus,
                    '                          True)
                    '        If intReturnValue = 0 Then
                    '            'PGV - Flag entfernen
                    '            row("booPGV") = False
                    '            intReturnValue = 5
                    '        Else
                    '            intReturnValue = 3
                    '        End If
                    '    Else
                    '        intReturnValue = 4
                    '    End If
                    'End If

                End If
                strBitLog += Trim(intReturnValue.ToString)


                'If row("booPGV") And intReturnValue = 0 Then
                '    intReturnValue = FcCheckPGVDate(row("datPGVFrom"),
                '                                    intAccounting)
                '    If intReturnValue <> 0 Then
                '        'Falls TP-Buchung in blockierter Periode dann probieren mit Valuta-Korrektur
                '        If intPGVMonths = 1 And booValutaCoorect Then
                '            row("datKredValDatum") = Format(datValutaCorrect, "Short Date")
                '            intReturnValue = FcChCeckDate(IIf(IsDBNull(row("datKredValDatum")), #1789-09-17#, row("datKredValDatum")),
                '                                          objdtInfo,
                '                                          datPeriodFrom,
                '                                          datPeriodTo,
                '                                          strPeriodStatus,
                '                                          True)
                '            If intReturnValue = 0 Then
                '                'PGV - Flag entfernen
                '                row("booPGV") = False
                '                intReturnValue = 5
                '            Else
                '                intReturnValue = 3
                '            End If
                '        Else
                '            intReturnValue = 4
                '        End If
                '    End If

                'End If

                'strBitLog += Trim(intReturnValue.ToString)

                'RG - Datum 11
                'intReturnValue = FcChCeckDate(IIf(IsDBNull(row("datKredRGDatum")), #1789-09-17#, row("datKredRGDatum")),
                '                              objdtInfo,
                '                              datPeriodFrom,
                '                              datPeriodTo,
                '                              strPeriodStatus,
                '                              True)

                intReturnValue = FcCheckDate2(IIf(IsDBNull(row("datKredRGDatum")), #1789-09-17#, row("datKredRGDatum")),
                                              strYear,
                                              objdtDates,
                                              False)


                'Falls Problem versuchen mit Valuta-Datum-Anpassung
                'If intReturnValue <> 0 And booValutaCoorect Then
                '    row("datKredRGDatum") = Format(datValutaCorrect, "Short Date")
                '    intReturnValue = FcChCeckDate(IIf(IsDBNull(row("datKredRGDatum")), #1789-09-17#, row("datKredRGDatum")),
                '                                  objdtInfo,
                '                                  datPeriodFrom,
                '                                  datPeriodTo,
                '                                  strPeriodStatus,
                '                                  True)
                '    If intReturnValue = 0 Then
                '        'Korrektur hat funktioniert, Wert auf 2 setzen
                '        intReturnValue = 2
                '    Else
                '        intReturnValue = 3
                '    End If
                'End If
                strBitLog += Trim(intReturnValue.ToString)

                ''Referenz 12
                If IsDBNull(row("strKredRef")) Then
                    row("strKredRef") = ""
                    intReturnValue = 1
                Else
                    If (Not String.IsNullOrEmpty(row("strKredRef"))) And (row("intPayType") = 3 Or row("intPayType") = 10) Then
                        If Val(Left(row("strKredRef"), Len(row("strKredRef")) - 1)) > 0 Then

                            'Prüfziffer korrekt?
                            If Right(row("strKredRef"), 1) <> Main.FcModulo10(Left(row("strKredRef"), Len(row("strKredRef")) - 1)) Then
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
                'intReturnValue = IIf(IsDBNull(row("strKredRef")), 1, 0)
                strBitLog += Trim(intReturnValue.ToString)

                'interne Bank 13
                intReturnValue = Main.FcCheckDebiIntBank(intAccounting,
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
                booPKPrivate = IIf(Main.FcReadFromSettingsII("Buchh_PKKrediTable", intAccounting) = "t_customer", True, False)
                'Kreditor 1
                If Left(strBitLog, 1) <> "0" Then
                    strStatus += "Kred"
                    If Left(strBitLog, 1) <> "2" Then
                        If booPKPrivate Then
                            intReturnValue = MainKreditor.FcIsPrivateKreditorCreatable(intKreditorNew,
                                                                                        objKrBuha,
                                                                                        objfiBuha,
                                                                                        IIf(IsDBNull(row("intPayType")), 3, row("intPayType")),
                                                                                        IIf(IsDBNull(row("strKredRef")), "", row("strKredRef")),
                                                                                        intintBank,
                                                                                        IIf(IsDBNull(row("strKrediBank")), "", row("strKrediBank")),
                                                                                        strcmbBuha,
                                                                                        intAccounting)
                        Else
                            intReturnValue = MainKreditor.FcIsKreditorCreatable(intKreditorNew,
                                                                            objKrBuha,
                                                                            objfiBuha,
                                                                            strcmbBuha,
                                                                            IIf(IsDBNull(row("intPayType")), 9, row("intPayType")),
                                                                            IIf(IsDBNull(row("strKredRef")), "", row("strKredRef")),
                                                                            intintBank,
                                                                            IIf(IsDBNull(row("strKrediBank")), "", row("strKrediBank")),
                                                                            intAccounting)

                        End If
                        If intReturnValue = 0 Then
                            strStatus += " erstellt"
                            intReturnValue = MainKreditor.FcReadKreditorName(objKrBuha,
                                                                                row("strKredBez"),
                                                                                intKreditorNew,
                                                                                row("strKredCur"))

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
                                                                        row("strKredBez"),
                                                                        intKreditorNew,
                                                                        row("strKredCur"))
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
                    row("strKredKtoBez") = MainDebitor.FcReadDebitorKName(objfiBuha, row("lngKredKtoNbr"))
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
                booDiffHeadText = IIf(Main.FcReadFromSettingsII("Buchh_KTextSpecial", intAccounting) = "0", False, True)
                If booDiffHeadText Then
                    strKrediHeadText = MainDebitor.FcSQLParse(Main.FcReadFromSettingsII("Buchh_KTextSpecialText",
                                                                                intAccounting),
                                                                                row("strKredRGNbr"),
                                                                            objdtKredits.Tables("tblKrediHeadsFromUser"),
                                                                            "C")
                    row("strKredText") = strKrediHeadText
                End If

                'Wird ein anderer Text in den Sub-Buchung gewünscht?
                'Soll der Gelesene Sub-Text bleiben?
                booLeaveSubText = IIf(Main.FcReadFromSettingsII("Buchh_KSubLeaveText", intAccounting) = "0", False, True)
                If Not booLeaveSubText Then
                    booDiffSubText = IIf(Main.FcReadFromSettingsII("Buchh_KSubTextSpecial", intAccounting) = "0", False, True)
                    If booDiffSubText Then
                        strKrediSubText = MainDebitor.FcSQLParse(Main.FcReadFromSettingsII("Buchh_KSubTextSpecialText",
                                                                                intAccounting),
                                                                                row("strKredRGNbr"),
                                                                           objdtKredits.Tables("tblKrediHeadsFromUser"),
                                                                           "C")
                    Else
                        strKrediSubText = row("strKredText")
                    End If
                    selsubrow = objdtKredits.Tables("tblKrediSubsFromUser").Select("lngKredID=" + row("lngKredID").ToString)
                    For Each subrow As DataRow In selsubrow
                        subrow("strKredSubText") = strKrediSubText
                    Next
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
                objdtKredits.Tables("tblKrediHeadsFromUser").AcceptChanges()

            Next

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Check-Kredit " + intKreditorNew.ToString + " ID " + lngKrediID.ToString)

        Finally

        End Try


    End Function

    Friend Function FcCheckKrediSubBookings(ByVal lngKredID As Int32,
                                              ByRef objDtKrediSub As DataTable,
                                              ByRef intSubNumber As Int16,
                                              ByRef dblSubBrutto As Double,
                                              ByRef dblSubNetto As Double,
                                              ByRef dblSubMwSt As Double,
                                              ByRef objFiBhg As SBSXASLib.AXiFBhg,
                                              ByRef objBebu As SBSXASLib.AXiBeBu,
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
                    intReturnValue = FcCheckMwSt(objFiBhg,
                                                 subrow("strMwStKey"),
                                                 subrow("dblMwStSatz"),
                                                 strStrStCodeSage200,
                                                 subrow("lngKto"))
                    If intReturnValue = 0 Then
                        subrow("strMwStKey") = strStrStCodeSage200
                        'Check ob korrekt berechnet
                        strSteuer = Split(objFiBhg.GetSteuerfeld(subrow("lngKto").ToString, "Zum Rechnen", subrow("dblBrutto").ToString, strStrStCodeSage200), "{<}")
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
                                                  objFiBhg,
                                                  IIf(IsDBNull(subrow("dblMwSt")), 0, subrow("dblMwSt")),
                                                  IIf(IsDBNull(subrow("lngKST")), 0, subrow("lngKST")),
                                                  False)
                    If intReturnValue = 0 Then
                        subrow("strKtoBez") = MainDebitor.FcReadDebitorKName(objFiBhg, subrow("lngKto"))
                    ElseIf intReturnValue = 2 Then
                        subrow("strKtoBez") = MainDebitor.FcReadDebitorKName(objFiBhg, subrow("lngKto")) + " MwSt!"
                    ElseIf intReturnValue = 3 Then
                        subrow("strKtoBez") = MainDebitor.FcReadDebitorKName(objFiBhg, subrow("lngKto")) + " NoKST"
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
                    intReturnValue = FcCheckKstKtr(subrow("lngKST"),
                                                   objFiBhg,
                                                   objBebu,
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
                If Left(strBitLog, 1) <> "0" Then
                    strStatusText += IIf(strStatusText <> "", ", ", "") + "MwSt"
                End If
                'Konto
                If Mid(strBitLog, 2, 1) <> "0" Then
                    If Left(strBitLog, 1) = "2" Then
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

        End Try

    End Function


End Class
