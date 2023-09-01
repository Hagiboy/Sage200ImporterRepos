Option Strict Off
Option Explicit On

Imports MySql.Data.MySqlClient
Imports System.Data.SqlClient
Imports System.Data.OracleClient
Imports System.Net
Imports System.IO
Imports System.Xml
Imports Microsoft.VisualBasic

'Imports System.Data.OleDb

Friend NotInheritable Class Main

    Public Shared Function tblDebitorenHead() As DataTable
        Dim DT As DataTable
        'Dim myNewRow As DataRow

        Try

            DT = New DataTable("tblDebitorenHead")
            Dim strDebRGNbr As DataColumn = New DataColumn("strDebRGNbr")
            strDebRGNbr.DataType = System.Type.[GetType]("System.String")
            strDebRGNbr.MaxLength = 50
            DT.Columns.Add(strDebRGNbr)
            'DT.PrimaryKey = New DataColumn() {DT.Columns("strDebRGNbr")}
            Dim intBuchhaltung As DataColumn = New DataColumn("intBuchhaltung")
            intBuchhaltung.DataType = System.Type.[GetType]("System.Int32")
            DT.Columns.Add(intBuchhaltung)
            Dim booDebBook As DataColumn = New DataColumn("booDebBook")
            booDebBook.DataType = System.Type.[GetType]("System.Boolean")
            DT.Columns.Add(booDebBook)
            Dim intBuchungsart As DataColumn = New DataColumn("intBuchungsart")
            intBuchungsart.DataType = System.Type.[GetType]("System.Int32")
            DT.Columns.Add(intBuchungsart)
            Dim intRGArt As DataColumn = New DataColumn("intRGArt")
            intRGArt.DataType = System.Type.[GetType]("System.Int32")
            DT.Columns.Add(intRGArt)
            Dim strRGArt As DataColumn = New DataColumn("strRGArt")
            strRGArt.DataType = System.Type.[GetType]("System.String")
            strRGArt.MaxLength = 50
            DT.Columns.Add(strRGArt)
            Dim lngLinkedRG As DataColumn = New DataColumn("lngLinkedRG")
            lngLinkedRG.DataType = System.Type.[GetType]("System.Int32")
            DT.Columns.Add(lngLinkedRG)
            Dim lngLinkedDeb As DataColumn = New DataColumn("lngLinkedDeb")
            lngLinkedDeb.DataType = System.Type.[GetType]("System.Int32")
            DT.Columns.Add(lngLinkedDeb)
            Dim booLinked As DataColumn = New DataColumn("booLinked")
            booLinked.DataType = System.Type.[GetType]("System.Boolean")
            DT.Columns.Add(booLinked)
            Dim booLinkedPayed As DataColumn = New DataColumn("booLinkedPayed")
            booLinkedPayed.DataType = System.Type.[GetType]("System.Boolean")
            DT.Columns.Add(booLinkedPayed)
            Dim strRGName As DataColumn = New DataColumn("strRGName")
            strRGName.DataType = System.Type.[GetType]("System.String")
            strRGName.MaxLength = 70
            DT.Columns.Add(strRGName)
            Dim strOPNr As DataColumn = New DataColumn("strOPNr")
            strOPNr.DataType = System.Type.[GetType]("System.String")
            strOPNr.MaxLength = 20
            DT.Columns.Add(strOPNr)
            Dim lngDebNbr As DataColumn = New DataColumn("lngDebNbr")
            lngDebNbr.DataType = System.Type.[GetType]("System.Int32")
            DT.Columns.Add(lngDebNbr)
            Dim strDebPKBez As DataColumn = New DataColumn("strDebBez")
            strDebPKBez.DataType = System.Type.[GetType]("System.String")
            strDebPKBez.MaxLength = 150
            DT.Columns.Add(strDebPKBez)
            Dim lngDebKtoNbr As DataColumn = New DataColumn("lngDebKtoNbr")
            lngDebKtoNbr.DataType = System.Type.[GetType]("System.Int32")
            DT.Columns.Add(lngDebKtoNbr)
            Dim strDebKtoBez As DataColumn = New DataColumn("strDebKtoBez")
            strDebKtoBez.DataType = System.Type.[GetType]("System.String")
            strDebKtoBez.MaxLength = 150
            DT.Columns.Add(strDebKtoBez)
            Dim strDebCur As DataColumn = New DataColumn("strDebCur")
            strDebCur.DataType = System.Type.[GetType]("System.String")
            strDebCur.MaxLength = 3
            DT.Columns.Add(strDebCur)
            Dim dblDebNetto As DataColumn = New DataColumn("dblDebNetto")
            dblDebNetto.DataType = System.Type.[GetType]("System.Double")
            DT.Columns.Add(dblDebNetto)
            Dim dblDebMwSt As DataColumn = New DataColumn("dblDebMwSt")
            dblDebMwSt.DataType = System.Type.[GetType]("System.Double")
            DT.Columns.Add(dblDebMwSt)
            Dim dblDebBrutto As DataColumn = New DataColumn("dblDebBrutto")
            dblDebBrutto.DataType = System.Type.[GetType]("System.Double")
            DT.Columns.Add(dblDebBrutto)
            Dim intSubBookings As DataColumn = New DataColumn("intSubBookings")
            intSubBookings.DataType = System.Type.[GetType]("System.Int32")
            DT.Columns.Add(intSubBookings)
            Dim dblSumSubBookings As DataColumn = New DataColumn("dblSumSubBookings")
            dblSumSubBookings.DataType = System.Type.[GetType]("System.Double")
            DT.Columns.Add(dblSumSubBookings)
            Dim lngDebIdentNbr As DataColumn = New DataColumn("lngDebIdentNbr")
            lngDebIdentNbr.DataType = System.Type.[GetType]("System.Int32")
            DT.Columns.Add(lngDebIdentNbr)
            Dim strDebIdentNbr2 As DataColumn = New DataColumn("strDebIdentNbr2")
            strDebIdentNbr2.DataType = System.Type.[GetType]("System.String")
            strDebIdentNbr2.MaxLength = 60
            DT.Columns.Add(strDebIdentNbr2)
            Dim strDebText As DataColumn = New DataColumn("strDebText")
            strDebText.DataType = System.Type.[GetType]("System.String")
            strDebText.MaxLength = 255
            DT.Columns.Add(strDebText)
            Dim strRGBemerkung As DataColumn = New DataColumn("strRGBemerkung")
            strRGBemerkung.DataType = System.Type.[GetType]("System.String")
            strRGBemerkung.MaxLength = 1024
            DT.Columns.Add(strRGBemerkung)
            Dim strDebReferenz As DataColumn = New DataColumn("strDebReferenz")
            strDebReferenz.DataType = System.Type.[GetType]("System.String")
            strDebReferenz.MaxLength = 50
            DT.Columns.Add(strDebReferenz)
            Dim datDebRGDatum As DataColumn = New DataColumn("datDebRGDatum")
            datDebRGDatum.DataType = System.Type.[GetType]("System.DateTime")
            DT.Columns.Add(datDebRGDatum)
            Dim datDebValDatum As DataColumn = New DataColumn("datDebValDatum")
            datDebValDatum.DataType = System.Type.[GetType]("System.DateTime")
            DT.Columns.Add(datDebValDatum)
            Dim datRGCreate As DataColumn = New DataColumn("datRGCreate")
            datRGCreate.DataType = System.Type.[GetType]("System.DateTime")
            DT.Columns.Add(datRGCreate)
            Dim datDebDue As DataColumn = New DataColumn("datDebDue")
            datDebDue.DataType = System.Type.[GetType]("System.DateTime")
            DT.Columns.Add(datDebDue)
            Dim intPayTYpe As DataColumn = New DataColumn("intPayType")
            intPayTYpe.DataType = System.Type.GetType("System.Int16")
            DT.Columns.Add(intPayTYpe)
            Dim strDebiBank As DataColumn = New DataColumn("strDebiBank")
            strDebiBank.DataType = System.Type.[GetType]("System.String")
            strDebiBank.MaxLength = 27
            DT.Columns.Add(strDebiBank)
            Dim strDebRef As DataColumn = New DataColumn("strDebRef")
            strDebRef.DataType = System.Type.[GetType]("System.String")
            strDebRef.MaxLength = 27
            DT.Columns.Add(strDebRef)
            Dim strZahlBed As DataColumn = New DataColumn("strZahlBed")
            strZahlBed.DataType = System.Type.[GetType]("System.String")
            strZahlBed.MaxLength = 5
            DT.Columns.Add(strZahlBed)
            Dim intintBank As DataColumn = New DataColumn("intintBank")
            intintBank.DataType = System.Type.[GetType]("System.Int16")
            intintBank.Caption = "IBank"
            DT.Columns.Add(intintBank)
            Dim strDebStatusBitLog As DataColumn = New DataColumn("strDebStatusBitLog")
            strDebStatusBitLog.DataType = System.Type.[GetType]("System.String")
            strDebStatusBitLog.MaxLength = 50
            DT.Columns.Add(strDebStatusBitLog)
            Dim strDebStatusText As DataColumn = New DataColumn("strDebStatusText")
            strDebStatusText.DataType = System.Type.[GetType]("System.String")
            strDebStatusText.MaxLength = 255
            DT.Columns.Add(strDebStatusText)
            Dim strDebBookStatus As DataColumn = New DataColumn("strDebBookStatus")
            strDebBookStatus.DataType = System.Type.[GetType]("System.String")
            strDebBookStatus.MaxLength = 50
            DT.Columns.Add(strDebBookStatus)
            Dim booBooked As DataColumn = New DataColumn("booBooked")
            booBooked.DataType = System.Type.[GetType]("System.Boolean")
            DT.Columns.Add(booBooked)
            Dim datBooked As DataColumn = New DataColumn("datBooked")
            datBooked.DataType = System.Type.[GetType]("System.DateTime")
            DT.Columns.Add(datBooked)
            Dim lngBelegNr As DataColumn = New DataColumn("lngBelegNr")
            lngBelegNr.DataType = System.Type.[GetType]("System.Int32")
            DT.Columns.Add(lngBelegNr)
            Dim lngDebiKST As DataColumn = New DataColumn("lngDebiKST")
            lngDebiKST.DataType = System.Type.[GetType]("System.Int32")
            DT.Columns.Add(lngDebiKST)
            Dim booCrToInv As DataColumn = New DataColumn("booCrToInv")
            booCrToInv.DataType = System.Type.[GetType]("System.Boolean")
            DT.Columns.Add(booCrToInv)
            Dim intKtoPayed As DataColumn = New DataColumn("intKtoPayed")
            intKtoPayed.DataType = System.Type.[GetType]("System.Int32")
            DT.Columns.Add(intKtoPayed)
            Dim booPGV As DataColumn = New DataColumn("booPGV")
            booPGV.DataType = System.Type.[GetType]("System.Boolean")
            DT.Columns.Add(booPGV)
            Dim strPGVType As DataColumn = New DataColumn("strPGVType")
            strPGVType.DataType = System.Type.[GetType]("System.String")
            strPGVType.MaxLength = 2
            DT.Columns.Add(strPGVType)
            Dim datPGVFrom As DataColumn = New DataColumn("datPGVFrom")
            datPGVFrom.DataType = System.Type.[GetType]("System.DateTime")
            DT.Columns.Add(datPGVFrom)
            Dim intPGVMthsAY As DataColumn = New DataColumn("intPGVMthsAY")
            intPGVMthsAY.DataType = System.Type.[GetType]("System.Int32")
            DT.Columns.Add(intPGVMthsAY)
            Dim datPGVTo As DataColumn = New DataColumn("datPGVTo")
            datPGVTo.DataType = System.Type.[GetType]("System.DateTime")
            DT.Columns.Add(datPGVTo)
            Dim intPGVMthsNY As DataColumn = New DataColumn("intPGVMthsNY")
            intPGVMthsNY.DataType = System.Type.[GetType]("System.Int32")
            DT.Columns.Add(intPGVMthsNY)
            Dim booDatChanged As DataColumn = New DataColumn("booDatChanged")
            booDatChanged.DataType = System.Type.[GetType]("System.Boolean")
            DT.Columns.Add(booDatChanged)
            Dim intZKond As DataColumn = New DataColumn("intZKond")
            intZKond.DataType = System.Type.[GetType]("System.Int16")
            DT.Columns.Add(intZKond)
            Dim intZKondT As DataColumn = New DataColumn("intZKondT")
            intZKondT.DataType = System.Type.[GetType]("System.Int16")
            DT.Columns.Add(intZKondT)
            Return DT

        Catch ex As Exception
            MessageBox.Show(ex.Message + "Debitoren-Head-Tabelle " + Err.Number.ToString)

        End Try


    End Function

    Public Shared Function tblDebitorenSub() As DataTable

        Dim DT As DataTable

        Try

            DT = New DataTable("tblDebitorenSub")
            Dim lngID As DataColumn = New DataColumn("lngID")
            lngID.DataType = System.Type.[GetType]("System.Int32")
            lngID.AutoIncrement = True
            lngID.AutoIncrementSeed = 1
            lngID.AutoIncrementStep = 1
            DT.Columns.Add(lngID)
            Dim strRGNr As DataColumn = New DataColumn("strRGNr")
            strRGNr.DataType = System.Type.[GetType]("System.String")
            strRGNr.MaxLength = 50
            strRGNr.Caption = "RG-Nr"
            DT.Columns.Add(strRGNr)
            Dim intSollHaben As DataColumn = New DataColumn("intSollHaben")
            intSollHaben.DataType = System.Type.[GetType]("System.Int16")
            intSollHaben.Caption = "S1/H0"
            DT.Columns.Add(intSollHaben)
            Dim lngKto As DataColumn = New DataColumn("lngKto")
            lngKto.DataType = System.Type.[GetType]("System.Int32")
            lngKto.Caption = "Konto"
            DT.Columns.Add(lngKto)
            Dim strKtoBez As DataColumn = New DataColumn("strKtoBez")
            strKtoBez.DataType = System.Type.[GetType]("System.String")
            strKtoBez.MaxLength = 150
            strKtoBez.Caption = "Bezeichnung"
            DT.Columns.Add(strKtoBez)
            Dim lngKST As DataColumn = New DataColumn("lngKST")
            lngKST.DataType = System.Type.[GetType]("System.Int32")
            lngKST.Caption = "KST"
            DT.Columns.Add(lngKST)
            Dim strKstBez As DataColumn = New DataColumn("strKstBez")
            strKstBez.DataType = System.Type.[GetType]("System.String")
            strKstBez.MaxLength = 150
            strKstBez.Caption = "Bez."
            DT.Columns.Add(strKstBez)
            Dim lngProj As DataColumn = New DataColumn("lngProj")
            lngProj.DataType = System.Type.[GetType]("System.Int32")
            lngProj.Caption = "Proj"
            DT.Columns.Add(lngProj)
            Dim strProjBez As DataColumn = New DataColumn("strProjBez")
            strProjBez.DataType = System.Type.[GetType]("System.String")
            strProjBez.MaxLength = 50
            strProjBez.Caption = "Pr-Bez."
            DT.Columns.Add(strProjBez)
            Dim dblNetto As DataColumn = New DataColumn("dblNetto")
            dblNetto.DataType = System.Type.[GetType]("System.Double")
            dblNetto.Caption = "Netto"
            DT.Columns.Add(dblNetto)
            Dim dblMwSt As DataColumn = New DataColumn("dblMwSt")
            dblMwSt.DataType = System.Type.[GetType]("System.Double")
            dblMwSt.Caption = "MwSt"
            DT.Columns.Add(dblMwSt)
            Dim dblBrutto As DataColumn = New DataColumn("dblBrutto")
            dblBrutto.DataType = System.Type.[GetType]("System.Double")
            dblBrutto.Caption = "Brutto"
            DT.Columns.Add(dblBrutto)
            Dim dblMwStSatz As DataColumn = New DataColumn("dblMwStSatz")
            dblMwStSatz.DataType = System.Type.[GetType]("System.Double")
            dblMwStSatz.Caption = "MwStS"
            DT.Columns.Add(dblMwStSatz)
            Dim strMwStKey As DataColumn = New DataColumn("strMwStKey")
            strMwStKey.DataType = System.Type.[GetType]("System.String")
            strMwStKey.MaxLength = 50
            DT.Columns.Add(strMwStKey)
            Dim strArtikel As DataColumn = New DataColumn("strArtikel")
            strArtikel.DataType = System.Type.[GetType]("System.String")
            strArtikel.MaxLength = 255
            DT.Columns.Add(strArtikel)
            Dim strDebSubText As DataColumn = New DataColumn("strDebSubText")
            strDebSubText.DataType = System.Type.[GetType]("System.String")
            strDebSubText.MaxLength = 255
            strDebSubText.Caption = "Buch-Text"
            DT.Columns.Add(strDebSubText)
            Dim strStatusUBBitLog As DataColumn = New DataColumn("strStatusUBBitLog")
            strStatusUBBitLog.DataType = System.Type.[GetType]("System.String")
            strStatusUBBitLog.MaxLength = 50
            DT.Columns.Add(strStatusUBBitLog)
            Dim strStatusUBText As DataColumn = New DataColumn("strStatusUBText")
            strStatusUBText.DataType = System.Type.[GetType]("System.String")
            strStatusUBText.MaxLength = 255
            strStatusUBText.Caption = "Status"
            DT.Columns.Add(strStatusUBText)
            Dim strDebBookStatus As DataColumn = New DataColumn("strDebBookStatus")
            strDebBookStatus.DataType = System.Type.[GetType]("System.String")
            strDebBookStatus.MaxLength = 50
            DT.Columns.Add(strDebBookStatus)
            Return DT

        Catch ex As Exception
            MessageBox.Show(ex.Message + "Debitoren-Sub-Tabelle " + Err.Number.ToString)
        End Try

    End Function

    Public Shared Function tblKreditorenHead() As DataTable

        Dim DT As DataTable

        Try

            'Dim myNewRow As DataRow
            DT = New DataTable("tblKreditorenHead")
            Dim lngKredID As DataColumn = New DataColumn("lngKredID")
            lngKredID.DataType = System.Type.[GetType]("System.Int32")
            DT.Columns.Add(lngKredID)
            'DT.PrimaryKey = New DataColumn() {DT.Columns("lngKredID")}
            Dim strKredRGNbr As DataColumn = New DataColumn("strKredRGNbr")
            strKredRGNbr.DataType = System.Type.[GetType]("System.String")
            strKredRGNbr.MaxLength = 50
            DT.Columns.Add(strKredRGNbr)
            'DT.PrimaryKey = New DataColumn() {DT.Columns("strKredRGNbr")}
            Dim intBuchhaltung As DataColumn = New DataColumn("intBuchhaltung")
            intBuchhaltung.DataType = System.Type.[GetType]("System.Int32")
            DT.Columns.Add(intBuchhaltung)
            Dim booDebBook As DataColumn = New DataColumn("booKredBook")
            booDebBook.DataType = System.Type.[GetType]("System.Boolean")
            DT.Columns.Add(booDebBook)
            Dim intBuchungsart As DataColumn = New DataColumn("intBuchungsart")
            intBuchungsart.DataType = System.Type.[GetType]("System.Int32")
            DT.Columns.Add(intBuchungsart)
            'Dim intRGArt As DataColumn = New DataColumn("intRGArt")
            'intRGArt.DataType = System.Type.[GetType]("System.Int32")
            'DT.Columns.Add(intRGArt)
            'Dim strRGArt As DataColumn = New DataColumn("strRGArt")
            'strRGArt.DataType = System.Type.[GetType]("System.String")
            'strRGArt.MaxLength = 50
            'DT.Columns.Add(strRGArt)
            'Dim lngLinkedRG As DataColumn = New DataColumn("lngLinkedRG")
            'lngLinkedRG.DataType = System.Type.[GetType]("System.Int32")
            'DT.Columns.Add(lngLinkedRG)
            'Dim booLinked As DataColumn = New DataColumn("booLinked")
            'booLinked.DataType = System.Type.[GetType]("System.Boolean")
            'DT.Columns.Add(booLinked)
            Dim strRGName As DataColumn = New DataColumn("strRGName")
            strRGName.DataType = System.Type.[GetType]("System.String")
            strRGName.MaxLength = 50
            DT.Columns.Add(strRGName)
            Dim strOPNr As DataColumn = New DataColumn("strOPNr")
            strOPNr.DataType = System.Type.[GetType]("System.String")
            strOPNr.MaxLength = 30
            DT.Columns.Add(strOPNr)
            Dim lngKredNbr As DataColumn = New DataColumn("lngKredNbr")
            lngKredNbr.DataType = System.Type.[GetType]("System.Int32")
            DT.Columns.Add(lngKredNbr)
            Dim strKredPKBez As DataColumn = New DataColumn("strKredBez")
            strKredPKBez.DataType = System.Type.[GetType]("System.String")
            strKredPKBez.MaxLength = 50
            DT.Columns.Add(strKredPKBez)
            Dim lngKredKtoNbr As DataColumn = New DataColumn("lngKredKtoNbr")
            lngKredKtoNbr.DataType = System.Type.[GetType]("System.Int32")
            DT.Columns.Add(lngKredKtoNbr)
            Dim strKredKtoBez As DataColumn = New DataColumn("strKredKtoBez")
            strKredKtoBez.DataType = System.Type.[GetType]("System.String")
            strKredKtoBez.MaxLength = 80
            DT.Columns.Add(strKredKtoBez)
            Dim strKredCur As DataColumn = New DataColumn("strKredCur")
            strKredCur.DataType = System.Type.[GetType]("System.String")
            strKredCur.MaxLength = 3
            DT.Columns.Add(strKredCur)
            Dim dblKredNetto As DataColumn = New DataColumn("dblKredNetto")
            dblKredNetto.DataType = System.Type.[GetType]("System.Double")
            DT.Columns.Add(dblKredNetto)
            Dim dblKredMwSt As DataColumn = New DataColumn("dblKredMwSt")
            dblKredMwSt.DataType = System.Type.[GetType]("System.Double")
            DT.Columns.Add(dblKredMwSt)
            Dim dblKredBrutto As DataColumn = New DataColumn("dblKredBrutto")
            dblKredBrutto.DataType = System.Type.[GetType]("System.Double")
            DT.Columns.Add(dblKredBrutto)
            Dim intSubBookings As DataColumn = New DataColumn("intSubBookings")
            intSubBookings.DataType = System.Type.[GetType]("System.Int32")
            DT.Columns.Add(intSubBookings)
            Dim dblSumSubBookings As DataColumn = New DataColumn("dblSumSubBookings")
            dblSumSubBookings.DataType = System.Type.[GetType]("System.Double")
            DT.Columns.Add(dblSumSubBookings)
            Dim lngKredIdentNbr As DataColumn = New DataColumn("lngKredIdentNbr")
            lngKredIdentNbr.DataType = System.Type.[GetType]("System.Int32")
            DT.Columns.Add(lngKredIdentNbr)
            Dim strKredIdentNbr2 As DataColumn = New DataColumn("strKredIdentNbr2")
            strKredIdentNbr2.DataType = System.Type.[GetType]("System.String")
            strKredIdentNbr2.MaxLength = 50
            DT.Columns.Add(strKredIdentNbr2)
            Dim strKredText As DataColumn = New DataColumn("strKredText")
            strKredText.DataType = System.Type.[GetType]("System.String")
            strKredText.MaxLength = 125
            DT.Columns.Add(strKredText)
            Dim strRGBemerkung As DataColumn = New DataColumn("strRGBemerkung")
            strRGBemerkung.DataType = System.Type.[GetType]("System.String")
            strRGBemerkung.MaxLength = 1024
            DT.Columns.Add(strRGBemerkung)
            Dim datKredRGDatum As DataColumn = New DataColumn("datKredRGDatum")
            datKredRGDatum.DataType = System.Type.[GetType]("System.DateTime")
            DT.Columns.Add(datKredRGDatum)
            Dim datKredValDatum As DataColumn = New DataColumn("datKredValDatum")
            datKredValDatum.DataType = System.Type.[GetType]("System.DateTime")
            DT.Columns.Add(datKredValDatum)
            Dim strKrediBank As DataColumn = New DataColumn("strKrediBank")
            strKrediBank.DataType = System.Type.[GetType]("System.String")
            strKrediBank.MaxLength = 50
            DT.Columns.Add(strKrediBank)
            Dim strKredRef As DataColumn = New DataColumn("strKredRef")
            strKredRef.DataType = System.Type.[GetType]("System.String")
            strKredRef.MaxLength = 30
            DT.Columns.Add(strKredRef)
            Dim strKrediBankInt As DataColumn = New DataColumn("strKrediBankInt")
            strKrediBankInt.DataType = System.Type.[GetType]("System.String")
            strKrediBankInt.MaxLength = 5
            DT.Columns.Add(strKrediBankInt)
            'Dim strZahlBed As DataColumn = New DataColumn("strZahlBed")
            'strZahlBed.DataType = System.Type.[GetType]("System.String")
            'strZahlBed.MaxLength = 5
            'DT.Columns.Add(strZahlBed)
            Dim intPayType As DataColumn = New DataColumn("intPayType")
            intPayType.DataType = System.Type.[GetType]("System.Int16")
            DT.Columns.Add(intPayType)
            Dim intintBank As DataColumn = New DataColumn("intintBank")
            intintBank.DataType = System.Type.[GetType]("System.Int16")
            intintBank.Caption = "IBank"
            DT.Columns.Add(intintBank)
            Dim intEBank As DataColumn = New DataColumn("intEBank")
            intEBank.DataType = System.Type.[GetType]("System.Int16")
            intEBank.Caption = "EBank"
            DT.Columns.Add(intEBank)
            Dim strKredStatusBitLog As DataColumn = New DataColumn("strKredStatusBitLog")
            strKredStatusBitLog.DataType = System.Type.[GetType]("System.String")
            strKredStatusBitLog.MaxLength = 50
            DT.Columns.Add(strKredStatusBitLog)
            Dim strKredStatusText As DataColumn = New DataColumn("strKredStatusText")
            strKredStatusText.DataType = System.Type.[GetType]("System.String")
            strKredStatusText.MaxLength = 255
            DT.Columns.Add(strKredStatusText)
            Dim strKredBookStatus As DataColumn = New DataColumn("strKredBookStatus")
            strKredBookStatus.DataType = System.Type.[GetType]("System.String")
            strKredBookStatus.MaxLength = 50
            DT.Columns.Add(strKredBookStatus)
            Dim booBooked As DataColumn = New DataColumn("booBooked")
            booBooked.DataType = System.Type.[GetType]("System.Boolean")
            DT.Columns.Add(booBooked)
            Dim datBooked As DataColumn = New DataColumn("datBooked")
            datBooked.DataType = System.Type.[GetType]("System.DateTime")
            DT.Columns.Add(datBooked)
            Dim lngBelegNr As DataColumn = New DataColumn("lngBelegNr")
            lngBelegNr.DataType = System.Type.[GetType]("System.Int32")
            DT.Columns.Add(lngBelegNr)
            Dim lngKrediKST As DataColumn = New DataColumn("lngKrediKST")
            lngKrediKST.DataType = System.Type.[GetType]("System.Int32")
            DT.Columns.Add(lngKrediKST)
            Dim intZKond As DataColumn = New DataColumn("intZKond")
            intZKond.DataType = System.Type.[GetType]("System.Int16")
            DT.Columns.Add(intZKond)
            Dim booPGV As DataColumn = New DataColumn("booPGV")
            booPGV.DataType = System.Type.[GetType]("System.Boolean")
            DT.Columns.Add(booPGV)
            Dim strPGVType As DataColumn = New DataColumn("strPGVType")
            strPGVType.DataType = System.Type.[GetType]("System.String")
            strPGVType.MaxLength = 2
            DT.Columns.Add(strPGVType)
            Dim datPGVFrom As DataColumn = New DataColumn("datPGVFrom")
            datPGVFrom.DataType = System.Type.[GetType]("System.DateTime")
            DT.Columns.Add(datPGVFrom)
            Dim intPGVMthsAY As DataColumn = New DataColumn("intPGVMthsAY")
            intPGVMthsAY.DataType = System.Type.[GetType]("System.Int32")
            DT.Columns.Add(intPGVMthsAY)
            Dim datPGVTo As DataColumn = New DataColumn("datPGVTo")
            datPGVTo.DataType = System.Type.[GetType]("System.DateTime")
            DT.Columns.Add(datPGVTo)
            Dim intPGVMthsNY As DataColumn = New DataColumn("intPGVMthsNY")
            intPGVMthsNY.DataType = System.Type.[GetType]("System.Int32")
            DT.Columns.Add(intPGVMthsNY)
            Return DT

        Catch ex As Exception
            MessageBox.Show(ex.Message + "Kreditoren-Head-Tabelle " + Err.Number.ToString)

        End Try


    End Function

    Public Shared Function tblKreditorenSub() As DataTable

        Dim DT As DataTable

        Try

            DT = New DataTable("tblKreditorenSub")
            Dim lngID As DataColumn = New DataColumn("lngID")
            lngID.DataType = System.Type.[GetType]("System.Int32")
            lngID.AutoIncrement = True
            lngID.AutoIncrementSeed = 1
            lngID.AutoIncrementStep = 1
            DT.Columns.Add(lngID)
            Dim lngKredID As DataColumn = New DataColumn("lngKredID")
            lngKredID.DataType = System.Type.[GetType]("System.Int32")
            lngKredID.Caption = "Kred-ID"
            DT.Columns.Add(lngKredID)
            Dim strRGNr As DataColumn = New DataColumn("strRGNr")
            strRGNr.DataType = System.Type.[GetType]("System.String")
            strRGNr.MaxLength = 50
            strRGNr.Caption = "RG-Nr"
            DT.Columns.Add(strRGNr)
            Dim intSollHaben As DataColumn = New DataColumn("intSollHaben")
            intSollHaben.DataType = System.Type.[GetType]("System.Int16")
            intSollHaben.Caption = "S1/H0"
            DT.Columns.Add(intSollHaben)
            Dim lngKto As DataColumn = New DataColumn("lngKto")
            lngKto.DataType = System.Type.[GetType]("System.Int32")
            lngKto.Caption = "Konto"
            DT.Columns.Add(lngKto)
            Dim strKtoBez As DataColumn = New DataColumn("strKtoBez")
            strKtoBez.DataType = System.Type.[GetType]("System.String")
            strKtoBez.MaxLength = 80
            strKtoBez.Caption = "Bezeichnung"
            DT.Columns.Add(strKtoBez)
            Dim lngKST As DataColumn = New DataColumn("lngKST")
            lngKST.DataType = System.Type.[GetType]("System.Int32")
            lngKST.Caption = "KST"
            DT.Columns.Add(lngKST)
            Dim strKstBez As DataColumn = New DataColumn("strKstBez")
            strKstBez.DataType = System.Type.[GetType]("System.String")
            strKstBez.MaxLength = 50
            strKstBez.Caption = "Bez."
            DT.Columns.Add(strKstBez)
            Dim dblNetto As DataColumn = New DataColumn("dblNetto")
            dblNetto.DataType = System.Type.[GetType]("System.Double")
            dblNetto.Caption = "Netto"
            DT.Columns.Add(dblNetto)
            Dim dblMwSt As DataColumn = New DataColumn("dblMwSt")
            dblMwSt.DataType = System.Type.[GetType]("System.Double")
            dblMwSt.Caption = "MwSt"
            DT.Columns.Add(dblMwSt)
            Dim dblBrutto As DataColumn = New DataColumn("dblBrutto")
            dblBrutto.DataType = System.Type.[GetType]("System.Double")
            dblBrutto.Caption = "Brutto"
            DT.Columns.Add(dblBrutto)
            Dim dblMwStSatz As DataColumn = New DataColumn("dblMwStSatz")
            dblMwStSatz.DataType = System.Type.[GetType]("System.Double")
            dblMwStSatz.Caption = "MwStS"
            DT.Columns.Add(dblMwStSatz)
            Dim strMwStKey As DataColumn = New DataColumn("strMwStKey")
            strMwStKey.DataType = System.Type.[GetType]("System.String")
            strMwStKey.MaxLength = 50
            DT.Columns.Add(strMwStKey)
            Dim strArtikel As DataColumn = New DataColumn("strArtikel")
            strArtikel.DataType = System.Type.[GetType]("System.String")
            strArtikel.MaxLength = 128
            DT.Columns.Add(strArtikel)
            Dim strKredSubText As DataColumn = New DataColumn("strKredSubText")
            strKredSubText.DataType = System.Type.[GetType]("System.String")
            strKredSubText.MaxLength = 125
            strKredSubText.Caption = "Buch-Text"
            DT.Columns.Add(strKredSubText)
            Dim strStatusUBBitLog As DataColumn = New DataColumn("strStatusUBBitLog")
            strStatusUBBitLog.DataType = System.Type.[GetType]("System.String")
            strStatusUBBitLog.MaxLength = 50
            DT.Columns.Add(strStatusUBBitLog)
            Dim strStatusUBText As DataColumn = New DataColumn("strStatusUBText")
            strStatusUBText.DataType = System.Type.[GetType]("System.String")
            strStatusUBText.MaxLength = 255
            strStatusUBText.Caption = "Status"
            DT.Columns.Add(strStatusUBText)
            Dim strKredBookStatus As DataColumn = New DataColumn("strKredBookStatus")
            strKredBookStatus.DataType = System.Type.[GetType]("System.String")
            strKredBookStatus.MaxLength = 50
            DT.Columns.Add(strKredBookStatus)
            Return DT

        Catch ex As Exception
            MessageBox.Show(ex.Message + "Kreditoren-Sub-Tabelle " + Err.Number.ToString)

        End Try
    End Function

    Public Shared Function tblInfo() As DataTable

        Try

            Dim DT As DataTable
            DT = New DataTable("tblDebitorenSub")
            Dim strInfoT As DataColumn = New DataColumn("strInfoT")
            strInfoT.DataType = System.Type.[GetType]("System.String")
            strInfoT.MaxLength = 50
            strInfoT.Caption = "Info-Titel"
            DT.Columns.Add(strInfoT)
            Dim strInfoV As DataColumn = New DataColumn("strInfoV")
            strInfoV.DataType = System.Type.[GetType]("System.String")
            strInfoV.MaxLength = 50
            strInfoV.Caption = "Info-Wert"
            DT.Columns.Add(strInfoV)
            Return DT

        Catch ex As Exception
            MessageBox.Show(ex.Message + "Info-Sub-Tabelle " + Err.Number.ToString)

        End Try


    End Function

    Public Shared Function FcLoginSage2(ByRef objdbconn As MySqlConnection,
                                       ByRef objsqlConn As SqlClient.SqlConnection,
                                       ByRef objsqlCom As SqlClient.SqlCommand,
                                       ByRef objFinanz As SBSXASLib.AXFinanz,
                                       ByRef objfiBuha As SBSXASLib.AXiFBhg,
                                       ByRef objdbBuha As SBSXASLib.AXiDbBhg,
                                       ByRef objdbPIFb As SBSXASLib.AXiPlFin,
                                       ByRef objFiBebu As SBSXASLib.AXiBeBu,
                                       ByRef objkrBuha As SBSXASLib.AXiKrBhg,
                                       ByVal intAccounting As Int16,
                                       ByRef objdtInfo As DataTable,
                                       ByVal strPeriod As String,
                                       ByRef strYear As String,
                                       ByRef intTeqNbr As Int16,
                                       ByRef intTeqNbrLY As Int16,
                                       ByRef intTeqNbrPLY As Int16,
                                       ByRef datPeriodFrom As Date,
                                       ByRef datPeriodTo As Date,
                                       ByRef strPeriodStatus As String) As Int16

        '0=ok, 1=Fibu nicht ok, 2=Debi nicht ok, 3=Debi nicht ok
        Dim booAccOk As Boolean
        Dim strMandant As String
        Dim strLogonInfo() As String
        Dim strPeriode() As String
        Dim FcReturns As Int16
        Dim intPeriodenNr As Int16
        Dim strPeriodenInfo As String
        Dim objdtPeriodeLY As New DataTable
        Dim strPeriodeLY As String
        Dim strPeriodePLY As String
        Dim objdbcmd As New MySqlCommand
        Dim dtPeriods As New DataTable


        Try

            objFinanz = Nothing
            objFinanz = New SBSXASLib.AXFinanz

            Application.DoEvents()

            'Login
            Call objFinanz.ConnectSBSdb(System.Configuration.ConfigurationManager.AppSettings("OwnSageServer"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageDB"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageID"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSagePsw"), "")

            'objdbconn.Open()
            strMandant = FcReadFromSettings(objdbconn,
                                            "Buchh200_Name",
                                            intAccounting)
            'objdbconn.Close()
            booAccOk = objFinanz.CheckMandant(strMandant)

            'Open Mandantg
            objFinanz.OpenMandant(strMandant, strPeriod)

            'Von Login aktuelle Periode auslesen
            strLogonInfo = Split(objFinanz.GetLogonInfo(), "{>}")
            objdtInfo.Rows.Add("Man/Periode", strMandant + "/" + strLogonInfo(7) + ", " + intAccounting.ToString)

            'Check Periode
            intPeriodenNr = objFinanz.ReadPeri(strMandant, strLogonInfo(7))
            strPeriodenInfo = objFinanz.GetPeriListe(0)

            strPeriode = Split(strPeriodenInfo, "{>}")

            'Teq-Nr von Vorjar lesen um in Suche nutzen zu können
            objdtPeriodeLY.Rows.Clear()
            strPeriodeLY = (Val(Left(strPeriode(4), 4)) - 1).ToString + Right(strPeriode(4), 4)
            objsqlCom.CommandText = "SELECT teqnbr FROM periode WHERE mandid='" + strMandant + "' AND dtebis='" + strPeriodeLY + "'"
            objsqlCom.Connection = objsqlConn
            objsqlConn.Open()
            objdtPeriodeLY.Load(objsqlCom.ExecuteReader)
            objsqlConn.Close()
            If objdtPeriodeLY.Rows.Count > 0 Then
                intTeqNbrLY = objdtPeriodeLY.Rows(0).Item("teqnbr")
            Else
                intTeqNbrLY = 0
            End If
            'Teq-Nr vom Vorvorjahr
            objdtPeriodeLY.Rows.Clear()
            strPeriodePLY = (Val(Left(strPeriode(4), 4)) - 2).ToString + Right(strPeriode(4), 4)
            objsqlCom.CommandText = "SELECT teqnbr FROM periode WHERE mandid='" + strMandant + "' AND dtebis='" + strPeriodePLY + "'"
            objsqlCom.Connection = objsqlConn
            objsqlConn.Open()
            objdtPeriodeLY.Load(objsqlCom.ExecuteReader)
            objsqlConn.Close()
            If objdtPeriodeLY.Rows.Count > 0 Then
                intTeqNbrPLY = objdtPeriodeLY.Rows(0).Item("teqnbr")
            Else
                intTeqNbrPLY = 0
            End If

            intTeqNbr = strPeriode(8)
            objdtInfo.Rows.Add("GeschäftsJ", strPeriode(3) + "-" + strPeriode(4) + ", teq: " + strPeriode(8).ToString + ", " + intTeqNbrLY.ToString + ", " + intTeqNbrPLY.ToString)
            objdtInfo.Rows.Add("Buchungen/ Status", strPeriode(5) + "-" + strPeriode(6) + "/ " + strPeriode(2))
            strYear = Strings.Left(strPeriode(4), 4)

            FcReturns = FcReadPeriodenDef(objsqlConn,
                                      objsqlCom,
                                      strPeriode(8),
                                      objdtInfo,
                                      strYear)

            'Perioden-Definition vom Tool einlesen
            'In einer ersten Phase nur erster DS einlesen
            objdbcmd.Connection = objdbconn
            objdbconn.Open()
            objdbcmd.CommandText = "SELECT * FROM t_sage_buchhaltungen_periods WHERE year=" + strYear + " AND refMandant=" + intAccounting.ToString
            dtPeriods.Load(objdbcmd.ExecuteReader)
            objdbconn.Close()
            If dtPeriods.Rows.Count > 0 Then
                datPeriodFrom = dtPeriods.Rows(0).Item("periodFrom")
                datPeriodTo = dtPeriods.Rows(0).Item("periodTo")
                strPeriodStatus = dtPeriods.Rows(0).Item("status")
            Else
                datPeriodFrom = Convert.ToDateTime(strYear + "-01-01 00:00:01")
                datPeriodTo = Convert.ToDateTime(strYear + "-12-31 23:59:59")
                strPeriodStatus = "O"
            End If
            objdtInfo.Rows.Add("Perioden", Format(datPeriodFrom, "dd.MM.yyyy hh:mm:ss") + " - " + Format(datPeriodTo, "dd.MM.yyyy hh:mm:ss") + "/ " + strPeriodStatus)

            'Finanz Buha öffnen
            If Not IsNothing(objfiBuha) Then
                objfiBuha = Nothing
            End If
            objfiBuha = New SBSXASLib.AXiFBhg
            objfiBuha = objFinanz.GetFibuObj()
            'Debitor öffnen
            If Not IsNothing(objdbBuha) Then
                objdbBuha = Nothing
            End If
            objdbBuha = New SBSXASLib.AXiDbBhg
            objdbBuha = objFinanz.GetDebiObj()
            If Not IsNothing(objdbPIFb) Then
                objdbPIFb = Nothing
            End If
            objdbPIFb = New SBSXASLib.AXiPlFin
            objdbPIFb = objfiBuha.GetCheckObj()
            If Not IsNothing(objFiBebu) Then
                objFiBebu = Nothing
            End If
            objFiBebu = New SBSXASLib.AXiBeBu
            objFiBebu = objFinanz.GetBeBuObj()
            'Kreditor
            If Not IsNothing(objkrBuha) Then
                objkrBuha = Nothing
            End If
            objkrBuha = New SBSXASLib.AXiKrBhg
            objkrBuha = objFinanz.GetKrediObj

            Application.DoEvents()

        Catch ex As Exception
            MsgBox("OpenMandant:" + vbCrLf + "Error" + vbCrLf + "Error # " + Str(Err.Number) + " was generated by " + Err.Source + vbCrLf + Err.Description + " Fehlernummer" & Str(Err.Number And 65535))
            Err.Clear()
            End

        Finally
            objdtPeriodeLY.Dispose()
            dtPeriods.Dispose()

        End Try

    End Function


    Public Shared Function FcLoginSage(ByRef objdbconn As MySqlConnection,
                                       ByRef objsqlConn As SqlClient.SqlConnection,
                                       ByRef objsqlCom As SqlClient.SqlCommand,
                                       ByRef objFinanz As SBSXASLib.AXFinanz,
                                       ByRef objfiBuha As SBSXASLib.AXiFBhg,
                                       ByRef objdbBuha As SBSXASLib.AXiDbBhg,
                                       ByRef objdbPIFb As SBSXASLib.AXiPlFin,
                                       ByRef objFiBebu As SBSXASLib.AXiBeBu,
                                       ByRef objkrBuha As SBSXASLib.AXiKrBhg,
                                       ByVal intAccounting As Int16,
                                       ByRef objdtInfo As DataTable,
                                       ByVal strPeriod As String,
                                       ByRef strYear As String,
                                       ByRef intTeqNbr As Int16,
                                       ByRef intTeqNbrLY As Int16) As Int16


        '0=ok, 1=Fibu nicht ok, 2=Debi nicht ok, 3=Debi nicht ok

        Dim booAccOk As Boolean
        Dim strMandant As String
        Dim b As Object
        Dim strLogonInfo() As String
        Dim strPeriode() As String
        Dim FcReturns As Int16
        Dim intPeriodenNr As Int16
        Dim strPeriodenInfo As String
        Dim objdtPeriodeLY As New DataTable
        Dim strPeriodeLY As String

        b = Nothing

        objFinanz = Nothing
        objFinanz = New SBSXASLib.AXFinanz


        On Error GoTo ErrorHandler

        'Login
        Call objFinanz.ConnectSBSdb(System.Configuration.ConfigurationManager.AppSettings("OwnSageServer"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageDB"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageID"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSagePsw"), "")

        objdbconn.Open()
        strMandant = FcReadFromSettings(objdbconn, "Buchh200_Name", intAccounting)
        objdbconn.Close()
        booAccOk = objFinanz.CheckMandant(strMandant)

        'Open Mandantg
        objFinanz.OpenMandant(strMandant, strPeriod)
        'Buha in Info schreiben
        'objdtInfo.Rows.Add("Buha", strMandant)

        'Von Login aktuelle Periode auslesen
        strLogonInfo = Split(objFinanz.GetLogonInfo(), "{>}")
        objdtInfo.Rows.Add("Man/Periode", strMandant + "/" + strLogonInfo(7))

        'Check Periode
        'booAccOk = objFinanz.CheckPeriode(strMandant, "2020")
        'strPeriodenInfo = objFinanz.GetLogonInfo()
        intPeriodenNr = objFinanz.ReadPeri(strMandant, strLogonInfo(7))
        'For intLooper As Int16 = 0 To intPeriodenNr
        strPeriodenInfo = objFinanz.GetPeriListe(0)
        'strPeriodenInfo = objFinanz.GetResource(intLooper)
        'Next

        strPeriode = Split(strPeriodenInfo, "{>}")
        'Teq-Nr von Vorjar lesen um in Suche nutzen zu können
        strPeriodeLY = (Val(Left(strPeriode(4), 4)) - 1).ToString + Right(strPeriode(4), 4)
        objsqlCom.CommandText = "SELECT teqnbr FROM periode WHERE mandid='" + strMandant + "' AND dtebis='" + strPeriodeLY + "'"
        objsqlCom.Connection = objsqlConn
        objsqlConn.Open()
        objdtPeriodeLY.Load(objsqlCom.ExecuteReader)
        objsqlConn.Close()
        'Variable übergeben, Achtung nicht definitiv. Situatin ist nicht klar wenn Vorjahr nicht existiert
        intTeqNbrLY = objdtPeriodeLY.Rows(0).Item("teqnbr")
        intTeqNbr = strPeriode(8)
        objdtInfo.Rows.Add("GeschäftsJ", strPeriode(3) + "-" + strPeriode(4) + ", teq: " + strPeriode(8).ToString + ", " + intTeqNbrLY.ToString)
        objdtInfo.Rows.Add("Buchungen/ Status", strPeriode(5) + "-" + strPeriode(6) + "/ " + strPeriode(2))
        strYear = Strings.Left(strPeriode(4), 4)
        'objdtInfo.Rows.Add("Status", strPeriode(2))
        'Debug.Print(FcReadPeriodenDef(objsqlConn, objsqlCom, strPeriode(8))(0))

        'objdtInfo.Rows.Add("Perioden-Def", FcReadPeriodenDef(objsqlConn, objsqlCom, strPeriode(8))(0))
        'objdtInfo.Rows.Add("Defintion von", FcReadPeriodenDef(objsqlConn, objsqlCom, strPeriode(8))(1))

        FcReturns = FcReadPeriodenDef(objsqlConn,
                                      objsqlCom,
                                      strPeriode(8),
                                      objdtInfo,
                                      strYear)


        If b = 0 Then GoTo isOk
        b = b - 200
        MsgBox("Mandant oder Periode falsch - Programm beendet", 0, "Fehler")
        objFinanz = Nothing
        End

isOk:
        'Finanz Buha öffnen
        objfiBuha = Nothing
        objfiBuha = New SBSXASLib.AXiFBhg
        objfiBuha = objFinanz.GetFibuObj()
        'Debitor öffnen
        objdbBuha = Nothing
        objdbBuha = New SBSXASLib.AXiDbBhg
        objdbBuha = objFinanz.GetDebiObj()
        objdbPIFb = Nothing
        objdbPIFb = objfiBuha.GetCheckObj()
        objFiBebu = Nothing
        objFiBebu = objFinanz.GetBeBuObj()
        'Kreditor
        objkrBuha = Nothing
        objkrBuha = New SBSXASLib.AXiKrBhg
        objkrBuha = objFinanz.GetKrediObj
        'db = Main_Renamed.Finanz.GetDebiObj
        'UPGRADE_WARNING: Couldn't resolve default property of object s. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        's = db.ReadDebitor3(CInt(DebitorID.Text), WhgID.Text)
        's = DbBhg.ReadDebitor3(-1000, "")
        'Debug.Print("Angaben Debitor " + s)
        'UPGRADE_NOTE: Object db may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'db = Nothing

        'UPGRADE_WARNING: Couldn't resolve default property of object s. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'AusgabeText.Text = "ReadDebitor3:" & Chr(13) & Chr(10) & s



        'MsgBox("OpenMandant:" & Chr(13) & Chr(10) & "Funktionierte")
        Exit Function


ErrorHandler:

        b = Err.Number And 65535
        MsgBox("OpenMandant:" & Chr(13) & Chr(10) & "Error" & Chr(13) & Chr(10) & "Error # " & Str(Err.Number) & " was generated by " & Err.Source & Chr(13) & Chr(10) & Err.Description & " Unsere Fehlernummer" & Str(b))
        Err.Clear()
        Resume Next

    End Function

    Public Shared Function FcReadPeriodsFromMandant(ByRef objdbconn As MySqlConnection,
                                                    ByRef objFinanz As SBSXASLib.AXFinanz,
                                                    ByVal intAccounting As Int16,
                                                    ByRef cmbPeriods As ComboBox) As Int16



        Dim strMandant As String
        Dim booAccOk As Int16
        Dim intLbNbr As Int16
        Dim strPeriodenListe As String = String.Empty
        Dim strPeriodeAr() As String
        Dim intLooper As Int16


        Try

            objFinanz = Nothing
            objFinanz = New SBSXASLib.AXFinanz

            'Login
            Call objFinanz.ConnectSBSdb(System.Configuration.ConfigurationManager.AppSettings("OwnSageServer"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageDB"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageID"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSagePsw"), "")

            'objdbconn.Open()
            strMandant = FcReadFromSettings(objdbconn, "Buchh200_Name", intAccounting)
            'objdbconn.Close()
            booAccOk = objFinanz.CheckMandant(strMandant)

            'Combo leeren
            cmbPeriods.Items.Clear()

            'GJ einlesen
            intLbNbr = objFinanz.ReadPeri(strMandant, "")
            Do Until strPeriodenListe = "EOF"
                strPeriodenListe = objFinanz.GetPeriListe(intLooper)
                strPeriodeAr = Split(strPeriodenListe, "{>}")
                If strPeriodenListe <> "EOF" Then
                    cmbPeriods.Items.Add(strPeriodeAr(0))
                End If
                intLooper += 1
            Loop

            'Auf aktuelles Jahr gehen
            'Bei Jahresanfang
            'cmbPeriods.SelectedIndex = cmbPeriods.Items.IndexOf((DateAndTime.Year("2022-12-31")).ToString)
            cmbPeriods.SelectedIndex = cmbPeriods.Items.IndexOf((DateAndTime.Year(DateAndTime.Now())).ToString)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Periodendefinition lesen")

        End Try


    End Function

    Public Shared Function FcReadPeriodenDef(ByRef objSQLConnection As SqlClient.SqlConnection,
                                             ByRef objSQLCommand As SqlClient.SqlCommand,
                                             ByVal intPeriodenNr As Int32,
                                             ByRef objdtInfo As DataTable,
                                             ByVal strYear As String) As Int16

        'Returns 0=definiert, 1=nicht defeniert, 9=Problem
        Dim objlocdtPeriDef As New DataTable
        Dim strPeriodenDef(4) As String


        Try

            objSQLConnection.Open()
            objSQLCommand.CommandText = "SELECT * FROM peridef WHERE teqnbr=" + intPeriodenNr.ToString
            objSQLCommand.Connection = objSQLConnection
            objlocdtPeriDef.Load(objSQLCommand.ExecuteReader)

            If objlocdtPeriDef.Rows.Count > 0 Then 'Perioden-Definition vorhanden

                strPeriodenDef(0) = objlocdtPeriDef.Rows(0).Item(2) 'Bezeichnung
                strPeriodenDef(1) = objlocdtPeriDef.Rows(0).Item(3).ToString  'Von
                strPeriodenDef(2) = objlocdtPeriDef.Rows(0).Item(4).ToString  'Bis
                strPeriodenDef(3) = objlocdtPeriDef.Rows(0).Item(5)  'Status

                objdtInfo.Rows.Add("Perioden S200", strPeriodenDef(0))
                objdtInfo.Rows.Add("Von - Bis/ Status", Format(Convert.ToDateTime(strPeriodenDef(1)), "dd.MM.yyyy hh:mm:ss") + " - " + Format(Convert.ToDateTime(strPeriodenDef(2)), "dd.MM.yyyy hh:mm:ss") + "/ " + strPeriodenDef(3))

                Return 0
            Else

                objdtInfo.Rows.Add("Perioden S200", "keine")
                objdtInfo.Rows.Add("Von - Bis/ Status", Format(Convert.ToDateTime("01.01." + strYear + " 00:00:00"), "dd.MM.yyyy hh:mm:ss") + " - " + Format(Convert.ToDateTime("31.12." + strYear + " 23:59:59"), "dd.MM.yyyy hh:mm:ss") + "/ " + "O")

                Return 1

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Periodendefinition lesen")
            Return 9

        Finally
            objSQLConnection.Close()
            objlocdtPeriDef.Constraints.Clear()
            objlocdtPeriDef.Clear()
            objlocdtPeriDef.Dispose()
            strPeriodenDef = Nothing

        End Try

    End Function

    Public Shared Function FcReadBankSettings(ByVal intAccounting As Int16,
                                              ByVal intPayType As Int16,
                                              ByVal strBank As String,
                                              ByRef objdbconn As MySqlConnection) As String

        Dim objlocdtBank As New DataTable("tbllocBank")
        Dim objlocMySQLcmd As New MySqlCommand

        Try


            If intPayType = 10 Then
                objlocMySQLcmd.CommandText = "SELECT strBLZ FROM t_sage_tblaccountingbank WHERE intAccountingID=" + intAccounting.ToString + " AND QRTNNR='" + strBank + "'"
            Else
                objlocMySQLcmd.CommandText = "SELECT strBLZ FROM t_sage_tblaccountingbank WHERE intAccountingID=" + intAccounting.ToString + " AND strBank='" + strBank + "'"
            End If

            objlocMySQLcmd.Connection = objdbconn
            objlocdtBank.Load(objlocMySQLcmd.ExecuteReader)

            Return objlocdtBank.Rows(0).Item(0).ToString


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Bankleitzahl suchen.")

        Finally
            objlocdtBank.Dispose()
            objlocMySQLcmd.Dispose()

        End Try


    End Function


    Public Shared Function FcReadFromSettings(ByRef objdbconn As MySqlConnection,
                                              ByVal strField As String,
                                              ByVal intMandant As Int16) As String

        Dim objlocdtSetting As New DataTable("tbllocSettings")
        Dim objlocMySQLcmd As New MySqlCommand

        Try

            objdbconn.Open()
            objlocMySQLcmd.CommandText = "SELECT t_sage_buchhaltungen." + strField + " FROM t_sage_buchhaltungen WHERE Buchh_Nr=" + intMandant.ToString
            'Debug.Print(objlocMySQLcmd.CommandText)
            objlocMySQLcmd.Connection = objdbconn
            objlocdtSetting.Load(objlocMySQLcmd.ExecuteReader)
            'Debug.Print("Records" + objlocdtSetting.Rows.Count.ToString)
            'Debug.Print("Return " + objlocdtSetting.Rows(0).Item(0).ToString)
            Return objlocdtSetting.Rows(0).Item(0).ToString


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Einstellung lesen")

        Finally
            objlocdtSetting.Dispose()
            objlocMySQLcmd.Dispose()
            objdbconn.Close()

        End Try


    End Function

    Public Shared Function FcReadFromSettingsNC(ByVal strField As String,
                                                ByVal intMandant As Int16) As String

        Dim objlocdtSetting As New DataTable("tbllocSettings")
        Dim objlocMySQLcmd As New MySqlCommand
        Dim objdbConnLoc As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString"))

        Try

            objdbconnLoc.Open()
            objlocMySQLcmd.CommandText = "SELECT t_sage_buchhaltungen." + strField + " FROM t_sage_buchhaltungen WHERE Buchh_Nr=" + intMandant.ToString
            'Debug.Print(objlocMySQLcmd.CommandText)
            objlocMySQLcmd.Connection = objdbConnLoc
            objlocdtSetting.Load(objlocMySQLcmd.ExecuteReader)
            'Debug.Print("Records" + objlocdtSetting.Rows.Count.ToString)
            'Debug.Print("Return " + objlocdtSetting.Rows(0).Item(0).ToString)
            Return objlocdtSetting.Rows(0).Item(0).ToString


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Einstellung lesen")

        Finally
            objlocdtSetting.Constraints.Clear()
            objlocdtSetting.Clear()
            objlocdtSetting = Nothing

            objlocMySQLcmd.Dispose()
            objdbConnLoc.Close()

        End Try


    End Function


    Public Shared Function FcCheckDebit(ByVal intAccounting As Integer,
                                        ByRef objdtDebits As DataTable,
                                        ByRef objdtDebitSubs As DataTable,
                                        ByRef objFinanz As SBSXASLib.AXFinanz,
                                        ByRef objfiBuha As SBSXASLib.AXiFBhg,
                                        ByRef objdbBuha As SBSXASLib.AXiDbBhg,
                                        ByRef objdbPIFb As SBSXASLib.AXiPlFin,
                                        ByRef objFiBebu As SBSXASLib.AXiBeBu,
                                        ByRef objdbconn As MySqlConnection,
                                        ByRef objdbconnZHDB02 As MySqlConnection,
                                        ByRef objsqlcommand As MySqlCommand,
                                        ByRef objsqlcommandZHDB02 As MySqlCommand,
                                        ByRef objOrdbconn As OracleClient.OracleConnection,
                                        ByRef objOrcommand As OracleClient.OracleCommand,
                                        ByRef objdbAccessConn As OleDb.OleDbConnection,
                                        ByRef objdbSQLConn As SqlConnection,
                                        ByRef objdbSQLCmd As SqlCommand,
                                        ByRef objdtInfo As DataTable,
                                        ByVal strcmbBuha As String,
                                        ByVal intTeqNbr As Int16,
                                        ByVal intTeqNbrLY As Int16,
                                        ByVal intTeqNbrPLY As Int16,
                                        ByVal strYear As String,
                                        ByVal strPeriode As String,
                                        ByVal datPeriodFrom As Date,
                                        ByVal datPeriodTo As Date,
                                        ByVal strPeriodStatus As String,
                                        ByVal booValutaCorrect As Boolean,
                                        ByVal datValutaCorrect As Date) As Integer

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


        'Dim objdrDebiSub As DataRow = objdtDebitSubs.NewRow

        Try

            'Teq-Nbr extrahieren
            'intTeqNbr = Conversion.Val(Strings.Right(objdtInfo.Rows(1).Item(1), 3))

            'objdbconn.Open()
            'objOrdbconn.Open()
            'objdbAccessConn.Open()

            For Each row As DataRow In objdtDebits.Rows

                'If row("strDebRGNbr") = "101261" Then Stop
                strRGNbr = row("strDebRGNbr") 'Für Error-Msg

                'Runden
                row("dblDebNetto") = Decimal.Round(row("dblDebNetto"), 4, MidpointRounding.AwayFromZero)
                row("dblDebMwSt") = Decimal.Round(row("dblDebMwst"), 2, MidpointRounding.AwayFromZero)
                row("dblDebBrutto") = Decimal.Round(row("dblDebBrutto"), 4, MidpointRounding.AwayFromZero)
                'OP - Nummer nicht numerische Zeichen entfernen
                'row("strOPNr") = Main.FcCleanRGNrStrict(row("strOPNr"))
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
                intReturnValue = MainDebitor.FcGetRefDebiNr(objdbconn,
                                                objdbconnZHDB02,
                                                objsqlcommand,
                                                objsqlcommandZHDB02,
                                                objOrdbconn,
                                                objOrcommand,
                                                objdbAccessConn,
                                                IIf(IsDBNull(row("lngDebNbr")), 0, row("lngDebNbr")),
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
                Application.DoEvents()

                'Kto 02
                'intReturnValue = FcCheckKonto(row("lngDebKtoNbr"), objfiBuha, row("dblDebMwSt"), 0)
                intReturnValue = 0
                strBitLog += Trim(intReturnValue.ToString)

                'Currency 03
                intReturnValue = FcCheckCurrency(row("strDebCur"), objfiBuha)
                strBitLog += Trim(intReturnValue.ToString)
                Application.DoEvents()

                'Sub 04
                booAutoCorrect = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettings(objdbconn, "Buchh_HeadAutoCorrect", intAccounting)))
                booCpyKSTToSub = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettings(objdbconn, "Buchh_KSTHeadToSub", intAccounting)))
                booSplittBill = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettings(objdbconn, "Buchh_LinkedBookings", intAccounting)))
                If booSplittBill And IIf(IsDBNull(row("intRGArt")), 0, row("intRGArt")) = 10 Then
                    row("booLinked") = True

                Else
                    row("booLinked") = False
                End If
                booCashSollCorrect = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettings(objdbconn, "Buchh_CashSollKontoKorr", intAccounting)))
                intReturnValue = FcCheckSubBookings(row("strDebRGNbr"),
                                                    objdtDebitSubs,
                                                    intSubNumber,
                                                    dblSubBrutto,
                                                    dblSubNetto,
                                                    dblSubMwSt,
                                                    objdbconn,
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

                strBitLog += Trim(intReturnValue.ToString)

                'Gibt es eine Bezahlbuchung zu erstellen? 
                booGeneratePymentBooking = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettings(objdbconn, "Buchh_GeneratePaymentBooking", intAccounting)))
                If booGeneratePymentBooking And row("intBuchungsart") <> 1 And row("intKtoPayed") > 0 Then
                    'Bedingungen erfüllt
                    Dim drPaymentBuchung As DataRow = objdtDebitSubs.NewRow
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
                    objdtDebitSubs.Rows.Add(drPaymentBuchung)
                    drPaymentBuchung = Nothing
                    'Summe der Sub-Buchungen anpassen
                    dblSubBrutto = Decimal.Round(dblSubBrutto + row("dblDebBrutto"), 2, MidpointRounding.AwayFromZero)

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
                        'row("dblDebNetto") = dblSubNetto * -1
                        'row("dblDebMwSt") = dblSubMwSt * -1
                        If IIf(IsDBNull(row("dblDebMwSt")), 0, row("dblDebMwSt")) <> dblSubMwSt * -1 Then
                            row("dblDebMwSt") = dblSubMwSt * -1
                        End If
                        If IIf(IsDBNull(row("dblDebNetto")), 0, row("dblDebNetto")) <> dblSubNetto * -1 Then
                            row("dblDebNetto") = dblSubNetto * -1
                        End If

                        'Für evtl. Rundungsdifferenzen einen Datensatz in die Sub-Tabelle hinzufügen
                        If IIf(IsDBNull(row("dblDebBrutto")), 0, row("dblDebBrutto")) + dblSubBrutto <> 0 Then '0 _
                            'Or IIf(IsDBNull(row("dblDebMwSt")), 0, row("dblDebMwSt")) + dblSubMwSt <> 0 _
                            'Or IIf(IsDBNull(row("dblDebNetto")), 0, row("dblDebNetto")) + dblSubNetto <> 0 Then

                            'row("dblDebNetto") = dblSubNetto * -1
                            'row("dblDebMwSt") = dblSubMwSt * -1


                            dblRDiffBrutto = Decimal.Round(dblSubBrutto + row("dblDebBrutto"), 4, MidpointRounding.AwayFromZero)
                            dblRDiffMwSt = 0 'row("dblDebMwSt") - Decimal.Round(dblSubMwSt, 2, MidpointRounding.AwayFromZero)
                            dblRDiffNetto = 0 ' row("dblDebNetto") - Decimal.Round(dblSubNetto, 2, MidpointRounding.AwayFromZero)

                            'Zu sub-Table hinzifügen
                            Dim objdrDebiSub As DataRow = objdtDebitSubs.NewRow
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
                            objdtDebitSubs.Rows.Add(objdrDebiSub)
                            objdrDebiSub = Nothing
                            'Summe der Sub-Buchungen anpassen
                            dblSubBrutto = Decimal.Round(dblSubBrutto - dblRDiffBrutto, 2, MidpointRounding.AwayFromZero)
                            'dblSubMwSt = Decimal.Round(dblSubMwSt - dblRDiffMwSt, 2, MidpointRounding.AwayFromZero)
                            'dblSubNetto = Decimal.Round(dblSubNetto - dblRDiffNetto, 2, MidpointRounding.AwayFromZero)
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
                    intReturnValue = FcCreateDebRef(objdbconn,
                                                    intAccounting,
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
                booPKPrivate = IIf(FcReadFromSettings(objdbconn, "Buchh_PKTable", intAccounting) = "t_customer", True, False)
                'Debitor
                If Left(strBitLog, 1) <> "0" Then
                    strStatus += "Deb"
                    If Left(strBitLog, 1) <> "2" Then
                        If booPKPrivate = True Then
                            intReturnValue = MainDebitor.FcIsPrivateDebitorCreatable(objdbconn,
                                                                                     objdbconnZHDB02,
                                                                                     objsqlcommandZHDB02,
                                                                                     intDebitorNew,
                                                                                     objdbBuha,
                                                                                     strcmbBuha,
                                                                                     intAccounting)
                        Else
                            intReturnValue = MainDebitor.FcIsDebitorCreatable(objdbconn,
                                                                              objdbconnZHDB02,
                                                                              objsqlcommand,
                                                                              objsqlcommandZHDB02,
                                                                              intDebitorNew,
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
                If Year(row("datDebRGDatum")) <> Year(row("datDebValDatum")) And Year(row("datDebValDatum")) >= 2022 Then
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
                            row("datPGVFrom") = Year(datValutaSave).ToString + "-" + Month(datValutaSave).ToString + "-" + Day(datValutaSave).ToString
                            row("datPGVTo") = Year(datValutaSave).ToString + "-" + Month(datValutaSave).ToString + "-" + Day(datValutaSave).ToString
                            row("datDebValDatum") = "2023-01-01"
                            booDateChanged = True
                        ElseIf row("strPGVType") = "RV" Then
                            row("datPGVFrom") = Year(datValutaSave).ToString + "-" + Month(datValutaSave).ToString + "-" + Day(datValutaSave).ToString
                            row("datPGVTo") = Year(datValutaSave).ToString + "-" + Month(datValutaSave).ToString + "-" + Day(datValutaSave).ToString
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
                intReturnValue = FcChCeckDate(IIf(IsDBNull(row("datDebValDatum")), #1789-09-17#, row("datDebValDatum")),
                                              objdtInfo,
                                              datPeriodFrom,
                                              datPeriodTo,
                                              strPeriodStatus,
                                              True)
                'Falls Problem versuchen mit Valuta-Datum-Anpassung
                If intReturnValue <> 0 And booValutaCorrect Then
                    row("datDebValDatum") = Format(datValutaCorrect, "Short Date")
                    booDateChanged = True
                    intReturnValue = FcChCeckDate(IIf(IsDBNull(row("datDebValDatum")), #1789-09-17#, row("datDebValDatum")),
                                              objdtInfo,
                                              datPeriodFrom,
                                              datPeriodTo,
                                              strPeriodStatus,
                                              True)
                    If intReturnValue = 0 Then
                        'Korrektur hat funktioniert Wert auf 2 setzen
                        intReturnValue = 2
                    Else
                        intReturnValue = 3
                    End If

                End If
                'Bei PGV checken ob PGV-Startdatum in blockierter Periode
                If row("booPGV") And intReturnValue = 0 Then
                    intReturnValue = FcCheckPGVDate(row("datPGVFrom"),
                                                    intAccounting,
                                                    objdbconnZHDB02)
                    If intReturnValue <> 0 Then
                        'Falls TA-Buchung in blockierter Periode probieren mit Valuta-Korrektur
                        If intPGVMonths = 1 And booValutaCorrect Then
                            row("datDebValDatum") = Format(datValutaCorrect, "Short Date")
                            booDateChanged = True
                            intReturnValue = FcChCeckDate(IIf(IsDBNull(row("datDebValDatum")), #1789-09-17#, row("datDebValDatum")),
                                              objdtInfo,
                                              datPeriodFrom,
                                              datPeriodTo,
                                              strPeriodStatus,
                                              True)
                            If intReturnValue = 0 Then
                                'PGV - Flag entfernen
                                row("booPGV") = False
                                intReturnValue = 5
                            Else
                                intReturnValue = 3
                            End If
                        Else
                            intReturnValue = 4
                        End If
                    End If

                End If
                strBitLog += Trim(intReturnValue.ToString)

                'RG - Datum 11
                intReturnValue = FcChCeckDate(IIf(IsDBNull(row("datDebRGDatum")), #1789-09-17#, row("datDebRGDatum")),
                                              objdtInfo,
                                              datPeriodFrom,
                                              datPeriodTo,
                                              strPeriodStatus,
                                              True)

                'Falls Problem versuchen mit Valuta-Datum-Anpassung
                If intReturnValue <> 0 And booValutaCorrect Then
                    row("datDebRGDatum") = datValutaCorrect
                    booDateChanged = True
                    intReturnValue = FcChCeckDate(IIf(IsDBNull(row("datDebRGDatum")), #1789-09-17#, row("datDebRGDatum")),
                                              objdtInfo,
                                              datPeriodFrom,
                                              datPeriodTo,
                                              strPeriodStatus,
                                              True)
                    If intReturnValue = 0 Then
                        'Korrektur hat funktioniert Wert auf 2 setzen
                        intReturnValue = 2
                    Else
                        intReturnValue = 3
                    End If

                End If
                strBitLog += Trim(intReturnValue.ToString)
                'Falls ein Datum geändert wurde dann Flag setzen
                If booDateChanged Then
                    row("booDatChanged") = True
                End If

                'Interne Bank 12
                'If row.Table.Columns("intPayType") Is Nothing Then
                '    row("intPayType") = 3
                'End If
                If IsDBNull(row("intPayType")) Then
                    row("intPayType") = 9
                End If
                intReturnValue = MainDebitor.FcCheckDebiIntBank(objdbconn,
                                                                intAccounting,
                                                                IIf(IsDBNull(row("strDebiBank")), "", row("strDebiBank")),
                                                                row("intPayType"),
                                                                intiBankSage200)
                strBitLog += Trim(intReturnValue.ToString)

                'Bei SplittBill: Existiert verlinkter Beleg? 13
                If row("booLinked") Then
                    'Zuerst Debitor von erstem Beleg suchen
                    intDebitorNew = MainDebitor.FcGetDebitorFromLinkedRG(objdbconn,
                                                                         objdbconnZHDB02,
                                                                         objsqlcommand,
                                                                         objsqlcommandZHDB02,
                                                                         objOrdbconn,
                                                                         objOrcommand,
                                                                         objdbAccessConn,
                                                                         objdbSQLConn,
                                                                         objdbSQLCmd,
                                                                         IIf(IsDBNull(row("lngLinkedRG")), 0, row("lngLinkedRG")),
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
                    selSBrows = objdtDebitSubs.Select("strRGNr='" + row("strDebRGNbr") + "'")

                    For Each SBsubrow As DataRow In selSBrows
                        SBsubrow.Delete()
                    Next

                    Dim drSBBuchung As DataRow = objdtDebitSubs.NewRow
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
                    objdtDebitSubs.Rows.Add(drSBBuchung)
                    drSBBuchung = Nothing

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
                    intReturnValue = MainDebitor.FcGetDZKondSageID(objdbconn,
                                                                   objdbconnZHDB02,
                                                                   objsqlcommandZHDB02,
                                                                   row("intZKond"),
                                                                   intDZKondS200)
                    row("intZKond") = intDZKondS200
                End If
                If row("intZKondT") = 1 And row("intZKond") = 0 Then
                    'Fall kein Privatekunde
                    If booPKPrivate = False Then
                        'Daten aus den Tab_Repbetriebe holen
                        intReturnValue = MainDebitor.FcGetDZkondFromRep(objdbconn,
                                                                    objdbconnZHDB02,
                                                                    objsqlcommandZHDB02,
                                                                    row("lngDebNbr"),
                                                                    intDZKond,
                                                                    intAccounting)
                    Else
                        'Daten aus der t_customer holen
                        intReturnValue = MainDebitor.FcGetDZkondFromCust(objdbconn,
                                                                         objdbconnZHDB02,
                                                                         objsqlcommandZHDB02,
                                                                         row("lngDebNbr"),
                                                                         intDZKond,
                                                                         intAccounting)
                    End If
                    row("intZKond") = intDZKond
                End If
                'Prüfem ob Zahlungs-Kondition - ID existiert in Sage 200 bei Mandant
                strMandant = FcReadFromSettings(objdbconn,
                                                "Buchh200_Name",
                                                intAccounting)
                intReturnValue = MainDebitor.FcCheckDZKond(objdbSQLConn,
                                                           objdbSQLCmd,
                                                           strMandant,
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
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "VDCor"
                        'Korrektur hat geklappt, Wert wieder auf 0 setzen
                        strBitLog = Left(strBitLog, 9) + "0" + Right(strBitLog, Len(strBitLog) - 10)
                    ElseIf Mid(strBitLog, 10, 1) = "3" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "VDCorNok"
                    ElseIf Mid(strBitLog, 10, 1) = "4" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "PGVDblckVD"
                    ElseIf Mid(strBitLog, 10, 1) = "5" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "PGVVDCor"
                        'Korrektur hat geklappt, Wert wieder auf 0 setzen
                        strBitLog = Left(strBitLog, 9) + "0" + Right(strBitLog, Len(strBitLog) - 10)
                    End If
                End If
                'RG Datum 
                If Mid(strBitLog, 11, 1) <> "0" Then
                    If Mid(strBitLog, 11, 1) = "1" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "RgD"
                    ElseIf Mid(strBitLog, 11, 1) = "2" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "RgDCor"
                        'Korrektur hat geklappt, Wert wieder auf 0 setzen
                        strBitLog = Left(strBitLog, 10) + "0" + Right(strBitLog, Len(strBitLog) - 11)
                    ElseIf Mid(strBitLog, 11, 1) = "3" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "RgDCorNok"
                    ElseIf Mid(strBitLog, 11, 1) = "4" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "PGVDblckRD"
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
                booDiffHeadText = IIf(FcReadFromSettings(objdbconn, "Buchh_TextSpecial", intAccounting) = "0", False, True)
                If booDiffHeadText Then
                    strDebiHeadText = MainDebitor.FcSQLParse(FcReadFromSettings(objdbconn,
                                                                                "Buchh_TextSpecialText",
                                                                                intAccounting),
                                                             row("strDebRGNbr"),
                                                             objdtDebits,
                                                             "D")
                    row("strDebText") = strDebiHeadText
                End If

                'Wird ein anderer Text in den Sub-Buchung gewünscht?
                booDiffSubText = IIf(FcReadFromSettings(objdbconn, "Buchh_SubTextSpecial", intAccounting) = "0", False, True)
                If booDiffSubText And Not row("booLinked") Then
                    strDebiSubText = MainDebitor.FcSQLParse(FcReadFromSettings(objdbconn,
                                                                               "Buchh_SubTextSpecialText",
                                                                               intAccounting),
                                                            row("strDebRGNbr"),
                                                            objdtDebits,
                                                            "D")
                Else
                    strDebiSubText = row("strDebText")
                End If
                'Falls nicht SB - Linked dann Text in SB ersetzen
                If Not row("booLinked") Then
                    selsubrow = objdtDebitSubs.Select("strRGNr='" + row("strDebRGNbr") + "'")
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

                Application.DoEvents()

            Next

        Catch ex As Exception
            MessageBox.Show(ex.Message + vbCrLf + "Auf RG " + strRGNbr, "Debitor Kopfdaten-Check", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        Finally

            If objOrdbconn.State = ConnectionState.Open Then
                objOrdbconn.Close()
            End If
            If objdbconn.State = ConnectionState.Open Then
                objdbconn.Close()
            End If
            If objdbAccessConn.State = ConnectionState.Open Then
                objdbAccessConn.Close()
            End If

            selsubrow = Nothing
            selSBrows = Nothing

        End Try


    End Function

    Public Shared Function FcCheckPGVDate(ByVal datPGVDateToCheck As Date,
                                          ByVal intAccounting As Int16,
                                          ByRef objdbconnZHDB02 As MySqlConnection)


        Dim tblPeriods As New DataTable
        Dim intYearToCheck As Int16
        Dim objsqlcommand As New MySqlCommand

        Try

            'Zuerst testen ob überhaupt eine Definition fürs Jahr existiert
            intYearToCheck = Year(datPGVDateToCheck)

            objsqlcommand.CommandText = "SELECT * FROM t_sage_buchhaltungen_periods WHERE year=" + intYearToCheck.ToString + " AND refMandant=" + intAccounting.ToString
            objsqlcommand.Connection = objdbconnZHDB02
            If objdbconnZHDB02.State = ConnectionState.Closed Then
                objdbconnZHDB02.Open()
            End If
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

        End Try

    End Function

    Public Shared Function FcChCeckDate(ByVal datDateToCheck As Date,
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
            Application.DoEvents()

        End Try


    End Function

    Public Shared Function FcCheckOPDouble(ByRef objdbBuha As SBSXASLib.AXiDbBhg,
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
            Application.DoEvents()

        End Try

    End Function


    Public Shared Function FcCreateDebRef(ByRef objdbconn As MySqlConnection,
                                          ByVal intAccounting As Integer,
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

                strRefFrom = FcReadFromSettings(objdbconn, "Buchh_ESRNrFrom", intAccounting)
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
                                              strBank,
                                              objdbconn)

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

        End Try


    End Function

    Public Shared Function FcModulo10(ByVal strNummer As String) As Integer

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


    Public Shared Function FcCleanRGNrStrict(ByVal strRGNrToClean As String) As String

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

    Public Shared Function FcCheckBelegHead(ByVal intBuchungsArt As Int16,
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

    Public Shared Function FcCheckProj(ByRef objFiBebu As SBSXASLib.AXiBeBu,
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


    Public Shared Function FcCheckMwSt(ByRef objdbconn As MySqlConnection,
                                       ByRef objFiBhg As SBSXASLib.AXiFBhg,
                                       ByVal strStrCode As String,
                                       ByVal dblStrWert As Double,
                                       ByRef strStrCode200 As String,
                                       ByVal intKonto As Int32) As Integer

        'returns 0=ok, 1=nicht gefunden

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

            objlocMySQLcmd.Connection = objdbconn
            objlocMySQLcmd.Connection.Open()
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
            objlocMySQLcmd.Connection.Close()

        End Try


    End Function

    Public Shared Function FcCheckSubBookings(ByVal strDebRgNbr As String,
                                              ByRef objDtDebiSub As DataTable,
                                              ByRef intSubNumber As Int16,
                                              ByRef dblSubBrutto As Double,
                                              ByRef dblSubNetto As Double,
                                              ByRef dblSubMwSt As Double,
                                              ByRef objdbconn As MySqlConnection,
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
                    intReturnValue = FcCheckMwSt(objdbconn,
                                                 objFiBhg,
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
                            'Stop
                            '                            If booAutoCorrect Then 'And Val(strSteuer(2)) - subrow("dblMwst") <= 1.5 Then
                            'Falls MwSt-Betrag nur in 3 und 4 Stelle anders, dann erfassten Betrag nehmen.
                            If Math.Abs(Val(strSteuer(2)) - IIf(IsDBNull(subrow("dblMwst")), 0, subrow("dblMwst"))) >= 0.01 Then
                                strStatusText += "MwSt " + subrow("dblMwst").ToString
                                subrow("dblMwst") = Val(strSteuer(2))
                                'subrow("dblMwStSatz") = Val(strSteuer(3))
                                'subrow("dblBrutto") = Decimal.Round(subrow("dblNetto") + subrow("dblMwSt"), 2, MidpointRounding.AwayFromZero)
                                'subrow("dblNetto") = Decimal.Round(subrow("dblBrutto") + subrow("dblMwSt"), 2, MidpointRounding.AwayFromZero)
                                strStatusText += " cor -> " + subrow("dblMwst").ToString + ", "
                                '                           Else
                                '                          If Val(strSteuer(2)) - subrow("dblMwst") > 10 Then
                                '                         strStatusText += " -> " + strSteuer(2).ToString + ", "
                                '                        intReturnValue = 1
                                '                   Else
                                '                      strStatusText += " Tol -> " + strSteuer(2).ToString + ", "
                                '                 End If
                                '                End If
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
                                                   objFiPI,
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
                    ElseIf mid(strBitLog, 2, 1) = "5" Then
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
                Application.DoEvents()

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

        End Try

    End Function

    Public Shared Function FcCheckKrediSubBookings(ByVal lngKredID As Int32,
                                              ByRef objDtKrediSub As DataTable,
                                              ByRef intSubNumber As Int16,
                                              ByRef dblSubBrutto As Double,
                                              ByRef dblSubNetto As Double,
                                              ByRef dblSubMwSt As Double,
                                              ByRef objdbconn As MySqlConnection,
                                              ByRef objFiBhg As SBSXASLib.AXiFBhg,
                                              ByRef objFiPI As SBSXASLib.AXiPlFin,
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

                Application.DoEvents()

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
                            subrow("dblMwStSatz") = 0
                            subrow("strMwStKey") = Nothing
                        End If
                    End If
                End If

                'MwSt prüfen 01
                If Not IsDBNull(subrow("strMwStKey")) Then
                    'If subrow("dblMwStSatz") > 0 And subrow("dblMwSt") = 0 Then Stop
                    If subrow("dblMwStSatz") > 0 And subrow("dblMwSt") > 0 And subrow("strMwStKey") = "ohne" Then Stop
                    If subrow("dblMwStSatz") = 0 And subrow("dblMwSt") = 0 And subrow("strMwStKey") <> "ohne" Then
                        subrow("strMwStKey") = "ohne"
                    End If
                    intReturnValue = FcCheckMwStToCorrect(objdbconn,
                                                          subrow("strMwStKey"),
                                                          subrow("dblMwStSatz"),
                                                          subrow("dblMwSt"))
                    intReturnValue = FcCheckMwSt(objdbconn,
                                                 objFiBhg,
                                                 subrow("strMwStKey"),
                                                 subrow("dblMwStSatz"),
                                                 strStrStCodeSage200,
                                                 subrow("lngKto"))
                    If intReturnValue = 0 Then
                        subrow("strMwStKey") = strStrStCodeSage200
                        'Check of korrekt berechnet
                        strSteuer = Split(objFiBhg.GetSteuerfeld(subrow("lngKto").ToString, "Zum Rechnen", subrow("dblBrutto").ToString, strStrStCodeSage200), "{<}")
                        If Val(strSteuer(2)) <> subrow("dblMwst") Then
                            'Im Fall von Auto-Korrekt anpassen
                            'Stop
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
                    intReturnValue = FcCheckKstKtr(subrow("lngKST"), objFiBhg, objFiPI, subrow("lngKto"), strKstKtrSage200)
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
                Application.DoEvents()

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


    Public Shared Function FcCheckMwStToCorrect(ByRef objdbconn As MySqlConnection,
                                                ByVal strStrCode As String,
                                                ByRef dblStrWert As Double,
                                                ByVal dblStrAmount As Double) As Integer

        Dim objlocdtMwSt As New DataTable("tbllocMwSt")
        Dim objlocMySQLcmd As New MySqlCommand
        Dim strSteuerRec As String = String.Empty

        Try

            'Sind die Angaben stimmig?
            If Len(strStrCode) > 0 And dblStrAmount <> 0 And dblStrWert = 0 Then 'MwSt Wert ist 0 obwohl Schlüssel und MwSt-Betrag

                objlocMySQLcmd.CommandText = "Select  * FROM t_sage_sage50mwst WHERE strKey='" + strStrCode + "'"

                objlocMySQLcmd.Connection = objdbconn
                objlocdtMwSt.Load(objlocMySQLcmd.ExecuteReader)

                If objlocdtMwSt.Rows.Count = 0 Then
                    MessageBox.Show("MwSt " + strStrCode + " ist nicht definiert. Korrektur von MwST-Satz nicht möglich.")
                    Return 1
                Else
                    'MwSt-Satz änern gemäss Tabelle
                    dblStrWert = objlocdtMwSt.Rows(0).Item("dblProzent")
                    Return 2

                End If

            ElseIf Len(strStrCode) > 0 And dblStrAmount = 0 And dblStrWert <> 0 Then 'MwSt Wert ist nicht 0 obwohl kein Betrag

                'Check was ist hinterlegt
                objlocMySQLcmd.CommandText = "SELECT  * FROM t_sage_sage50mwst WHERE strKey='" + strStrCode + "'"

                objlocMySQLcmd.Connection = objdbconn
                objlocdtMwSt.Load(objlocMySQLcmd.ExecuteReader)

                If objlocdtMwSt.Rows.Count = 0 Then
                    MessageBox.Show("MwSt " + strStrCode + " ist nicht definiert. Korrektur von MwST-Satz nicht möglich.")
                    Return 1
                Else
                    If objlocdtMwSt.Rows(0).Item("dblProzent") <> dblStrWert Then
                        dblStrWert = objlocdtMwSt.Rows(0).Item("dblProzent")
                        Return 2
                    End If

                End If

            ElseIf strStrCode = "ohne" Or strStrCode = "frei" Then
                'Check was ist hinterlegt
                objlocMySQLcmd.CommandText = "SELECT  * FROM t_sage_sage50mwst WHERE strKey='" + strStrCode + "'"

                objlocMySQLcmd.Connection = objdbconn
                objlocdtMwSt.Load(objlocMySQLcmd.ExecuteReader)

                If objlocdtMwSt.Rows.Count = 0 Then
                    MessageBox.Show("MwSt " + strStrCode + " ist nicht definiert. Korrektur von MwST-Satz nicht möglich.")
                    Return 1
                Else
                    If objlocdtMwSt.Rows(0).Item("dblProzent") <> dblStrWert Then
                        dblStrWert = objlocdtMwSt.Rows(0).Item("dblProzent")
                        Return 2
                    End If

                End If

            Else
                Return 0

            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message)

        End Try

    End Function

    Public Shared Function FcCheckKstKtr(ByVal lngKST As Long,
                                         ByRef objFiBhg As SBSXASLib.AXiFBhg,
                                         ByRef objFiPI As SBSXASLib.AXiPlFin,
                                         ByVal lngKonto As Long,
                                         ByRef strKstKtrSage200 As String) As Int16

        'return 0=ok, 1=Kst existiert kene Kostenart, 2=Kst nicht defniert, 3=nicht auf Konto anwendbar 1000 - 2999

        Dim strReturn As String
        Dim strReturnAr() As String
        Dim booKstKAok As Boolean
        Dim strKst, strKA As String

        booKstKAok = False
        objFiPI = Nothing
        objFiPI = objFiBhg.GetCheckObj

        Try
            'If CInt(Left(lngKonto.ToString, 1)) >= 3 Then
            strReturn = objFiBhg.GetKstKtrInfo(lngKST.ToString)
            If strReturn = "EOF" Then
                Return 2
            Else
                strReturnAr = Split(strReturn, "{>}")
                strKstKtrSage200 = strReturnAr(1)
                strKst = Convert.ToString(lngKST)
                strKA = Convert.ToString(lngKonto)
                'Ist Kst auf Kostenbart definiert?
                booKstKAok = objFiPI.CheckKstKtr(strKst, strKA)

                If booKstKAok Then
                    Return 0
                Else
                    Return 1
                End If
            End If
            'Else
            'Return 3
            'End If


        Catch ex As Exception
            Return 1

        End Try

    End Function

    Public Shared Function FcGetPKNewFromRep(ByRef objdbconnZHDB02 As MySqlConnection,
                                             ByRef objsqlcommandZHDB02 As MySqlCommand,
                                             ByVal intPKRefField As Int32,
                                             ByVal strMode As String) As Int32

        'Aus Tabelle Rep_Betriebe auf ZHDB02 auslesen 
        Dim objdtRepBetrieb As New DataTable

        Try

            objdbconnZHDB02.Open()
            objsqlcommandZHDB02.Connection = objdbconnZHDB02
            If strMode = "P" Then
                objsqlcommandZHDB02.CommandText = "SELECT PKNr From t_customer WHERE ID=" + intPKRefField.ToString
            Else
                objsqlcommandZHDB02.CommandText = "SELECT PKNr From tab_repbetriebe WHERE Rep_Nr=" + intPKRefField.ToString
            End If
            objdtRepBetrieb.Load(objsqlcommandZHDB02.ExecuteReader)
            If (objdtRepBetrieb.Rows.Count > 0) Then
                If Not IsDBNull(objdtRepBetrieb.Rows(0).Item("PKNr")) Then
                    Return objdtRepBetrieb.Rows(0).Item("PKNr")
                Else
                    Return 0
                End If
            Else
                Return 0
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem Neue PK-Nr.")
            Return 0

        Finally
            objdbconnZHDB02.Close()

        End Try


    End Function


    Public Shared Function FcCheckCurrency(ByVal strCurrency As String, ByRef objfiBuha As SBSXASLib.AXiFBhg) As Integer

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
                Application.DoEvents()
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

    Public Shared Function FcCheckKonto(ByVal lngKtoNbr As Long,
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
                'If dblMwSt = 0 Then
                'Return 0
                strKontoInfo = Split(objfiBuha.GetKontoInfo(lngKtoNbr.ToString), "{>}")
                If booExistanceOnly Then
                    Return 0
                End If
                'KST?
                'strKontoInfo = Split(objfiBuha.GetKontoInfo(lngKtoNbr.ToString), "{>}")
                If lngKST > 0 Then
                    If CInt(Left(lngKtoNbr.ToString, 1)) >= 3 Then
                        'strKontoInfo = Split(objfiBuha.GetKontoInfo(lngKtoNbr.ToString), "{>}")
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

                            'Else
                            'Steuerpflichtig?
                            'strKontoInfo = Split(objfiBuha.GetKontoInfo(lngKtoNbr.ToString), "{>}")
                            'If strKontoInfo(26) = "" Then
                            'Return 2
                            'Else
                            'Return 0
                            'End If
                            'End If
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


    Public Shared Function InsertDataTableColumnName(ByRef dtSouce As DataTable,
                                                     ByRef dtResult As DataTable) As Boolean

        Dim rowResult As DataRow
        Dim Result As Boolean = True

        Try

            For Each rowSource As DataRow In dtSouce.Rows
                rowResult = dtResult.NewRow()

                For Each ColumnSource As DataColumn In rowSource.Table.Columns
                    Dim ColumnResult As DataColumnCollection = dtResult.Columns

                    If ColumnResult.Contains(ColumnSource.ColumnName) Then
                        rowResult(ColumnSource.ColumnName) = rowSource(ColumnSource.ColumnName)
                    End If
                Next
                dtResult.Rows.Add(rowResult)
            Next

            Return Result

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Result = False
            Return Result

        Finally
            rowResult = Nothing

        End Try

    End Function



    Public Shared Function FcGetSteuerFeld(ByRef objFBhg As SBSXASLib.AXiFBhg, ByVal lngKto As Long, ByVal strDebiSubText As String, ByVal dblBrutto As Double, ByVal strMwStKey As String, ByVal dblMwSt As Double) As String

        Dim strSteuerFeld As String = String.Empty

        Try

            If dblMwSt <> 0 Then

                strSteuerFeld = objFBhg.GetSteuerfeld(lngKto.ToString, strDebiSubText, dblBrutto.ToString, strMwStKey, dblMwSt.ToString)

            Else

                strSteuerFeld = objFBhg.GetSteuerfeld(lngKto.ToString, strDebiSubText, dblBrutto.ToString, strMwStKey)

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)

        End Try

        Return strSteuerFeld

    End Function

    Public Shared Function FcGetKurs(ByVal strCurrency As String,
                                     ByVal strDateValuta As String,
                                     ByRef objFBhg As SBSXASLib.AXiFBhg,
                                     ByVal Optional intKonto As Integer = 0) As Double

        'Konzept: Falls ein Konto mitgegeben wird, wird überprüft ob auf dem Konto die mitgegebene Währung Leitwärhung ist. Falls ja wird der Kurs 1 zurück gegeben

        Dim strKursZeile As String = String.Empty
        Dim strKursZeileAr() As String
        Dim strKontoInfo() As String

        objFBhg.ReadKurse(strCurrency, "", "J")

        Do While strKursZeile <> "EOF"
            strKursZeile = objFBhg.GetKursZeile()
            If strKursZeile <> "EOF" Then
                strKursZeileAr = Split(strKursZeile, "{>}")
                If strKursZeileAr(0) = strCurrency Then
                    'If strKursZeileAr(0) = "EUR" Then Stop
                    'Prüfen ob Currency Leitwährung auf Konto. Falls ja Return 1
                    If intKonto <> 0 Then
                        strKontoInfo = Split(objFBhg.GetKontoInfo(intKonto.ToString), "{>}")
                        If strKontoInfo(7) = strCurrency Then
                            Return 1
                        Else
                            Return strKursZeileAr(4)
                            Return 0
                        End If
                    Else
                        Return strKursZeileAr(4)
                    End If
                End If
            Else
                Return 1 'Kurs nicht gefunden
            End If
        Loop

    End Function

    Public Shared Function FcCheckKredit(ByVal intAccounting As Integer,
                                        ByRef objdtKredits As DataTable,
                                        ByRef objdtKreditSubs As DataTable,
                                        ByRef objFinanz As SBSXASLib.AXFinanz,
                                        ByRef objfiBuha As SBSXASLib.AXiFBhg,
                                        ByRef objKrBuha As SBSXASLib.AXiKrBhg,
                                        ByRef objdbPIFb As SBSXASLib.AXiPlFin,
                                        ByRef objdbconn As MySqlConnection,
                                        ByRef objdbconnZHDB02 As MySqlConnection,
                                        ByRef objsqlcommand As MySqlCommand,
                                        ByRef objsqlcommandZHDB02 As MySqlCommand,
                                        ByRef objOrdbconn As OracleClient.OracleConnection,
                                        ByRef objOrcommand As OracleClient.OracleCommand,
                                        ByRef objdbAccessConn As OleDb.OleDbConnection,
                                        ByRef objdtInfo As DataTable,
                                        ByVal strcmbBuha As String,
                                        ByVal strYear As String,
                                        ByVal strPeriode As String,
                                        ByVal datPeriodFrom As Date,
                                        ByVal datPeriodTo As Date,
                                        ByVal strPeriodStatus As String,
                                        ByVal booValutaCoorect As Boolean,
                                        ByVal datValutaCorrect As Date) As Integer

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

            For Each row As DataRow In objdtKredits.Rows


                'If row("lngKredID") = "6978" Then Stop
                'Runden
                row("dblKredNetto") = Decimal.Round(row("dblKredNetto"), 2, MidpointRounding.AwayFromZero)
                row("dblKredMwSt") = Decimal.Round(row("dblKredMwst"), 2, MidpointRounding.AwayFromZero)
                row("dblKredBrutto") = Decimal.Round(row("dblKredBrutto"), 2, MidpointRounding.AwayFromZero)
                'lngKrediID = row("lngKredID")

                'Status-String erstellen
                'Kreditor 01
                intReturnValue = MainKreditor.FcGetRefKrediNr(objdbconn,
                                                 objdbconnZHDB02,
                                                 objsqlcommand,
                                                 objsqlcommandZHDB02,
                                                 objOrdbconn,
                                                 objOrcommand,
                                                 objdbAccessConn,
                                                 IIf(IsDBNull(row("lngKredNbr")), 0, row("lngKredNbr")),
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
                booAutoCorrect = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettings(objdbconn, "Buchh_HeadKAutoCorrect", intAccounting)))
                'booAutoCorrect = False
                booCpyKSTToSub = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettings(objdbconn, "Buchh_KKSTHeadToSub", intAccounting)))
                intReturnValue = FcCheckKrediSubBookings(row("lngKredID"),
                                                         objdtKreditSubs,
                                                         intSubNumber,
                                                         dblSubBrutto,
                                                         dblSubNetto,
                                                         dblSubMwSt,
                                                         objdbconn,
                                                         objfiBuha,
                                                         objdbPIFb,
                                                         row("intBuchungsart"),
                                                         booAutoCorrect,
                                                         booCpyKSTToSub,
                                                         row("lngKrediKST"),
                                                         row("intPayType"),
                                                         row("strKrediBank"))

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
                            Dim objdrKrediSub As DataRow = objdtKreditSubs.NewRow
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
                            objdtKreditSubs.Rows.Add(objdrKrediSub)
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

                Application.DoEvents()

                'PGV => Prüfung vor Valuta-Datum da Valuta-Datum verändert wird. PGV soll nur möglich sein wenn rebilled
                If Not IsDBNull(row("datPGVFrom")) And MainKreditor.FcIsAllKrediRebilled(objdtKreditSubs, row("lngKredID")) = 0 Then
                    row("booPGV") = True
                ElseIf Not IsDBNull(row("datPGVFrom")) And MainKreditor.FcIsAllKrediRebilled(objdtKreditSubs, row("lngKredID")) = 1 Then
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
                If Year(row("datKredRGDatum")) <> Year(row("datKredValDatum")) And Year(row("datKredValDatum")) >= 2022 Then

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
                            row("datPGVFrom") = Year(datValutaSave).ToString + "-" + Month(datValutaSave).ToString + "-" + Day(datValutaSave).ToString
                            row("datPGVTo") = Year(datValutaSave).ToString + "-" + Month(datValutaSave).ToString + "-" + Day(datValutaSave).ToString
                            row("datKredValDatum") = "2023-01-01" ' Year(row("datKredRGDatum")).ToString + "-01-01"
                        ElseIf row("strPGVType") = "RV" Then
                            row("datPGVFrom") = Year(datValutaSave).ToString + "-" + Month(datValutaSave).ToString + "-" + Day(datValutaSave).ToString
                            row("datPGVTo") = Year(datValutaSave).ToString + "-" + Month(datValutaSave).ToString + "-" + Day(datValutaSave).ToString
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
                intReturnValue = FcChCeckDate(IIf(IsDBNull(row("datKredValDatum")), row("datKredRGDatum"), row("datKredValDatum")),
                                              objdtInfo,
                                              datPeriodFrom,
                                              datPeriodTo,
                                              strPeriodStatus,
                                              True)

                'Falls Problem versuchen mit Valuta-Datum-Anpassung
                If intReturnValue <> 0 And booValutaCoorect Then
                    row("datKredValDatum") = Format(datValutaCorrect, "Short Date")
                    intReturnValue = FcChCeckDate(IIf(IsDBNull(row("datKredValDatum")), #1789-09-17#, row("datKredValDatum")),
                                                  objdtInfo,
                                                  datPeriodFrom,
                                                  datPeriodTo,
                                                  strPeriodStatus,
                                                  True)
                    If intReturnValue = 0 Then
                        intReturnValue = 2
                    Else
                        intReturnValue = 3
                    End If

                End If

                'Bei PGV checken ob PGV-Startdatum in blockierter Periode
                If row("booPGV") And intReturnValue = 0 Then
                    intReturnValue = FcCheckPGVDate(row("datPGVFrom"),
                                                    intAccounting,
                                                    objdbconnZHDB02)
                    If intReturnValue <> 0 Then
                        'Falls TP-Buchung in blockierter Periode dann probieren mit Valuta-Korrektur
                        If intPGVMonths = 1 And booValutaCoorect Then
                            row("datKredValDatum") = Format(datValutaCorrect, "Short Date")
                            intReturnValue = FcChCeckDate(IIf(IsDBNull(row("datKredValDatum")), #1789-09-17#, row("datKredValDatum")),
                                                          objdtInfo,
                                                          datPeriodFrom,
                                                          datPeriodTo,
                                                          strPeriodStatus,
                                                          True)
                            If intReturnValue = 0 Then
                                'PGV - Flag entfernen
                                row("booPGV") = False
                                intReturnValue = 5
                            Else
                                intReturnValue = 3
                            End If
                        Else
                            intReturnValue = 4
                        End If
                    End If

                End If
                strBitLog += Trim(intReturnValue.ToString)

                'RG - Datum 11
                intReturnValue = FcChCeckDate(IIf(IsDBNull(row("datKredRGDatum")), #1789-09-17#, row("datKredRGDatum")),
                                              objdtInfo,
                                              datPeriodFrom,
                                              datPeriodTo,
                                              strPeriodStatus,
                                              True)

                'Falls Problem versuchen mit Valuta-Datum-Anpassung
                If intReturnValue <> 0 And booValutaCoorect Then
                    row("datKredRGDatum") = Format(datValutaCorrect, "Short Date")
                    intReturnValue = FcChCeckDate(IIf(IsDBNull(row("datKredRGDatum")), #1789-09-17#, row("datKredRGDatum")),
                                                  objdtInfo,
                                                  datPeriodFrom,
                                                  datPeriodTo,
                                                  strPeriodStatus,
                                                  True)
                    If intReturnValue = 0 Then
                        'Korrektur hat funktioniert, Wert auf 2 setzen
                        intReturnValue = 2
                    Else
                        intReturnValue = 3
                    End If
                End If
                strBitLog += Trim(intReturnValue.ToString)

                ''Referenz 12
                If (Not String.IsNullOrEmpty(row("strKredRef"))) And (row("intPayType") = 3 Or row("intPayType") = 10) Then
                    If Val(Left(row("strKredRef"), Len(row("strKredRef")) - 1)) > 0 Then

                        If Right(row("strKredRef"), 1) <> Main.FcModulo10(Left(row("strKredRef"), Len(row("strKredRef")) - 1)) Then
                            strBitLog += "1"
                        Else
                            strBitLog += "0"
                        End If

                    Else
                        strBitLog += "1"
                    End If
                Else
                    strBitLog += "0"
                End If
                'Debug.Print("Erfasste Prüfziffer " + Right(row("strKredRef"), 1) + ", kalkuliert " + Main.FcModulo10(Left(row("strKredRef"), Len(row("strKredRef")) - 1)).ToString)
                'intReturnValue = IIf(IsDBNull(row("strKredRef")), 1, 0)
                'strBitLog += Trim(intReturnValue.ToString)
                'interne Bank 13
                intReturnValue = Main.FcCheckDebiIntBank(objdbconn,
                                                         intAccounting,
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
                booPKPrivate = IIf(FcReadFromSettings(objdbconn, "Buchh_PKKrediTable", intAccounting) = "t_customer", True, False)
                'Kreditor 1
                If Left(strBitLog, 1) <> "0" Then
                    strStatus += "Kred"
                    If Left(strBitLog, 1) <> "2" Then
                        If booPKPrivate Then
                            intReturnValue = MainKreditor.FcIsPrivateKreditorCreatable(objdbconn,
                                                                                        objdbconnZHDB02,
                                                                                        objsqlcommandZHDB02,
                                                                                        intKreditorNew,
                                                                                        objKrBuha,
                                                                                        objfiBuha,
                                                                                        IIf(IsDBNull(row("intPayType")), 3, row("intPayType")),
                                                                                        IIf(IsDBNull(row("strKredRef")), "", row("strKredRef")),
                                                                                        intintBank,
                                                                                        IIf(IsDBNull(row("strKrediBank")), "", row("strKrediBank")),
                                                                                        strcmbBuha,
                                                                                        intAccounting)
                        Else
                            intReturnValue = MainKreditor.FcIsKreditorCreatable(objdbconn,
                                                                            objdbconnZHDB02,
                                                                            objsqlcommandZHDB02,
                                                                            intKreditorNew,
                                                                            objKrBuha,
                                                                            objfiBuha,
                                                                            strcmbBuha,
                                                                            IIf(IsDBNull(row("intPayType")), 3, row("intPayType")),
                                                                            IIf(IsDBNull(row("strKredRef")), "", row("strKredRef")),
                                                                            intintBank,
                                                                            IIf(IsDBNull(row("strKrediBank")), "", row("strKrediBank")),
                                                                            intAccounting)

                        End If
                        If intReturnValue = 0 Then
                            strStatus += " erstellt"
                            row("strKredBez") = MainKreditor.FcReadKreditorName(objKrBuha, intKreditorNew, row("strKredCur"))

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
                    row("strKredBez") = MainKreditor.FcReadKreditorName(objKrBuha, intKreditorNew, row("strKredCur"))
                    row("lngKredNbr") = intKreditorNew
                    row("intEBank") = 0
                    If row("intPayType") = 9 Then
                        strIBANToPass = row("strKredRef")
                    ElseIf row("intPayType") = 10 Then
                        strIBANToPass = row("strKrediBank")
                    End If
                    If row("intPayType") = 9 Or row("intPayType") = 10 Then
                        intReturnValue = MainKreditor.FcCheckKreditBank(objdbconn,
                                                       objdbconnZHDB02,
                                                       objKrBuha,
                                                       intKreditorNew,
                                                       IIf(IsDBNull(row("intPayType")), 3, row("intPayType")),
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
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "VDCor"
                        'Korrektur hat geklappt, Wert wieder auf 0 setzen
                        strBitLog = Left(strBitLog, 9) + "0" + Right(strBitLog, Len(strBitLog) - 10)
                    ElseIf Mid(strBitLog, 10, 1) = "3" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "VDCorNok"
                    ElseIf Mid(strBitLog, 10, 1) = "4" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "PGVDblck"
                    ElseIf Mid(strBitLog, 10, 1) = "5" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "PGVVDCor"
                        'Korrektur hat geklappt, Wert wieder auf 0 setzen
                        strBitLog = Left(strBitLog, 9) + "0" + Right(strBitLog, Len(strBitLog) - 10)
                    End If
                End If
                'RG Datum 11
                If Mid(strBitLog, 11, 1) <> "0" Then
                    If Mid(strBitLog, 11, 1) = "1" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "RgD"
                    ElseIf Mid(strBitLog, 11, 1) = "2" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "RgDCor"
                        'Korrektur hat geklappt, Wert wieder auf 0 setzen
                        strBitLog = Left(strBitLog, 10) + "0" + Right(strBitLog, Len(strBitLog) - 11)
                    ElseIf Mid(strBitLog, 11, 1) = "3" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "RgDCorNok"
                    ElseIf Mid(strBitLog, 11, 1) = "4" Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "PGVDblck"
                    End If
                End If
                'Referenz 12
                If Mid(strBitLog, 12, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Ref "
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
                If objdbconn.State = ConnectionState.Closed Then
                    objdbconn.Open()
                End If
                booDiffHeadText = IIf(FcReadFromSettings(objdbconn, "Buchh_KTextSpecial", intAccounting) = "0", False, True)
                If booDiffHeadText Then
                    strKrediHeadText = MainDebitor.FcSQLParse(FcReadFromSettings(objdbconn,
                                                                                "Buchh_KTextSpecialText",
                                                                                intAccounting),
                                                                                row("strKredRGNbr"),
                                                                            objdtKredits,
                                                                            "C")
                    row("strKredText") = strKrediHeadText
                End If

                'Wird ein anderer Text in den Sub-Buchung gewünscht?
                'Soll der Gelesene Sub-Text bleiben?
                booLeaveSubText = IIf(FcReadFromSettings(objdbconn, "Buchh_KSubLeaveText", intAccounting) = "0", False, True)
                If Not booLeaveSubText Then
                    booDiffSubText = IIf(FcReadFromSettings(objdbconn, "Buchh_KSubTextSpecial", intAccounting) = "0", False, True)
                    If booDiffSubText Then
                        strKrediSubText = MainDebitor.FcSQLParse(FcReadFromSettings(objdbconn,
                                                                                "Buchh_KSubTextSpecialText",
                                                                                intAccounting),
                                                                                row("strKredRGNbr"),
                                                                           objdtKredits,
                                                                           "C")
                    Else
                        strKrediSubText = row("strKredText")
                    End If
                    selsubrow = objdtKreditSubs.Select("lngKredID=" + row("lngKredID").ToString)
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

                Application.DoEvents()

            Next

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Check-Kredit " + intKreditorNew.ToString + " ID " + lngKrediID.ToString)

        Finally
            If objOrdbconn.State = ConnectionState.Open Then
                objOrdbconn.Close()
            End If
            If objdbconn.State = ConnectionState.Open Then
                objdbconn.Close()
            End If

        End Try


    End Function

    Public Shared Function FcCheckPayType(ByRef intPayType As Int16,
                                          ByVal strReferenz As String,
                                          ByVal strKrediBank As String) As Int16

        '0=ok, 1=IBAN Nr. aber nicht IBAN-Typ, 6=ESR-Nr aber keine Bank oder ungültige, 4=keine Referenz, 5=keine korrekte QR-IBAN 2=QR-ESR, 6=ESR Bank-Referenz nicht korrekt, 7=IBAN ist QR-IBAN, 9=Problem

        Try

            If Len(strReferenz) > 0 Then
                'Wurde eine IBAN - Nr. übergeben aber Typ ist nicht IBAN
                If Len(strReferenz) >= 21 Then ' And intPayType <> 9 Then
                    ''Sind die ersten 2 Positionen nicht numerisch?
                    'If Strings.Asc(Left(strReferenz, 1)) < 48 And Strings.Asc(Left(strReferenz, 1)) > 57 Then '1 Zeichen nicht numerisch
                    '    If Strings.Asc(Mid(strReferenz, 2, 1)) < 48 And Strings.Asc(Mid(strReferenz, 2, 1)) > 57 Then '2 Zeichen nicht numerisch
                    '        intPayType = 9
                    '        Return 1
                    '    End If
                    'End If
                    If Main.FcAreFirst2Chars(strReferenz) = 0 And intPayType <> 9 And Mid(strReferenz, 5, 1) <> "3" Then 'Falscher PayType bei IBAN-Nr.
                        intPayType = 9
                        Return 1
                    End If
                    'QR-ESR?
                    'Bank - Referenz IBAN?
                    If Main.FcAreFirst2Chars(strReferenz) = 0 Then 'IBAN - Referenz
                        'If Main.FcAreFirst2Chars(strKrediBank) = 0 Then
                        'intPayType = 9
                        'Return 0
                        'Else
                        'normale IBAN
                        'Check ob nicht QR-IBAN als Zahl-IBAN erfasst
                        If Mid(strReferenz, 5, 1) = "3" And Left(strReferenz, 2) = "CH" Then
                            intPayType = 9
                            Return 7
                        Else
                            intPayType = 9
                            Return 0
                        End If
                        'End If
                    Else 'QR-ESR ?
                        If Main.FcAreFirst2Chars(IIf(strKrediBank = "", "00", strKrediBank)) = 0 Then 'IBAN als Bank
                            'QR-IBAN?
                            If Mid(strKrediBank, 5, 1) = "3" Then
                                intPayType = 10
                                Return 2
                            Else
                                'keine QR-IBAN-ESR-Ref
                                'intPayType = 3
                                Return 5
                            End If
                        Else

                            If Len(strKrediBank) <> 9 Then 'ESR aber keine gültige Bank
                                'ESR, falsch deklariert
                                If intPayType <> 3 Then
                                    intPayType = 3
                                End If
                                Return 6
                            Else
                                'Debug.Print("Checksum " + Strings.Left(strKrediBank, 8) + " " + Strings.Right(strKrediBank, 1) + ", " + Main.FcModulo10(Strings.Left(strKrediBank, 8)).ToString)
                                If Main.FcModulo10(Strings.Left(strKrediBank, 8)).ToString <> Strings.Right(strKrediBank, 1) Then
                                    Return 6
                                Else
                                    Return 0 'Bank ok
                                End If

                            End If
                        End If
                    End If
                ElseIf intPayType = 0 Then
                    Return 9
                End If
                'If Len(strKrediBank) <> 9 Then 'ESR aber keine gültige Bank
                '    Return 3
                'Else
                '    Return 0 'Bank ok
                'End If

                'Else
            Else
                If intPayType = 9 And Len(strReferenz) = 0 Then
                    intPayType = 3 'Nicht IBAN
                    Return 4
                    'ElseIf intPayType = 0 Then
                    '    Return 9
                End If

            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Fc CheckPayType")
            Return 9

        Finally

        End Try

    End Function

    Public Shared Function FcAreFirst2Chars(ByVal strToCheck As String) As Int16

        '0=Nicht numerisch, 1=numerisch, 9=Problem

        Try
            'Sind die ersten 2 Positionen nicht numerisch?
            If Asc(Left(strToCheck, 1)) < 48 Or Asc(Left(strToCheck, 1)) > 57 Then '1 Zeichen nicht numerisch
                If Asc(Mid(strToCheck, 2, 1)) < 48 Or Asc(Mid(strToCheck, 2, 1)) > 57 Then '2 Zeichen nicht numerisch
                    Return 0
                Else
                    Return 1
                End If
            Else
                Return 1
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return 9

        Finally

        End Try

    End Function

    Public Shared Function fcCheckTransitorischeDebit(ByVal intAccounting As Int16,
                                                      ByRef objdbconn As MySqlConnection,
                                                      ByRef objdbAccessConn As OleDb.OleDbConnection)

        Dim strSQLMan As String
        'Dim strSQLSub As String
        Dim strRGTableType As String
        Dim objRGMySQLConn As New MySqlConnection
        Dim objlocMySQLcmd As New MySqlCommand
        Dim objlocOLEdbcmd As New OleDb.OleDbCommand
        Dim booTransits As Boolean
        Dim intAffected As Int16

        Dim tblCompute As New DataTable()
        Dim booTransitcond As Boolean


        Dim objDTTransitDebits As New DataTable
        Dim strMDBName As String


        Try

            objdbconn.Open()
            'Gibt es transitorische Buchungen?
            booTransits = CBool(FcReadFromSettings(objdbconn, "Buchh_Transit", intAccounting))

            If booTransits Then

                'Table - Art lesen
                strRGTableType = FcReadFromSettings(objdbconn, "Buchh_RGTableType", intAccounting)
                'Debitoren - Table Name lesen
                strMDBName = FcReadFromSettings(objdbconn, "Buchh_RGTableMDB", intAccounting)

                'Debitzoren Transit-Queries für Mandant einlesen
                strSQLMan = "Select * FROM t_sage_buchhaltungen_sub WHERE strType='D' AND refMandant=" + intAccounting.ToString
                        objRGMySQLConn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString")
                objlocMySQLcmd.Connection = objRGMySQLConn
                objlocMySQLcmd.CommandText = strSQLMan
                objRGMySQLConn.Open()
                objDTTransitDebits.Load(objlocMySQLcmd.ExecuteReader)
                objRGMySQLConn.Close()

                For Each rowdebitquery As DataRow In objDTTransitDebits.Rows

                    If IIf(IsDBNull(rowdebitquery("strCondition")), "", rowdebitquery("strCondition")) <> "" Then
                        'Es wurde eine Bedingung definiert
                        booTransitcond = Convert.ToBoolean(tblCompute.Compute("#" + DateTime.Now.ToString("yyyy-MM-dd") + "#" + rowdebitquery("strCondition"), Nothing))
                        'Debug.Print("Result " + "#" + DateTime.Now.ToString("yyyy-MM-dd") + "#" + rowdebitquery("strCondition") + ", " + booTransitcond.ToString)
                    Else
                        booTransitcond = True
                    End If

                    If booTransitcond Then
                        'Debug.Print("Running Query " + rowdebitquery("strSQL"))
                        If strRGTableType = "A" Then
                            'Access
                            Call FcInitAccessConnecation(objdbAccessConn, strMDBName)
                            objdbAccessConn.Open()
                            objlocOLEdbcmd.Connection = objdbAccessConn
                            objlocOLEdbcmd.CommandText = rowdebitquery("strSQL")
                            intAffected = objlocOLEdbcmd.ExecuteNonQuery()
                            objdbAccessConn.Close()
                        ElseIf strRGTableType = "M" Then
                            'MySQL
                            objRGMySQLConn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(strMDBName)
                            objRGMySQLConn.Open()
                            objlocMySQLcmd.Connection = objRGMySQLConn
                            objlocMySQLcmd.CommandText = rowdebitquery("strSQL")
                            intAffected = objlocMySQLcmd.ExecuteNonQuery()
                            objRGMySQLConn.Close()
                        End If
                    End If

                Next


            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)

        Finally
            If objRGMySQLConn.State = ConnectionState.Open Then
                objRGMySQLConn.Close()
            End If
            If objdbconn.State = ConnectionState.Open Then
                objdbconn.Close()
            End If
            If objdbAccessConn.State = ConnectionState.Open Then
                objdbAccessConn.Close()
            End If

        End Try


    End Function

    Public Shared Function fcCheckTransitorischeKredit(ByVal intAccounting As Int16,
                                                       ByRef objdbconn As MySqlConnection,
                                                       ByRef objdbAccessConn As OleDb.OleDbConnection)

        Dim strSQLMan As String
        'Dim strSQLSub As String
        Dim strRGTableType As String
        Dim objRGMySQLConn As New MySqlConnection
        Dim objlocMySQLcmd As New MySqlCommand
        Dim objlocOLEdbcmd As New OleDb.OleDbCommand
        Dim booTransits As Boolean
        Dim intAffected As Int16

        Dim tblCompute As New DataTable()
        Dim booTransitcond As Boolean


        Dim objDTTransitDebits As New DataTable
        Dim strMDBName As String


        Try

            objdbconn.Open()
            'Gibt es transitorische Buchungen?
            booTransits = CBool(FcReadFromSettings(objdbconn, "Buchh_Transit", intAccounting))

            If booTransits Then

                'Table - Art lesen
                strRGTableType = FcReadFromSettings(objdbconn, "Buchh_KRGTableType", intAccounting)
                'Debitoren - Table Name lesen
                strMDBName = FcReadFromSettings(objdbconn, "Buchh_KRGTableMDB", intAccounting)

                'Debitzoren Transit-Queries für Mandant einlesen
                strSQLMan = "SELECT * FROM t_sage_buchhaltungen_sub WHERE strType='K' AND refMandant=" + intAccounting.ToString
                objRGMySQLConn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString")
                objlocMySQLcmd.Connection = objRGMySQLConn
                objlocMySQLcmd.CommandText = strSQLMan
                objRGMySQLConn.Open()
                objDTTransitDebits.Load(objlocMySQLcmd.ExecuteReader)
                objRGMySQLConn.Close()

                For Each rowdebitquery As DataRow In objDTTransitDebits.Rows

                    If IIf(IsDBNull(rowdebitquery("strCondition")), "", rowdebitquery("strCondition")) <> "" Then
                        'Es wurde eine Bedingung definiert
                        booTransitcond = Convert.ToBoolean(tblCompute.Compute("#" + DateTime.Now.ToString("yyyy-MM-dd") + "#" + rowdebitquery("strCondition"), Nothing))
                        'Debug.Print("Result " + "#" + DateTime.Now.ToString("yyyy-MM-dd") + "#" + rowdebitquery("strCondition") + ", " + booTransitcond.ToString)
                    Else
                        booTransitcond = True
                    End If

                    If booTransitcond Then
                        'Debug.Print("Running Query " + rowdebitquery("strSQL"))
                        If strRGTableType = "A" Then
                            'Access
                            Call FcInitAccessConnecation(objdbAccessConn, strMDBName)
                            objdbAccessConn.Open()
                            objlocOLEdbcmd.Connection = objdbAccessConn
                            objlocOLEdbcmd.CommandText = rowdebitquery("strSQL")
                            intAffected = objlocOLEdbcmd.ExecuteNonQuery()
                            objdbAccessConn.Close()
                        ElseIf strRGTableType = "M" Then
                            'MySQL
                            objRGMySQLConn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(strMDBName)
                            objRGMySQLConn.Open()
                            objlocMySQLcmd.Connection = objRGMySQLConn
                            objlocMySQLcmd.CommandText = rowdebitquery("strSQL")
                            intAffected = objlocMySQLcmd.ExecuteNonQuery()
                            objRGMySQLConn.Close()
                        End If
                    End If

                Next


            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Transitorisch-Check Kreditoren", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        Finally
            If objRGMySQLConn.State = ConnectionState.Open Then
                objRGMySQLConn.Close()
            End If
            If objdbconn.State = ConnectionState.Open Then
                objdbconn.Close()
            End If
            If objdbAccessConn.State = ConnectionState.Open Then
                objdbAccessConn.Close()
            End If

        End Try


    End Function

    Public Shared Function FcInitAccessConnecation(ByRef objaccesscon As OleDb.OleDbConnection,
                                                   ByVal strMDBName As String) As Int16

        'Access - Connection soll initialisiert werden
        '0 = ok, 1 = nicht ok

        Dim dbProvider, dbSource, dbPathAndFile As String

        Try

            dbProvider = "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
            dbSource = "Data Source="
            dbPathAndFile = "\\sdlc.mssag.ch\Apps\Backends\" + strMDBName + ";Jet OLEDB:System Database=\\sdlc.mssag.ch\Apps\Backends\Workbench.mdw;User ID=HagerR;"
            objaccesscon.ConnectionString = dbProvider + dbSource + dbPathAndFile
            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return 1

        End Try


    End Function

    Public Shared Function FcNextPKNr(ByRef objdbconnZHDB02 As MySqlConnection,
                                      ByVal intRepNr As Int32,
                                      ByRef intNewPKNr As Int32,
                                      ByVal intAccounting As Int16,
                                      ByVal strMode As String) As Int16

        '0=ok, 1=Rep - Nr. existiert nicht, 2=Bereich voll, 3=keine Bereichdefinition 9=Problem

        'PK - Nummer soll der Funktion gegeben werden, Funktion sucht sich dann die PK_Gruppe 
        'Konzept: Tabelle füllen und dann durchsteppen
        Dim objsqlcommand As New MySqlCommand
        Dim objdtPKNr As New DataTable
        Dim intPKNrGuppenID As Int16
        Dim intRangeStart, intRangeEnd, i, intRecordCounter As Int32
        Dim objdsPKNbrs As New DataSet
        Dim objDAPKNbrs As New MySqlDataAdapter
        Dim objdbconn As New MySqlConnection


        Try

            'Wo ist die RepBetriebe?
            If objdbconnZHDB02.State = ConnectionState.Closed Then
                objdbconnZHDB02.Open()
            End If
            If strMode = "D" Then
                objdbconn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(Main.FcReadFromSettings(objdbconnZHDB02, "Buchh_PKTableConnection", intAccounting))
            Else
                objdbconn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(Main.FcReadFromSettings(objdbconnZHDB02, "Buchh_PKKrediTableConnection", intAccounting))
            End If

            If objdbconn.State = ConnectionState.Closed Then
                objdbconn.Open()
            End If
            objsqlcommand.Connection = objdbconn
            objsqlcommand.CommandText = "SELECT PKNrGruppeID FROM tab_repbetriebe WHERE Rep_Nr=" + intRepNr.ToString
            objdtPKNr.Load(objsqlcommand.ExecuteReader)

            If objdtPKNr.Rows.Count > 0 Then 'Rep_Betrieb gefunden
                intPKNrGuppenID = IIf(IsDBNull(objdtPKNr.Rows(0).Item("PKNrGruppeID")), 2, objdtPKNr.Rows(0).Item("PKNrGruppeID"))
                'Start und End des Bereichs setzen
                objdtPKNr.Clear()
                objsqlcommand.CommandText = "SELECT RangeStart, RangeEnd " +
                                            "FROM tab_repbetriebe_pknrgruppe " +
                                            "WHERE ID=" + intPKNrGuppenID.ToString + " AND ID<5"
                objdtPKNr.Load(objsqlcommand.ExecuteReader)
                If objdtPKNr.Rows.Count > 0 Then 'Bereichsdefinition gefunden
                    intRangeStart = objdtPKNr.Rows(0).Item("RangeStart")
                    intRangeEnd = objdtPKNr.Rows(0).Item("RangeEnd")
                    'PK - Bereich laden und durchsteppen und Lücke oder nächste PK-Nr suchen
                    'Muss über Dataset gehen da Datatable ein Fehler bringt
                    'objdtPKNr.Clear()

                    objsqlcommand.CommandText = "SELECT PKNr " +
                                                "FROM tab_repbetriebe " +
                                                "WHERE PKNr BETWEEN " + intRangeStart.ToString + " AND " + intRangeEnd.ToString + " " +
                                                "ORDER BY PKNr"
                    'objdtPKNr.Load(objsqlcommand.ExecuteReader)
                    objDAPKNbrs.SelectCommand = objsqlcommand
                    objdsPKNbrs.EnforceConstraints = False
                    objDAPKNbrs.Fill(objdsPKNbrs)

                    intNewPKNr = 0
                    i = intRangeStart
                    If objdsPKNbrs.Tables(0).Rows.Count = 0 Then
                        intNewPKNr = i
                    Else
                        intRecordCounter = 0
                        Do Until intRecordCounter = objdsPKNbrs.Tables(0).Rows.Count
                            If Not objdsPKNbrs.Tables(0).Rows(intRecordCounter).Item("PKNr") = i Then
                                intNewPKNr = i
                                Return 0
                            End If
                            i += 1
                            intRecordCounter += 1
                        Loop
                        If i <= intRangeEnd Then
                            intNewPKNr = i
                        End If
                    End If
                    If intNewPKNr = 0 Then
                        Return 2
                    End If
                Else
                    Return 3
                End If
            Else
                Return 1
            End If

        Catch ex As InvalidCastException
            MessageBox.Show("Rep_Nr " + intRepNr.ToString + " ist keiner Gruppe zugewiesen. Erstellung nicht möglich.", "Gruppe fehlt", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 9

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Problem Debitoren-Nummer-Vergabe Rep_Nr " + intRepNr.ToString, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return 9

        Finally
            If objdbconnZHDB02.State = ConnectionState.Open Then
                'objdbconnZHDB02.Close()
            End If
            objDAPKNbrs.Dispose()
            objdsPKNbrs.Dispose()
            objsqlcommand = Nothing
            objdtPKNr.Dispose()

        End Try


        'Dim db As DAO.DATABASE, RS As DAO.Recordset, RangeStart As Long, RangeEnd As Long, i As Long

        'On Error GoTo ErrHandler
        'Set db = CurrentDb()
        'Set RS = db.OpenRecordset("SELECT RangeStart, RangeEnd" _
        '                            & " FROM tab_repbetriebe_pknrgruppe" _
        '                            & " WHERE ID=" & PKNrGruppeID, dbOpenSnapshot)
        '    If Not RS.EOF Then
        '            RangeStart = RS(0)
        '            RangeEnd = RS(1)
        '        Set RS = db.OpenRecordset("SELECT PKNr" _
        '                                & " FROM Tab_Repbetriebe" _
        '                                & " WHERE PKNr BETWEEN " & RangeStart & " AND " & RangeEnd _
        '                                & " ORDER BY PKNr", dbOpenSnapshot)
        '        PKNr = 0
        '            i = RangeStart
        '            If RS.EOF Then
        '                PKNr = i
        '            Else
        '                Do Until RS.EOF
        '                    If Not RS(0) = i Then
        '                        PKNr = i
        '                        Exit Do
        '                    End If
        '                    i = i + 1
        '                    RS.MoveNext
        '                Loop
        '                If i <= RangeEnd Then PKNr = i
        '            End If
        '            If PKNr = 0 Then
        '                ErrNumber = 2
        '                ErrDescription = "Achtung! Für diese Gruppe ist der Nummerkreis erschöpft (max " & RangeEnd & ")."
        '            Else
        '                NextPKNr = True
        '            End If
        '        Else
        '            ErrNumber = 3
        '            ErrDescription = "Achtung! Für Gruppe mit der ID '" & PKNrGruppeID & "' wurde keine Nummerkreis-Definition gefunden."
        '        End If

        'ExitProc:
        '        On Error Resume Next
        '        RS.Close
        '        'Set RS = Nothing
        '        'Set db = Nothing
        '        Exit Function

        'ErrHandler:
        '        ErrDescription = Err.Description
        '        ErrNumber = Err.Number
        '        If Not errSilent Then ShowErr Err.Number, Err.Description, Err.Source, vbCritical, "Bei Order-Erstellung"
        '        Resume ExitProc

    End Function

    Public Shared Function FcGetIBANDetails(ByRef objdbconn As MySqlConnection,
                                           ByVal strIBAN As String,
                                           ByRef strBankName As String,
                                           ByRef strBankAddress1 As String,
                                           ByRef strBankAddress2 As String,
                                           ByRef strBankBIC As String,
                                           ByRef strBankCountry As String,
                                           ByRef strBankClearing As String) As Int16

        Dim objIBANReq As HttpWebRequest
        Dim objdtIBAN As New DataTable
        'Dim striBANURI As New Uri("https://rest.sepatools.eu/validate_iban_dummy/AL90208110080000001039531801")
        Dim strIBANURI As New Uri("https://ssl.ibanrechner.de/http.html?function=validate_iban&iban=" + strIBAN + "&user=MSSAGSchweiz&password=6ux!mCXiS6EmCiA")
        Dim strResponse As String
        Dim objResponse As HttpWebResponse
        Dim objXMLDoc As New XmlDocument
        Dim objXMLNodeList As XmlNodeList
        Dim strXMLTag(10) As String
        Dim strXMLText(10) As String
        Dim strXMLAddress() As String
        Dim strBalance As String

        Dim objmysqlcom As New MySqlCommand

        Dim intRecAffected As Integer

        Try

            'Zuerst prüfen ob IBAN nicht schon in der Tabelle der bekannten existiert
            objmysqlcom.Connection = objdbconn
            objmysqlcom.CommandText = "SELECT * FROM t_sage_tbliban WHERE strIBANNr='" + strIBAN + "'"
            objdtIBAN.Load(objmysqlcom.ExecuteReader)
            If objdtIBAN.Rows.Count = 0 Then

                objIBANReq = DirectCast(HttpWebRequest.Create(strIBANURI), HttpWebRequest)
                If (objIBANReq.GetResponse().ContentLength > 0) Then
                    objResponse = objIBANReq.GetResponse()
                    'Dim objStreamReader As New StreamReader(objIBANReq.GetResponse().GetResponseStream())
                    Dim objStreamReader As New StreamReader(objResponse.GetResponseStream())
                    'strResponse = objStreamReader.ReadToEnd()
                    objXMLDoc.LoadXml(objStreamReader.ReadToEnd())
                    'Antwort der Funktion
                    objXMLNodeList = objXMLDoc.SelectNodes("/result")
                    For Each objXMLNode As XmlNode In objXMLNodeList
                        'result
                        strXMLTag(0) = objXMLNode.ChildNodes.Item(1).Name
                        strXMLText(0) = objXMLNode.ChildNodes.Item(1).InnerText
                        'return code
                        strXMLTag(1) = objXMLNode.ChildNodes.Item(2).Name
                        strXMLText(1) = objXMLNode.ChildNodes.Item(2).InnerText
                        'country
                        strXMLTag(2) = objXMLNode.ChildNodes.Item(6).Name
                        strXMLText(2) = objXMLNode.ChildNodes.Item(6).InnerText
                        'bank-code
                        strXMLTag(3) = objXMLNode.ChildNodes.Item(7).Name
                        strXMLText(3) = objXMLNode.ChildNodes.Item(7).InnerText
                        'bank
                        strXMLTag(4) = objXMLNode.ChildNodes.Item(8).Name
                        strXMLText(4) = objXMLNode.ChildNodes.Item(8).InnerText
                        'bank address
                        strXMLTag(5) = objXMLNode.ChildNodes.Item(9).Name
                        strXMLAddress = Split(objXMLNode.ChildNodes.Item(9).InnerText, vbLf)
                        If strXMLAddress.Count = 2 Then
                            strXMLText(5) = strXMLAddress(0)
                            strXMLTag(6) = "bank_address2"
                            strXMLText(6) = strXMLAddress(1)
                        ElseIf strXMLAddress.Count = 3 Then
                            strXMLText(5) = strXMLAddress(1)
                            strXMLTag(6) = "bank_address2"
                            strXMLText(6) = strXMLAddress(2)
                        End If
                        strXMLTag(7) = objXMLNode.ChildNodes.Item(39).Name
                        strXMLText(7) = objXMLNode.ChildNodes.Item(39).InnerText
                    Next
                    'BIC
                    objXMLNodeList = objXMLDoc.SelectNodes("/result/bic_candidates-list/bic_candidates")
                    For Each objXMLNode As XmlNode In objXMLNodeList
                        'result
                        strXMLTag(8) = objXMLNode.ChildNodes.Item(0).Name
                        strXMLText(8) = objXMLNode.ChildNodes.Item(0).InnerText
                    Next

                    'objXMLDoc.Load(strResponse)
                    objStreamReader.Close()
                    objResponse.Close()
                    strBankName = Trim(strXMLText(4))
                    strBankAddress1 = Trim(strXMLText(5))
                    strBankAddress2 = Trim(strXMLText(6))
                    strBankCountry = Trim(strXMLText(2))
                    strBankClearing = Trim(strXMLText(3))
                    strBankBIC = Trim(strXMLText(8))

                    'in IBAN-Tabelle schreiben
                    objmysqlcom.CommandText = "INSERT INTO t_sage_tbliban (strIBANNr, 
                                                                        strIBANBankName, 
                                                                        strIBANBankAddress1, 
                                                                        strIBANBankAddress2, 
                                                                        strIBANBankBIC, 
                                                                        strIBANBankCountry, 
                                                                        strIBANBankClearing) " +
                                                            "VALUES('" + strIBAN + "', '" +
                                                            Replace(strBankName, "'", "`") + "', '" +
                                                            Replace(strBankAddress1, "'", "`") + "', '" +
                                                            Replace(strBankAddress2, "'", "`") + "', '" +
                                                            strBankBIC + "', '" +
                                                            strBankCountry + "', '" +
                                                            strBankClearing + "')"
                    intRecAffected = objmysqlcom.ExecuteNonQuery()

                    Return 0

                End If
            Else
                'Aus Tabelle zurückgeben
                strBankName = objdtIBAN.Rows(0).Item("strIBANBankName")
                strBankAddress1 = objdtIBAN.Rows(0).Item("strIBANBankAddress1")
                strBankAddress2 = objdtIBAN.Rows(0).Item("strIBANBankAddress2")
                strBankCountry = objdtIBAN.Rows(0).Item("strIBANBankCountry")
                strBankClearing = objdtIBAN.Rows(0).Item("strIBANBankClearing")
                strBankBIC = objdtIBAN.Rows(0).Item("strIBANBankBIC")

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Fehler auf IBAN-Check " + strIBAN)
            Return 9

        Finally
            objdtIBAN.Dispose()
            objmysqlcom.Dispose()
            objXMLDoc = Nothing
            objResponse = Nothing
            objXMLNodeList = Nothing

        End Try

    End Function

    Public Shared Function FcWriteNewDebToRepbetrieb(ByRef objdbconnZHDB02 As MySqlConnection,
                                                     ByVal intRepNr As Int32,
                                                     intNewDebNr As Int32,
                                                     intAccounting As Int16,
                                                     strMode As String) As Int16

        '0=Update ok, 1=Update hat nicht geklappt, 9=Error

        Dim strSQL As String
        Dim objmysqlcmd As New MySqlCommand
        Dim objdbconn As New MySqlConnection
        Dim intAffected As Int16

        Try

            'Wo ist die Rep_Betriebe?
            If objdbconnZHDB02.State = ConnectionState.Closed Then
                objdbconnZHDB02.Open()
            End If
            If strMode = "D" Then
                objdbconn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(Main.FcReadFromSettings(objdbconnZHDB02, "Buchh_PKTableConnection", intAccounting))
            Else
                objdbconn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(Main.FcReadFromSettings(objdbconnZHDB02, "Buchh_PKKrediTableConnection", intAccounting))
            End If

            If objdbconn.State = ConnectionState.Closed Then
                objdbconn.Open()
            End If
            strSQL = "UPDATE tab_repbetriebe SET PKNr=" + intNewDebNr.ToString + " WHERE Rep_Nr=" + intRepNr.ToString
            objmysqlcmd.Connection = objdbconn
            objmysqlcmd.CommandText = strSQL
            intAffected = objmysqlcmd.ExecuteNonQuery()
            If intAffected <> 1 Then
                Return 1
            Else
                Return 0
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return 9

        Finally
            If objdbconnZHDB02.State = ConnectionState.Open Then
                'objdbconnZHDB02.Close()
            End If
            If objdbconn.State = ConnectionState.Open Then
                objdbconn.Close()
            End If
            objdbconn.Dispose()

        End Try

    End Function

    Public Shared Function FcCheckDebiIntBank(ByRef objdbconn As MySqlConnection,
                                              ByVal intAccounting As Integer,
                                              ByVal striBankS50 As String,
                                              ByRef intIBankS200 As String) As Int16

        '0=ok, 1=Sage50 iBank nicht gefunden, 2=Kein Standard gesetzt, 3=Nichts angegeben, auf Standard gesetzt, 9=Problem

        Dim objdbcommand As New MySqlCommand
        Dim objdtiBank As New DataTable

        Try

            If objdbconn.State = ConnectionState.Closed Then
                objdbconn.Open()
            End If
            'wurde i Bank definiert?
            If striBankS50 <> "" Then
                'Sage 50 - Bank suchen
                objdbcommand.Connection = objdbconn
                'objdbconn.Open()
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
            If objdbconn.State = ConnectionState.Open Then
                'objdbconn.Close()
            End If
            objdtiBank.Constraints.Clear()
            objdtiBank.Clear()
            objdtiBank.Dispose()

        End Try

    End Function

    Public Shared Function FcNextPrivatePKNr(ByRef objdbconnZHDB02 As MySqlConnection,
                                             ByVal intPersNr As Int32,
                                             ByRef intNewPKNr As Int32) As Int16

        '0=ok, 1=Rep - Nr. existiert nicht, 2=Bereich voll, 3=keine Bereichdefinition 9=Problem

        'PK - Nummer soll der Funktion gegeben werden, Funktion sucht sich dann die PK_Gruppe 
        'Konzept: Tabelle füllen und dann durchsteppen
        Dim objsqlcommand As New MySqlCommand
        Dim objdtPKNr As New DataTable
        Dim intPKNrGuppenID As Int16
        Dim intRangeStart, intRangeEnd, i, intRecordCounter As Int32
        Dim objdsPKNbrs As New DataSet
        Dim objDAPKNbrs As New MySqlDataAdapter
        Dim objDAPersons As New MySqlDataAdapter
        Dim objdsPersons As New DataSet

        Try

            If objdbconnZHDB02.State = ConnectionState.Closed Then
                objdbconnZHDB02.Open()
            End If
            objsqlcommand.Connection = objdbconnZHDB02
            objsqlcommand.CommandText = "SELECT PKNrGruppeID FROM t_customer WHERE ID=" + intPersNr.ToString
            objDAPersons.SelectCommand = objsqlcommand
            objdsPersons.EnforceConstraints = False
            objDAPersons.Fill(objdsPersons)

            If objdsPersons.Tables(0).Rows.Count > 0 Then 'Person gefunden
                intPKNrGuppenID = objdsPersons.Tables(0).Rows(0).Item("PKNrGruppeID")
                'Start und End des Bereichs setzen
                objdtPKNr.Clear()
                objsqlcommand.CommandText = "SELECT RangeStart, RangeEnd " +
                                            "FROM tab_repbetriebe_pknrgruppe " +
                                            "WHERE ID=" + intPKNrGuppenID.ToString
                objdtPKNr.Load(objsqlcommand.ExecuteReader)
                If objdtPKNr.Rows.Count > 0 Then 'Bereichsdefinition gefunden
                    intRangeStart = objdtPKNr.Rows(0).Item("RangeStart")
                    intRangeEnd = objdtPKNr.Rows(0).Item("RangeEnd")
                    'PK - Bereich laden und durchsteppen und Lücke oder nächste PK-Nr suchen
                    'Muss über Dataset gehen da Datatable ein Fehler bringt
                    'objdtPKNr.Clear()

                    objsqlcommand.CommandText = "SELECT PKNr " +
                                                "FROM t_customer " +
                                                "WHERE PKNr BETWEEN " + intRangeStart.ToString + " AND " + intRangeEnd.ToString + " " +
                                                "ORDER BY PKNr"
                    'objdtPKNr.Load(objsqlcommand.ExecuteReader)
                    objDAPKNbrs.SelectCommand = objsqlcommand
                    objdsPKNbrs.EnforceConstraints = False
                    objDAPKNbrs.Fill(objdsPKNbrs)

                    intNewPKNr = 0
                    i = intRangeStart
                    If objdsPKNbrs.Tables(0).Rows.Count = 0 Then
                        intNewPKNr = i
                    Else
                        intRecordCounter = 0
                        Do Until intRecordCounter = objdsPKNbrs.Tables(0).Rows.Count
                            If Not objdsPKNbrs.Tables(0).Rows(intRecordCounter).Item("PKNr") = i Then
                                intNewPKNr = i
                                Return 0
                            End If
                            i += 1
                            intRecordCounter += 1
                        Loop
                        If i <= intRangeEnd Then
                            intNewPKNr = i
                        End If
                    End If
                    If intNewPKNr = 0 Then
                        Return 2
                    End If
                Else
                    Return 3
                End If
            Else
                Return 1
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return 9

        Finally
            If objdbconnZHDB02.State = ConnectionState.Open Then
                objdbconnZHDB02.Close()
            End If
            objDAPKNbrs.Dispose()
            objdsPKNbrs.Dispose()
            objsqlcommand = Nothing
            objdtPKNr.Dispose()

        End Try

    End Function

    Public Shared Function FcWriteNewPrivateDebToRepbetrieb(ByRef objdbconnZHDB02 As MySqlConnection,
                                                            ByVal intPersNr As Int32,
                                                            intNewDebNr As Int32) As Int16

        '0=Update ok, 1=Update hat nicht geklappt, 9=Error

        Dim strSQL As String
        Dim objmysqlcmd As New MySqlCommand
        Dim intAffected As Int16

        Try

            strSQL = "UPDATE t_customer SET PKNr=" + intNewDebNr.ToString + " WHERE ID=" + intPersNr.ToString
            objdbconnZHDB02.Open()
            objmysqlcmd.Connection = objdbconnZHDB02
            objmysqlcmd.CommandText = strSQL
            intAffected = objmysqlcmd.ExecuteNonQuery()
            If intAffected <> 1 Then
                Return 1
            Else
                Return 0
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return 9

        Finally
            If objdbconnZHDB02.State = ConnectionState.Open Then
                'objdbconnZHDB02.Close()
            End If

        End Try

    End Function

End Class
