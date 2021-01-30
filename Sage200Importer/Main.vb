Option Strict Off
Option Explicit On

Imports MySql.Data.MySqlClient
Imports System.Data.OracleClient
Imports System.Net
Imports System.IO
Imports System.Xml

'Imports System.Data.OleDb

Friend NotInheritable Class Main

    Public Shared Function tblDebitorenHead() As DataTable
        Dim DT As DataTable
        'Dim myNewRow As DataRow
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
        Dim booLinked As DataColumn = New DataColumn("booLinked")
        booLinked.DataType = System.Type.[GetType]("System.Boolean")
        DT.Columns.Add(booLinked)
        Dim strRGName As DataColumn = New DataColumn("strRGName")
        strRGName.DataType = System.Type.[GetType]("System.String")
        strRGName.MaxLength = 50
        DT.Columns.Add(strRGName)
        Dim strOPNr As DataColumn = New DataColumn("strOPNr")
        strOPNr.DataType = System.Type.[GetType]("System.String")
        strOPNr.MaxLength = 13
        DT.Columns.Add(strOPNr)
        Dim lngDebNbr As DataColumn = New DataColumn("lngDebNbr")
        lngDebNbr.DataType = System.Type.[GetType]("System.Int32")
        DT.Columns.Add(lngDebNbr)
        Dim strDebPKBez As DataColumn = New DataColumn("strDebBez")
        strDebPKBez.DataType = System.Type.[GetType]("System.String")
        strDebPKBez.MaxLength = 50
        DT.Columns.Add(strDebPKBez)
        Dim lngDebKtoNbr As DataColumn = New DataColumn("lngDebKtoNbr")
        lngDebKtoNbr.DataType = System.Type.[GetType]("System.Int32")
        DT.Columns.Add(lngDebKtoNbr)
        Dim strDebKtoBez As DataColumn = New DataColumn("strDebKtoBez")
        strDebKtoBez.DataType = System.Type.[GetType]("System.String")
        strDebKtoBez.MaxLength = 50
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
        Dim datDebRGDatum As DataColumn = New DataColumn("datDebRGDatum")
        datDebRGDatum.DataType = System.Type.[GetType]("System.DateTime")
        DT.Columns.Add(datDebRGDatum)
        Dim datDebValDatum As DataColumn = New DataColumn("datDebValDatum")
        datDebValDatum.DataType = System.Type.[GetType]("System.DateTime")
        DT.Columns.Add(datDebValDatum)
        Dim strDebiBank As DataColumn = New DataColumn("strDebiBank")
        strDebiBank.DataType = System.Type.[GetType]("System.String")
        strDebiBank.MaxLength = 5
        DT.Columns.Add(strDebiBank)
        Dim strDebRef As DataColumn = New DataColumn("strDebRef")
        strDebRef.DataType = System.Type.[GetType]("System.String")
        strDebRef.MaxLength = 27
        DT.Columns.Add(strDebRef)
        Dim strZahlBed As DataColumn = New DataColumn("strZahlBed")
        strZahlBed.DataType = System.Type.[GetType]("System.String")
        strZahlBed.MaxLength = 5
        DT.Columns.Add(strZahlBed)
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
        Return DT
    End Function

    Public Shared Function tblDebitorenSub() As DataTable
        Dim DT As DataTable
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
        strKtoBez.MaxLength = 50
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
    End Function

    Public Shared Function tblKreditorenHead() As DataTable
        Dim DT As DataTable
        'Dim myNewRow As DataRow
        DT = New DataTable("tblKreditorenHead")
        Dim lngKredID As DataColumn = New DataColumn("lngKredID")
        lngKredID.DataType = System.Type.[GetType]("System.Int32")
        DT.Columns.Add(lngKredID)
        DT.PrimaryKey = New DataColumn() {DT.Columns("lngKredID")}
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
        strKredRef.MaxLength = 27
        DT.Columns.Add(strKredRef)
        'Dim strZahlBed As DataColumn = New DataColumn("strZahlBed")
        'strZahlBed.DataType = System.Type.[GetType]("System.String")
        'strZahlBed.MaxLength = 5
        'DT.Columns.Add(strZahlBed)
        Dim intPayType As DataColumn = New DataColumn("intPayType")
        intPayType.DataType = System.Type.[GetType]("System.Int16")
        DT.Columns.Add(intPayType)
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
        Return DT

    End Function

    Public Shared Function tblKreditorenSub() As DataTable
        Dim DT As DataTable
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
    End Function

    Public Shared Function tblInfo() As DataTable

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

    End Function


    Public Shared Function FcLoginSage(ByRef objdbconn As MySqlConnection,
                                       ByRef objsqlConn As SqlClient.SqlConnection,
                                       ByRef objsqlCom As SqlClient.SqlCommand,
                                       ByRef objFinanz As SBSXASLib.AXFinanz,
                                       ByRef objfiBuha As SBSXASLib.AXiFBhg,
                                       ByRef objdbBuha As SBSXASLib.AXiDbBhg,
                                       ByRef objdbPIFb As SBSXASLib.AXiPlFin,
                                       ByRef objkrBuha As SBSXASLib.AXiKrBhg,
                                       ByVal intAccounting As Int16,
                                       ByRef objdtInfo As DataTable,
                                       ByVal strPeriod As String) As Int16


        '0=ok, 1=Fibu nicht ok, 2=Debi nicht ok, 3=Debi nicht ok

        Dim booAccOk As Boolean
        Dim strMandant As String
        Dim b As Object
        Dim strLogonInfo() As String
        Dim strPeriode() As String
        Dim FcReturns As Int16

        b = Nothing

        objFinanz = Nothing
        objFinanz = New SBSXASLib.AXFinanz


        On Error GoTo ErrorHandler

        'Loign
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
        Dim intPeriodenNr As Int16
        Dim strPeriodenInfo As String
        'strPeriodenInfo = objFinanz.GetLogonInfo()
        intPeriodenNr = objFinanz.ReadPeri(strMandant, strLogonInfo(7))
        'For intLooper As Int16 = 0 To intPeriodenNr
        strPeriodenInfo = objFinanz.GetPeriListe(0)
        'strPeriodenInfo = objFinanz.GetResource(intLooper)
        'Next
        strPeriode = Split(strPeriodenInfo, "{>}")
        objdtInfo.Rows.Add("GeschäftsJ", strPeriode(3) + "-" + strPeriode(4))
        objdtInfo.Rows.Add("Buchungen/ Status", strPeriode(5) + "-" + strPeriode(6) + "/ " + strPeriode(2))
        'objdtInfo.Rows.Add("Status", strPeriode(2))
        'Debug.Print(FcReadPeriodenDef(objsqlConn, objsqlCom, strPeriode(8))(0))

        'objdtInfo.Rows.Add("Perioden-Def", FcReadPeriodenDef(objsqlConn, objsqlCom, strPeriode(8))(0))
        'objdtInfo.Rows.Add("Defintion von", FcReadPeriodenDef(objsqlConn, objsqlCom, strPeriode(8))(1))

        FcReturns = FcReadPeriodenDef(objsqlConn, objsqlCom, strPeriode(8), objdtInfo)


        If b = 0 Then GoTo isOk
        b = b - 200
        MsgBox("Mandant oder Periode falsch - Programm beendet", 0, "Fehler")
        objFinanz = Nothing
        End

isOk:
        'Finanz Buha öffnen
        objfiBuha = Nothing
        objfiBuha = New SBSXASLib.AXiFBhg
        objfiBuha = objFinanz.GetFibuObj
        'Debitor öffnen
        objdbBuha = Nothing
        objdbBuha = New SBSXASLib.AXiDbBhg
        objdbBuha = objFinanz.GetDebiObj
        objdbPIFb = Nothing
        objdbPIFb = objfiBuha.GetCheckObj
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
        Dim strPeriodenListe As String = ""
        Dim strPeriodeAr() As String
        Dim intLooper As Int16

        objFinanz = Nothing
        objFinanz = New SBSXASLib.AXFinanz


        'Loign
        Call objFinanz.ConnectSBSdb(System.Configuration.ConfigurationManager.AppSettings("OwnSageServer"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageDB"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSageID"),
                                    System.Configuration.ConfigurationManager.AppSettings("OwnSagePsw"), "")

        objdbconn.Open()
        strMandant = FcReadFromSettings(objdbconn, "Buchh200_Name", intAccounting)
        objdbconn.Close()
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


    End Function

    Public Shared Function FcReadPeriodenDef(ByRef objSQLConnection As SqlClient.SqlConnection, ByRef objSQLCommand As SqlClient.SqlCommand, ByVal intPeriodenNr As Int32, ByRef objdtInfo As DataTable) As Int16

        'Returns 0=definiert, 1=nicht defeniert, 9=Problem

        Dim objlocdtPeriDef As New DataTable
        Dim strPeriodenDef(4) As String

        Try

            objSQLConnection.Open()
            objSQLCommand.CommandText = "SELECT * FROM peridef WHERE teqnbr=" + intPeriodenNr.ToString
            objSQLCommand.Connection = objSQLConnection
            objlocdtPeriDef.Load(objSQLCommand.ExecuteReader)

            If objlocdtPeriDef.Rows.Count = 1 Then 'Perioden-Definition vorhanden
                strPeriodenDef(0) = objlocdtPeriDef.Rows(0).Item(2) 'Bezeichnung
                strPeriodenDef(1) = objlocdtPeriDef.Rows(0).Item(3).ToString  'Von
                strPeriodenDef(2) = objlocdtPeriDef.Rows(0).Item(4).ToString  'Bis
                strPeriodenDef(3) = objlocdtPeriDef.Rows(0).Item(5)  'Status

                objdtInfo.Rows.Add("Perioden-Def", strPeriodenDef(0))
                objdtInfo.Rows.Add("Von - Bis/ Status", strPeriodenDef(1) + " - " + strPeriodenDef(2) + "/ " + strPeriodenDef(3))

                Return 0
            Else

                objdtInfo.Rows.Add("Perioden-Def", "keine")
                objdtInfo.Rows.Add("Von - Bis/ Status", "01.01." + Year(Today()).ToString + " 00:00:00 - " + "31.12." + Year(Today()).ToString + " 23:59:59/ " + "O")

                Return 1

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Periodendefinition lesen")
            Return 9

        Finally
            objSQLConnection.Close()

        End Try

    End Function

    Public Shared Function FcReadBankSettings(ByVal intAccounting As Int16, ByVal strBank As String, ByRef objdbconn As MySqlConnection) As String

        Dim objlocdtBank As New DataTable("tbllocBank")
        Dim objlocMySQLcmd As New MySqlCommand

        Try
            objlocMySQLcmd.CommandText = "SELECT strBLZ FROM t_sage_tblaccountingbank WHERE intAccountingID=" + intAccounting.ToString + " AND strBank='" + strBank + "'"
            objlocMySQLcmd.Connection = objdbconn
            objlocdtBank.Load(objlocMySQLcmd.ExecuteReader)


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Bankleitzahl suchen.")

        End Try

        Return objlocdtBank.Rows(0).Item(0).ToString

    End Function


    Public Shared Function FcReadFromSettings(ByRef objdbconn As MySqlConnection, ByVal strField As String, ByVal intMandant As Int16) As String

        Dim objlocdtSetting As New DataTable("tbllocSettings")
        Dim objlocMySQLcmd As New MySqlCommand

        Try

            objlocMySQLcmd.CommandText = "SELECT t_sage_buchhaltungen." + strField + " FROM t_sage_buchhaltungen WHERE Buchh_Nr=" + intMandant.ToString
            'Debug.Print(objlocMySQLcmd.CommandText)
            objlocMySQLcmd.Connection = objdbconn
            objlocdtSetting.Load(objlocMySQLcmd.ExecuteReader)
            'Debug.Print("Records" + objlocdtSetting.Rows.Count.ToString)


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Einstellung lesen")

        End Try

        'Debug.Print("Return " + objlocdtSetting.Rows(0).Item(0).ToString)
        Return objlocdtSetting.Rows(0).Item(0).ToString

    End Function

    Public Shared Function FcCheckDebit(ByVal intAccounting As Integer,
                                        ByRef objdtDebits As DataTable,
                                        ByRef objdtDebitSubs As DataTable,
                                        ByRef objFinanz As SBSXASLib.AXFinanz,
                                        ByRef objfiBuha As SBSXASLib.AXiFBhg,
                                        ByRef objdbBuha As SBSXASLib.AXiDbBhg,
                                        ByRef objdbPIFb As SBSXASLib.AXiPlFin,
                                        ByRef objdbconn As MySqlConnection,
                                        ByRef objdbconnZHDB02 As MySqlConnection,
                                        ByRef objsqlcommand As MySqlCommand,
                                        ByRef objsqlcommandZHDB02 As MySqlCommand,
                                        ByRef objOrdbconn As OracleClient.OracleConnection,
                                        ByRef objOrcommand As OracleClient.OracleCommand,
                                        ByRef objdbAccessConn As OleDb.OleDbConnection,
                                        ByRef objdtInfo As DataTable,
                                        ByVal strcmbBuha As String) As Integer

        'DebiBitLog 1=PK, 2=Konto, 3=Währung, 4=interne Bank, 5=OP Kopf, 6=RG-Datum, 7=Valuta Datum, 8=Subs, 9=OP doppelt
        Dim strBitLog As String = ""
        Dim intReturnValue As Integer
        Dim strStatus As String = ""
        Dim intSubNumber As Int16
        Dim dblSubNetto As Double
        Dim dblSubMwSt As Double
        Dim dblSubBrutto As Double
        Dim booAutoCorrect As Boolean
        Dim selsubrow() As DataRow
        Dim strDebiReferenz As String = ""
        Dim booDiffHeadText As Boolean
        Dim strDebiHeadText As String
        Dim booDiffSubText As Boolean
        Dim strDebiSubText As String
        Dim intDebitorNew As Int32
        Dim intiBankSage200 As Int16
        Dim dblRDiffNetto As Double
        Dim dblRDiffMwSt As Double
        Dim dblRDiffBrutto As Double
        'Dim objdrDebiSub As DataRow = objdtDebitSubs.NewRow

        Try

            objdbconn.Open()
            objOrdbconn.Open()
            'objdbAccessConn.Open()

            For Each row As DataRow In objdtDebits.Rows

                'If row("strDebRGNbr") = "106473" Then Stop

                'Runden
                row("dblDebNetto") = Decimal.Round(row("dblDebNetto"), 2, MidpointRounding.AwayFromZero)
                row("dblDebMwSt") = Decimal.Round(row("dblDebMwst"), 2, MidpointRounding.AwayFromZero)
                row("dblDebBrutto") = Decimal.Round(row("dblDebBrutto"), 2, MidpointRounding.AwayFromZero)

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
                If intReturnValue = 1 Then 'Neue Debi wurde angelegt
                    strStatus = "NDeb "
                End If
                If intDebitorNew <> 0 Then
                    intReturnValue = MainDebitor.FcCheckDebitor(intDebitorNew, row("intBuchungsart"), objdbBuha)
                Else
                    intReturnValue = 2
                End If
                strBitLog = Trim(intReturnValue.ToString)

                'Kto 02
                intReturnValue = FcCheckKonto(row("lngDebKtoNbr"), objfiBuha, row("dblDebMwSt"), 0)
                strBitLog += Trim(intReturnValue.ToString)

                'Currency 03
                intReturnValue = FcCheckCurrency(row("strDebCur"), objfiBuha)
                strBitLog += Trim(intReturnValue.ToString)

                'Sub 04
                booAutoCorrect = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettings(objdbconn, "Buchh_HeadAutoCorrect", intAccounting)))
                intReturnValue = FcCheckSubBookings(row("strDebRGNbr"), objdtDebitSubs, intSubNumber, dblSubBrutto, dblSubNetto, dblSubMwSt, objdbconn, objfiBuha, objdbPIFb, row("intBuchungsart"), booAutoCorrect)
                strBitLog += Trim(intReturnValue.ToString)

                'Autokorrektur 05
                'booAutoCorrect = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettings(objdbconn, "Buchh_HeadAutoCorrect", intAccounting)))
                If booAutoCorrect And row("intBuchungsart") = 1 Then
                    'Git es etwas zu korrigieren?
                    If IIf(IsDBNull(row("dblDebBrutto")), 0, row("dblDebBrutto")) <> dblSubBrutto Or
                        IIf(IsDBNull(row("dblDebNetto")), 0, row("dblDebNetto")) <> dblSubNetto Or
                        IIf(IsDBNull(row("dblDebMwSt")), 0, row("dblDebMwSt")) <> dblSubMwSt Then
                        row("dblDebBrutto") = dblSubBrutto * -1
                        row("dblDebNetto") = dblSubNetto * -1
                        row("dblDebMwSt") = dblSubMwSt * -1
                        ''In Sub korrigieren
                        'selsubrow = objdtDebitSubs.Select("strRGNr='" + row("strDebRGNbr") + "' AND intSollHaben=2")
                        'If selsubrow.Length = 1 Then
                        '    selsubrow(0).Item("dblBrutto") = dblSubBrutto * -1
                        '    selsubrow(0).Item("dblMwSt") = dblSubMwSt * -1
                        '    selsubrow(0).Item("dblNetto") = dblSubNetto * -1
                        'End If
                        strBitLog += "1"
                    Else
                        strBitLog += "0"
                    End If
                Else
                    If row("intBuchungsart") = 1 Then

                        dblRDiffBrutto = 0
                        row("dblDebNetto") = dblSubNetto * -1
                        row("dblDebMwSt") = dblSubMwSt * -1

                        'Für evtl. Rundungsdifferenzen einen Datensatz in die Sub-Tabelle hinzufügen
                        If IIf(IsDBNull(row("dblDebBrutto")), 0, row("dblDebBrutto")) + dblSubBrutto <> 0 Then
                            'Or IIf(IsDBNull(row("dblDebMwSt")), 0, row("dblDebMwSt")) + dblSubMwSt <> 0 _
                            'Or IIf(IsDBNull(row("dblDebNetto")), 0, row("dblDebNetto")) + dblSubNetto <> 0 Then

                            'row("dblDebNetto") = dblSubNetto * -1
                            'row("dblDebMwSt") = dblSubMwSt * -1

                            dblRDiffBrutto = Decimal.Round(row("dblDebBrutto") + dblSubBrutto, 2, MidpointRounding.AwayFromZero)
                            dblRDiffMwSt = 0 'Decimal.Round(row("dblDebMwSt") + dblSubMwSt, 2, MidpointRounding.AwayFromZero)
                            dblRDiffNetto = 0 'Decimal.Round(row("dblDebNetto") + dblSubNetto, 2, MidpointRounding.AwayFromZero)

                            'Zu sub-Table hinzifügen
                            Dim objdrDebiSub As DataRow = objdtDebitSubs.NewRow
                            objdrDebiSub("strRGNr") = row("strDebRGNbr")
                            objdrDebiSub("intSollHaben") = 1
                            objdrDebiSub("lngKto") = 6906
                            objdrDebiSub("strKtoBez") = "Rundungsdifferenzen"
                            objdrDebiSub("lngKST") = 999999
                            objdrDebiSub("strKstBez") = "SystemKST"
                            objdrDebiSub("dblNetto") = dblRDiffNetto
                            objdrDebiSub("dblMwSt") = dblRDiffMwSt
                            objdrDebiSub("dblBrutto") = dblRDiffBrutto
                            objdrDebiSub("dblMwStSatz") = 0
                            objdrDebiSub("strMwStKey") = "null"
                            objdrDebiSub("strArtikel") = "Rundungsdifferenz"
                            objdrDebiSub("strDebSubText") = "Eingefügt"
                            objdrDebiSub("strStatusUBBitLog") = "00000000"
                            If Math.Abs(dblRDiffBrutto) > 1 Then
                                objdrDebiSub("strStatusUBText") = "Rund > 1"
                            Else
                                objdrDebiSub("strStatusUBText") = "ok"
                            End If
                            objdtDebitSubs.Rows.Add(objdrDebiSub)
                            'Summe der Sub-Buchungen anpassen
                            dblSubBrutto = Decimal.Round(dblSubBrutto - dblRDiffBrutto, 2, MidpointRounding.AwayFromZero)
                            'dblSubMwSt = Decimal.Round(dblSubMwSt - dblRDiffMwSt, 2, MidpointRounding.AwayFromZero)
                            'dblSubNetto = Decimal.Round(dblSubNetto - dblRDiffNetto, 2, MidpointRounding.AwayFromZero)
                            If Math.Abs(dblRDiffBrutto) > 1 Then
                                strBitLog += "1"
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
                intReturnValue = FcCreateDebRef(objdbconn, intAccounting, row("strDebiBank"), row("strDebRGNbr"), row("strOPNr"), row("intBuchungsart"), strDebiReferenz)
                strBitLog += Trim(intReturnValue.ToString)

                'Status-String auswerten, vorziehen um neue PK - Nummer auszulesen
                'Debitor
                If Left(strBitLog, 1) <> "0" Then
                    strStatus += "Deb"
                    If Left(strBitLog, 1) <> "2" Then
                        intReturnValue = MainDebitor.FcIsDebitorCreatable(objdbconn, objdbconnZHDB02, objsqlcommandZHDB02, intDebitorNew, objdbBuha, strcmbBuha, intAccounting)
                        If intReturnValue = 0 Then
                            strStatus += " erstellt"
                        Else
                            strStatus += " nicht erstellt."
                        End If
                        row("strDebBez") = MainDebitor.FcReadDebitorName(objdbBuha, intDebitorNew, row("strDebCur"))
                        row("lngDebNbr") = intDebitorNew
                    Else
                        strStatus += " keine Ref"
                        row("strDebBez") = "n/a"
                    End If
                Else
                    If row("intBuchungsart") = 1 Then
                        row("strDebBez") = MainDebitor.FcReadDebitorName(objdbBuha, intDebitorNew, row("strDebCur"))
                    Else
                        row("strDebBez") = "Nicht relevant"
                    End If
                    row("lngDebNbr") = intDebitorNew
                End If

                'OP - Verdopplung 09
                intReturnValue = FcCheckOPDouble(objdbBuha, IIf(IsDBNull(row("lngDebNbr")), 0, row("lngDebNbr")), row("strDebRGNbr"))
                strBitLog += Trim(intReturnValue.ToString)
                'Valuta - Datum 10
                intReturnValue = FcChCeckDate(IIf(IsDBNull(row("datDebValDatum")), #1789-09-17#, row("datDebValDatum")), objdtInfo)
                strBitLog += Trim(intReturnValue.ToString)
                'RG - Datum 11
                intReturnValue = FcChCeckDate(IIf(IsDBNull(row("datDebRGDatum")), #1789-09-17#, row("datDebRGDatum")), objdtInfo)
                strBitLog += Trim(intReturnValue.ToString)
                'Interne Bank 12
                intReturnValue = MainDebitor.FcCheckDebiIntBank(objdbconn, intAccounting, IIf(IsDBNull(row("strDebiBank")), "", row("strDebiBank")), intiBankSage200)
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
                    row("strDebKtoBez") = MainDebitor.FcReadDebitorKName(objfiBuha, row("lngDebKtoNbr"))
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
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "AutoC"
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
                Else
                    row("strDebRef") = strDebiReferenz
                End If
                'OP
                If Mid(strBitLog, 9, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "OPDbl"
                    'Else
                    '    row("strDebRef") = strDebiReferenz
                End If
                'Valuta Datum 
                If Mid(strBitLog, 10, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "ValD"
                    'Else
                    '    row("strDebRef") = strDebiReferenz
                End If
                'RG Datum 
                If Mid(strBitLog, 11, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "RgD"
                    'Else
                    '    row("strDebRef") = strDebiReferenz
                End If
                'interne Bank
                If Mid(strBitLog, 12, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "iBnk"
                Else
                    row("strDebiBank") = intiBankSage200
                End If

                'Status schreiben
                If Val(strBitLog) = 0 Or Val(strBitLog) = 10000000 Then
                    row("booDebBook") = True
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "ok"
                End If
                row("strDebStatusText") = strStatus
                row("strDebStatusBitLog") = strBitLog

                'Wird ein anderer Text in der Head-Buchung gewünscht?
                booDiffHeadText = IIf(FcReadFromSettings(objdbconn, "Buchh_TextSpecial", intAccounting) = "0", False, True)
                If booDiffHeadText Then
                    strDebiHeadText = MainDebitor.FcSQLParse(FcReadFromSettings(objdbconn, "Buchh_TextSpecialText", intAccounting), row("strDebRGNbr"), objdtDebits, objOrdbconn, objOrcommand)
                    row("strDebText") = strDebiHeadText
                End If

                'Wird ein anderer Text in den Sub-Buchung gewünscht?
                booDiffSubText = IIf(FcReadFromSettings(objdbconn, "Buchh_SubTextSpecial", intAccounting) = "0", False, True)
                If booDiffSubText Then
                    strDebiSubText = MainDebitor.FcSQLParse(FcReadFromSettings(objdbconn, "Buchh_SubTextSpecialText", intAccounting), row("strDebRGNbr"), objdtDebits, objOrdbconn, objOrcommand)
                Else
                    strDebiSubText = row("strDebText")
                End If
                selsubrow = objdtDebitSubs.Select("strRGNr='" + row("strDebRGNbr") + "'")
                            For Each subrow In selsubrow
                    subrow("strDebSubText") = strDebiSubText
                Next

                'Init
                strBitLog = ""
                strStatus = ""
                intSubNumber = 0
                dblSubBrutto = 0
                dblSubNetto = 0
                dblSubMwSt = 0

            Next

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Debitor Kopfdaten-Check")

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
        End Try


    End Function


    Public Shared Function FcChCeckDate(ByVal datDateToCheck As Date, ByRef objdtInfo As DataTable) As Int16

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
                    If objdtInfo.Rows.Count > 3 Then
                        intActualLine = 4
                        booPeriodeOpen = True
                        Do While intActualLine < objdtInfo.Rows.Count
                            'Wurden zusätzliche Perioden defniert und falls ja, ist der Status offen?
                            datPerVon = Convert.ToDateTime(Left(objdtInfo.Rows(intActualLine).Item(1), 10) + " 00:00:01")
                            datPerBis = Convert.ToDateTime(Mid(objdtInfo.Rows(intActualLine).Item(1), 23, 10) + " 23:59:59")
                            booBuhaOpen = IIf(Right(objdtInfo.Rows(intActualLine).Item(1), 1) = "O", True, False)
                            If datDateToCheck >= datPerVon And datDateToCheck <= datPerBis Then
                                If booBuhaOpen Then
                                    booPeriodeOpen = True
                                Else
                                    booPeriodeOpen = False
                                End If
                            End If
                            intActualLine += 2
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
                    Return 1
                End If
            Else
                Return 1
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Datumscheck")
            Return 9

        End Try


    End Function

    Public Shared Function FcCheckOPDouble(ByRef objdbBuha As SBSXASLib.AXiDbBhg, ByVal strDebitor As String, ByVal strOPNr As String) As Int16

        'Return 0=ok, 1=Beleg existiert, 9=Problem

        Dim intBelegReturn As Int16

        Try
            intBelegReturn = objdbBuha.doesBelegExist(strDebitor, "CHF", strOPNr, "0", "", "")
            If intBelegReturn = 0 Then
                Return 0
            Else
                Return 1
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Check doppelte OP - Nr.")
            Return 9

        End Try

    End Function


    Public Shared Function FcCreateDebRef(ByRef objdbconn As MySqlConnection,
                                          ByVal intAccounting As Integer,
                                          ByVal strBank As String,
                                          ByVal strRGNr As String,
                                          ByVal strOPNr As String,
                                          ByVal intBuchungsArt As Integer,
                                          ByRef strReferenz As String) As Integer

        'Return 0=ok oder nicht nötig, 1=keine Angaben hinterlegt, 2=Berechnung hat nicht geklappt

        Dim strTLNNr As String
        Dim strCleanedNr As String = ""
        Dim strRefFrom As String

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
                    Case "O"
                        strCleanedNr = strOPNr

                End Select

                strTLNNr = FcReadBankSettings(intAccounting, strBank, objdbconn)

                strCleanedNr = FcCleanRGNrStrict(strCleanedNr)

                strReferenz = strTLNNr + StrDup(20 - Len(strCleanedNr), "0") + strCleanedNr + Trim(CStr(FcModulo10(strTLNNr + StrDup(20 - Len(strCleanedNr), "0") + strCleanedNr)))

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

        intTabelle(0) = 0 : intTabelle(1) = 9
        intTabelle(2) = 4 : intTabelle(3) = 6
        intTabelle(4) = 8 : intTabelle(5) = 2
        intTabelle(6) = 7 : intTabelle(7) = 1
        intTabelle(8) = 3 : intTabelle(9) = 5

        For intIndex = 1 To Len(strNummer)
            intÜbertrag = intTabelle((intÜbertrag + Mid(strNummer, intIndex, 1)) Mod 10)
        Next

        Return (10 - intÜbertrag) Mod 10

    End Function


    Public Shared Function FcCleanRGNrStrict(ByVal strRGNrToClean As String) As String

        Dim intCounter As Int16
        Dim strCleanRGNr As String = ""

        For intCounter = 1 To Len(strRGNrToClean)
            If Mid(strRGNrToClean, intCounter, 1) = "0" Or Val(Mid(strRGNrToClean, intCounter, 1)) > 0 Then
                strCleanRGNr += Mid(strRGNrToClean, intCounter, 1)
            End If

        Next

        Return strCleanRGNr

    End Function

    Public Shared Function FcCheckBelegHead(ByVal intBuchungsArt As Int16,
                                            ByVal dblBrutto As Double,
                                            ByVal dblNetto As Double,
                                            ByVal dblMwSt As Double,
                                            ByVal dblRDiff As Double) As Int16

        'Returns 0=ok oder nicht wichtig, 1=Brutto, 2=Netto, 3=Beide, 4=Diff

        If intBuchungsArt = 1 Then
            If dblBrutto = 0 And dblNetto = 0 Then
                Return 3
            ElseIf dblBrutto = 0 Then
                Return 1
            ElseIf dblNetto = 0 Then
                Return 2
            ElseIf Math.Round(dblBrutto - dblRDiff - dblMwSt, 2, MidpointRounding.AwayFromZero) <> Math.Round(dblNetto, 2, MidpointRounding.AwayFromZero) Then
                Return 4
            Else
                Return 0
            End If
        Else
            Return 0
        End If

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
        Dim strSteuerRec As String = ""
        'Dim strSteuerRecAr() As String
        Dim intLooper As Int16 = 0

        Try

            'Falls MwStKey 'ohne' und Konto >= 3000 und 3999 dann ohne = frei
            If strStrCode = "ohne" Then
                If intKonto >= 3000 And intKonto <= 3999 Then
                    strStrCode = "frei"
                End If
            End If

            'Besprechung mit Muhi 20201209 => Es soll eine fixe Vergabe des MStSchlüssels passieren 
            objlocMySQLcmd.CommandText = "SELECT  * FROM t_sage_sage50mwst WHERE strKey='" + strStrCode + "'"

            objlocMySQLcmd.Connection = objdbconn
            objlocdtMwSt.Load(objlocMySQLcmd.ExecuteReader)

            If objlocdtMwSt.Rows.Count = 0 Then
                MessageBox.Show("MwSt " + strStrCode + " ist nicht definiert für Sage 50 MsSt-Key " + strStrCode + ".")
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
                                              ByVal intBuchungsArt As Int32,
                                              ByVal booAutoCorrect As Boolean) As Int16

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
        Dim strStrStCodeSage200 As String = ""
        Dim strKstKtrSage200 As String = ""
        Dim selsubrow() As DataRow
        Dim strStatusOverAll As String = "0000000"
        Dim strSteuer() As String

        'Summen bilden und Angaben prüfen
        intSubNumber = 0
        dblSubNetto = 0
        dblSubMwSt = 0
        dblSubBrutto = 0

        selsubrow = objDtDebiSub.Select("strRGNr='" + strDebRgNbr + "'")

        For Each subrow As DataRow In selsubrow

            strBitLog = ""

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

            'Zuerst evtl. falsch gesetzte KTR oder Steuer - Sätze prüfen
            If subrow("lngKto") < 3000 Then
                subrow("strMwStKey") = Nothing
                subrow("lngKST") = 0
            End If

            'MwSt prüfen
            If Not IsDBNull(subrow("strMwStKey")) Then
                intReturnValue = FcCheckMwSt(objdbconn, objFiBhg, subrow("strMwStKey"), IIf(IsDBNull(subrow("dblMwStSatz")), 0, subrow("dblMwStSatz")), strStrStCodeSage200, subrow("lngKto"))
                If intReturnValue = 0 Then
                    subrow("strMwStKey") = strStrStCodeSage200
                    'Check ob korrekt berechnet
                    strSteuer = Split(objFiBhg.GetSteuerfeld(subrow("lngKto").ToString, "Zum Rechnen", subrow("dblBrutto").ToString, strStrStCodeSage200), "{<}")
                    If Val(strSteuer(2)) <> IIf(IsDBNull(subrow("dblMwst")), 0, subrow("dblMwst")) Then
                        'Im Fall von Auto-Korrekt anpassen
                        'Stop
                        'If booAutoCorrect Then
                        strStatusText += "MwSt " + subrow("dblMwst").ToString
                        subrow("dblMwst") = Val(strSteuer(2))
                        subrow("dblBrutto") = Decimal.Round(subrow("dblNetto") + subrow("dblMwSt"), 2, MidpointRounding.AwayFromZero)
                        'subrow("dblNetto") = Decimal.Round(subrow("dblBrutto") + subrow("dblMwSt"), 2, MidpointRounding.AwayFromZero)
                        strStatusText += " -> " + subrow("dblMwst").ToString + ", "
                        'Else
                        'intReturnValue = 1
                        'End If
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
            If subrow("intSollHaben") = 1 Then
                dblSubNetto += IIf(IsDBNull(subrow("dblNetto")), 0, subrow("dblNetto")) * -1
                dblSubMwSt += IIf(IsDBNull(subrow("dblMwSt")), 0, subrow("dblMwSt")) * -1
                dblSubBrutto += IIf(IsDBNull(subrow("dblBrutto")), 0, subrow("dblBrutto")) * -1
            Else
                dblSubNetto += subrow("dblNetto")
                dblSubMwSt += subrow("dblMwSt")
                dblSubBrutto += subrow("dblBrutto")
            End If

            'Runden
            dblSubNetto = Decimal.Round(dblSubNetto, 2, MidpointRounding.AwayFromZero)
            dblSubMwSt = Decimal.Round(dblSubMwSt, 2, MidpointRounding.AwayFromZero)
            dblSubBrutto = Decimal.Round(dblSubBrutto, 2, MidpointRounding.AwayFromZero)

            'Konto prüfen
            If IIf(IsDBNull(subrow("lngKto")), 0, subrow("lngKto")) Then
                intReturnValue = FcCheckKonto(subrow("lngKto"), objFiBhg, IIf(IsDBNull(subrow("dblMwSt")), 0, subrow("dblMwSt")), IIf(IsDBNull(subrow("lngKST")), 0, subrow("lngKST")))
                If intReturnValue = 0 Then
                    subrow("strKtoBez") = MainDebitor.FcReadDebitorKName(objFiBhg, subrow("lngKto"))
                ElseIf intReturnValue = 2 Then
                    subrow("strKtoBez") = MainDebitor.FcReadDebitorKName(objFiBhg, subrow("lngKto")) + " MwSt!"
                ElseIf intReturnValue = 3 Then
                    subrow("strKtoBez") = MainDebitor.FcReadDebitorKName(objFiBhg, subrow("lngKto")) + " NoKST"
                ElseIf intReturnValue = 4 Then
                    subrow("strKtoBez") = MainDebitor.FcReadDebitorKName(objFiBhg, subrow("lngKto")) + " K<3KST"
                    subrow("lngKST") = 0
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
                intReturnValue = FcCheckKstKtr(IIf(IsDBNull(subrow("lngKST")), 0, subrow("lngKST")), objFiBhg, objFiPI, subrow("lngKto"), strKstKtrSage200)
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
                intReturnValue = 1

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
            strStatusText = ""

            strStatusOverAll = strStatusOverAll Or strBitLog

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
                                              ByVal booAutoCorrect As Boolean) As Int16

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
        Dim strStrStCodeSage200 As String = ""
        Dim strKstKtrSage200 As String = ""
        Dim selsubrow() As DataRow
        Dim strStatusOverAll As String = "0000000"
        Dim strSteuer() As String

        'Summen bilden und Angaben prüfen
        intSubNumber = 0
        dblSubNetto = 0
        dblSubMwSt = 0
        dblSubBrutto = 0

        selsubrow = objDtKrediSub.Select("lngKredID=" + lngKredID.ToString)

        For Each subrow As DataRow In selsubrow

            strBitLog = ""
            'Runden
            subrow("dblNetto") = IIf(IsDBNull(subrow("dblNetto")), 0, Decimal.Round(subrow("dblNetto"), 2, MidpointRounding.AwayFromZero))
            subrow("dblMwSt") = IIf(IsDBNull(subrow("dblMwst")), 0, Decimal.Round(subrow("dblMwst"), 2, MidpointRounding.AwayFromZero))
            subrow("dblBrutto") = IIf(IsDBNull(subrow("dblBrutto")), 0, Decimal.Round(subrow("dblBrutto"), 2, MidpointRounding.AwayFromZero))
            subrow("dblMwStSatz") = IIf(IsDBNull(subrow("dblMwStSatz")), 0, Decimal.Round(subrow("dblMwStSatz"), 1, MidpointRounding.AwayFromZero))

            'Zuerst evtl. falsch gesetzte KTR oder Steuer - Sätze prüfen
            If subrow("lngKto") < 3000 Then
                subrow("strMwStKey") = Nothing
                subrow("lngKST") = 0
            End If

            'MwSt prüfen
            If Not IsDBNull(subrow("strMwStKey")) Then
                intReturnValue = FcCheckMwStToCorrect(objdbconn, subrow("strMwStKey"), subrow("dblMwStSatz"), subrow("dblMwSt"))
                intReturnValue = FcCheckMwSt(objdbconn, objFiBhg, subrow("strMwStKey"), subrow("dblMwStSatz"), strStrStCodeSage200, subrow("lngKto"))
                If intReturnValue = 0 Then
                    subrow("strMwStKey") = strStrStCodeSage200
                    'Check of korrekt berechnet
                    strSteuer = Split(objFiBhg.GetSteuerfeld(subrow("lngKto").ToString, "Zum Rechnen", subrow("dblBrutto").ToString, strStrStCodeSage200), "{<}")
                    If Val(strSteuer(2)) <> subrow("dblMwst") Then
                        'Im Fall von Auto-Korrekt anpassen
                        'Stop
                        If booAutoCorrect Then
                            subrow("dblMwst") = Val(strSteuer(2))
                            subrow("dblBrutto") = subrow("dblNetto") + subrow("dblMwSt")
                        Else
                            intReturnValue = 1
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
                dblSubNetto += IIf(IsDBNull(subrow("dblNetto")), 0, subrow("dblNetto")) * -1
                dblSubMwSt += IIf(IsDBNull(subrow("dblMwSt")), 0, subrow("dblMwSt")) * -1
                dblSubBrutto += IIf(IsDBNull(subrow("dblBrutto")), 0, subrow("dblBrutto")) * -1
            Else
                dblSubNetto += subrow("dblNetto")
                dblSubMwSt += subrow("dblMwSt")
                dblSubBrutto += subrow("dblBrutto")
            End If

            'Konto prüfen
            If IIf(IsDBNull(subrow("lngKto")), 0, subrow("lngKTo")) > 0 Then
                intReturnValue = FcCheckKonto(subrow("lngKto"), objFiBhg, IIf(IsDBNull(subrow("dblMwSt")), 0, subrow("dblMwSt")), IIf(IsDBNull(subrow("lngKST")), 0, subrow("lngKST")))
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
            strStatusText = ""
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
                strStatusText = "ok"
            End If

            'BitLog und Text schreiben
            subrow("strStatusUBBitLog") = strBitLog
            subrow("strStatusUBText") = strStatusText

            strStatusOverAll = strStatusOverAll Or strBitLog

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

    End Function


    Public Shared Function FcCheckMwStToCorrect(ByRef objdbconn As MySqlConnection,
                                                ByVal strStrCode As String,
                                                ByRef dblStrWert As Double,
                                                ByVal dblStrAmount As Double) As Integer

        Dim objlocdtMwSt As New DataTable("tbllocMwSt")
        Dim objlocMySQLcmd As New MySqlCommand
        Dim strSteuerRec As String = ""

        Try

            'Sind die Angaben stimmig?
            If Len(strStrCode) > 0 And dblStrAmount <> 0 And dblStrWert = 0 Then 'MwSt Wert ist 0 obwohl Schlüssel und MwSt-Betrag

                objlocMySQLcmd.CommandText = "SELECT  * FROM t_sage_sage50mwst WHERE strKey='" + strStrCode + "'"

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

    Public Shared Function FcCheckKstKtr(ByVal lngKST As Long, objFiBhg As SBSXASLib.AXiFBhg, ByRef objFiPI As SBSXASLib.AXiPlFin, ByVal lngKonto As Long, ByRef strKstKtrSage200 As String) As Int16

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

    Public Shared Function FcGetPKNewFromRep(ByRef objdbconnZHDB02 As MySqlConnection, ByRef objsqlcommandZHDB02 As MySqlCommand, ByVal intPKRefField As Int32) As Int32

        'Aus Tabelle Rep_Betriebe auf ZHDB02 auslesen 
        Dim objdtRepBetrieb As New DataTable

        Try

            objdbconnZHDB02.Open()
            objsqlcommandZHDB02.Connection = objdbconnZHDB02
            objsqlcommandZHDB02.CommandText = "SELECT PKNr From tab_repbetriebe WHERE Rep_Nr=" + intPKRefField.ToString
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

        booFoundCurrency = False
        strReturn = ""

        Call objfiBuha.ReadWhg()

        'If strCurrency = "EUR" Then Stop

        strReturn = objfiBuha.GetWhgZeile()
        Do While strReturn <> "EOF"
            If Left(strReturn, 3) = strCurrency Then
                'If strCurrency = "EUR" Then Stop
                booFoundCurrency = True
            End If
            strReturn = objfiBuha.GetWhgZeile()
        Loop

        If booFoundCurrency Then
            Return 0
        Else
            Return 1
        End If


    End Function

    Public Shared Function FcCheckKonto(ByVal lngKtoNbr As Long, ByRef objfiBuha As SBSXASLib.AXiFBhg, ByVal dblMwSt As Double, ByVal lngKST As Int32) As Integer

        'Returns 0=ok, 1=existiert nicht, 2=existiert aber keine KST erlaubt, 3=KST nicht auf Konto definiert, 4=KST auf Konto > 3

        Dim strReturn As String
        Dim strKontoInfo() As String

        strReturn = objfiBuha.GetKontoInfo(lngKtoNbr.ToString)
        If strReturn = "EOF" Then
            Return 1
        Else
            'If dblMwSt = 0 Then
            'Return 0
            'KST?
            If lngKST > 0 Then
                If CInt(Left(lngKtoNbr.ToString, 1)) >= 3 Then
                    strKontoInfo = Split(objfiBuha.GetKontoInfo(lngKtoNbr.ToString), "{>}")
                    If strKontoInfo(22) = "" Then
                        Return 3
                    Else
                        Return 0

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
            End If
        End If

    End Function


    Public Shared Function InsertDataTableColumnName(ByRef dtSouce As DataTable, ByRef dtResult As DataTable) As Boolean
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
        End Try

    End Function


    'Public Shared Function FcSetBuchMode(ByRef objdbBuha As SBSXASLib.AXiDbBhg, ByVal strMode As String) As Int16

    '    objdbBuha.SetBuchMode(strMode)

    '    Return 0

    'End Function

    'Public Shared Function FcSetBelegKopf4(ByRef objdbBuha As SBSXASLib.AXiDbBhg,
    '                                       ByVal lngBelegNr As Long,
    '                                       ByVal strValutaDatum As String,
    '                                       ByVal lngDebitor As Long,
    '                                       ByVal strBelegTyp As String,
    '                                       ByVal strBelegDatum As String,
    '                                       ByVal strVerFallDatum As String,
    '                                       ByVal strBelegText As String,
    '                                       ByVal strReferenz As String,
    '                                       ByVal lngKondition As Long,
    '                                       ByVal strSachbearbeiter As String,
    '                                       ByVal strVerkaeufer As String,
    '                                       ByVal strMahnSperre As String,
    '                                       ByVal shrMahnstufe As Short,
    '                                       ByVal strBetraBrutto As String,
    '                                       ByVal strKurs As String,
    '                                       ByVal strBelegExt As String,
    '                                       ByVal strSKonto As String,
    '                                       ByVal strDebiCur As String,
    '                                       ByVal strSammelKonto As String,
    '                                       ByVal strVerzugsZ As String,
    '                                       ByVal strZusatzText As String,
    '                                       ByVal strEBankKonto As String,
    '                                       ByVal strIkoDebitor As String) As Integer


    '    'Zuerst prüfen ob Zwingende Werte angegeben worden sind

    '    'Ausführung
    '    objdbBuha.SetBelegKopf4(lngBelegNr, strValutaDatum, lngDebitor, strBelegTyp, strBelegDatum, strVerFallDatum, strBelegText, strReferenz, lngKondition, strSachbearbeiter, strVerkaeufer, strMahnSperre, shrMahnstufe, strBetraBrutto,
    '                            strKurs, strBelegExt, strSKonto, strDebiCur, strSammelKonto, strVerzugsZ, strZusatzText, strEBankKonto, strIkoDebitor)

    '    Return 0

    'End Function

    'Public Shared Function FcSetVerteilung(ByRef objdbBuha As SBSXASLib.AXiDbBhg,
    '                                       ByVal strGegenKonto As String,
    '                                       ByVal strFibuText As String,
    '                                       ByVal strNettoBetrag As String,
    '                                       ByVal strArraySteuer As String,
    '                                       ByVal strArrayKST As String,
    '                                       ByVal strArrayKSTE As String) As Integer

    '    'Prüfen ob Daten vollständig

    '    'Ausführung
    '    objdbBuha.SetVerteilung(strGegenKonto, strFibuText, strNettoBetrag, strArraySteuer, strArrayKST, strArrayKSTE)

    '    Return 0

    'End Function

    'Public Shared Function FcWriteBuchung(ByRef objdbBuha As SBSXASLib.AXiDbBhg) As Integer

    '    'Ausführung
    '    objdbBuha.WriteBuchung()

    '    Return 0

    'End Function

    Public Shared Function FcGetSteuerFeld(ByRef objFBhg As SBSXASLib.AXiFBhg, ByVal lngKto As Long, ByVal strDebiSubText As String, ByVal dblBrutto As Double, ByVal strMwStKey As String, ByVal dblMwSt As Double) As String

        Dim strSteuerFeld As String = ""

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

    Public Shared Function FcGetKurs(ByVal strCurrency As String, ByVal strDateValuta As String, ByRef objFBhg As SBSXASLib.AXiFBhg, ByVal Optional intKonto As Integer = 0) As Double

        'Konzept: Falls ein Konto mitgegeben wird, wird überprüft ob auf dem Konto die mitgegebene Währung Leitwärhung ist. Falls ja wird der Kurs 1 zurück gegeben

        Dim strKursZeile As String = ""
        Dim strKursZeileAr() As String
        Dim strKontoInfo() As String

        objFBhg.ReadKurse(strCurrency, "", "J")

        Do While strKursZeile <> "EOF"
            strKursZeile = objFBhg.GetKursZeile()
            If strKursZeile <> "EOF" Then
                strKursZeileAr = Split(strKursZeile, "{>}")
                If strKursZeileAr(0) = strCurrency Then
                    If strKursZeileAr(0) = "EUR" Then Stop
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
                                        ByVal strcmbBuha As String) As Integer

        'DebiBitLog 1=PK, 2=Konto, 3=Währung, 4=interne Bank, 5=OP Kopf, 6=RG-Datum, 7=Valuta Datum, 8=Subs, 9=OP doppelt
        Dim strBitLog As String = ""
        Dim intReturnValue As Integer
        Dim strStatus As String = ""
        Dim intSubNumber As Int16
        Dim dblSubNetto As Double
        Dim dblSubMwSt As Double
        Dim dblSubBrutto As Double
        Dim booAutoCorrect As Boolean
        Dim selsubrow() As DataRow
        Dim strKrediReferenz As String
        Dim booDiffHeadText As Boolean
        Dim strKrediiHeadText As String
        Dim booDiffSubText As Boolean
        Dim strKrediSubText As String
        Dim intKreditorNew As Int32
        Dim strCleanOPNbr As String

        Try

            objdbconn.Open()
            objOrdbconn.Open()

            For Each row As DataRow In objdtKredits.Rows

                '
                'If row("lngKredID") = "1103800" Then Stop
                'Runden
                row("dblKredNetto") = Decimal.Round(row("dblKredNetto"), 2, MidpointRounding.AwayFromZero)
                row("dblKredMwSt") = Decimal.Round(row("dblKredMwst"), 2, MidpointRounding.AwayFromZero)
                row("dblKredBrutto") = Decimal.Round(row("dblKredBrutto"), 2, MidpointRounding.AwayFromZero)
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

                strBitLog += Trim(intReturnValue.ToString)
                If intKreditorNew <> 0 Then
                    intReturnValue = MainKreditor.FcCheckKreditor(intKreditorNew, row("intBuchungsart"), objKrBuha)
                    'intReturnValue = FcCheckKreditBank(objKrBuha, intKreditorNew, row("intPayType"), row("strKredRef"), row("strKrediBank"), objdbconnZHDB02)
                    'intReturnValue = 3
                Else
                    intReturnValue = 2
                End If
                strBitLog = Trim(intReturnValue.ToString)

                'Kto 02
                intReturnValue = FcCheckKonto(row("lngKredKtoNbr"), objfiBuha, row("dblKredMwSt"), 0)
                strBitLog += Trim(intReturnValue.ToString)

                'Currency 03
                intReturnValue = FcCheckCurrency(row("strKredCur"), objfiBuha)
                strBitLog += Trim(intReturnValue.ToString)

                'Sub 04
                booAutoCorrect = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettings(objdbconn, "Buchh_HeadAutoCorrect", intAccounting)))
                'booAutoCorrect = False
                intReturnValue = FcCheckKrediSubBookings(row("lngKredID"), objdtKreditSubs, intSubNumber, dblSubBrutto, dblSubNetto, dblSubMwSt, objdbconn, objfiBuha, objdbPIFb, row("intBuchungsart"), booAutoCorrect)
                strBitLog += Trim(intReturnValue.ToString)

                'Autokorrektur 05
                'booAutoCorrect = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettings(objdbconn, "Buchh_HeadAutoCorrect", intAccounting)))
                'booAutoCorrect = False
                If booAutoCorrect Then
                    'Git es etwas zu korrigieren?
                    If IIf(IsDBNull(row("dblKredBrutto")), 0, row("dblKredBrutto")) <> dblSubBrutto Or
                        IIf(IsDBNull(row("dblKredNetto")), 0, row("dblKredNetto")) <> dblSubNetto Or
                        IIf(IsDBNull(row("dblKredMwSt")), 0, row("dblKredMwSt")) <> dblSubMwSt Then
                        row("dblKredBrutto") = Math.Round(dblSubBrutto * -1, 2, MidpointRounding.AwayFromZero)
                        row("dblKredNetto") = dblSubNetto * -1
                        row("dblKredMwSt") = dblSubMwSt * -1
                        ''In Sub korrigieren
                        'selsubrow = objdtDebitSubs.Select("strRGNr='" + row("strDebRGNbr") + "' AND intSollHaben=2")
                        'If selsubrow.Length = 1 Then
                        '    selsubrow(0).Item("dblBrutto") = dblSubBrutto * -1
                        '    selsubrow(0).Item("dblMwSt") = dblSubMwSt * -1
                        '    selsubrow(0).Item("dblNetto") = dblSubNetto * -1
                        'End If
                        strBitLog += "1"
                    Else
                        strBitLog += "0"
                    End If
                Else
                    strBitLog += "0"
                End If

                'Diff Kopf - Sub? 06
                If row("intBuchungsart") = 1 Then 'OP
                    If IIf(IsDBNull(row("dblKredBrutto")), 0, row("dblKredBrutto")) + dblSubBrutto <> 0 _
                        Or IIf(IsDBNull(row("dblKredMwSt")), 0, row("dblKredMwSt")) + dblSubMwSt <> 0 _
                        Or IIf(IsDBNull(row("dblKredNetto")), 0, row("dblKredNetto")) + dblSubNetto <> 0 Then
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
                                                  0)
                strBitLog += Trim(intReturnValue.ToString)
                'OP - Nummer prüfen 08
                'intReturnValue = FcCreateDebRef(objdbconn, intAccounting, row("strDebiBank"), row("strDebRGNbr"), row("intBuchungsart"), strDebiReferenz)
                strCleanOPNbr = IIf(IsDBNull(row("strOPNr")), "", row("strOPNr"))
                intReturnValue = MainKreditor.FcChCeckKredOP(strCleanOPNbr, IIf(IsDBNull(row("strKredRGNbr")), "", row("strKredRGNbr")))
                row("strOPNr") = strCleanOPNbr
                strBitLog += Trim(intReturnValue.ToString)
                'OP - Verdopplung 09
                intReturnValue = MainKreditor.FcCheckKrediOPDouble(objKrBuha, IIf(IsDBNull(row("lngKredNbr")), 0, row("lngKredNbr")), row("strKredRGNbr"))
                strBitLog += Trim(intReturnValue.ToString)
                'Valuta - Datum 10
                intReturnValue = FcChCeckDate(row("datKredValDatum"), objdtInfo)
                strBitLog += Trim(intReturnValue.ToString)
                'RG - Datum 11
                intReturnValue = FcChCeckDate(row("datKredRGDatum"), objdtInfo)
                strBitLog += Trim(intReturnValue.ToString)
                ''Referenz 12
                'intReturnValue = IIf(IsDBNull(row("strKredRef")), 1, 0)
                'strBitLog += Trim(intReturnValue.ToString)

                'Status-String auswerten
                'Kreditor
                If Left(strBitLog, 1) <> "0" Then
                    strStatus = "Kred"
                    If Left(strBitLog, 1) <> "2" Then
                        intReturnValue = MainKreditor.FcIsKreditorCreatable(objdbconn,
                                                                            objdbconnZHDB02,
                                                                            objsqlcommandZHDB02,
                                                                            intKreditorNew,
                                                                            objKrBuha,
                                                                            strcmbBuha,
                                                                            IIf(IsDBNull(row("intPayType")), 3, row("intPayType")),
                                                                            row("strKredRef"))
                        If intReturnValue = 0 Then
                            strStatus += " erstellt"
                            row("strKredBez") = MainKreditor.FcReadKreditorName(objKrBuha, intKreditorNew, row("strKredCur"))

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
                    intReturnValue = MainKreditor.FcCheckKreditBank(objdbconn,
                                                       objdbconnZHDB02,
                                                       objKrBuha,
                                                       intKreditorNew,
                                                       IIf(IsDBNull(row("intPayType")), 3, row("intPayType")),
                                                       IIf(IsDBNull(row("strKredRef")), "", row("strKredRef")),
                                                       IIf(IsDBNull(row("strKredRef")), "", row("strKredRef")),
                                                       row("strKredCur"))
                End If
                'Konto
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
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "AutoC"
                End If
                'Diff zu Subbuchungen
                If Mid(strBitLog, 6, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "DiffS"
                End If
                'OP Kopf
                If Mid(strBitLog, 7, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "BelK"
                End If
                'OP Nummer
                If Mid(strBitLog, 8, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "OPNbr"
                End If
                'OP Doppelt
                If Mid(strBitLog, 9, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "OPDbl"
                    'Else
                    '   row("strDebRef") = strDebiReferenz
                End If
                'Valuta Datum 
                If Mid(strBitLog, 10, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "ValD"
                    'Else
                    '    row("strDebRef") = strDebiReferenz
                End If
                'RG Datum 
                If Mid(strBitLog, 11, 1) <> "0" Then
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "RgD"
                    'Else
                    '    row("strDebRef") = strDebiReferenz
                End If
                'Referenz
                'If Left(strBitLog, 12, 1) <> "0" Then
                '    strStatus = "KRef "
                '    If intKreditorNew > 0 Then 'Neue Kreditoren-Nr bekannt Versuch IBAN von Standard zu lesen


                '    End If
                'End If

                'Status schreiben
                If Val(strBitLog) = 0 Or Val(strBitLog) = 1000000 Then
                    row("booKredBook") = True
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "ok"
                End If
                row("strKredStatusText") = strStatus
                row("strKredStatusBitLog") = strBitLog

                ''Wird ein anderer Text in der Head-Buchung gewünscht?
                'booDiffHeadText = IIf(FcReadFromSettings(objdbconn, "Buchh_TextSpecial", intAccounting) = "0", False, True)
                'If booDiffHeadText Then
                '    strDebiHeadText = FcSQLParse(FcReadFromSettings(objdbconn, "Buchh_TextSpecialText", intAccounting), row("strDebRGNbr"), objdtDebits)
                '    row("strDebText") = strDebiHeadText
                'End If

                ''Wird ein anderer Text in den Sub-Buchung gewünscht?
                'booDiffSubText = IIf(FcReadFromSettings(objdbconn, "Buchh_SubTextSpecial", intAccounting) = "0", False, True)
                'If booDiffSubText Then
                '    strDebiSubText = FcSQLParse(FcReadFromSettings(objdbconn, "Buchh_SubTextSpecialText", intAccounting), row("strDebRGNbr"), objdtDebits)
                'Else
                '    strDebiSubText = row("strDebText")
                'End If
                'selsubrow = objdtDebitSubs.Select("strRGNr='" + row("strDebRGNbr") + "'")
                        'For Each subrow In selsubrow
                        '    subrow("strDebSubText") = strDebiSubText
                        'Next

                        'Init
                        strBitLog = ""
                strStatus = ""
                intSubNumber = 0
                dblSubBrutto = 0
                dblSubNetto = 0
                dblSubMwSt = 0
                intKreditorNew = 0

            Next

        Catch ex As Exception
            MessageBox.Show(ex.Message)

        Finally
            If objOrdbconn.State = ConnectionState.Open Then
                objOrdbconn.Close()
            End If
            If objdbconn.State = ConnectionState.Open Then
                objdbconn.Close()
            End If

        End Try


    End Function

    Public Shared Function fcCheckTransitorischeDebit(ByVal intAccounting As Int16, ByRef objdbconn As MySqlConnection,
                                       ByRef objdbAccessConn As OleDb.OleDbConnection)

        Dim strSQL As String
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
                strSQL = "SELECT * FROM t_sage_buchhaltungen_sub WHERE strType='D' AND refMandant=" + intAccounting.ToString
                objRGMySQLConn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString")
                objlocMySQLcmd.Connection = objRGMySQLConn
                objlocMySQLcmd.CommandText = strSQL
                objRGMySQLConn.Open()
                objDTTransitDebits.Load(objlocMySQLcmd.ExecuteReader)
                objRGMySQLConn.Close()

                For Each rowdebitquery As DataRow In objDTTransitDebits.Rows

                    If Not IsDBNull(rowdebitquery("strCondition")) Then
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

    Public Shared Function FcInitAccessConnecation(ByRef objaccesscon As OleDb.OleDbConnection, ByVal strMDBName As String) As Int16

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

    Public Shared Function FcNextPKNr(ByRef objdbconnZHDB02 As MySqlConnection, ByVal intRepNr As Int32, ByRef intNewPKNr As Int32) As Int16

        '0=ok, 1=Rep - Nr. existiert nicht, 2=Bereich voll, 3=keine Bereichdefinition 9=Problem

        'PK - Nummer soll der Funktion gegeben werden, Funktion sucht sich dann die PK_Gruppe 
        'Konzept: Tabelle füllen und dann durchsteppen
        Dim objsqlcommand As New MySqlCommand
        Dim objdtPKNr As New DataTable
        Dim intPKNrGuppenID As Int16
        Dim intRangeStart, intRangeEnd, i, intRecordCounter As Int32
        Dim objdsPKNbrs As New DataSet
        Dim objDAPKNbrs As New MySqlDataAdapter


        Try

            objdbconnZHDB02.Open()
            objsqlcommand.Connection = objdbconnZHDB02
            objsqlcommand.CommandText = "SELECT PKNrGruppeID FROM tab_repbetriebe WHERE Rep_Nr=" + intRepNr.ToString
            objdtPKNr.Load(objsqlcommand.ExecuteReader)

            If objdtPKNr.Rows.Count > 0 Then 'Rep_Betrieb gefunden
                intPKNrGuppenID = objdtPKNr.Rows(0).Item("PKNrGruppeID")
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
            MessageBox.Show(ex.Message)
            Return 9

        Finally
            objdtIBAN.Dispose()
            objmysqlcom.Dispose()
            objXMLDoc = Nothing
            objResponse = Nothing
            objXMLNodeList = Nothing

        End Try

    End Function

    Public Shared Function FcWriteNewDebToRepbetrieb(ByRef objdbconnZHDB02 As MySqlConnection, ByVal intRepNr As Int32, intNewDebNr As Int32) As Int16

        '0=Update ok, 1=Update hat nicht geklappt, 9=Error

        Dim strSQL As String
        Dim objmysqlcmd As New MySqlCommand
        Dim intAffected As Int16

        Try

            strSQL = "UPDATE tab_repbetriebe SET PKNr=" + intNewDebNr.ToString + " WHERE Rep_Nr=" + intRepNr.ToString
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
                objdbconnZHDB02.Close()
            End If

        End Try

    End Function
End Class
