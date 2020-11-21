Option Strict Off
Option Explicit On

Imports MySql.Data.MySqlClient
Imports System.Data.OracleClient
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
        DT.PrimaryKey = New DataColumn() {DT.Columns("strDebRGNbr")}
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
        strDebIdentNbr2.MaxLength = 50
        DT.Columns.Add(strDebIdentNbr2)
        Dim strDebText As DataColumn = New DataColumn("strDebText")
        strDebText.DataType = System.Type.[GetType]("System.String")
        strDebText.MaxLength = 50
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
        strArtikel.MaxLength = 128
        DT.Columns.Add(strArtikel)
        Dim strDebSubText As DataColumn = New DataColumn("strDebSubText")
        strDebSubText.DataType = System.Type.[GetType]("System.String")
        strDebSubText.MaxLength = 50
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
        strOPNr.MaxLength = 13
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
        strKredKtoBez.MaxLength = 50
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
                                       ByRef objdtInfo As DataTable) As Int16


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


        'On Error GoTo ErrorHandler

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
        objFinanz.OpenMandant(strMandant, "")
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
        MsgBox("OpenMandant:" & Chr(13) & Chr(10) & "Error" & Chr(13) & Chr(10) & "Die Button auf dem Main wurden ausgeschaltet !!!" & Chr(13) & Chr(10) & "Error # " & Str(Err.Number) & " was generated by " & Err.Source & Chr(13) & Chr(10) & Err.Description & " Unsere Fehlernummer" & Str(b))
        Err.Clear()

    End Function

    Public Shared Function FcFillDebit(ByVal intAccounting As Integer,
                                       ByRef objdtHead As DataTable,
                                       ByRef objdtSub As DataTable,
                                       ByRef objdbconn As MySqlConnection,
                                       ByRef objdbAccessConn As OleDb.OleDbConnection) As Integer

        Dim strSQL As String
        Dim strSQLSub As String
        Dim strRGTableType As String
        Dim objRGMySQLConn As New MySqlConnection
        Dim objlocMySQLcmd As New MySqlCommand
        Dim objlocOLEdbcmd As New OleDb.OleDbCommand

        Dim objDTDebiHead As New DataTable
        Dim dbProvider, dbSource, dbPathAndFile, strMDBName As String
        Dim objdrSub As DataRow
        Dim intFcReturns As Int16

        objdbconn.Open()

        strMDBName = FcReadFromSettings(objdbconn, "Buchh_RGTableMDB", intAccounting)
        dbProvider = "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
        dbSource = "Data Source="
        dbPathAndFile = "\\sdlc.mssag.ch\Apps\Backends\" + strMDBName + ";Jet OLEDB:System Database=\\sdlc.mssag.ch\Apps\Backends\Workbench.mdw;User ID=HagerR;"

        'Head Debitoren löschen
        objdtHead.Clear()
        strSQL = FcReadFromSettings(objdbconn, "Buchh_SQLHead", intAccounting)
        strRGTableType = FcReadFromSettings(objdbconn, "Buchh_RGTableType", intAccounting)

        Try

            'objlocMySQLcmd.CommandText = strSQL
            If strRGTableType = "A" Then
                'Access
                objdbAccessConn.ConnectionString = dbProvider + dbSource + dbPathAndFile
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
                strSQLSub = FcSQLParse(FcReadFromSettings(objdbconn, "Buchh_SQLDetail", intAccounting), row("strDebRGNbr"), objdtHead)
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
            MessageBox.Show(ex.Message)

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

    Public Shared Function FcRoundInTable(ByRef objdt As DataTable, ByVal strColumnName As String, ByVal intDecimals As Int16) As Int16

        Try

            For Each row As DataRow In objdt.Rows

                row.Item(strColumnName) = Math.Round(row.Item(strColumnName), 2, MidpointRounding.AwayFromZero)

            Next
            Return 0

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return 1

        End Try

    End Function

    Public Shared Function FcSQLParse(ByVal strSQLToParse As String, ByVal strRGNbr As String, ByVal objdtDebi As DataTable) As String

        'Funktion setzt in eingelesenem SQL wieder Variablen ein
        Dim intPipePositionBegin, intPipePositionEnd As Integer
        Dim strWork, strField As String
        Dim RowDebi() As DataRow

        'Zuerst Datensatz in Debi-Head suchen
        RowDebi = objdtDebi.Select("strDebRGNbr='" + strRGNbr + "'")

        '| suchen
        If InStr(strSQLToParse, "|") > 0 Then
            'Vorkommen gefunden
            intPipePositionBegin = InStr(strSQLToParse, "|")
            intPipePositionEnd = InStr(intPipePositionBegin + 1, strSQLToParse, "|")
            Do Until intPipePositionBegin = 0
                strField = Mid(strSQLToParse, intPipePositionBegin + 1, intPipePositionEnd - intPipePositionBegin - 1)
                Select Case strField
                    Case "rsDebi.Fields(""RGNr"")"
                        strField = RowDebi(0).Item("strDebRGNbr")
                    Case "rsDebiTemp.Fields([strDebPKBez])"
                        strField = RowDebi(0).Item("strDebBez")
                    Case "rsDebiTemp.Fields([lngDebIdentNbr])"
                        strField = RowDebi(0).Item("lngDebIdentNbr")
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
            MessageBox.Show(ex.Message)
            Return 9

        Finally
            objSQLConnection.Close()

        End Try

    End Function

    Public Shared Function FcReadBankSettings(ByVal intAccounting As Int16, ByVal strBank As String, ByRef objdbconn As MySqlConnection) As String

        Dim objlocdtBank As New DataTable("tbllocBank")
        Dim objlocMySQLcmd As New MySqlCommand

        Try
            objlocMySQLcmd.CommandText = "SELECT strBLZ FROM tblAccountingBank WHERE intAccountingID=" + intAccounting.ToString + " AND strBank='" + strBank + "'"
            objlocMySQLcmd.Connection = objdbconn
            objlocdtBank.Load(objlocMySQLcmd.ExecuteReader)


        Catch ex As Exception
            MessageBox.Show(ex.Message)

        End Try

        Return objlocdtBank.Rows(0).Item(0).ToString

    End Function


    Public Shared Function FcReadFromSettings(ByRef objdbconn As MySqlConnection, ByVal strField As String, ByVal intMandant As Int16) As String

        Dim objlocdtSetting As New DataTable("tbllocSettings")
        Dim objlocMySQLcmd As New MySqlCommand

        Try

            objlocMySQLcmd.CommandText = "SELECT buchhaltungen." + strField + " FROM buchhaltungen WHERE Buchh_Nr=" + intMandant.ToString
            'Debug.Print(objlocMySQLcmd.CommandText)
            objlocMySQLcmd.Connection = objdbconn
            objlocdtSetting.Load(objlocMySQLcmd.ExecuteReader)
            'Debug.Print("Records" + objlocdtSetting.Rows.Count.ToString)


        Catch ex As Exception
            MessageBox.Show(ex.Message)

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
                                        ByRef objdtInfo As DataTable) As Integer

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
        Dim strDebiReferenz As String
        Dim booDiffHeadText As Boolean
        Dim strDebiHeadText As String
        Dim booDiffSubText As Boolean
        Dim strDebiSubText As String
        Dim intDebitorNew As Int32

        Try

            objdbconn.Open()
            objOrdbconn.Open()

            For Each row As DataRow In objdtDebits.Rows

                '
                If row("strDebRGNbr") = "57976" Then Stop

                'Status-String erstellen
                'Debitor 01
                intReturnValue = FcGetRefDebiNr(objdbconn, objdbconnZHDB02, objsqlcommand, objsqlcommandZHDB02, objOrdbconn, objOrcommand, IIf(IsDBNull(row("lngDebNbr")), 0, row("lngDebNbr")), intAccounting, intDebitorNew)
                strBitLog += Trim(intReturnValue.ToString)
                If intDebitorNew <> 0 Then
                    intReturnValue = FcCheckDebitor(intDebitorNew, row("intBuchungsart"), objdbBuha)
                Else
                    intReturnValue = 2
                End If
                strBitLog = Trim(intReturnValue.ToString)

                'Kto 02
                intReturnValue = FcCheckKonto(row("lngDebKtoNbr"), objfiBuha, row("dblDebMwSt"))
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
                If booAutoCorrect Then
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
                    strBitLog += "0"
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
                intReturnValue = FcCheckBelegHead(row("intBuchungsart"), IIf(IsDBNull(row("dblDebBrutto")), 0, row("dblDebBrutto")), IIf(IsDBNull(row("dblDebNetto")), 0, row("dblDebNetto")), IIf(IsDBNull(row("dblDebMwSt")), 0, row("dblDebMwSt")))
                strBitLog += Trim(intReturnValue.ToString)
                'Referenz 08
                intReturnValue = FcCreateDebRef(objdbconn, intAccounting, row("strDebiBank"), row("strDebRGNbr"), row("intBuchungsart"), strDebiReferenz)
                strBitLog += Trim(intReturnValue.ToString)
                'OP - Verdopplung 09
                intReturnValue = FcCheckOPDouble(objdbBuha, IIf(IsDBNull(row("lngDebNbr")), 0, row("lngDebNbr")), row("strDebRGNbr"))
                strBitLog += Trim(intReturnValue.ToString)
                'Valuta - Datum 10
                intReturnValue = FcChCeckDate(row("datDebValDatum"), objdtInfo)
                strBitLog += Trim(intReturnValue.ToString)
                'RG - Datum 11
                intReturnValue = FcChCeckDate(row("datDebRGDatum"), objdtInfo)
                strBitLog += Trim(intReturnValue.ToString)
                'intReturnValue = fcCheckIntBank()

                'Status-String auswerten
                'Debitor
                If Left(strBitLog, 1) <> "0" Then
                    strStatus = "Deb"
                    If Left(strBitLog, 1) <> "2" Then
                        intReturnValue = FcIsDebitorCreatable(objdbconnZHDB02, objsqlcommandZHDB02, intDebitorNew, objdbBuha)
                        If intReturnValue = 0 Then
                            strStatus += " erstellt"
                        Else
                            strStatus += " nicht erstellt."
                        End If
                        row("strDebBez") = FcReadDebitorName(objdbBuha, intDebitorNew, row("strDebCur"))
                        row("lngDebNbr") = intDebitorNew
                    Else
                        strStatus += " keine Ref"
                        row("strDebBez") = "n/a"
                    End If
                Else
                    row("strDebBez") = FcReadDebitorName(objdbBuha, intDebitorNew, row("strDebCur"))
                    row("lngDebNbr") = intDebitorNew
                End If
                'Konto
                If Mid(strBitLog, 2, 1) <> "0" Then
                    If Mid(strBitLog, 2, 1) <> 2 Then
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Kto"
                    Else
                        strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "Kto MwSt"
                    End If
                    row("strDebKtoBez") = "n/a"
                Else
                    row("strDebKtoBez") = FcReadDebitorKName(objfiBuha, row("lngDebKtoNbr"))
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

                'Status schreiben
                If Val(strBitLog) = 0 Or Val(strBitLog) = 1000000 Then
                    row("booDebBook") = True
                    strStatus = strStatus + IIf(strStatus <> "", ", ", "") + "ok"
                End If
                row("strDebStatusText") = strStatus
                row("strDebStatusBitLog") = strBitLog

                'Wird ein anderer Text in der Head-Buchung gewünscht?
                booDiffHeadText = IIf(FcReadFromSettings(objdbconn, "Buchh_TextSpecial", intAccounting) = "0", False, True)
                If booDiffHeadText Then
                    strDebiHeadText = FcSQLParse(FcReadFromSettings(objdbconn, "Buchh_TextSpecialText", intAccounting), row("strDebRGNbr"), objdtDebits)
                    row("strDebText") = strDebiHeadText
                End If

                'Wird ein anderer Text in den Sub-Buchung gewünscht?
                booDiffSubText = IIf(FcReadFromSettings(objdbconn, "Buchh_SubTextSpecial", intAccounting) = "0", False, True)
                If booDiffSubText Then
                    strDebiSubText = FcSQLParse(FcReadFromSettings(objdbconn, "Buchh_SubTextSpecialText", intAccounting), row("strDebRGNbr"), objdtDebits)
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
            MessageBox.Show(ex.Message)

        Finally
            objOrdbconn.Close()
            objdbconn.Close()

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
            MessageBox.Show(ex.Message)
            Return 9

        End Try


    End Function

    Public Shared Function FcCheckOPDouble(ByRef objdbBuha As SBSXASLib.AXiDbBhg, ByVal strDebitor As String, ByVal strOPNr As String) As Int16

        'Return 0=ok, 1=Beleg existiert, 9=Problem

        Dim intBelegReturn As Int16

        Try
            intBelegReturn = objdbBuha.doesBelegExist(strDebitor, "CHF", strOPNr, "", "", "")
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


    Public Shared Function FcCreateDebRef(ByRef objdbconn As MySqlConnection, ByVal intAccounting As Integer, ByVal strBank As String, ByVal strRGNr As String, ByVal intBuchungsArt As Integer, ByRef strReferenz As String) As Integer

        'Return 0=ok oder nicht nötig, 1=keine Angaben hinterlegt, 2=Berechnung hat nicht geklappt

        Dim strTLNNr As String
        Dim strCleanedRGNr As String

        Try

            If intBuchungsArt = 1 Then
                strTLNNr = FcReadBankSettings(intAccounting, strBank, objdbconn)
                strCleanedRGNr = FcCleanRGNrStrict(strRGNr)

                strReferenz = strTLNNr + StrDup(20 - Len(strCleanedRGNr), "0") + strCleanedRGNr + Trim(CStr(FcModulo10(strTLNNr + StrDup(20 - Len(strCleanedRGNr), "0") + strCleanedRGNr)))
                Return 0

            Else
                Return 0
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
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

    Public Shared Function FcCheckBelegHead(ByVal intBuchungsArt As Int16, ByVal dblBrutto As Double, ByVal dblNetto As Double, ByVal dblMwSt As Double) As Int16

        'Returns 0=ok oder nicht wichtig, 1=Brutto, 2=Netto, 3=Beide, 4=Diff

        If intBuchungsArt = 1 Then
            If dblBrutto = 0 And dblNetto = 0 Then
                Return 3
            ElseIf dblBrutto = 0 Then
                Return 1
            ElseIf dblNetto = 0 Then
                Return 2
            ElseIf Math.Round(dblBrutto - dblMwSt, 2, MidpointRounding.AwayFromZero) <> Math.Round(dblNetto, 2, MidpointRounding.AwayFromZero) Then
                Return 4
            Else
                Return 0
            End If
        Else
            Return 0
        End If

    End Function

    Public Shared Function FcCheckMwSt(ByRef objdbconn As MySqlConnection, ByRef objFiBhg As SBSXASLib.AXiFBhg, ByVal strStrCode As String, ByVal dblStrWert As Double, ByRef strStrCode200 As String) As Integer

        'returns 0=ok, 1=nicht gefunden

        Dim objlocdtMwSt As New DataTable("tbllocMwSt")
        Dim objlocMySQLcmd As New MySqlCommand
        Dim strSteuerRec As String = ""
        Dim strSteuerRecAr() As String
        Dim intLooper As Int16 = 0

        Try

            objlocMySQLcmd.CommandText = "SELECT  * FROM sage50mwst WHERE strKey='" + strStrCode + "' AND dblProzent=" + dblStrWert.ToString

            objlocMySQLcmd.Connection = objdbconn
            objlocdtMwSt.Load(objlocMySQLcmd.ExecuteReader)

            If objlocdtMwSt.Rows.Count = 0 Then
                MessageBox.Show("MwSt " + strStrCode + " ist nicht definiert für " + dblStrWert.ToString + ".")
                Return 1
            Else
                'In Sage 200 suchen
                Do Until strSteuerRec = "EOF"
                    strSteuerRec = objFiBhg.GetStIDListe(intLooper)
                    If strSteuerRec <> "EOF" Then
                        strSteuerRecAr = Split(strSteuerRec, "{>}")
                        'Gefunden?
                        If strSteuerRecAr(3) = dblStrWert And strSteuerRecAr(6) = objlocdtMwSt.Rows(0).Item("strBruttoNetto") And strSteuerRecAr(7) = objlocdtMwSt.Rows(0).Item("strGegenKonto") Then
                            'Debug.Print("Found " + strSteuerRecAr(0).ToString)
                            strStrCode200 = strSteuerRecAr(0)
                            Return 0
                        End If
                    Else
                        Return 1
                    End If
                    intLooper += 1
                Loop
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message)

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

            'MwSt prüfen
            If Not IsDBNull(subrow("strMwStKey")) Then
                intReturnValue = FcCheckMwSt(objdbconn, objFiBhg, subrow("strMwStKey"), IIf(IsDBNull(subrow("dblMwStSatz")), 0, subrow("dblMwStSatz")), strStrStCodeSage200)
                If intReturnValue = 0 Then
                    subrow("strMwStKey") = strStrStCodeSage200
                    'Check of korrekt berechnet
                    strSteuer = Split(objFiBhg.GetSteuerfeld(subrow("lngKto").ToString, "Zum Rechnen", subrow("dblBrutto").ToString, strStrStCodeSage200), "{<}")
                    If Val(strSteuer(2)) <> IIf(IsDBNull(subrow("dblMwst")), 0, subrow("dblMwst")) Then
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

            'Konto prüfen
            If Not IsDBNull(subrow("lngKto")) Then
                intReturnValue = FcCheckKonto(subrow("lngKto"), objFiBhg, IIf(IsDBNull(subrow("dblMwSt")), 0, subrow("dblMwSt")))
                If intReturnValue = 0 Then
                    subrow("strKtoBez") = FcReadDebitorKName(objFiBhg, subrow("lngKto"))
                ElseIf intReturnValue = 2 Then
                    subrow("strKtoBez") = FcReadDebitorKName(objFiBhg, subrow("lngKto")) + " MwSt!"
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

            'MwSt prüfen
            If Not IsDBNull(subrow("strMwStKey")) Then
                intReturnValue = FcCheckMwSt(objdbconn, objFiBhg, subrow("strMwStKey"), Decimal.Round(subrow("dblMwStSatz"), 1, MidpointRounding.AwayFromZero), strStrStCodeSage200)
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
            If Not IsDBNull(subrow("lngKto")) Then
                intReturnValue = FcCheckKonto(subrow("lngKto"), objFiBhg, IIf(IsDBNull(subrow("dblMwSt")), 0, subrow("dblMwSt")))
                If intReturnValue = 0 Then
                    subrow("strKtoBez") = FcReadDebitorKName(objFiBhg, subrow("lngKto"))
                ElseIf intReturnValue = 2 Then
                    subrow("strKtoBez") = FcReadDebitorKName(objFiBhg, subrow("lngKto")) + " MwSt!"
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


    Public Shared Function FcCheckKstKtr(ByVal lngKST As Long, objFiBhg As SBSXASLib.AXiFBhg, ByRef objFiPI As SBSXASLib.AXiPlFin, ByVal lngKonto As Long, ByRef strKstKtrSage200 As String) As Int16

        'return 0=ok, 1=Kst existiert kene Kostenart, 2=Kst nicht defniert

        Dim strReturn As String
        Dim strReturnAr() As String
        Dim booKstKAok As Boolean
        Dim strKst, strKA As String

        booKstKAok = False
        objFiPI = Nothing
        objFiPI = objFiBhg.GetCheckObj

        Try
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

        Catch ex As Exception
            Return 1

        End Try

    End Function

    Public Shared Function FcReadDebitorKName(ByRef objfiBuha As SBSXASLib.AXiFBhg, ByVal lngDebKtoNbr As Long) As String

        Dim strDebitorKName As String
        Dim strDebitorKAr() As String

        strDebitorKName = objfiBuha.GetKontoInfo(lngDebKtoNbr)

        strDebitorKAr = Split(strDebitorKName, "{>}")

        Return strDebitorKAr(8)

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


    Public Shared Function FcGetRefDebiNr(ByRef objdbconn As MySqlConnection,
                                          ByRef objdbconnZHDB02 As MySqlConnection,
                                          ByRef objsqlcommand As MySqlCommand,
                                          ByRef objsqlcommandZHDB02 As MySqlCommand,
                                          ByRef objOrdbconn As OracleClient.OracleConnection,
                                          ByRef objOrcommand As OracleClient.OracleCommand,
                                          ByVal lngDebiNbr As Int32,
                                          ByVal intAccounting As Int32,
                                          ByRef intDebiNew As Int32) As Int16

        'Return 0=ok, 1=noch nicht implementiert, 2=Rep_Ref nicht definiert, 3=Nicht in Tab_Repbetriebe, 4=keine Angaben in Tab_Repbetriebe

        Dim strTableName, strTableType, strDebFieldName, strDebNewField, strDebNewFieldType, strCompFieldName, strStreetFieldName, strZIPFieldName, strTownFieldName, strSageName, strDebiAccField As String
        'Dim intCreatable As Int16
        Dim objdtDebitor As New DataTable
        Dim intPKNewField As Int32
        Dim objdbConnDeb As New MySqlConnection
        Dim objsqlCommDeb As New MySqlCommand

        strTableName = FcReadFromSettings(objdbconn, "Buchh_PKTable", intAccounting)
        strTableType = FcReadFromSettings(objdbconn, "Buchh_PKTableType", intAccounting)
        strDebFieldName = FcReadFromSettings(objdbconn, "Buchh_PKField", intAccounting)
        strDebNewField = FcReadFromSettings(objdbconn, "Buchh_PKNewField", intAccounting)
        strDebNewFieldType = FcReadFromSettings(objdbconn, "Buchh_PKNewFType", intAccounting)
        strCompFieldName = FcReadFromSettings(objdbconn, "Buchh_PKCompany", intAccounting)
        strStreetFieldName = FcReadFromSettings(objdbconn, "Buchh_PKStreet", intAccounting)
        strZIPFieldName = FcReadFromSettings(objdbconn, "Buchh_PKZIP", intAccounting)
        strTownFieldName = FcReadFromSettings(objdbconn, "Buchh_PKTown", intAccounting)
        strSageName = FcReadFromSettings(objdbconn, "Buchh_PKSageName", intAccounting)
        strDebiAccField = FcReadFromSettings(objdbconn, "Buchh_DPKAccount", intAccounting)

        If strTableName <> "" And strDebFieldName <> "" Then

            If strTableType = "O" Then 'Oracle
                'objOrdbconn.Open()
                objOrcommand.CommandText = "SELECT " + strDebFieldName + ", " + strDebNewField + ", " + strCompFieldName + ", " + strStreetFieldName + ", " + strZIPFieldName + ", " + strTownFieldName + ", " + strSageName + ", " + strDebiAccField +
                                            " FROM " + strTableName + " WHERE " + strDebFieldName + "=" + lngDebiNbr.ToString
                objdtDebitor.Load(objOrcommand.ExecuteReader)
                'Ist DebiNrNew Linked oder Direkt
                'If strDebNewFieldType = "D" Then

                'objOrdbconn.Close()
            ElseIf strTableType = "M" Then 'MySQL
                intDebiNew = 0
                'MySQL - Tabelle einlesen
                objdbConnDeb.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(FcReadFromSettings(objdbconn, "Buchh_PKTableConnection", intAccounting))
                objdbConnDeb.Open()
                objsqlCommDeb.CommandText = "SELECT " + strDebFieldName + ", " + strDebNewField + ", " + strCompFieldName + ", " + strStreetFieldName + ", " + strZIPFieldName + ", " + strTownFieldName + ", " + strSageName + ", " + strDebiAccField +
                                            " FROM " + strTableName + " WHERE " + strDebFieldName + "=" + lngDebiNbr.ToString
                objsqlCommDeb.Connection = objdbConnDeb
                objdtDebitor.Load(objsqlCommDeb.ExecuteReader)
                objdbConnDeb.Close()

            End If

            If IsDBNull(objdtDebitor.Rows(0).Item(strDebNewField)) Then
                intDebiNew = 0
                Return 2
            Else
                intPKNewField = objdtDebitor.Rows(0).Item(strDebNewField)
                intPKNewField = FcGetPKNewFromRep(objdbconnZHDB02, objsqlcommandZHDB02, objdtDebitor.Rows(0).Item(strDebNewField))
                If intPKNewField = 0 Then
                    intDebiNew = 0
                    Return 3
                Else
                    intDebiNew = intPKNewField
                    Return 0
                End If
            End If


        End If

        Return intPKNewField


    End Function

    Public Shared Function FcGetRefKrediNr(ByRef objdbconn As MySqlConnection,
                                          ByRef objdbconnZHDB02 As MySqlConnection,
                                          ByRef objsqlcommand As MySqlCommand,
                                          ByRef objsqlcommandZHDB02 As MySqlCommand,
                                          ByRef objOrdbconn As OracleClient.OracleConnection,
                                          ByRef objOrcommand As OracleClient.OracleCommand,
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

        strTableName = FcReadFromSettings(objdbconn, "Buchh_PKKrediTable", intAccounting)
        strTableType = FcReadFromSettings(objdbconn, "Buchh_PKKrediTableType", intAccounting)
        strKredFieldName = FcReadFromSettings(objdbconn, "Buchh_PKKrediField", intAccounting)
        strKredNewField = FcReadFromSettings(objdbconn, "Buchh_PKKrediNewField", intAccounting)
        strKredNewFieldType = FcReadFromSettings(objdbconn, "Buchh_PKKrediNewFType", intAccounting)
        strCompFieldName = FcReadFromSettings(objdbconn, "Buchh_PKKrediCompany", intAccounting)
        strStreetFieldName = FcReadFromSettings(objdbconn, "Buchh_PKKrediStreet", intAccounting)
        strZIPFieldName = FcReadFromSettings(objdbconn, "Buchh_PKKrediZIP", intAccounting)
        strTownFieldName = FcReadFromSettings(objdbconn, "Buchh_PKKrediTown", intAccounting)
        strSageName = FcReadFromSettings(objdbconn, "Buchh_PKKrediSageName", intAccounting)
        strKredAccField = FcReadFromSettings(objdbconn, "Buchh_PKKrediAccount", intAccounting)

        If strTableName <> "" And strKredFieldName <> "" Then

            If strTableType = "O" Then 'Oracle
                'objOrdbconn.Open()
                objOrcommand.CommandText = "SELECT " + strKredFieldName + ", " + strKredNewField + ", " + strCompFieldName + ", " + strStreetFieldName + ", " + strZIPFieldName + ", " + strTownFieldName + ", " + strSageName + ", " + strKredAccField +
                                            " FROM " + strTableName + " WHERE " + strKredFieldName + "=" + lngKrediNbr.ToString
                objdtKreditor.Load(objOrcommand.ExecuteReader)
                'Ist DebiNrNew Linked oder Direkt
                'If strDebNewFieldType = "D" Then

                'objOrdbconn.Close()
            ElseIf strTableType = "M" Then 'MySQL
                intKrediNew = 0
                'MySQL - Tabelle einlesen
                objdbConnKred.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(FcReadFromSettings(objdbconn, "Buchh_PKKrediTableConnection", intAccounting))
                objdbConnKred.Open()
                objsqlCommKred.CommandText = "SELECT " + strKredFieldName + ", " + strKredNewField + ", " + strCompFieldName + ", " + strStreetFieldName + ", " + strZIPFieldName + ", " + strTownFieldName + ", " + strSageName + ", " + strKredAccField +
                                            " FROM " + strTableName + " WHERE " + strKredFieldName + "=" + lngKrediNbr.ToString
                objsqlCommKred.Connection = objdbConnKred
                objdtKreditor.Load(objsqlCommKred.ExecuteReader)
                objdbConnKred.Close()

            End If

            'If IsDBNull(objdtKreditor.Rows(0).Item(strKredNewField)) Then
            If objdtKreditor.Rows.Count = 0 Then
                intKrediNew = 0
                Return 2
            Else
                intPKNewField = IIf(IsDBNull(objdtKreditor.Rows(0).Item(strKredNewField)), 0, objdtKreditor.Rows(0).Item(strKredNewField))
                'intPKNewField = FcGetPKNewFromRep(objdbconnZHDB02, objsqlcommandZHDB02, objdtKreditor.Rows(0).Item(strKredNewField))
                If intPKNewField = 0 Then
                    intKrediNew = 0
                    Return 3
                Else
                    intKrediNew = intPKNewField
                    Return 0
                End If
            End If


        End If

        Return intPKNewField


    End Function


    Public Shared Function FcIsDebitorCreatable(ByRef objdbconnZHDB02 As MySqlConnection,
                                                ByRef objsqlcommandZHDB02 As MySqlCommand,
                                                ByVal lngDebiNbr As Long,
                                                ByRef objDbBhg As SBSXASLib.AXiDbBhg) As Int16

        'Return: 0=creatable und erstellt, 3=Sage - Suchtext nicht erfasst, 4=Betrieb nicht gefunden, 9=Nicht hinterlegt

        Dim intCreatable As Int16
        Dim objdtDebitor As New DataTable
        Dim strLand As String
        Dim intLangauage As Int32
        'Dim intPKNewField As Int32
        Dim strSQL As String
        Dim intAffected As Int16

        Try

            'Angaben einlesen
            objdbconnZHDB02.Open()
            objsqlcommandZHDB02.CommandText = "SELECT Rep_Firma, Rep_Strasse, Rep_PLZ, Rep_Ort, Rep_DebiKonto, Rep_Gruppe, Rep_Vertretung, Rep_Ansprechpartner, Rep_Land, Rep_Tel1, Rep_Fax, Rep_Mail, " +
                                                "Rep_Language, Rep_Kredi_MWSTNr, Rep_Kreditlimite, Rep_Kred_Pay_Def, Rep_Kred_Bank_Name, Rep_Kred_Bank_PLZ, Rep_Kred_Bank_Ort, Rep_Kred_IBAN, Rep_Kred_Bank_BIC, " +
                                                "Rep_Kred_Currency FROM Tab_Repbetriebe WHERE PKNr=" + lngDebiNbr.ToString
            objdtDebitor.Load(objsqlcommandZHDB02.ExecuteReader)

            'Gefunden?
            If objdtDebitor.Rows.Count > 0 Then
                'Debug.Print("Gefunden, kann erstellt werden")

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
                        strLand = "NA"
                End Select

                'Sprache zuweisen von 1-Stelligem String nach Sage 200 Regionen
                Select Case IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Language")), "D", objdtDebitor.Rows(0).Item("Rep_Language"))
                    Case "D"
                        intLangauage = 2055
                    Case "F"
                        intLangauage = 4108
                    Case "I"
                        intLangauage = 2064
                    Case Else
                        intLangauage = 2057 'Englisch
                End Select

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
                                          IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Kred_Pay_Def")), 0, objdtDebitor.Rows(0).Item("Rep_Kred_Pay_Def")),
                                          IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Kred_Bank_Name")), "", objdtDebitor.Rows(0).Item("Rep_Kred_Bank_Name")),
                                          IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Kred_Bank_PLZ")), "", objdtDebitor.Rows(0).Item("Rep_Kred_Bank_PLZ")),
                                          IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Kred_Bank_Ort")), "", objdtDebitor.Rows(0).Item("Rep_Kred_Bank_Ort")),
                                          IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Kred_IBAN")), "", objdtDebitor.Rows(0).Item("Rep_Kred_IBAN")),
                                          IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Kred_Bank_BIC")), "", objdtDebitor.Rows(0).Item("Rep_Kred_Bank_BIC")),
                                          IIf(IsDBNull(objdtDebitor.Rows(0).Item("Rep_Kred_Currency")), "CHF", objdtDebitor.Rows(0).Item("Rep_Kred_Currency")))

                If intCreatable = 0 Then
                    'MySQL
                    strSQL = "INSERT INTO Tbl_RTFAutomail (RGNbr, MailCreateDate, MailCreateWho, MailTo, MailSender, MailTitle, MAilMsg, MailSent) VALUES (" +
                                                         lngDebiNbr.ToString + ", Date('" + Format(Today(), "yyyy-MM-dd").ToString + "'), 'Sage200Imp', " +
                                                         "'rene.hager@mssag.ch', 'Sage200@mssag.ch', 'Debitor " +
                                                         lngDebiNbr.ToString + " wurde erstell im Mandant EE', 'Bitte kontrollieren und Daten erg&auml;nzen.', false)"
                    ' objlocMySQLRGConn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(strMDBName)
                    'objlocMySQLRGConn.Open()
                    'objlocMySQLRGcmd.Connection = objlocMySQLRGConn
                    objsqlcommandZHDB02.CommandText = strSQL
                    intAffected = objsqlcommandZHDB02.ExecuteNonQuery()

                End If


                Return 0
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

    Public Shared Function FcIsKreditorCreatable(ByRef objdbconnZHDB02 As MySqlConnection,
                                                ByRef objsqlcommandZHDB02 As MySqlCommand,
                                                ByVal lngKrediNbr As Long,
                                                ByRef objKrBhg As SBSXASLib.AXiKrBhg) As Int16

        'Return: 0=creatable und erstellt, 3=Sage - Suchtext nicht erfasst, 4=Betrieb nicht gefunden, 9=Nicht hinterlegt

        Dim intCreatable As Int16
        Dim objdtKreditor As New DataTable
        Dim strLand As String
        Dim intLangauage As Int32
        'Dim intPKNewField As Int32
        Dim strSQL As String
        Dim intAffected As Int16

        Try

            'Angaben einlesen
            objdbconnZHDB02.Open()
            objsqlcommandZHDB02.CommandText = "SELECT Rep_Firma, Rep_Strasse, Rep_PLZ, Rep_Ort, Rep_KredGegenKonto, Rep_Gruppe, Rep_Vertretung, Rep_Ansprechpartner, Rep_Land, Rep_Tel1, Rep_Fax, Rep_Mail, " +
                                                "Rep_Language, Rep_Kredi_MWSTNr, Rep_Kreditlimite, Rep_Kred_Pay_Def, Rep_Kred_Bank_Name, Rep_Kred_Bank_PLZ, Rep_Kred_Bank_Ort, Rep_Kred_IBAN, Rep_Kred_Bank_BIC, " +
                                                "Rep_Kred_Currency FROM Tab_Repbetriebe WHERE PKNr=" + lngKrediNbr.ToString
            objsqlcommandZHDB02.Connection = objdbconnZHDB02
            objdtKreditor.Load(objsqlcommandZHDB02.ExecuteReader)

            'Gefunden?
            If objdtKreditor.Rows.Count > 0 Then
                'Debug.Print("Gefunden, kann erstellt werden")

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
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Kred_Pay_Def")), 0, objdtKreditor.Rows(0).Item("Rep_Kred_Pay_Def")),
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Kred_Bank_Name")), "", objdtKreditor.Rows(0).Item("Rep_Kred_Bank_Name")),
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Kred_Bank_PLZ")), "", objdtKreditor.Rows(0).Item("Rep_Kred_Bank_PLZ")),
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Kred_Bank_Ort")), "", objdtKreditor.Rows(0).Item("Rep_Kred_Bank_Ort")),
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Kred_IBAN")), "", objdtKreditor.Rows(0).Item("Rep_Kred_IBAN")),
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Kred_Bank_BIC")), "", objdtKreditor.Rows(0).Item("Rep_Kred_Bank_BIC")),
                                          IIf(IsDBNull(objdtKreditor.Rows(0).Item("Rep_Kred_Currency")), "CHF", objdtKreditor.Rows(0).Item("Rep_Kred_Currency")))

                If intCreatable = 0 Then
                    'MySQL
                    strSQL = "INSERT INTO Tbl_RTFAutomail (RGNbr, MailCreateDate, MailCreateWho, MailTo, MailSender, MailTitle, MAilMsg, MailSent) VALUES (" +
                                                         lngKrediNbr.ToString + ", Date('" + Format(Today(), "yyyy-MM-dd").ToString + "'), 'Sage200Imp', " +
                                                         "'rene.hager@mssag.ch', 'Sage200@mssag.ch', 'Kreditor " +
                                                         lngKrediNbr.ToString + " wurde erstell im Mandant EE', 'Bitte kontrollieren und Daten erg&auml;nzen.', false)"
                    ' objlocMySQLRGConn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings(strMDBName)
                    'objlocMySQLRGConn.Open()
                    'objlocMySQLRGcmd.Connection = objlocMySQLRGConn
                    objsqlcommandZHDB02.CommandText = strSQL
                    intAffected = objsqlcommandZHDB02.ExecuteNonQuery()

                End If


                Return 0
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


    Public Shared Function FcGetPKNewFromRep(ByRef objdbconnZHDB02 As MySqlConnection, ByRef objsqlcommandZHDB02 As MySqlCommand, ByVal intPKRefField As Int32) As Int32

        'Aus Tabelle Rep_Betriebe auf ZHDB02 auslesen
        Dim objdtRepBetrieb As New DataTable

        Try

            objdbconnZHDB02.Open()
            objsqlcommandZHDB02.Connection = objdbconnZHDB02
            objsqlcommandZHDB02.CommandText = "SELECT PKNr From tab_repbetriebe WHERE Rep_Nr=" + intPKRefField.ToString
            objdtRepBetrieb.Load(objsqlcommandZHDB02.ExecuteReader)
            If objdtRepBetrieb.Rows.Count > 0 Then
                Return objdtRepBetrieb.Rows(0).Item("PKNr")
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
                                           ByVal strCurrency As String) As Int16

        Dim strDebCountry As String = strLand
        Dim strDebCurrency As String = strCurrency
        Dim strDebSprachCode As String = intLangauage.ToString
        Dim strDebSperren As String = "N"
        Dim intDebErlKto As Integer = 3200
        Dim shrDebZahlK As Short = 1 'Wird für EE fix auf 30 Tage Netto gesetzt
        Dim intDebToleranzNbr As Integer = 1
        Dim intDebMahnGroup As Integer = 1
        Dim strDebWerbung As String = "N"
        Dim strText As String
        Dim strTelefon1 As String
        Dim strTelefax As String

        strText = IIf(strGruppe = "", "", "Gruppe: " + strGruppe) + IIf(strVertretung = "" Or "0", "", strText + vbCrLf + "Vertretung: " + strVertretung)
        strTelefon1 = IIf(strTel = "" Or strTel = "0", "", strTel)
        strTelefax = IIf(strFax = "" Or strFax = "0", "", strFax)

        'Debitor erstellen

        Try

            Call objDbBhg.SetCommonInfo2(intDebitorNewNbr, strDebName, "", strDebStreet, "", "", "", strDebCountry, strDebPLZ, strDebOrt, strTelefon1, "", strTelefax, strMail, "", strDebCurrency, "", "", strAnsprechpartner, strDebSprachCode, strText)
            Call objDbBhg.SetExtendedInfo8(strDebSperren, strKreditLimite, intDebSammelKto.ToString, intDebErlKto.ToString, "", "", "", shrDebZahlK.ToString, intDebToleranzNbr.ToString, intDebMahnGroup.ToString, "", "", strDebWerbung, "", "", strMwStNr)
            If intPayDefault = 9 Then 'IBAN
                If Len(strZVIBAN) > 15 Then
                    Call objDbBhg.SetZahlungsverbindung("B", "", strZVBankName, "", "", strZVBankPLZ.ToString, strZVBankOrt, Left(strZVIBAN, 2), Mid(strZVIBAN, 5, 5), "J", strZVBIC, "", "", "", strZVIBAN, "")
                End If
            End If
            Call objDbBhg.WriteDebitor3(0)

            'Mail über Erstellung absetzen


            Return 0
            'intDebAdrLaufN = DbBhg.GetAdressLaufnr()
            'intDebBankLaufNr = DbBhg.GetZahlungsverbLaufnr()

        Catch ex As Exception
            MessageBox.Show(ex.Message)

            Return 1

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
                                           ByVal strCurrency As String) As Int16

        Dim strKredCountry As String = strLand
        Dim strKredCurrency As String = strCurrency
        Dim strKredSprachCode As String = intLangauage.ToString
        Dim strKredSperren As String = "N"
        'Dim intKredErlKto As Integer = 2000
        Dim intKredVorErfKto As Int32 = 2040
        Dim intKredAufwandKto As Int32 = 4200
        Dim shrKredZahlK As Short = 1 'Wird für EE fix auf 30 Tage Netto gesetzt
        Dim intKredToleranzNbr As Integer = 1
        Dim intKredMahnGroup As Integer = 1
        Dim strKredWerbung As String = "N"
        Dim strText As String
        Dim strTelefon1 As String
        Dim strTelefax As String

        strText = IIf(strGruppe = "", "", "Gruppe: " + strGruppe) + IIf(strVertretung = "" Or "0", "", strText + vbCrLf + "Vertretung: " + strVertretung)
        strTelefon1 = IIf(strTel = "" Or strTel = "0", "", strTel)
        strTelefax = IIf(strFax = "" Or strFax = "0", "", strFax)

        'Debitor erstellen

        Try

            Call objKrBhg.SetCommonInfo2(intKreditorNewNbr,
                                         strKredName,
                                         "",
                                         strKredStreet,
                                         "",
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
                                           intKredAufwandKto.ToString,
                                           "",
                                           "",
                                           "",
                                           shrKredZahlK.ToString,
                                           "",
                                           strKredWerbung)
            If intPayDefault = 9 Then 'IBAN
                If Len(strZVIBAN) > 15 Then
                    Call objKrBhg.SetZahlungsverbindung("B", strZVIBAN, strZVBankName, "", "", strZVBankPLZ.ToString, strZVBankOrt, Left(strZVIBAN, 2), Mid(strZVIBAN, 5, 5), "J", strZVBIC, "", "", "", strZVIBAN, "")
                End If
            End If
            Call objKrBhg.WriteKreditor3(0)

            'Mail über Erstellung absetzen


            Return 0
            'intDebAdrLaufN = DbBhg.GetAdressLaufnr()
            'intDebBankLaufNr = DbBhg.GetZahlungsverbLaufnr()

        Catch ex As Exception
            MessageBox.Show(ex.Message)

            Return 1

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

    Public Shared Function FcCheckKonto(ByVal lngKtoNbr As Long, ByRef objfiBuha As SBSXASLib.AXiFBhg, ByVal dblMwSt As Double) As Integer

        'Returns 0=ok, 1=existiert nicht, 2=existiert aber keine Steuern

        Dim strReturn As String
        Dim strKontoInfo() As String

        strReturn = objfiBuha.GetKontoInfo(lngKtoNbr.ToString)
        If strReturn = "EOF" Then
            Return 1
        Else
            'If dblMwSt = 0 Then
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

    Public Shared Function FcCheckDebitor(ByVal lngDebitor As Long, ByVal intBuchungsart As Integer, ByRef objdbBuha As SBSXASLib.AXiDbBhg) As Integer

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
        Dim strdbProvider, strdbSource, strdbPathAndFile As String


        objMySQLConn.Open()

        strMDBName = FcReadFromSettings(objMySQLConn, "Buchh_RGTableMDB", intMandant)
        strRGTableType = FcReadFromSettings(objMySQLConn, "Buchh_RGTableType", intMandant)
        strNameRGTable = FcReadFromSettings(objMySQLConn, "Buchh_TableDeb", intMandant)
        strBelegNrName = FcReadFromSettings(objMySQLConn, "Buchh_TableRGBelegNrName", intMandant)
        strRGNbrFieldName = FcReadFromSettings(objMySQLConn, "Buchh_TableRGNbrFieldName", intMandant)
        'strSQL = "UPDATE " + strNameRGTable + " SET gebucht=true, gebuchtDatum=#" + Format(datDate, "yyyy-MM-dd").ToString + "#, " + strBelegNrName + "=" + intBelegNr.ToString + " WHERE " + strRGNbrFieldName + "=" + strRGNbr

        Try

            If strRGTableType = "A" Then
                'Access
                strdbProvider = "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
                strdbSource = "Data Source="
                strdbPathAndFile = "\\sdlc.mssag.ch\Apps\Backends\" + strMDBName + ";Jet OLEDB:System Database=\\sdlc.mssag.ch\Apps\Backends\Workbench.mdw;User ID=HagerR;"
                strSQL = "UPDATE " + strNameRGTable + " SET gebucht=true, gebuchtDatum=#" + Format(datDate, "yyyy-MM-dd").ToString + "#, " + strBelegNrName + "=" + intBelegNr.ToString + " WHERE " + strRGNbrFieldName + "=" + strRGNbr

                objdbAccessConn.ConnectionString = strdbProvider + strdbSource + strdbPathAndFile
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
        Dim strdbProvider, strdbSource, strdbPathAndFile As String


        objMySQLConn.Open()

        strMDBName = FcReadFromSettings(objMySQLConn, "Buchh_KRGTableMDB", intMandant)
        strKRGTableType = FcReadFromSettings(objMySQLConn, "Buchh_KRGTableType", intMandant)
        strNameKRGTable = FcReadFromSettings(objMySQLConn, "Buchh_TableKred", intMandant)
        strBelegNrName = FcReadFromSettings(objMySQLConn, "Buchh_TableKRGBelegNrName", intMandant)
        strKRGNbrFieldName = FcReadFromSettings(objMySQLConn, "Buchh_TableKRGNbrFieldName", intMandant)
        'strSQL = "UPDATE " + strNameRGTable + " SET gebucht=true, gebuchtDatum=#" + Format(datDate, "yyyy-MM-dd").ToString + "#, " + strBelegNrName + "=" + intBelegNr.ToString + " WHERE " + strRGNbrFieldName + "=" + strRGNbr

        Try

            If strKRGTableType = "A" Then
                'Access
                strdbProvider = "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
                strdbSource = "Data Source="
                strdbPathAndFile = "\\sdlc.mssag.ch\Apps\Backends\" + strMDBName + ";Jet OLEDB:System Database=\\sdlc.mssag.ch\Apps\Backends\Workbench.mdw;User ID=HagerR;"
                strSQL = "UPDATE " + strNameKRGTable + " SET Kredigebucht=true, KredigebuchtDatum=#" + Format(datDate, "yyyy-MM-dd").ToString + "#, " + strBelegNrName + "='" + intBelegNr.ToString + "' WHERE " + strKRGNbrFieldName + "=" + lngKredID.ToString

                objdbAccessConn.ConnectionString = strdbProvider + strdbSource + strdbPathAndFile
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


    Public Shared Function FcSetBuchMode(ByRef objdbBuha As SBSXASLib.AXiDbBhg, ByVal strMode As String) As Int16

        objdbBuha.SetBuchMode(strMode)

        Return 0

    End Function

    Public Shared Function FcSetBelegKopf4(ByRef objdbBuha As SBSXASLib.AXiDbBhg,
                                           ByVal lngBelegNr As Long,
                                           ByVal strValutaDatum As String,
                                           ByVal lngDebitor As Long,
                                           ByVal strBelegTyp As String,
                                           ByVal strBelegDatum As String,
                                           ByVal strVerFallDatum As String,
                                           ByVal strBelegText As String,
                                           ByVal strReferenz As String,
                                           ByVal lngKondition As Long,
                                           ByVal strSachbearbeiter As String,
                                           ByVal strVerkaeufer As String,
                                           ByVal strMahnSperre As String,
                                           ByVal shrMahnstufe As Short,
                                           ByVal strBetraBrutto As String,
                                           ByVal strKurs As String,
                                           ByVal strBelegExt As String,
                                           ByVal strSKonto As String,
                                           ByVal strDebiCur As String,
                                           ByVal strSammelKonto As String,
                                           ByVal strVerzugsZ As String,
                                           ByVal strZusatzText As String,
                                           ByVal strEBankKonto As String,
                                           ByVal strIkoDebitor As String) As Integer


        'Zuerst prüfen ob Zwingende Werte angegeben worden sind

        'Ausführung
        objdbBuha.SetBelegKopf4(lngBelegNr, strValutaDatum, lngDebitor, strBelegTyp, strBelegDatum, strVerFallDatum, strBelegText, strReferenz, lngKondition, strSachbearbeiter, strVerkaeufer, strMahnSperre, shrMahnstufe, strBetraBrutto,
                                strKurs, strBelegExt, strSKonto, strDebiCur, strSammelKonto, strVerzugsZ, strZusatzText, strEBankKonto, strIkoDebitor)

        Return 0

    End Function

    Public Shared Function FcSetVerteilung(ByRef objdbBuha As SBSXASLib.AXiDbBhg,
                                           ByVal strGegenKonto As String,
                                           ByVal strFibuText As String,
                                           ByVal strNettoBetrag As String,
                                           ByVal strArraySteuer As String,
                                           ByVal strArrayKST As String,
                                           ByVal strArrayKSTE As String) As Integer

        'Prüfen ob Daten vollständig

        'Ausführung
        objdbBuha.SetVerteilung(strGegenKonto, strFibuText, strNettoBetrag, strArraySteuer, strArrayKST, strArrayKSTE)

        Return 0

    End Function

    Public Shared Function FcWriteBuchung(ByRef objdbBuha As SBSXASLib.AXiDbBhg) As Integer

        'Ausführung
        objdbBuha.WriteBuchung()

        Return 0

    End Function

    Public Shared Function FcGetSteuerFeld(ByRef objFBhg As SBSXASLib.AXiFBhg, ByVal lngKto As Long, ByVal strDebiSubText As String, ByVal dblBrutto As Double, ByVal strMwStKey As String, ByVal dblMwSt As Double) As String

        Dim strSteuerFeld As String = ""

        Try

            If dblMwSt > 0 Then

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

    Public Shared Function FcFillKredit(ByVal intAccounting As Integer,
                                       ByRef objdtHead As DataTable,
                                       ByRef objdtSub As DataTable,
                                       ByRef objdbconn As MySqlConnection,
                                       ByRef objdbAccessConn As OleDb.OleDbConnection) As Integer

        Dim strSQL As String
        Dim strSQLSub As String
        Dim strKRGTableType As String
        Dim objRGMySQLConn As New MySqlConnection
        Dim objlocMySQLcmd As New MySqlCommand
        Dim objlocOLEdbcmd As New OleDb.OleDbCommand

        Dim objDTDebiHead As New DataTable
        Dim dbProvider, dbSource, dbPathAndFile, strMDBName As String
        Dim objdrSub As DataRow
        Dim intFcReturns As Int16

        objdbconn.Open()

        strMDBName = FcReadFromSettings(objdbconn, "Buchh_KRGTableMDB", intAccounting)
        dbProvider = "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
        dbSource = "Data Source="
        dbPathAndFile = "\\sdlc.mssag.ch\Apps\Backends\" + strMDBName + ";Jet OLEDB:System Database=\\sdlc.mssag.ch\Apps\Backends\Workbench.mdw;User ID=HagerR;"

        'Head Debitzoren löschen
        objdtHead.Clear()
        strSQL = FcReadFromSettings(objdbconn, "Buchh_SQLHeadKred", intAccounting)
        strKRGTableType = FcReadFromSettings(objdbconn, "Buchh_KRGTableType", intAccounting)

        Try

            'objlocMySQLcmd.CommandText = strSQL
            If strKRGTableType = "A" Then
                'Access
                objdbAccessConn.ConnectionString = dbProvider + dbSource + dbPathAndFile
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
                strSQLSub = FcSQLParseKredi(FcReadFromSettings(objdbconn, "Buchh_SQLDetailKred", intAccounting), row("lngKredID"), objdtHead)
                If strKRGTableType = "A" Then
                    objlocOLEdbcmd.CommandText = strSQLSub
                    objdtSub.Load(objlocOLEdbcmd.ExecuteReader)
                ElseIf strKRGTableType = "M" Then
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
            MessageBox.Show(ex.Message)

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
                                        ByRef objdtInfo As DataTable) As Integer

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
                'If row("strDebRGNbr") = "44208" Then Stop
                'Runden
                row("dblKredNetto") = Decimal.Round(row("dblKredNetto"), 2, MidpointRounding.AwayFromZero)
                row("dblKredMwSt") = Decimal.Round(row("dblKredMwst"), 2, MidpointRounding.AwayFromZero)
                row("dblKredBrutto") = Decimal.Round(row("dblKredBrutto"), 2, MidpointRounding.AwayFromZero)
                'Status-String erstellen
                'Kreditor 01
                intReturnValue = FcGetRefKrediNr(objdbconn, objdbconnZHDB02, objsqlcommand, objsqlcommandZHDB02, objOrdbconn, objOrcommand, IIf(IsDBNull(row("lngKredNbr")), 0, row("lngKredNbr")), intAccounting, intKreditorNew)
                strBitLog += Trim(intReturnValue.ToString)
                If intKreditorNew <> 0 Then
                    intReturnValue = FcCheckKreditor(intKreditorNew, row("intBuchungsart"), objKrBuha)
                Else
                    intReturnValue = 2
                End If
                strBitLog = Trim(intReturnValue.ToString)

                'Kto 02
                intReturnValue = FcCheckKonto(row("lngKredKtoNbr"), objfiBuha, row("dblKredMwSt"))
                strBitLog += Trim(intReturnValue.ToString)

                'Currency 03
                intReturnValue = FcCheckCurrency(row("strKredCur"), objfiBuha)
                strBitLog += Trim(intReturnValue.ToString)

                'Sub 04
                'booAutoCorrect = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettings(objdbconn, "Buchh_HeadAutoCorrect", intAccounting)))
                booAutoCorrect = False
                intReturnValue = FcCheckKrediSubBookings(row("lngKredID"), objdtKreditSubs, intSubNumber, dblSubBrutto, dblSubNetto, dblSubMwSt, objdbconn, objfiBuha, objdbPIFb, row("intBuchungsart"), booAutoCorrect)
                strBitLog += Trim(intReturnValue.ToString)

                ''Autokorrektur 05
                ''booAutoCorrect = Convert.ToBoolean(Convert.ToInt16(FcReadFromSettings(objdbconn, "Buchh_HeadAutoCorrect", intAccounting)))
                booAutoCorrect = False
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
                intReturnValue = FcCheckBelegHead(row("intBuchungsart"), IIf(IsDBNull(row("dblKredBrutto")), 0, row("dblKredBrutto")), IIf(IsDBNull(row("dblKredNetto")), 0, row("dblKredNetto")), IIf(IsDBNull(row("dblKredMwSt")), 0, row("dblKredMwSt")))
                strBitLog += Trim(intReturnValue.ToString)
                'OP - Nummer prüfen 08
                'intReturnValue = FcCreateDebRef(objdbconn, intAccounting, row("strDebiBank"), row("strDebRGNbr"), row("intBuchungsart"), strDebiReferenz)
                strCleanOPNbr = IIf(IsDBNull(row("strOPNr")), "", row("strOPNr"))
                intReturnValue = FcChCeckKredOP(strCleanOPNbr, IIf(IsDBNull(row("strKredRGNbr")), "", row("strKredRGNbr")))
                row("strOPNr") = strCleanOPNbr
                strBitLog += Trim(intReturnValue.ToString)
                'OP - Verdopplung 09
                intReturnValue = FcCheckKrediOPDouble(objKrBuha, IIf(IsDBNull(row("lngKredNbr")), 0, row("lngKredNbr")), row("strKredRGNbr"))
                strBitLog += Trim(intReturnValue.ToString)
                'Valuta - Datum 10
                intReturnValue = FcChCeckDate(row("datKredValDatum"), objdtInfo)
                strBitLog += Trim(intReturnValue.ToString)
                'RG - Datum 11
                intReturnValue = FcChCeckDate(row("datKredRGDatum"), objdtInfo)
                strBitLog += Trim(intReturnValue.ToString)
                ''intReturnValue = fcCheckIntBank()


                'Status-String auswerten
                'Kreditor
                If Left(strBitLog, 1) <> "0" Then
                    strStatus = "Kred"
                    If Left(strBitLog, 1) <> "2" Then
                        intReturnValue = FcIsKreditorCreatable(objdbconnZHDB02, objsqlcommandZHDB02, intKreditorNew, objKrBuha)
                        If intReturnValue = 0 Then
                            strStatus += " erstellt"
                        Else
                            strStatus += " nicht erstellt."
                        End If
                        row("strKredBez") = FcReadKreditorName(objKrBuha, intKreditorNew, row("strKredCur"))
                        row("lngKredNbr") = intKreditorNew
                    Else
                        strStatus += " keine Ref"
                        row("strKredBez") = "n/a"
                    End If
                Else
                    row("strKredBez") = FcReadKreditorName(objKrBuha, intKreditorNew, row("strKredCur"))
                    row("lngKredNbr") = intKreditorNew
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
                    row("strKredKtoBez") = FcReadDebitorKName(objfiBuha, row("lngKredKtoNbr"))
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
                'OP - Nr.

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
        Dim dbProvider, dbSource, dbPathAndFile, strMDBName As String


        Try

            objdbconn.Open()
            'Gibt es transitorische Buchungen?
            booTransits = CBool(FcReadFromSettings(objdbconn, "Buchh_Transit", intAccounting))

            If booTransits Then

                'Table - Art lesen
                strRGTableType = FcReadFromSettings(objdbconn, "Buchh_RGTableType", intAccounting)
                'Debitoren - Table Name lesen
                strMDBName = FcReadFromSettings(objdbconn, "Buchh_RGTableMDB", intAccounting)
                ''Access
                'dbProvider = "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
                'dbSource = "Data Source="
                'dbPathAndFile = "\\sdlc.mssag.ch\Apps\Backends\" + strMDBName + ";Jet OLEDB:System Database=\\sdlc.mssag.ch\Apps\Backends\Workbench.mdw;User ID=HagerR;"

                'Debitzoren Transit-Queries für Mandant einlesen
                strSQL = "SELECT * FROM buchhaltungen_sub WHERE strType='D' AND refMandant=" + intAccounting.ToString
                objRGMySQLConn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString")
                objlocMySQLcmd.Connection = objRGMySQLConn
                objlocMySQLcmd.CommandText = strSQL
                objRGMySQLConn.Open()
                objDTTransitDebits.Load(objlocMySQLcmd.ExecuteReader)

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
                            objdbAccessConn.Open()
                            objlocOLEdbcmd.Connection = objdbAccessConn
                            objlocOLEdbcmd.CommandText = rowdebitquery("strSQL")
                            intAffected = objlocOLEdbcmd.ExecuteNonQuery()
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

End Class
