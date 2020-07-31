Option Strict Off
Option Explicit On

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
        intBuchungsart.DataType = System.Type.[GetType]("System.Int16")
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
        Dim lngDebPKNbr As DataColumn = New DataColumn("lngDebPKNbr")
        lngDebPKNbr.DataType = System.Type.[GetType]("System.Int32")
        DT.Columns.Add(lngDebPKNbr)
        Dim strDebPKBez As DataColumn = New DataColumn("strDebPKBez")
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
        strRGBemerkung.MaxLength = 50
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
        booLinked.DataType = System.Type.[GetType]("System.Boolean")
        DT.Columns.Add(booBooked)
        Dim datBooked As DataColumn = New DataColumn("datBooked")
        datBooked.DataType = System.Type.[GetType]("System.DateTime")
        DT.Columns.Add(datBooked)
        Dim lngBelegNr As DataColumn = New DataColumn("lngBelegNr")
        lngBelegNr.DataType = System.Type.[GetType]("System.Int32")
        DT.Columns.Add(lngBelegNr)
        Return DT
    End Function

End Class
