Option Strict Off
Option Explicit On

Imports MySql.Data.MySqlClient

Friend Class frmImportMain

    Public Finanz As SBSXASLib.AXFinanz
    Public FBhg As SBSXASLib.AXiFBhg
    Public DbBhg As SBSXASLib.AXiDbBhg
    Public KrBhg As SBSXASLib.AXiKrBhg
    Public BsExt As SBSXASLib.AXiBSExt
    Public Adr As SBSXASLib.AXiAdr
    Public BeBu As SBSXASLib.AXiBeBu
    Public PIFin As SBSXASLib.AXiPlFin

    Public Methode As String
    Public DidOpenmandant As Boolean

    Public FELD_SEP As String
    Public REC_SEP As String
    Public KSTKTR_SEP As String
    Public FELD_SEP_OUT As String
    Public REC_SEP_OUT As String
    Public nID As String

    Public objdbConn As New MySqlConnection(System.Configuration.ConfigurationManager.AppSettings("OwnConnectionString"))
    Public objDABuchhaltungen As New MySqlDataAdapter("SELECT * FROM buchhaltungen WHERE NOT Buchh200_Name IS NULL", objdbConn)
    'Public objDACarsGrid As New MySqlDataAdapter("SELECT tblcars.idCar, tblunits.strUnit, tblplates.strPlate, tblcars.strVIN, tblmodelle.strModell FROM tblcars LEFT JOIN tblunits ON tblcars.refUnit = tblunits.idUnit LEFT JOIN tblplates ON tblcars.refPlate = tblplates.idPlate LEFT JOIN tblmodelle ON tblcars.refModell = tblmodelle.idModell", objdbConn)
    'Public objdtDebitor As New DataTable("tbliDebitor")
    Public objdtBuchhaltungen As New DataTable("tbliBuchhaltungen")
    Public objdtDebitorenHead As New DataTable("tbliDebitorenHead")



    Public Sub InitVar()
        'UPGRADE_NOTE: Object PIFin may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        PIFin = Nothing
        'UPGRADE_NOTE: Object KrBhg may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        KrBhg = Nothing
        'UPGRADE_NOTE: Object FBhg may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        FBhg = Nothing
        'UPGRADE_NOTE: Object DbBhg may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        DbBhg = Nothing
        'UPGRADE_NOTE: Object BsExt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        BsExt = Nothing
        'UPGRADE_NOTE: Object BeBu may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        BeBu = Nothing
        'UPGRADE_NOTE: Object Adr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Adr = Nothing
        'UPGRADE_NOTE: Object Finanz may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Finanz = Nothing

        'Call Check_CheckStateChanged(Check, New System.EventArgs())

        FELD_SEP = "{<}"
        REC_SEP = "{>}"
        KSTKTR_SEP = "{-}"

        FELD_SEP_OUT = "{>}"
        REC_SEP_OUT = "{<}"

        'AXFinanzForm.rec1.Text = REC_SEP
        'AXFinanzForm.feld1.Text = FELD_SEP
        'AXFinanzForm.kst1.Text = KSTKTR_SEP

        'AXFinanzForm.rec2.Text = REC_SEP_OUT
        'AXFinanzForm.feld2.Text = FELD_SEP_OUT

        'lblVersion.Text = "SBSxas V-" & Version
    End Sub

    Private Sub butDebitoren_Click(sender As Object, e As EventArgs) Handles butDebitoren.Click

        Dim booAccOk As Boolean
        Dim strMandant As String
        Dim b As Object
        Dim s As Object
        b = Nothing
        On Error GoTo ErrorHandler

        Call InitVar()

        'Vorübergehend
        strMandant = "ZZ"

        Finanz = Nothing
        Finanz = New SBSXASLib.AXFinanz

        'Loign
        'Call Finanz.ConnectSBSdbNoPrompt("sage_Sage200", "Sage200", "sage200admin", "sage200", "")
        Call Finanz.ConnectSBSdb("ZHAP03", "Sage200", "sage200admin", "sage200", "")

        'Check Mandant
        booAccOk = Finanz.CheckMandant(strMandant)
        'Debug.Print("Ok " + booAccOk.ToString)

        'Check Access-Level
        booAccOk = Finanz.CheckAccess(0, strMandant) 'admin
        booAccOk = Finanz.CheckAccess(1, strMandant) 'hauptbuch
        booAccOk = Finanz.CheckAccess(2, strMandant) 'debi
        booAccOk = Finanz.CheckAccess(3, strMandant) 'kredi
        booAccOk = Finanz.CheckAccess(4, strMandant) 'lohn
        booAccOk = Finanz.CheckAccess(14, strMandant) 'darlehen

        'Check Periode
        booAccOk = Finanz.CheckPeriode(strMandant, "2020")

        'Open Mandantg
        Finanz.OpenMandant(strMandant, "2020")

        If b = 0 Then GoTo isOk
        b = b - 200
        MsgBox("Mandant oder Periode falsch - Programm beendet", 0, "Fehler")
        Finanz = Nothing
        End

isOk:
        'Debitor öffnen und Konto überprüfen
        'Dim db As SBSXASLib.AXiDbBhg
        DbBhg = Nothing
        DbBhg = Finanz.GetDebiObj
        'db = Main_Renamed.Finanz.GetDebiObj
        'UPGRADE_WARNING: Couldn't resolve default property of object s. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        's = db.ReadDebitor3(CInt(DebitorID.Text), WhgID.Text)
        s = DbBhg.ReadDebitor3(-1000, "")
        Debug.Print("Angaben Debitor " + s)
        'UPGRADE_NOTE: Object db may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'db = Nothing

        'UPGRADE_WARNING: Couldn't resolve default property of object s. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'AusgabeText.Text = "ReadDebitor3:" & Chr(13) & Chr(10) & s



        'MsgBox("OpenMandant:" & Chr(13) & Chr(10) & "Funktionierte")
        Exit Sub

ErrorHandler:
        'UPGRADE_WARNING: Couldn't resolve default property of object b. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        b = Err.Number And 65535
        'UPGRADE_WARNING: Couldn't resolve default property of object b. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        MsgBox("OpenMandant:" & Chr(13) & Chr(10) & "Error" & Chr(13) & Chr(10) & "Die Button auf dem Main wurden ausgeschaltet !!!" & Chr(13) & Chr(10) & "Error # " & Str(Err.Number) & " was generated by " & Err.Source & Chr(13) & Chr(10) & Err.Description & " Unsere Fehlernummer" & Str(b))
        Err.Clear()


    End Sub

    Private Sub frmImportMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Comboxen
        objdtBuchhaltungen.Clear()
        objDABuchhaltungen.Fill(objdtBuchhaltungen)
        'cmbMarken.Sorted = True
        cmbBuha.DataSource = objdtBuchhaltungen
        cmbBuha.DisplayMember = "Buchh_Bez"
        cmbBuha.ValueMember = "Buchh_Nr"

        'Tabelle Head erstellen
        objdtDebitorenHead = Main.tblDebitorenHead()

        dgvDebitoren.DataSource = objdtDebitorenHead

    End Sub
End Class
