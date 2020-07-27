Option Strict Off
Option Explicit On

Public Class Main

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

    Public strShowHiddenMethods As String
    'Public gCountFields As Boolean

    'Variablen initialisieren
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


    Private Sub changeDate()
        Dim tmpDatum As Object

        On Error Resume Next

        ''AXiBebuForm
        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'tmpDatum = VB.Right(AXiBeBuForm.Datum.Text, 4)
        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'AXiBeBuForm.Datum.Text = Jahr.Text & tmpDatum

        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'tmpDatum = VB.Right(AXiBeBuForm.Datum2.Text, 4)
        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'AXiBeBuForm.Datum2.Text = Jahr.Text & tmpDatum

        ''AXiBSEExtFrom
        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'tmpDatum = VB.Right(AXiBSExtForm.ValuaDatumBox.Text, 4)
        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'AXiBSExtForm.ValuaDatumBox.Text = Jahr.Text & tmpDatum

        ''AXiDbBhgForm
        ''tmpDatum = Right(AXiDbBhgForm.ValDte.Text, 4)
        ''AXiDbBhgForm.ValDte = Jahr.Text & tmpDatum

        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'tmpDatum = VB.Right(AXiDbBhgForm.StichDatum.Text, 4)
        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'AXiDbBhgForm.StichDatum.Text = Jahr.Text & tmpDatum

        ''AXiFBhgForm
        ''tmpDatum = Right(AXiFBhgForm.ValDate.Text, 4)
        ''AXiFBhgForm.ValDate = Jahr.Text & tmpDatum

        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'tmpDatum = VB.Right(AXiFBhgForm.bisdtein.Text, 4)
        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'AXiFBhgForm.bisdtein.Text = Jahr.Text & tmpDatum

        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'tmpDatum = VB.Right(AXiFBhgForm.innergemvaldte.Text, 4)
        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'AXiFBhgForm.innergemvaldte.Text = Jahr.Text & tmpDatum

        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'tmpDatum = VB.Right(AXiFBhgForm.ValDteBis.Text, 4)
        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'AXiFBhgForm.ValDteBis.Text = Jahr.Text & tmpDatum

        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'tmpDatum = VB.Right(AXiFBhgForm.InkraftDte.Text, 4)
        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'AXiFBhgForm.InkraftDte.Text = Jahr.Text & tmpDatum

        ''ZAXiDbBhgForm1
        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'tmpDatum = VB.Right(ZAXiDbBhgForm1.Valuta.Text, 4)
        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'ZAXiDbBhgForm1.Valuta.Text = Jahr.Text & tmpDatum

        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'tmpDatum = VB.Right(ZAXiDbBhgForm1.Belegdatum.Text, 4)
        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'ZAXiDbBhgForm1.Belegdatum.Text = Jahr.Text & tmpDatum

        ''ZAXiKrBhgForm1
        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'tmpDatum = VB.Right(ZAXiKrBhgForm1.Valuta.Text, 4)
        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'ZAXiKrBhgForm1.Valuta.Text = Jahr.Text & tmpDatum

        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'tmpDatum = VB.Right(ZAXiKrBhgForm1.BelDatum.Text, 4)
        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'ZAXiKrBhgForm1.BelDatum.Text = Jahr.Text & tmpDatum

        ''ZKrediWriteZahlung
        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'tmpDatum = VB.Right(ZKrediWriteZahlung.Datum1.Text, 4)
        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'ZKrediWriteZahlung.Datum1.Text = Jahr.Text & tmpDatum

        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'tmpDatum = VB.Right(ZKrediWriteZahlung.Datum.Text, 4)
        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'ZKrediWriteZahlung.Datum.Text = Jahr.Text & tmpDatum

        ''ZSammel_BuchungForm
        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'tmpDatum = VB.Right(ZSammel_BuchungForm.Datum.Text, 4)
        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'ZSammel_BuchungForm.Datum.Text = Jahr.Text & tmpDatum

        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'tmpDatum = VB.Right(ZSammel_BuchungForm.belgDat.Text, 4)
        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'ZSammel_BuchungForm.belgDat.Text = Jahr.Text & tmpDatum

        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'tmpDatum = VB.Right(ZSammel_BuchungForm.belgDat2.Text, 4)
        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'ZSammel_BuchungForm.belgDat2.Text = Jahr.Text & tmpDatum

        ''ZSetWriteZahlungForm
        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'tmpDatum = VB.Right(ZSetWriteZahlungForm.ValutaDatum.Text, 4)
        ''UPGRADE_WARNING: Couldn't resolve default property of object tmpDatum. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'ZSetWriteZahlungForm.ValutaDatum.Text = Jahr.Text & tmpDatum

    End Sub

    Function BoolToString(ByRef ok As Boolean) As String
        'UPGRADE_NOTE: str was upgraded to str_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        Dim str_Renamed As String
        If (ok = True) Then
            str_Renamed = " Ja "
        Else
            str_Renamed = "Nein"
        End If
        BoolToString = str_Renamed
    End Function


    'Private Sub AXFinanzCommand_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles AXFinanzCommand.Click
    '    AXFinanzForm.Show()
    '    Me.Hide()
    'End Sub

    'Private Sub AXiAdrCommand_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles AXiAdrCommand.Click
    '    AXiAdrForm.Show()
    '    Me.Hide()
    'End Sub

    'Private Sub AXiBeBuCommand_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles AXiBeBuCommand.Click
    '    AXiBeBuForm.Show()
    '    Me.Hide()
    'End Sub

    'Private Sub AXiBSExtCommand_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles AXiBSExtCommand.Click
    '    AXiBSExtForm.Show()
    '    Me.Hide()
    'End Sub

    'Private Sub AXiDbBhgCommand_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles AXiDbBhgCommand.Click
    '    AXiDbBhgForm.Show()
    '    Me.Hide()
    'End Sub

    'Private Sub AXiFBhgCommand_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles AXiFBhgCommand.Click
    '    AXiFBhgForm.Show()
    '    Me.Hide()
    'End Sub

    'Private Sub AXiKrBhgCommand_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles AXiKrBhgCommand.Click
    '    AXiKrBhgForm.Show()
    '    Me.Hide()
    'End Sub

    'Private Sub AXiPIFinCommand_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles AXiPIFinCommand.Click
    '    AXiPIFinForm.Show()
    '    Me.Hide()
    'End Sub




    'UPGRADE_WARNING: Event Check.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    'Private Sub Check_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles 'Check.CheckStateChanged
    'ZAXiDbBhgForm1.CheckMawisperreButton.Visible = False
    'ZAXiKrBhgForm1.CheckMawisperreButton.Visible = False

    'AXFinanzForm.cmdCheckLizenz.Visible = False
    'AXFinanzForm.Applikation.Visible = False
    'AXFinanzForm.Modul.Visible = False

    'If check.CheckState = 1 Then
    '    ZAXiKrBhgForm1Zusatz.SetHypArchButton.Enabled = True

    '    AXFinanzForm.cmdCheckLizenz.Visible = True
    '    AXFinanzForm.Applikation.Visible = True
    '    AXFinanzForm.Modul.Visible = True


    '    If strShowHiddenMethods = "1" Then
    '        ZAXiDbBhgForm1.CheckMawisperreButton.Visible = True
    '        ZAXiKrBhgForm1.CheckMawisperreButton.Visible = True
    '    End If
    '    ZAXiKrBhgForm1Zusatz.SetHypArchButton.Enabled = False
    'Else
    'End If
    'End Sub

    'UPGRADE_WARNING: Event bCountFields.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    'Private Sub bCountFields_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles bCountFields.CheckStateChanged
    'gCountFields = False
    'If bCountFields.CheckState = 1 Then
    '    gCountFields = True
    'End If
    'End Sub

    '    Private Sub Connect_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Connect.Click
    '        Dim MSG As Object

    '        On Error Resume Next
    '        Dim a As Object
    '        Dim b As Single

    '        'Finanz.ShowODBDDialogByNTSecurity = cbShowODBDDialogByNTSecurity.CheckState
    '        'Call Finanz.ConnectSBSdb(dsn.Text, dbName.Text, uid.Text, pwd.Text, locale.Text, Applikationsname.Text)
    '        'AXFinanzCommand.Enabled = True

    '        b = Err.Number And 65535

    '        If b = 0 Then GoTo isOk
    '        '        AXFinanzCommand.Enabled = False
    '        'UPGRADE_WARNING: Couldn't resolve default property of object MSG. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        MSG = "Error # " & Str(Err.Number) & " was generated by " & Err.Source & Chr(13) & Chr(10) & Err.Description & " Unsere Fehlernummer" & Str(b)
    '        Err.Clear()

    '        MsgBox(MSG)
    '        Exit Sub
    '        End
    'isOk:
    '        'AXFinanzCommand.Enabled = True
    '    End Sub

    '    Private Sub ConnectNoPromptButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ConnectNoPromptButton.Click
    '        Dim MSG As Object

    '        On Error Resume Next
    '        Dim a As Object
    '        Dim b As Single


    '        Call Finanz.ConnectSBSdbNoPrompt(dsn.Text, dbName.Text, uid.Text, pwd.Text, locale.Text, Applikationsname.Text)
    '        'AXFinanzCommand.Enabled = True

    '        b = Err.Number And 65535

    '        If b = 0 Then GoTo isOk
    '        '        AXFinanzCommand.Enabled = False
    '        'UPGRADE_WARNING: Couldn't resolve default property of object MSG. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        MSG = "Error # " & Str(Err.Number) & " was generated by " & Err.Source & Chr(13) & Chr(10) & Err.Description & " Unsere Fehlernummer" & Str(b)
    '        Err.Clear()

    '        MsgBox(MSG)
    '        Exit Sub
    '        End
    'isOk:
    '        'AXFinanzCommand.Enabled = True
    '    End Sub

    'Private Sub Exit_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Exit_Renamed.Click
    '    Call InitVar()
    '    End
    'End Sub

    '    Private Sub Main_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
    '        Dim MSG As Object
    '        Dim myerr As Object
    '        Dim Pfad As Object
    '        'bCountFields_CheckStateChanged(bCountFields, New System.EventArgs())
    '        'ChangeColor(Me)

    '        DidOpenmandant = False
    '        'Call Check_CheckStateChanged(Check, New System.EventArgs())
    '        On Error GoTo ErrorHandler
    '        strShowHiddenMethods = "0"
    '        'Call Check_CheckStateChanged(Check, New System.EventArgs())
    '        Dim m_str As String
    '        Dim errPos As Short
    '        errPos = 0

    '        Call InitVar()

    '        FELD_SEP = "{<}"
    '        REC_SEP = "{>}"
    '        KSTKTR_SEP = "{-}"

    '        FELD_SEP_OUT = "{>}"
    '        REC_SEP_OUT = "{<}"

    '        'AXFinanzCommand.Enabled = False
    '        'AXiAdrCommand.Enabled = False
    '        'AXiBeBuCommand.Enabled = False
    '        'AXiBSExtCommand.Enabled = False
    '        'AXiDbBhgCommand.Enabled = False
    '        'AXiFBhgCommand.Enabled = False
    '        'AXiKrBhgCommand.Enabled = False
    '        'AXiPIFinCommand.Enabled = False

    '        If (Environment.Is64BitProcess) Then
    '            'Me.Text += " - (64-Bit Testprogram)"
    '        Else
    '            'Me.Text += " - (32-Bit Testprogram)"
    '        End If



    '        'UPGRADE_NOTE: Object Finanz may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    '        Finanz = Nothing
    '        Finanz = New SBSXASLib.AXFinanz
    '        If Finanz.Is64Bit Then
    '            'lblVersion.Text += vbCrLf + "64-Bit"
    '        Else
    '            'lblVersion.Text += vbCrLf + "32-Bit"
    '        End If
    '        'gVersion = Finanz.GetVersion

    '        m_str = ""

    '        Adr = Finanz.GetAdrObj
    '        m_str = m_str & "GetAdrObj: Ja" & ", "
    '        BeBu = Finanz.GetBeBuObj
    '        m_str = m_str & "GetBeBuObj: Ja" & ", "
    '        BsExt = Finanz.GetBSExtensionObj
    '        m_str = m_str & "GetBSExtensionObj: Ja" & ", "
    '        DbBhg = Finanz.GetDebiObj
    '        m_str = m_str & "GetDebiObj : Ja" & ", "
    '        FBhg = Finanz.GetFibuObj
    '        m_str = m_str & "GetFibuObj : Ja" & ", "
    '        KrBhg = Finanz.GetKrediObj
    '        m_str = m_str & "GetKrediObj: Ja" & ", "
    '        PIFin = FBhg.GetCheckObj
    '        m_str = m_str & "GetCheckObj: Ja" & ", "

    '        'Ausgabefeld.Text = m_str

    '        errPos = 1
    '        Pfad = Nothing
    '        Dim Zeile As Object
    '        ''UPGRADE_WARNING: Couldn't resolve default property of object Pfad. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        'FileOpen(1, Pfad & "SBSACTIVEX.ini", OpenMode.Input)
    '        'Zeile = LineInput(1)
    '        ''UPGRADE_WARNING: Couldn't resolve default property of object Zeile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        'dsn.Text = Zeile
    '        'Zeile = LineInput(1)
    '        ''UPGRADE_WARNING: Couldn't resolve default property of object Zeile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        'dbName.Text = Zeile
    '        'Zeile = LineInput(1)
    '        ''UPGRADE_WARNING: Couldn't resolve default property of object Zeile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        'uid.Text = Zeile
    '        'Zeile = LineInput(1)
    '        ''UPGRADE_WARNING: Couldn't resolve default property of object Zeile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        'pwd.Text = Zeile
    '        'Zeile = LineInput(1)
    '        ''UPGRADE_WARNING: Couldn't resolve default property of object Zeile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        'locale.Text = Zeile
    '        'Zeile = LineInput(1)
    '        ''UPGRADE_WARNING: Couldn't resolve default property of object Zeile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        'Applikationsname.Text = Zeile
    '        'Zeile = LineInput(1)
    '        ''UPGRADE_WARNING: Couldn't resolve default property of object Zeile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        'AXFinanzForm.mandantID.Text = Zeile
    '        'Zeile = LineInput(1)
    '        ''UPGRADE_WARNING: Couldn't resolve default property of object Zeile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        'AXFinanzForm.Periode.Text = Zeile
    '        'Zeile = LineInput(1)
    '        ''UPGRADE_WARNING: Couldn't resolve default property of object Zeile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        'Jahr.Text = Zeile
    '        'Call changeDate()
    '        'Zeile = LineInput(1)
    '        ''UPGRADE_WARNING: Couldn't resolve default property of object Zeile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        'strShowHiddenMethods = Zeile
    '        'FileClose(1)
    'ende:

    '        Exit Sub

    'ErrorHandler:

    '        If errPos = 1 Then
    '            FileClose(1)
    '            Exit Sub
    '        End If

    '        If Err.Number <> 0 Then
    '            If Err.Number = 62 Then
    '                Err.Clear()
    '                GoTo ende
    '            End If

    '            If Err.Number = 53 Then
    '                Err.Clear()
    '                GoTo ende
    '            End If

    '            'UPGRADE_WARNING: Couldn't resolve default property of object myerr. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '            myerr = Err.Number
    '            'UPGRADE_WARNING: Couldn't resolve default property of object MSG. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '            MSG = "There was an error attempting to open the Automation server!" & Chr(13) & Chr(10) & "Error # " & Err.Number & " " & Err.Description

    '            MsgBox(MSG, , "Programm beendet")
    '            End
    '        End If
    '        On Error GoTo 0

    '    End Sub

    'UPGRADE_NOTE: Form_Terminate was upgraded to Form_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    'UPGRADE_WARNING: Main event Form.Terminate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
    Private Sub Form_Terminate_Renamed()
        'Call Exit_Renamed_Click(Exit_Renamed, New System.EventArgs())
    End Sub

    'Private Sub Main_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    '    'Call Exit_Renamed_Click(Exit_Renamed, New System.EventArgs())
    'End Sub


    '    Private Sub InitSBSxasButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles InitSBSxasButton.Click
    '        Dim MSG As Object
    '        Dim myerr As Object
    '        Dim Pfad As Object
    '        DidOpenmandant = False
    '        'Call Check_CheckStateChanged(Check, New System.EventArgs())
    '        On Error GoTo ErrorHandler

    '        Dim m_str As String
    '        Dim errPos As Short
    '        errPos = 0

    '        'Variablen initialisieren
    '        Call InitVar()

    '        Finanz = New SBSXASLib.AXFinanz()

    '        If Finanz.Is64Bit Then
    '            'lblVersion.Text += vbCrLf + "64-Bit"
    '        Else
    '            'lblVersion.Text += vbCrLf + "32-Bit"
    '        End If


    '        'AXFinanzCommand.Enabled = False
    '        'AXiAdrCommand.Enabled = False
    '        'AXiBeBuCommand.Enabled = False
    '        'AXiBSExtCommand.Enabled = False
    '        'AXiDbBhgCommand.Enabled = False
    '        'AXiFBhgCommand.Enabled = False
    '        'AXiKrBhgCommand.Enabled = False
    '        'AXiPIFinCommand.Enabled = False

    '        m_str = ""

    '        Adr = Finanz.GetAdrObj
    '        m_str = m_str & "GetAdrObj        : Ja" & ", "
    '        BeBu = Finanz.GetBeBuObj
    '        m_str = m_str & "GetBeBuObj       : Ja" & ", "
    '        BsExt = Finanz.GetBSExtensionObj
    '        m_str = m_str & "GetBSExtensionObj: Ja" & ", "
    '        DbBhg = Finanz.GetDebiObj
    '        m_str = m_str & "GetDebiObj : Ja" & ", "
    '        FBhg = Finanz.GetFibuObj
    '        m_str = m_str & "GetFibuObj : Ja" & ", "
    '        KrBhg = Finanz.GetKrediObj
    '        m_str = m_str & "GetKrediObj: Ja" & ", "
    '        PIFin = FBhg.GetCheckObj
    '        m_str = m_str & "GetCheckObj: Ja" & ", "


    '        'Ausgabefeld.Text = m_str

    '        'AXFinanzForm.SetDelimiters2Button_Click(Nothing, New System.EventArgs())
    '        'AXFinanzForm.SetOutDelimitersButton_Click(Nothing, New System.EventArgs())

    '        errPos = 1
    '        Pfad = Nothing

    '        Dim Zeile As Object
    '        ''UPGRADE_WARNING: Couldn't resolve default property of object Pfad. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        'FileOpen(1, Pfad & "SBSACTIVEX.ini", OpenMode.Input)
    '        'Zeile = LineInput(1)
    '        ''UPGRADE_WARNING: Couldn't resolve default property of object Zeile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        'dsn.Text = Zeile
    '        'Zeile = LineInput(1)
    '        ''UPGRADE_WARNING: Couldn't resolve default property of object Zeile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        'dbName.Text = Zeile
    '        'Zeile = LineInput(1)
    '        ''UPGRADE_WARNING: Couldn't resolve default property of object Zeile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        'uid.Text = Zeile
    '        'Zeile = LineInput(1)
    '        ''UPGRADE_WARNING: Couldn't resolve default property of object Zeile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        'pwd.Text = Zeile
    '        'Zeile = LineInput(1)
    '        ''UPGRADE_WARNING: Couldn't resolve default property of object Zeile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        'locale.Text = Zeile
    '        'Zeile = LineInput(1)
    '        ''UPGRADE_WARNING: Couldn't resolve default property of object Zeile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        'Applikationsname.Text = Zeile
    '        'Zeile = LineInput(1)
    '        ''UPGRADE_WARNING: Couldn't resolve default property of object Zeile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        'AXFinanzForm.mandantID.Text = Zeile
    '        'Zeile = LineInput(1)
    '        ''UPGRADE_WARNING: Couldn't resolve default property of object Zeile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        'AXFinanzForm.Periode.Text = Zeile
    '        'Zeile = LineInput(1)
    '        ''UPGRADE_WARNING: Couldn't resolve default property of object Zeile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        'Jahr.Text = Zeile
    '        'FileClose(1)

    '        Call changeDate()

    'ende:

    '        Exit Sub

    'ErrorHandler:

    '        If errPos = 1 Then
    '            FileClose(1)
    '            Exit Sub
    '        End If

    '        If Err.Number <> 0 Then
    '            If Err.Number = 62 Then
    '                Err.Clear()
    '                GoTo ende
    '            End If

    '            If Err.Number = 53 Then
    '                Err.Clear()
    '                GoTo ende
    '            End If

    '            'UPGRADE_WARNING: Couldn't resolve default property of object myerr. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '            myerr = Err.Number
    '            'UPGRADE_WARNING: Couldn't resolve default property of object MSG. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '            MSG = "There was an error attempting to open the Automation server!"
    '            MsgBox(MSG, , "Programm beendet")
    '            End
    '        End If
    '        On Error GoTo 0


    '    End Sub



    '    Private Sub mnuGetFields_Click()
    '        Dim Pfad As Object
    '        On Error GoTo ErrorHandler
    '        Dim Zeile As String

    '        Pfad = Nothing

    '        'UPGRADE_WARNING: Couldn't resolve default property of object Pfad. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        FileOpen(1, Pfad & "Ausgabe.txt", OpenMode.Input)
    '        Zeile = LineInput(1) 'Immer Titel
    '        Zeile = LineInput(1)

    '        FileClose(1)

    '        Dim anz As Integer
    '        anz = 0
    '        Do While 1
    '            If Zeile = "" Then
    '                Exit Do
    '            End If
    '            'If InStr(Zeile, AXFinanzForm.feld2.Text) = 0 Then
    '            '    Exit Do
    '            'End If

    '            'Call headFromList(Zeile, (AXFinanzForm.feld2).Text)
    '            anz = anz + 1
    '        Loop

    '        MsgBox("Anzahl Felder: " & Str(anz))
    '        Exit Sub

    'ErrorHandler:
    '        Err.Clear()
    '    End Sub

    '    Private Sub save_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles save.Click
    '        Dim Pfad As Object
    '        On Error GoTo ErrorHandler

    '        Pfad = Nothing

    '        'UPGRADE_WARNING: Couldn't resolve default property of object Pfad. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '        FileOpen(1, Pfad & "SBSACTIVEX.ini", OpenMode.Output)
    '        'PrintLine(1, dsn.Text)
    '        'PrintLine(1, dbName.Text)
    '        'PrintLine(1, uid.Text)
    '        'PrintLine(1, pwd.Text)
    '        'PrintLine(1, locale.Text)
    '        'PrintLine(1, Applikationsname.Text)
    '        'PrintLine(1, AXFinanzForm.mandantID.Text)
    '        'PrintLine(1, AXFinanzForm.Periode.Text)
    '        'PrintLine(1, Jahr.Text)
    '        'PrintLine(1, strShowHiddenMethods)
    '        FileClose(1)

    '        Call changeDate()

    '        Exit Sub

    'ErrorHandler:
    '        If Err.Number <> 0 Then
    '            MsgBox("Die Daten konnten nicht gesichert werden", MsgBoxStyle.Information, "Achtung")
    '        End If

    '    End Sub

    'Private Sub Applikationsname_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Applikationsname.Enter
    '    'SelectAll(Applikationsname)
    'End Sub

    'Private Sub dbName_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dbName.Enter
    '    'SelectAll(dbName)
    'End Sub

    'Private Sub dsn_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dsn.Enter
    '    'SelectAll(dsn)
    'End Sub
    'Private Sub Jahr_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Jahr.Enter
    '    'SelectAll(Jahr)
    'End Sub

    'Private Sub locale_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles locale.Enter
    '    'SelectAll(locale)
    'End Sub

    'Private Sub pwd_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles pwd.Enter
    '    'SelectAll(pwd)
    'End Sub

    'Private Sub uid_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles uid.Enter
    '    'SelectAll(uid)
    'End Sub

End Class
