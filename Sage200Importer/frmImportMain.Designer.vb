<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmImportMain
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmImportMain))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.butKreditoren = New System.Windows.Forms.Button()
        Me.butDblKredis = New System.Windows.Forms.Button()
        Me.butDebitoren = New System.Windows.Forms.Button()
        Me.butDblDebis = New System.Windows.Forms.Button()
        Me.chkValutaCorrect = New System.Windows.Forms.CheckBox()
        Me.dtpValutaCorrect = New System.Windows.Forms.DateTimePicker()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.dsDebitoren = New System.Data.DataSet()
        Me.MySQLdaDebitoren = New MySqlConnector.MySqlDataAdapter()
        Me.mysqlcmdDebDel = New MySqlConnector.MySqlCommand()
        Me.mysqlconn = New MySqlConnector.MySqlConnection()
        Me.mysqlcmdDebRead = New MySqlConnector.MySqlCommand()
        Me.mysqlcmdDebIns = New MySqlConnector.MySqlCommand()
        Me.mysqlcmbld = New MySqlConnector.MySqlCommandBuilder()
        Me.MySQLdaDebitorenSub = New MySqlConnector.MySqlDataAdapter()
        Me.mysqlcmdDebSubDel = New MySqlConnector.MySqlCommand()
        Me.mysqlcmdDebSubIns = New MySqlConnector.MySqlCommand()
        Me.mysqlcmdDebSubRead = New MySqlConnector.MySqlCommand()
        Me.mysqlcongen = New MySqlConnector.MySqlConnection()
        Me.mysqlcmdgen = New MySqlConnector.MySqlCommand()
        Me.ToolStripContainer1 = New System.Windows.Forms.ToolStripContainer()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.chkValutaEndCorrect = New System.Windows.Forms.CheckBox()
        Me.dtpValutaEndCorrect = New System.Windows.Forms.DateTimePicker()
        Me.lstBoxPerioden = New System.Windows.Forms.ListBox()
        Me.lstBoxMandant = New System.Windows.Forms.ListBox()
        Me.LblTaskID = New System.Windows.Forms.Label()
        Me.LblIdentity = New System.Windows.Forms.Label()
        Me.LblVersion = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.dsDebitoren, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ToolStripContainer1.ContentPanel.SuspendLayout()
        Me.ToolStripContainer1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.butKreditoren)
        Me.GroupBox1.Controls.Add(Me.butDblKredis)
        Me.GroupBox1.Controls.Add(Me.butDebitoren)
        Me.GroupBox1.Controls.Add(Me.butDblDebis)
        Me.GroupBox1.Location = New System.Drawing.Point(458, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(281, 62)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Modus"
        '
        'butKreditoren
        '
        Me.butKreditoren.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.butKreditoren.Location = New System.Drawing.Point(139, 19)
        Me.butKreditoren.Name = "butKreditoren"
        Me.butKreditoren.Size = New System.Drawing.Size(131, 41)
        Me.butKreditoren.TabIndex = 1
        Me.butKreditoren.Text = "&Kreditoren"
        Me.butKreditoren.UseVisualStyleBackColor = True
        '
        'butDblKredis
        '
        Me.butDblKredis.Location = New System.Drawing.Point(51, 0)
        Me.butDblKredis.Name = "butDblKredis"
        Me.butDblKredis.Size = New System.Drawing.Size(41, 22)
        Me.butDblKredis.TabIndex = 14
        Me.butDblKredis.Text = "DK"
        Me.butDblKredis.UseVisualStyleBackColor = True
        Me.butDblKredis.Visible = False
        '
        'butDebitoren
        '
        Me.butDebitoren.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.butDebitoren.Location = New System.Drawing.Point(16, 19)
        Me.butDebitoren.Name = "butDebitoren"
        Me.butDebitoren.Size = New System.Drawing.Size(117, 41)
        Me.butDebitoren.TabIndex = 0
        Me.butDebitoren.Text = "&Debitoren"
        Me.butDebitoren.UseVisualStyleBackColor = True
        '
        'butDblDebis
        '
        Me.butDblDebis.Location = New System.Drawing.Point(98, 0)
        Me.butDblDebis.Name = "butDblDebis"
        Me.butDblDebis.Size = New System.Drawing.Size(41, 22)
        Me.butDblDebis.TabIndex = 13
        Me.butDblDebis.Text = "DD"
        Me.butDblDebis.UseVisualStyleBackColor = True
        Me.butDblDebis.Visible = False
        '
        'chkValutaCorrect
        '
        Me.chkValutaCorrect.AutoSize = True
        Me.chkValutaCorrect.Location = New System.Drawing.Point(6, 19)
        Me.chkValutaCorrect.Name = "chkValutaCorrect"
        Me.chkValutaCorrect.Size = New System.Drawing.Size(118, 17)
        Me.chkValutaCorrect.TabIndex = 15
        Me.chkValutaCorrect.Text = "Valuta-Anpassung?"
        Me.chkValutaCorrect.UseVisualStyleBackColor = True
        '
        'dtpValutaCorrect
        '
        Me.dtpValutaCorrect.Enabled = False
        Me.dtpValutaCorrect.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpValutaCorrect.Location = New System.Drawing.Point(5, 37)
        Me.dtpValutaCorrect.MinDate = New Date(2020, 1, 1, 0, 0, 0, 0)
        Me.dtpValutaCorrect.Name = "dtpValutaCorrect"
        Me.dtpValutaCorrect.Size = New System.Drawing.Size(90, 20)
        Me.dtpValutaCorrect.TabIndex = 16
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.chkValutaCorrect)
        Me.GroupBox2.Controls.Add(Me.dtpValutaCorrect)
        Me.GroupBox2.Location = New System.Drawing.Point(756, 4)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(140, 69)
        Me.GroupBox2.TabIndex = 17
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Optionen Startdatum"
        '
        'dsDebitoren
        '
        Me.dsDebitoren.DataSetName = "dsDebiitorenHead"
        Me.dsDebitoren.EnforceConstraints = False
        '
        'MySQLdaDebitoren
        '
        Me.MySQLdaDebitoren.AcceptChangesDuringFill = False
        Me.MySQLdaDebitoren.AcceptChangesDuringUpdate = False
        Me.MySQLdaDebitoren.DeleteCommand = Me.mysqlcmdDebDel
        Me.MySQLdaDebitoren.InsertCommand = Nothing
        Me.MySQLdaDebitoren.SelectCommand = Me.mysqlcmdDebRead
        Me.MySQLdaDebitoren.UpdateBatchSize = 0
        Me.MySQLdaDebitoren.UpdateCommand = Nothing
        '
        'mysqlcmdDebDel
        '
        Me.mysqlcmdDebDel.CommandTimeout = 30
        Me.mysqlcmdDebDel.Connection = Me.mysqlconn
        Me.mysqlcmdDebDel.Transaction = Nothing
        Me.mysqlcmdDebDel.UpdatedRowSource = System.Data.UpdateRowSource.None
        '
        'mysqlconn
        '
        Me.mysqlconn.ProvideClientCertificatesCallback = Nothing
        Me.mysqlconn.ProvidePasswordCallback = Nothing
        Me.mysqlconn.RemoteCertificateValidationCallback = Nothing
        '
        'mysqlcmdDebRead
        '
        Me.mysqlcmdDebRead.CommandTimeout = 30
        Me.mysqlcmdDebRead.Connection = Me.mysqlconn
        Me.mysqlcmdDebRead.Transaction = Nothing
        Me.mysqlcmdDebRead.UpdatedRowSource = System.Data.UpdateRowSource.None
        '
        'mysqlcmdDebIns
        '
        Me.mysqlcmdDebIns.CommandTimeout = 30
        Me.mysqlcmdDebIns.Connection = Me.mysqlconn
        Me.mysqlcmdDebIns.Transaction = Nothing
        Me.mysqlcmdDebIns.UpdatedRowSource = System.Data.UpdateRowSource.None
        '
        'mysqlcmbld
        '
        Me.mysqlcmbld.DataAdapter = Me.MySQLdaDebitoren
        Me.mysqlcmbld.QuotePrefix = "`"
        Me.mysqlcmbld.QuoteSuffix = "`"
        '
        'MySQLdaDebitorenSub
        '
        Me.MySQLdaDebitorenSub.DeleteCommand = Me.mysqlcmdDebSubDel
        Me.MySQLdaDebitorenSub.InsertCommand = Me.mysqlcmdDebSubIns
        Me.MySQLdaDebitorenSub.SelectCommand = Me.mysqlcmdDebSubRead
        Me.MySQLdaDebitorenSub.UpdateBatchSize = 0
        Me.MySQLdaDebitorenSub.UpdateCommand = Nothing
        '
        'mysqlcmdDebSubDel
        '
        Me.mysqlcmdDebSubDel.CommandTimeout = 30
        Me.mysqlcmdDebSubDel.Connection = Me.mysqlconn
        Me.mysqlcmdDebSubDel.Transaction = Nothing
        Me.mysqlcmdDebSubDel.UpdatedRowSource = System.Data.UpdateRowSource.None
        '
        'mysqlcmdDebSubIns
        '
        Me.mysqlcmdDebSubIns.CommandTimeout = 0
        Me.mysqlcmdDebSubIns.Connection = Me.mysqlconn
        Me.mysqlcmdDebSubIns.Transaction = Nothing
        Me.mysqlcmdDebSubIns.UpdatedRowSource = System.Data.UpdateRowSource.None
        '
        'mysqlcmdDebSubRead
        '
        Me.mysqlcmdDebSubRead.CommandTimeout = 30
        Me.mysqlcmdDebSubRead.Connection = Me.mysqlconn
        Me.mysqlcmdDebSubRead.Transaction = Nothing
        Me.mysqlcmdDebSubRead.UpdatedRowSource = System.Data.UpdateRowSource.None
        '
        'mysqlcongen
        '
        Me.mysqlcongen.ProvideClientCertificatesCallback = Nothing
        Me.mysqlcongen.ProvidePasswordCallback = Nothing
        Me.mysqlcongen.RemoteCertificateValidationCallback = Nothing
        '
        'mysqlcmdgen
        '
        Me.mysqlcmdgen.CommandTimeout = 30
        Me.mysqlcmdgen.Connection = Me.mysqlcongen
        Me.mysqlcmdgen.Transaction = Nothing
        Me.mysqlcmdgen.UpdatedRowSource = System.Data.UpdateRowSource.None
        '
        'ToolStripContainer1
        '
        Me.ToolStripContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        '
        'ToolStripContainer1.ContentPanel
        '
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.GroupBox3)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.lstBoxPerioden)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.lstBoxMandant)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.LblTaskID)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.LblIdentity)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.LblVersion)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.GroupBox2)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.GroupBox1)
        Me.ToolStripContainer1.ContentPanel.Size = New System.Drawing.Size(1688, 80)
        Me.ToolStripContainer1.Location = New System.Drawing.Point(2, 1)
        Me.ToolStripContainer1.Name = "ToolStripContainer1"
        Me.ToolStripContainer1.Size = New System.Drawing.Size(1688, 105)
        Me.ToolStripContainer1.TabIndex = 24
        Me.ToolStripContainer1.Text = "ToolStripContainer1"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.chkValutaEndCorrect)
        Me.GroupBox3.Controls.Add(Me.dtpValutaEndCorrect)
        Me.GroupBox3.Location = New System.Drawing.Point(907, 4)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(140, 69)
        Me.GroupBox3.TabIndex = 23
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Optionen Enddatum"
        '
        'chkValutaEndCorrect
        '
        Me.chkValutaEndCorrect.AutoSize = True
        Me.chkValutaEndCorrect.Location = New System.Drawing.Point(6, 19)
        Me.chkValutaEndCorrect.Name = "chkValutaEndCorrect"
        Me.chkValutaEndCorrect.Size = New System.Drawing.Size(118, 17)
        Me.chkValutaEndCorrect.TabIndex = 15
        Me.chkValutaEndCorrect.Text = "Valuta-Anpassung?"
        Me.chkValutaEndCorrect.UseVisualStyleBackColor = True
        '
        'dtpValutaEndCorrect
        '
        Me.dtpValutaEndCorrect.Enabled = False
        Me.dtpValutaEndCorrect.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpValutaEndCorrect.Location = New System.Drawing.Point(5, 37)
        Me.dtpValutaEndCorrect.MinDate = New Date(2020, 1, 1, 0, 0, 0, 0)
        Me.dtpValutaEndCorrect.Name = "dtpValutaEndCorrect"
        Me.dtpValutaEndCorrect.Size = New System.Drawing.Size(90, 20)
        Me.dtpValutaEndCorrect.TabIndex = 16
        '
        'lstBoxPerioden
        '
        Me.lstBoxPerioden.FormattingEnabled = True
        Me.lstBoxPerioden.Location = New System.Drawing.Point(278, 4)
        Me.lstBoxPerioden.Name = "lstBoxPerioden"
        Me.lstBoxPerioden.Size = New System.Drawing.Size(127, 69)
        Me.lstBoxPerioden.TabIndex = 22
        '
        'lstBoxMandant
        '
        Me.lstBoxMandant.AllowDrop = True
        Me.lstBoxMandant.FormattingEnabled = True
        Me.lstBoxMandant.Location = New System.Drawing.Point(10, 4)
        Me.lstBoxMandant.Name = "lstBoxMandant"
        Me.lstBoxMandant.Size = New System.Drawing.Size(262, 69)
        Me.lstBoxMandant.TabIndex = 21
        '
        'LblTaskID
        '
        Me.LblTaskID.AutoSize = True
        Me.LblTaskID.Location = New System.Drawing.Point(1447, 27)
        Me.LblTaskID.Name = "LblTaskID"
        Me.LblTaskID.Size = New System.Drawing.Size(42, 13)
        Me.LblTaskID.TabIndex = 20
        Me.LblTaskID.Text = "TaskID"
        '
        'LblIdentity
        '
        Me.LblIdentity.AutoSize = True
        Me.LblIdentity.Location = New System.Drawing.Point(1333, 27)
        Me.LblIdentity.Name = "LblIdentity"
        Me.LblIdentity.Size = New System.Drawing.Size(41, 13)
        Me.LblIdentity.TabIndex = 19
        Me.LblIdentity.Text = "Identity"
        '
        'LblVersion
        '
        Me.LblVersion.AutoSize = True
        Me.LblVersion.Location = New System.Drawing.Point(1271, 27)
        Me.LblVersion.Name = "LblVersion"
        Me.LblVersion.Size = New System.Drawing.Size(14, 13)
        Me.LblVersion.TabIndex = 18
        Me.LblVersion.Text = "V"
        '
        'frmImportMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1689, 795)
        Me.Controls.Add(Me.ToolStripContainer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.Name = "frmImportMain"
        Me.Text = "Sage200 - Importer"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.dsDebitoren, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ToolStripContainer1.ContentPanel.ResumeLayout(False)
        Me.ToolStripContainer1.ContentPanel.PerformLayout()
        Me.ToolStripContainer1.ResumeLayout(False)
        Me.ToolStripContainer1.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents butKreditoren As Button
    Friend WithEvents butDebitoren As Button
    Friend WithEvents butDblDebis As Button
    Friend WithEvents butDblKredis As Button
    Friend WithEvents chkValutaCorrect As CheckBox
    Friend WithEvents dtpValutaCorrect As DateTimePicker
    Friend WithEvents GroupBox2 As GroupBox
    Public WithEvents MySQLdaDebitoren As MySqlConnector.MySqlDataAdapter
    Public WithEvents dsDebitoren As DataSet
    Public WithEvents mysqlconn As MySqlConnector.MySqlConnection
    Public WithEvents mysqlcmdDebRead As MySqlConnector.MySqlCommand
    Friend WithEvents mysqlcmbld As MySqlConnector.MySqlCommandBuilder
    Public WithEvents mysqlcmdDebIns As MySqlConnector.MySqlCommand
    Public WithEvents mysqlcmdDebDel As MySqlConnector.MySqlCommand
    Public WithEvents MySQLdaDebitorenSub As MySqlConnector.MySqlDataAdapter
    Public WithEvents mysqlcmdDebSubRead As MySqlConnector.MySqlCommand
    Public WithEvents mysqlcmdDebSubIns As MySqlConnector.MySqlCommand
    Friend WithEvents mysqlcmdDebSubDel As MySqlConnector.MySqlCommand
    Public WithEvents mysqlcongen As MySqlConnector.MySqlConnection
    Public WithEvents mysqlcmdgen As MySqlConnector.MySqlCommand
    Friend WithEvents ToolStripContainer1 As ToolStripContainer
    Friend WithEvents LblTaskID As Label
    Friend WithEvents LblIdentity As Label
    Friend WithEvents LblVersion As Label
    Friend WithEvents lstBoxMandant As ListBox
    Public WithEvents lstBoxPerioden As ListBox
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents chkValutaEndCorrect As CheckBox
    Friend WithEvents dtpValutaEndCorrect As DateTimePicker
End Class
