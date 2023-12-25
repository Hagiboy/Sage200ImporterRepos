<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmKredDisp
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
        Dim DataGridViewCellStyle55 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle56 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle57 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle58 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle59 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle60 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle61 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle62 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle63 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmKredDisp))
        Me.dgvInfo = New System.Windows.Forms.DataGridView()
        Me.dgvBookingSub = New System.Windows.Forms.DataGridView()
        Me.dgvBookings = New System.Windows.Forms.DataGridView()
        Me.MySQLdaKreditoren = New MySqlConnector.MySqlDataAdapter()
        Me.mysqlcmdKredDel = New MySqlConnector.MySqlCommand()
        Me.mysqlconn = New MySqlConnector.MySqlConnection()
        Me.mysqlcmdKredRead = New MySqlConnector.MySqlCommand()
        Me.MySQLdaKreditorenSub = New MySqlConnector.MySqlDataAdapter()
        Me.mysqlcmdKredSubDel = New MySqlConnector.MySqlCommand()
        Me.mysqlcmdKredSubRead = New MySqlConnector.MySqlCommand()
        Me.dsKreditoren = New System.Data.DataSet()
        Me.butImport = New System.Windows.Forms.Button()
        Me.dgvDates = New System.Windows.Forms.DataGridView()
        Me.BgWLoadKredi = New System.ComponentModel.BackgroundWorker()
        Me.BgWCheckKredi = New System.ComponentModel.BackgroundWorker()
        Me.BgWImportKredi = New System.ComponentModel.BackgroundWorker()
        Me.lstBoxPerioden = New System.Windows.Forms.ListBox()
        Me.butCheclLred = New System.Windows.Forms.Button()
        Me.TSKredi = New System.Windows.Forms.ToolStrip()
        Me.ButDeselect = New System.Windows.Forms.ToolStripButton()
        Me.PRKredi = New System.Windows.Forms.ToolStripProgressBar()
        Me.TSLblType = New System.Windows.Forms.ToolStripLabel()
        Me.TSLblNmbr = New System.Windows.Forms.ToolStripLabel()
        CType(Me.dgvInfo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvBookingSub, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvBookings, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dsKreditoren, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvDates, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TSKredi.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvInfo
        '
        DataGridViewCellStyle55.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle55.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle55.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle55.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle55.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle55.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle55.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvInfo.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle55
        Me.dgvInfo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle56.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle56.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle56.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle56.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle56.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle56.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle56.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvInfo.DefaultCellStyle = DataGridViewCellStyle56
        Me.dgvInfo.Location = New System.Drawing.Point(915, 2)
        Me.dgvInfo.Name = "dgvInfo"
        DataGridViewCellStyle57.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle57.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle57.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle57.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle57.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle57.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle57.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvInfo.RowHeadersDefaultCellStyle = DataGridViewCellStyle57
        Me.dgvInfo.Size = New System.Drawing.Size(308, 119)
        Me.dgvInfo.TabIndex = 5
        '
        'dgvBookingSub
        '
        DataGridViewCellStyle58.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle58.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle58.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle58.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle58.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle58.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle58.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvBookingSub.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle58
        Me.dgvBookingSub.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle59.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle59.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle59.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle59.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle59.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle59.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle59.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvBookingSub.DefaultCellStyle = DataGridViewCellStyle59
        Me.dgvBookingSub.Location = New System.Drawing.Point(12, 2)
        Me.dgvBookingSub.Name = "dgvBookingSub"
        DataGridViewCellStyle60.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle60.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle60.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle60.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle60.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle60.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle60.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvBookingSub.RowHeadersDefaultCellStyle = DataGridViewCellStyle60
        Me.dgvBookingSub.Size = New System.Drawing.Size(897, 119)
        Me.dgvBookingSub.TabIndex = 4
        '
        'dgvBookings
        '
        DataGridViewCellStyle61.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle61.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle61.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle61.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle61.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle61.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle61.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvBookings.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle61
        Me.dgvBookings.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle62.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle62.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle62.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle62.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle62.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle62.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle62.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvBookings.DefaultCellStyle = DataGridViewCellStyle62
        Me.dgvBookings.Location = New System.Drawing.Point(12, 127)
        Me.dgvBookings.Name = "dgvBookings"
        DataGridViewCellStyle63.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle63.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle63.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle63.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle63.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle63.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle63.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvBookings.RowHeadersDefaultCellStyle = DataGridViewCellStyle63
        Me.dgvBookings.Size = New System.Drawing.Size(1611, 479)
        Me.dgvBookings.TabIndex = 3
        '
        'MySQLdaKreditoren
        '
        Me.MySQLdaKreditoren.AcceptChangesDuringFill = False
        Me.MySQLdaKreditoren.AcceptChangesDuringUpdate = False
        Me.MySQLdaKreditoren.DeleteCommand = Me.mysqlcmdKredDel
        Me.MySQLdaKreditoren.InsertCommand = Nothing
        Me.MySQLdaKreditoren.SelectCommand = Me.mysqlcmdKredRead
        Me.MySQLdaKreditoren.UpdateBatchSize = 0
        Me.MySQLdaKreditoren.UpdateCommand = Nothing
        '
        'mysqlcmdKredDel
        '
        Me.mysqlcmdKredDel.CommandTimeout = 30
        Me.mysqlcmdKredDel.Connection = Me.mysqlconn
        Me.mysqlcmdKredDel.Transaction = Nothing
        Me.mysqlcmdKredDel.UpdatedRowSource = System.Data.UpdateRowSource.None
        '
        'mysqlconn
        '
        Me.mysqlconn.ProvideClientCertificatesCallback = Nothing
        Me.mysqlconn.ProvidePasswordCallback = Nothing
        Me.mysqlconn.RemoteCertificateValidationCallback = Nothing
        '
        'mysqlcmdKredRead
        '
        Me.mysqlcmdKredRead.CommandTimeout = 30
        Me.mysqlcmdKredRead.Connection = Me.mysqlconn
        Me.mysqlcmdKredRead.Transaction = Nothing
        Me.mysqlcmdKredRead.UpdatedRowSource = System.Data.UpdateRowSource.None
        '
        'MySQLdaKreditorenSub
        '
        Me.MySQLdaKreditorenSub.DeleteCommand = Me.mysqlcmdKredSubDel
        Me.MySQLdaKreditorenSub.InsertCommand = Nothing
        Me.MySQLdaKreditorenSub.SelectCommand = Me.mysqlcmdKredSubRead
        Me.MySQLdaKreditorenSub.UpdateBatchSize = 0
        Me.MySQLdaKreditorenSub.UpdateCommand = Nothing
        '
        'mysqlcmdKredSubDel
        '
        Me.mysqlcmdKredSubDel.CommandTimeout = 30
        Me.mysqlcmdKredSubDel.Connection = Me.mysqlconn
        Me.mysqlcmdKredSubDel.Transaction = Nothing
        Me.mysqlcmdKredSubDel.UpdatedRowSource = System.Data.UpdateRowSource.None
        '
        'mysqlcmdKredSubRead
        '
        Me.mysqlcmdKredSubRead.CommandTimeout = 30
        Me.mysqlcmdKredSubRead.Connection = Me.mysqlconn
        Me.mysqlcmdKredSubRead.Transaction = Nothing
        Me.mysqlcmdKredSubRead.UpdatedRowSource = System.Data.UpdateRowSource.None
        '
        'dsKreditoren
        '
        Me.dsKreditoren.DataSetName = "dsDebiitorenHead"
        Me.dsKreditoren.EnforceConstraints = False
        '
        'butImport
        '
        Me.butImport.Location = New System.Drawing.Point(1529, 65)
        Me.butImport.Name = "butImport"
        Me.butImport.Size = New System.Drawing.Size(84, 44)
        Me.butImport.TabIndex = 6
        Me.butImport.Text = "Import"
        Me.butImport.UseVisualStyleBackColor = True
        '
        'dgvDates
        '
        Me.dgvDates.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDates.Location = New System.Drawing.Point(1229, 1)
        Me.dgvDates.Name = "dgvDates"
        Me.dgvDates.Size = New System.Drawing.Size(292, 120)
        Me.dgvDates.TabIndex = 12
        '
        'BgWLoadKredi
        '
        '
        'BgWCheckKredi
        '
        '
        'BgWImportKredi
        '
        '
        'lstBoxPerioden
        '
        Me.lstBoxPerioden.FormattingEnabled = True
        Me.lstBoxPerioden.Location = New System.Drawing.Point(1529, 5)
        Me.lstBoxPerioden.Name = "lstBoxPerioden"
        Me.lstBoxPerioden.Size = New System.Drawing.Size(25, 30)
        Me.lstBoxPerioden.TabIndex = 14
        Me.lstBoxPerioden.Visible = False
        '
        'butCheclLred
        '
        Me.butCheclLred.Location = New System.Drawing.Point(1529, 21)
        Me.butCheclLred.Name = "butCheclLred"
        Me.butCheclLred.Size = New System.Drawing.Size(84, 38)
        Me.butCheclLred.TabIndex = 15
        Me.butCheclLred.Text = "Check"
        Me.butCheclLred.UseVisualStyleBackColor = True
        '
        'TSKredi
        '
        Me.TSKredi.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.TSKredi.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ButDeselect, Me.PRKredi, Me.TSLblNmbr, Me.TSLblType})
        Me.TSKredi.Location = New System.Drawing.Point(0, 606)
        Me.TSKredi.Name = "TSKredi"
        Me.TSKredi.Size = New System.Drawing.Size(1628, 28)
        Me.TSKredi.Stretch = True
        Me.TSKredi.TabIndex = 16
        Me.TSKredi.Text = "TSKredi"
        '
        'ButDeselect
        '
        Me.ButDeselect.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ButDeselect.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButDeselect.Image = CType(resources.GetObject("ButDeselect.Image"), System.Drawing.Image)
        Me.ButDeselect.ImageAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.ButDeselect.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ButDeselect.Name = "ButDeselect"
        Me.ButDeselect.Size = New System.Drawing.Size(24, 25)
        Me.ButDeselect.Text = "X"
        '
        'PRKredi
        '
        Me.PRKredi.Name = "PRKredi"
        Me.PRKredi.Size = New System.Drawing.Size(600, 25)
        '
        'TSLblType
        '
        Me.TSLblType.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.TSLblType.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.TSLblType.Name = "TSLblType"
        Me.TSLblType.Size = New System.Drawing.Size(59, 25)
        Me.TSLblType.Text = "TSLblType"
        '
        'TSLblNmbr
        '
        Me.TSLblNmbr.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.TSLblNmbr.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.TSLblNmbr.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TSLblNmbr.Name = "TSLblNmbr"
        Me.TSLblNmbr.Size = New System.Drawing.Size(18, 25)
        Me.TSLblNmbr.Text = "0"
        '
        'frmKredDisp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1628, 634)
        Me.Controls.Add(Me.TSKredi)
        Me.Controls.Add(Me.butCheclLred)
        Me.Controls.Add(Me.lstBoxPerioden)
        Me.Controls.Add(Me.dgvDates)
        Me.Controls.Add(Me.butImport)
        Me.Controls.Add(Me.dgvInfo)
        Me.Controls.Add(Me.dgvBookingSub)
        Me.Controls.Add(Me.dgvBookings)
        Me.Name = "frmKredDisp"
        Me.Text = "frmKredDisp"
        CType(Me.dgvInfo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvBookingSub, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvBookings, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dsKreditoren, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvDates, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TSKredi.ResumeLayout(False)
        Me.TSKredi.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents dgvInfo As DataGridView
    Friend WithEvents dgvBookingSub As DataGridView
    Friend WithEvents dgvBookings As DataGridView
    Friend WithEvents MySQLdaKreditoren As MySqlConnector.MySqlDataAdapter
    Public WithEvents mysqlcmdKredDel As MySqlConnector.MySqlCommand
    Friend WithEvents mysqlconn As MySqlConnector.MySqlConnection
    Friend WithEvents mysqlcmdKredRead As MySqlConnector.MySqlCommand
    Public WithEvents MySQLdaKreditorenSub As MySqlConnector.MySqlDataAdapter
    Friend WithEvents mysqlcmdKredSubDel As MySqlConnector.MySqlCommand
    Public WithEvents mysqlcmdKredSubRead As MySqlConnector.MySqlCommand
    Public WithEvents dsKreditoren As DataSet
    Friend WithEvents butImport As Button
    Friend WithEvents dgvDates As DataGridView
    Friend WithEvents BgWLoadKredi As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgWCheckKredi As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgWImportKredi As System.ComponentModel.BackgroundWorker
    Friend WithEvents lstBoxPerioden As ListBox
    Friend WithEvents butCheclLred As Button
    Friend WithEvents TSKredi As ToolStrip
    Friend WithEvents ButDeselect As ToolStripButton
    Friend WithEvents PRKredi As ToolStripProgressBar
    Friend WithEvents TSLblType As ToolStripLabel
    Friend WithEvents TSLblNmbr As ToolStripLabel
End Class
