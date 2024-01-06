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
        Dim DataGridViewCellStyle37 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle38 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle39 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle40 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle41 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle42 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle43 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle44 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle45 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
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
        Me.TSLblNmbr = New System.Windows.Forms.ToolStripLabel()
        Me.TSLblType = New System.Windows.Forms.ToolStripLabel()
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
        DataGridViewCellStyle37.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle37.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle37.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle37.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle37.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle37.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle37.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvInfo.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle37
        Me.dgvInfo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle38.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle38.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle38.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle38.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle38.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle38.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle38.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvInfo.DefaultCellStyle = DataGridViewCellStyle38
        Me.dgvInfo.Location = New System.Drawing.Point(915, 2)
        Me.dgvInfo.Name = "dgvInfo"
        DataGridViewCellStyle39.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle39.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle39.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle39.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle39.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle39.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle39.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvInfo.RowHeadersDefaultCellStyle = DataGridViewCellStyle39
        Me.dgvInfo.Size = New System.Drawing.Size(308, 119)
        Me.dgvInfo.TabIndex = 5
        '
        'dgvBookingSub
        '
        DataGridViewCellStyle40.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle40.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle40.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle40.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle40.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle40.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle40.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvBookingSub.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle40
        Me.dgvBookingSub.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle41.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle41.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle41.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle41.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle41.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle41.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle41.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvBookingSub.DefaultCellStyle = DataGridViewCellStyle41
        Me.dgvBookingSub.Location = New System.Drawing.Point(12, 2)
        Me.dgvBookingSub.Name = "dgvBookingSub"
        DataGridViewCellStyle42.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle42.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle42.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle42.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle42.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle42.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle42.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvBookingSub.RowHeadersDefaultCellStyle = DataGridViewCellStyle42
        Me.dgvBookingSub.Size = New System.Drawing.Size(897, 119)
        Me.dgvBookingSub.TabIndex = 4
        '
        'dgvBookings
        '
        DataGridViewCellStyle43.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle43.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle43.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle43.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle43.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle43.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle43.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvBookings.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle43
        Me.dgvBookings.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle44.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle44.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle44.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle44.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle44.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle44.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle44.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvBookings.DefaultCellStyle = DataGridViewCellStyle44
        Me.dgvBookings.Location = New System.Drawing.Point(12, 127)
        Me.dgvBookings.Name = "dgvBookings"
        DataGridViewCellStyle45.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle45.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle45.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle45.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle45.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle45.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle45.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvBookings.RowHeadersDefaultCellStyle = DataGridViewCellStyle45
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
        Me.butImport.Location = New System.Drawing.Point(1527, 63)
        Me.butImport.Name = "butImport"
        Me.butImport.Size = New System.Drawing.Size(84, 46)
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
        Me.BgWCheckKredi.WorkerReportsProgress = True
        '
        'BgWImportKredi
        '
        Me.BgWImportKredi.WorkerReportsProgress = True
        '
        'lstBoxPerioden
        '
        Me.lstBoxPerioden.FormattingEnabled = True
        Me.lstBoxPerioden.Location = New System.Drawing.Point(1524, 2)
        Me.lstBoxPerioden.Name = "lstBoxPerioden"
        Me.lstBoxPerioden.Size = New System.Drawing.Size(87, 17)
        Me.lstBoxPerioden.TabIndex = 14
        Me.lstBoxPerioden.Visible = False
        '
        'butCheclLred
        '
        Me.butCheclLred.Location = New System.Drawing.Point(1527, 12)
        Me.butCheclLred.Name = "butCheclLred"
        Me.butCheclLred.Size = New System.Drawing.Size(84, 45)
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
        'TSLblNmbr
        '
        Me.TSLblNmbr.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.TSLblNmbr.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.TSLblNmbr.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TSLblNmbr.Name = "TSLblNmbr"
        Me.TSLblNmbr.Size = New System.Drawing.Size(18, 25)
        Me.TSLblNmbr.Text = "0"
        '
        'TSLblType
        '
        Me.TSLblType.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.TSLblType.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.TSLblType.Name = "TSLblType"
        Me.TSLblType.Size = New System.Drawing.Size(59, 25)
        Me.TSLblType.Text = "TSLblType"
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
