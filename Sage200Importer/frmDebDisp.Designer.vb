<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDebDisp
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDebDisp))
        Me.dgvBookings = New System.Windows.Forms.DataGridView()
        Me.dgvBookingSub = New System.Windows.Forms.DataGridView()
        Me.dgvInfo = New System.Windows.Forms.DataGridView()
        Me.MySQLdaDebitoren = New MySqlConnector.MySqlDataAdapter()
        Me.mysqlcmdDebDel = New MySqlConnector.MySqlCommand()
        Me.mysqlconn = New MySqlConnector.MySqlConnection()
        Me.MySqlcmdDebIns = New MySqlConnector.MySqlCommand()
        Me.mysqlcmdDebRead = New MySqlConnector.MySqlCommand()
        Me.mysqlcmdDebUpd = New MySqlConnector.MySqlCommand()
        Me.MySQLdaDebitorenSub = New MySqlConnector.MySqlDataAdapter()
        Me.mysqlcmdDebSubDel = New MySqlConnector.MySqlCommand()
        Me.mysqlcmdDebSubRead = New MySqlConnector.MySqlCommand()
        Me.dsDebitoren = New System.Data.DataSet()
        Me.butImport = New System.Windows.Forms.Button()
        Me.BgWLoadDebi = New System.ComponentModel.BackgroundWorker()
        Me.dgvDates = New System.Windows.Forms.DataGridView()
        Me.BgWCheckDebi = New System.ComponentModel.BackgroundWorker()
        Me.BgWImportDebi = New System.ComponentModel.BackgroundWorker()
        Me.butCheckDeb = New System.Windows.Forms.Button()
        Me.lstBoxPerioden = New System.Windows.Forms.ListBox()
        Me.TSDebis = New System.Windows.Forms.ToolStrip()
        Me.ButDeselect = New System.Windows.Forms.ToolStripButton()
        Me.PRKredi = New System.Windows.Forms.ToolStripProgressBar()
        Me.TSLblDebType = New System.Windows.Forms.ToolStripLabel()
        Me.TSLblNmbr = New System.Windows.Forms.ToolStripLabel()
        CType(Me.dgvBookings, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvBookingSub, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvInfo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dsDebitoren, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvDates, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TSDebis.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvBookings
        '
        Me.dgvBookings.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvBookings.Location = New System.Drawing.Point(11, 137)
        Me.dgvBookings.Name = "dgvBookings"
        Me.dgvBookings.Size = New System.Drawing.Size(1571, 479)
        Me.dgvBookings.TabIndex = 0
        '
        'dgvBookingSub
        '
        Me.dgvBookingSub.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvBookingSub.Location = New System.Drawing.Point(12, 12)
        Me.dgvBookingSub.Name = "dgvBookingSub"
        Me.dgvBookingSub.Size = New System.Drawing.Size(830, 119)
        Me.dgvBookingSub.TabIndex = 1
        '
        'dgvInfo
        '
        Me.dgvInfo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvInfo.Location = New System.Drawing.Point(848, 12)
        Me.dgvInfo.Name = "dgvInfo"
        Me.dgvInfo.Size = New System.Drawing.Size(322, 119)
        Me.dgvInfo.TabIndex = 2
        '
        'MySQLdaDebitoren
        '
        Me.MySQLdaDebitoren.AcceptChangesDuringFill = False
        Me.MySQLdaDebitoren.AcceptChangesDuringUpdate = False
        Me.MySQLdaDebitoren.DeleteCommand = Me.mysqlcmdDebDel
        Me.MySQLdaDebitoren.InsertCommand = Me.MySqlcmdDebIns
        Me.MySQLdaDebitoren.SelectCommand = Me.mysqlcmdDebRead
        Me.MySQLdaDebitoren.UpdateBatchSize = 0
        Me.MySQLdaDebitoren.UpdateCommand = Me.mysqlcmdDebUpd
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
        'MySqlcmdDebIns
        '
        Me.MySqlcmdDebIns.CommandTimeout = 30
        Me.MySqlcmdDebIns.Connection = Me.mysqlconn
        Me.MySqlcmdDebIns.Transaction = Nothing
        Me.MySqlcmdDebIns.UpdatedRowSource = System.Data.UpdateRowSource.None
        '
        'mysqlcmdDebRead
        '
        Me.mysqlcmdDebRead.CommandTimeout = 30
        Me.mysqlcmdDebRead.Connection = Me.mysqlconn
        Me.mysqlcmdDebRead.Transaction = Nothing
        Me.mysqlcmdDebRead.UpdatedRowSource = System.Data.UpdateRowSource.None
        '
        'mysqlcmdDebUpd
        '
        Me.mysqlcmdDebUpd.CommandTimeout = 30
        Me.mysqlcmdDebUpd.Connection = Me.mysqlconn
        Me.mysqlcmdDebUpd.Transaction = Nothing
        Me.mysqlcmdDebUpd.UpdatedRowSource = System.Data.UpdateRowSource.None
        '
        'MySQLdaDebitorenSub
        '
        Me.MySQLdaDebitorenSub.DeleteCommand = Me.mysqlcmdDebSubDel
        Me.MySQLdaDebitorenSub.InsertCommand = Nothing
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
        'mysqlcmdDebSubRead
        '
        Me.mysqlcmdDebSubRead.CommandTimeout = 30
        Me.mysqlcmdDebSubRead.Connection = Me.mysqlconn
        Me.mysqlcmdDebSubRead.Transaction = Nothing
        Me.mysqlcmdDebSubRead.UpdatedRowSource = System.Data.UpdateRowSource.None
        '
        'dsDebitoren
        '
        Me.dsDebitoren.DataSetName = "dsDebiitorenHead"
        Me.dsDebitoren.EnforceConstraints = False
        '
        'butImport
        '
        Me.butImport.Location = New System.Drawing.Point(1483, 79)
        Me.butImport.Name = "butImport"
        Me.butImport.Size = New System.Drawing.Size(86, 39)
        Me.butImport.TabIndex = 3
        Me.butImport.Text = "Import"
        Me.butImport.UseVisualStyleBackColor = True
        '
        'BgWLoadDebi
        '
        '
        'dgvDates
        '
        Me.dgvDates.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDates.Location = New System.Drawing.Point(1176, 12)
        Me.dgvDates.Name = "dgvDates"
        Me.dgvDates.Size = New System.Drawing.Size(292, 118)
        Me.dgvDates.TabIndex = 11
        '
        'BgWCheckDebi
        '
        '
        'BgWImportDebi
        '
        '
        'butCheckDeb
        '
        Me.butCheckDeb.Location = New System.Drawing.Point(1483, 26)
        Me.butCheckDeb.Name = "butCheckDeb"
        Me.butCheckDeb.Size = New System.Drawing.Size(86, 35)
        Me.butCheckDeb.TabIndex = 13
        Me.butCheckDeb.Text = "Check"
        Me.butCheckDeb.UseVisualStyleBackColor = True
        '
        'lstBoxPerioden
        '
        Me.lstBoxPerioden.FormattingEnabled = True
        Me.lstBoxPerioden.Location = New System.Drawing.Point(1477, 14)
        Me.lstBoxPerioden.Name = "lstBoxPerioden"
        Me.lstBoxPerioden.Size = New System.Drawing.Size(32, 30)
        Me.lstBoxPerioden.TabIndex = 14
        Me.lstBoxPerioden.Visible = False
        '
        'TSDebis
        '
        Me.TSDebis.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.TSDebis.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ButDeselect, Me.PRKredi, Me.TSLblDebType, Me.TSLblNmbr})
        Me.TSDebis.Location = New System.Drawing.Point(0, 596)
        Me.TSDebis.Name = "TSDebis"
        Me.TSDebis.Size = New System.Drawing.Size(1581, 27)
        Me.TSDebis.Stretch = True
        Me.TSDebis.TabIndex = 15
        Me.TSDebis.Text = "TSDebis"
        '
        'ButDeselect
        '
        Me.ButDeselect.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ButDeselect.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButDeselect.Image = CType(resources.GetObject("ButDeselect.Image"), System.Drawing.Image)
        Me.ButDeselect.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ButDeselect.Name = "ButDeselect"
        Me.ButDeselect.Size = New System.Drawing.Size(23, 24)
        Me.ButDeselect.Text = "X"
        '
        'PRKredi
        '
        Me.PRKredi.Name = "PRKredi"
        Me.PRKredi.Size = New System.Drawing.Size(600, 24)
        '
        'TSLblDebType
        '
        Me.TSLblDebType.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.TSLblDebType.Name = "TSLblDebType"
        Me.TSLblDebType.Size = New System.Drawing.Size(80, 24)
        Me.TSLblDebType.Text = "TSLblDebType"
        '
        'TSLblNmbr
        '
        Me.TSLblNmbr.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.TSLblNmbr.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TSLblNmbr.Name = "TSLblNmbr"
        Me.TSLblNmbr.Size = New System.Drawing.Size(18, 24)
        Me.TSLblNmbr.Text = "0"
        '
        'frmDebDisp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1581, 623)
        Me.Controls.Add(Me.TSDebis)
        Me.Controls.Add(Me.lstBoxPerioden)
        Me.Controls.Add(Me.butCheckDeb)
        Me.Controls.Add(Me.dgvDates)
        Me.Controls.Add(Me.butImport)
        Me.Controls.Add(Me.dgvInfo)
        Me.Controls.Add(Me.dgvBookingSub)
        Me.Controls.Add(Me.dgvBookings)
        Me.Name = "frmDebDisp"
        Me.Text = "Deb"
        CType(Me.dgvBookings, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvBookingSub, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvInfo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dsDebitoren, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvDates, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TSDebis.ResumeLayout(False)
        Me.TSDebis.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents dgvBookings As DataGridView
    Friend WithEvents dgvBookingSub As DataGridView
    Friend WithEvents dgvInfo As DataGridView
    Friend WithEvents MySQLdaDebitoren As MySqlConnector.MySqlDataAdapter
    Friend WithEvents mysqlcmdDebRead As MySqlConnector.MySqlCommand
    Friend WithEvents mysqlconn As MySqlConnector.MySqlConnection
    Public WithEvents mysqlcmdDebDel As MySqlConnector.MySqlCommand
    Public WithEvents MySQLdaDebitorenSub As MySqlConnector.MySqlDataAdapter
    Public WithEvents mysqlcmdDebSubRead As MySqlConnector.MySqlCommand
    Friend WithEvents mysqlcmdDebSubDel As MySqlConnector.MySqlCommand
    Public WithEvents dsDebitoren As DataSet
    Friend WithEvents butImport As Button
    Friend WithEvents BgWLoadDebi As System.ComponentModel.BackgroundWorker
    Friend WithEvents dgvDates As DataGridView
    Friend WithEvents BgWCheckDebi As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgWImportDebi As System.ComponentModel.BackgroundWorker
    Friend WithEvents mysqlcmdDebUpd As MySqlConnector.MySqlCommand
    Friend WithEvents MySqlcmdDebIns As MySqlConnector.MySqlCommand
    Friend WithEvents butCheckDeb As Button
    Friend WithEvents lstBoxPerioden As ListBox
    Friend WithEvents TSDebis As ToolStrip
    Friend WithEvents ButDeselect As ToolStripButton
    Friend WithEvents PRKredi As ToolStripProgressBar
    Friend WithEvents TSLblDebType As ToolStripLabel
    Friend WithEvents TSLblNmbr As ToolStripLabel
End Class
