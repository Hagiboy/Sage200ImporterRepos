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
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle13 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle14 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle15 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle16 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle17 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle18 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
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
        Me.txtNumber = New System.Windows.Forms.TextBox()
        Me.lblDB = New System.Windows.Forms.Label()
        Me.dgvDates = New System.Windows.Forms.DataGridView()
        Me.BgWLoadKredi = New System.ComponentModel.BackgroundWorker()
        Me.BgWCheckKredi = New System.ComponentModel.BackgroundWorker()
        Me.BgWImportKredi = New System.ComponentModel.BackgroundWorker()
        Me.butDeSeöect = New System.Windows.Forms.Button()
        CType(Me.dgvInfo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvBookingSub, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvBookings, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dsKreditoren, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvDates, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvInfo
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle10.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle10.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle10.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle10.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle10.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle10.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvInfo.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle10
        Me.dgvInfo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle11.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle11.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle11.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle11.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle11.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvInfo.DefaultCellStyle = DataGridViewCellStyle11
        Me.dgvInfo.Location = New System.Drawing.Point(915, 2)
        Me.dgvInfo.Name = "dgvInfo"
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle12.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle12.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle12.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle12.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle12.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle12.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvInfo.RowHeadersDefaultCellStyle = DataGridViewCellStyle12
        Me.dgvInfo.Size = New System.Drawing.Size(308, 119)
        Me.dgvInfo.TabIndex = 5
        '
        'dgvBookingSub
        '
        DataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle13.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle13.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle13.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle13.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle13.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle13.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvBookingSub.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle13
        Me.dgvBookingSub.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle14.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle14.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle14.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle14.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle14.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle14.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvBookingSub.DefaultCellStyle = DataGridViewCellStyle14
        Me.dgvBookingSub.Location = New System.Drawing.Point(12, 2)
        Me.dgvBookingSub.Name = "dgvBookingSub"
        DataGridViewCellStyle15.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle15.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle15.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle15.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle15.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle15.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle15.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvBookingSub.RowHeadersDefaultCellStyle = DataGridViewCellStyle15
        Me.dgvBookingSub.Size = New System.Drawing.Size(897, 119)
        Me.dgvBookingSub.TabIndex = 4
        '
        'dgvBookings
        '
        DataGridViewCellStyle16.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle16.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle16.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle16.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle16.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle16.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle16.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvBookings.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle16
        Me.dgvBookings.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle17.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle17.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle17.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle17.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle17.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle17.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle17.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvBookings.DefaultCellStyle = DataGridViewCellStyle17
        Me.dgvBookings.Location = New System.Drawing.Point(12, 127)
        Me.dgvBookings.Name = "dgvBookings"
        DataGridViewCellStyle18.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle18.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle18.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle18.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle18.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle18.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle18.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvBookings.RowHeadersDefaultCellStyle = DataGridViewCellStyle18
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
        Me.butImport.Location = New System.Drawing.Point(1527, 40)
        Me.butImport.Name = "butImport"
        Me.butImport.Size = New System.Drawing.Size(86, 53)
        Me.butImport.TabIndex = 6
        Me.butImport.Text = "Import"
        Me.butImport.UseVisualStyleBackColor = True
        '
        'txtNumber
        '
        Me.txtNumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNumber.Location = New System.Drawing.Point(1540, 5)
        Me.txtNumber.Name = "txtNumber"
        Me.txtNumber.Size = New System.Drawing.Size(60, 29)
        Me.txtNumber.TabIndex = 10
        '
        'lblDB
        '
        Me.lblDB.AutoSize = True
        Me.lblDB.Location = New System.Drawing.Point(1606, 16)
        Me.lblDB.Name = "lblDB"
        Me.lblDB.Size = New System.Drawing.Size(22, 13)
        Me.lblDB.TabIndex = 11
        Me.lblDB.Text = "DB"
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
        'butDeSeöect
        '
        Me.butDeSeöect.Location = New System.Drawing.Point(1540, 101)
        Me.butDeSeöect.Name = "butDeSeöect"
        Me.butDeSeöect.Size = New System.Drawing.Size(31, 20)
        Me.butDeSeöect.TabIndex = 13
        Me.butDeSeöect.Text = "X"
        Me.butDeSeöect.UseVisualStyleBackColor = True
        '
        'frmKredDisp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1628, 609)
        Me.Controls.Add(Me.butDeSeöect)
        Me.Controls.Add(Me.dgvDates)
        Me.Controls.Add(Me.lblDB)
        Me.Controls.Add(Me.txtNumber)
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
    Friend WithEvents txtNumber As TextBox
    Friend WithEvents lblDB As Label
    Friend WithEvents dgvDates As DataGridView
    Friend WithEvents BgWLoadKredi As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgWCheckKredi As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgWImportKredi As System.ComponentModel.BackgroundWorker
    Friend WithEvents butDeSeöect As Button
End Class
