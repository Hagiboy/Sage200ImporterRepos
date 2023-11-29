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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
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
        CType(Me.dgvInfo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvBookingSub, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvBookings, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dsKreditoren, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvDates, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvInfo
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvInfo.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvInfo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvInfo.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgvInfo.Location = New System.Drawing.Point(915, 2)
        Me.dgvInfo.Name = "dgvInfo"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvInfo.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvInfo.Size = New System.Drawing.Size(308, 119)
        Me.dgvInfo.TabIndex = 5
        '
        'dgvBookingSub
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvBookingSub.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.dgvBookingSub.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvBookingSub.DefaultCellStyle = DataGridViewCellStyle5
        Me.dgvBookingSub.Location = New System.Drawing.Point(12, 2)
        Me.dgvBookingSub.Name = "dgvBookingSub"
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvBookingSub.RowHeadersDefaultCellStyle = DataGridViewCellStyle6
        Me.dgvBookingSub.Size = New System.Drawing.Size(897, 119)
        Me.dgvBookingSub.TabIndex = 4
        '
        'dgvBookings
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvBookings.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle7
        Me.dgvBookings.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvBookings.DefaultCellStyle = DataGridViewCellStyle8
        Me.dgvBookings.Location = New System.Drawing.Point(12, 127)
        Me.dgvBookings.Name = "dgvBookings"
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle9.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle9.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle9.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvBookings.RowHeadersDefaultCellStyle = DataGridViewCellStyle9
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
        Me.butImport.Location = New System.Drawing.Point(1527, 51)
        Me.butImport.Name = "butImport"
        Me.butImport.Size = New System.Drawing.Size(86, 53)
        Me.butImport.TabIndex = 6
        Me.butImport.Text = "Import"
        Me.butImport.UseVisualStyleBackColor = True
        '
        'txtNumber
        '
        Me.txtNumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNumber.Location = New System.Drawing.Point(1540, 16)
        Me.txtNumber.Name = "txtNumber"
        Me.txtNumber.Size = New System.Drawing.Size(60, 29)
        Me.txtNumber.TabIndex = 10
        '
        'lblDB
        '
        Me.lblDB.AutoSize = True
        Me.lblDB.Location = New System.Drawing.Point(1606, 27)
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
        'frmKredDisp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1628, 609)
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
End Class
