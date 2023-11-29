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
        Me.dgvBookings = New System.Windows.Forms.DataGridView()
        Me.dgvBookingSub = New System.Windows.Forms.DataGridView()
        Me.dgvInfo = New System.Windows.Forms.DataGridView()
        Me.MySQLdaDebitoren = New MySqlConnector.MySqlDataAdapter()
        Me.mysqlcmdDebDel = New MySqlConnector.MySqlCommand()
        Me.mysqlconn = New MySqlConnector.MySqlConnection()
        Me.mysqlcmdDebRead = New MySqlConnector.MySqlCommand()
        Me.MySQLdaDebitorenSub = New MySqlConnector.MySqlDataAdapter()
        Me.mysqlcmdDebSubDel = New MySqlConnector.MySqlCommand()
        Me.mysqlcmdDebSubRead = New MySqlConnector.MySqlCommand()
        Me.dsDebitoren = New System.Data.DataSet()
        Me.butImport = New System.Windows.Forms.Button()
        Me.txtNumber = New System.Windows.Forms.TextBox()
        Me.lblDB = New System.Windows.Forms.Label()
        Me.BgWLoadDebi = New System.ComponentModel.BackgroundWorker()
        Me.dgvDates = New System.Windows.Forms.DataGridView()
        CType(Me.dgvBookings, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvBookingSub, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvInfo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dsDebitoren, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvDates, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.butImport.Location = New System.Drawing.Point(1474, 59)
        Me.butImport.Name = "butImport"
        Me.butImport.Size = New System.Drawing.Size(86, 53)
        Me.butImport.TabIndex = 3
        Me.butImport.Text = "Import"
        Me.butImport.UseVisualStyleBackColor = True
        '
        'txtNumber
        '
        Me.txtNumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNumber.Location = New System.Drawing.Point(1489, 24)
        Me.txtNumber.Name = "txtNumber"
        Me.txtNumber.Size = New System.Drawing.Size(60, 29)
        Me.txtNumber.TabIndex = 9
        '
        'lblDB
        '
        Me.lblDB.AutoSize = True
        Me.lblDB.Location = New System.Drawing.Point(1555, 35)
        Me.lblDB.Name = "lblDB"
        Me.lblDB.Size = New System.Drawing.Size(22, 13)
        Me.lblDB.TabIndex = 10
        Me.lblDB.Text = "DB"
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
        'frmDebDisp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1581, 623)
        Me.Controls.Add(Me.dgvDates)
        Me.Controls.Add(Me.lblDB)
        Me.Controls.Add(Me.txtNumber)
        Me.Controls.Add(Me.butImport)
        Me.Controls.Add(Me.dgvInfo)
        Me.Controls.Add(Me.dgvBookingSub)
        Me.Controls.Add(Me.dgvBookings)
        Me.Name = "frmDebDisp"
        Me.Text = "frmDebDisp"
        CType(Me.dgvBookings, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvBookingSub, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvInfo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dsDebitoren, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvDates, System.ComponentModel.ISupportInitialize).EndInit()
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
    Friend WithEvents txtNumber As TextBox
    Friend WithEvents lblDB As Label
    Friend WithEvents BgWLoadDebi As System.ComponentModel.BackgroundWorker
    Friend WithEvents dgvDates As DataGridView
End Class
