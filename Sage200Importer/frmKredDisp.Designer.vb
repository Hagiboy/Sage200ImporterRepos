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
        CType(Me.dgvInfo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvBookingSub, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvBookings, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dsKreditoren, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvInfo
        '
        Me.dgvInfo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvInfo.Location = New System.Drawing.Point(1021, 2)
        Me.dgvInfo.Name = "dgvInfo"
        Me.dgvInfo.Size = New System.Drawing.Size(391, 119)
        Me.dgvInfo.TabIndex = 5
        '
        'dgvBookingSub
        '
        Me.dgvBookingSub.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvBookingSub.Location = New System.Drawing.Point(12, 2)
        Me.dgvBookingSub.Name = "dgvBookingSub"
        Me.dgvBookingSub.Size = New System.Drawing.Size(1003, 119)
        Me.dgvBookingSub.TabIndex = 4
        '
        'dgvBookings
        '
        Me.dgvBookings.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvBookings.Location = New System.Drawing.Point(12, 127)
        Me.dgvBookings.Name = "dgvBookings"
        Me.dgvBookings.Size = New System.Drawing.Size(1553, 479)
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
        Me.butImport.Location = New System.Drawing.Point(1447, 51)
        Me.butImport.Name = "butImport"
        Me.butImport.Size = New System.Drawing.Size(86, 53)
        Me.butImport.TabIndex = 6
        Me.butImport.Text = "Import"
        Me.butImport.UseVisualStyleBackColor = True
        '
        'txtNumber
        '
        Me.txtNumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNumber.Location = New System.Drawing.Point(1460, 16)
        Me.txtNumber.Name = "txtNumber"
        Me.txtNumber.Size = New System.Drawing.Size(60, 29)
        Me.txtNumber.TabIndex = 10
        '
        'frmKredDisp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1577, 609)
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
End Class
