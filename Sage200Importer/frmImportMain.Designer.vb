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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmImportMain))
        Me.cmbBuha = New System.Windows.Forms.ComboBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.butKreditoren = New System.Windows.Forms.Button()
        Me.butDblKredis = New System.Windows.Forms.Button()
        Me.butDebitoren = New System.Windows.Forms.Button()
        Me.dgvBookings = New System.Windows.Forms.DataGridView()
        Me.butImport = New System.Windows.Forms.Button()
        Me.dgvBookingSub = New System.Windows.Forms.DataGridView()
        Me.dgvInfo = New System.Windows.Forms.DataGridView()
        Me.txtNumber = New System.Windows.Forms.TextBox()
        Me.butImportK = New System.Windows.Forms.Button()
        Me.cmbPerioden = New System.Windows.Forms.ComboBox()
        Me.lblVersion = New System.Windows.Forms.Label()
        Me.butMail = New System.Windows.Forms.Button()
        Me.butDblDebis = New System.Windows.Forms.Button()
        Me.chkValutaCorrect = New System.Windows.Forms.CheckBox()
        Me.dtpValutaCorrect = New System.Windows.Forms.DateTimePicker()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvBookings, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvBookingSub, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvInfo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmbBuha
        '
        Me.cmbBuha.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbBuha.FormattingEnabled = True
        Me.cmbBuha.Location = New System.Drawing.Point(23, 24)
        Me.cmbBuha.Name = "cmbBuha"
        Me.cmbBuha.Size = New System.Drawing.Size(270, 28)
        Me.cmbBuha.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.butKreditoren)
        Me.GroupBox1.Controls.Add(Me.butDblKredis)
        Me.GroupBox1.Controls.Add(Me.butDebitoren)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 85)
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
        'dgvBookings
        '
        Me.dgvBookings.AllowUserToOrderColumns = True
        Me.dgvBookings.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvBookings.Location = New System.Drawing.Point(12, 158)
        Me.dgvBookings.Name = "dgvBookings"
        Me.dgvBookings.Size = New System.Drawing.Size(1656, 467)
        Me.dgvBookings.TabIndex = 3
        '
        'butImport
        '
        Me.butImport.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.butImport.Location = New System.Drawing.Point(1565, 69)
        Me.butImport.Name = "butImport"
        Me.butImport.Size = New System.Drawing.Size(103, 41)
        Me.butImport.TabIndex = 4
        Me.butImport.Text = "&Import D"
        Me.butImport.UseVisualStyleBackColor = True
        '
        'dgvBookingSub
        '
        Me.dgvBookingSub.AllowUserToOrderColumns = True
        Me.dgvBookingSub.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvBookingSub.DefaultCellStyle = DataGridViewCellStyle1
        Me.dgvBookingSub.Location = New System.Drawing.Point(299, 19)
        Me.dgvBookingSub.Name = "dgvBookingSub"
        Me.dgvBookingSub.Size = New System.Drawing.Size(754, 133)
        Me.dgvBookingSub.TabIndex = 5
        '
        'dgvInfo
        '
        Me.dgvInfo.AllowUserToOrderColumns = True
        Me.dgvInfo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvInfo.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgvInfo.Location = New System.Drawing.Point(1059, 19)
        Me.dgvInfo.Name = "dgvInfo"
        Me.dgvInfo.RowTemplate.Height = 18
        Me.dgvInfo.Size = New System.Drawing.Size(354, 133)
        Me.dgvInfo.TabIndex = 7
        '
        'txtNumber
        '
        Me.txtNumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNumber.Location = New System.Drawing.Point(1608, 21)
        Me.txtNumber.Name = "txtNumber"
        Me.txtNumber.Size = New System.Drawing.Size(60, 29)
        Me.txtNumber.TabIndex = 8
        '
        'butImportK
        '
        Me.butImportK.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.butImportK.Location = New System.Drawing.Point(1565, 111)
        Me.butImportK.Name = "butImportK"
        Me.butImportK.Size = New System.Drawing.Size(103, 41)
        Me.butImportK.TabIndex = 9
        Me.butImportK.Text = "&Import K"
        Me.butImportK.UseVisualStyleBackColor = True
        '
        'cmbPerioden
        '
        Me.cmbPerioden.FormattingEnabled = True
        Me.cmbPerioden.Location = New System.Drawing.Point(23, 58)
        Me.cmbPerioden.Name = "cmbPerioden"
        Me.cmbPerioden.Size = New System.Drawing.Size(92, 21)
        Me.cmbPerioden.TabIndex = 10
        '
        'lblVersion
        '
        Me.lblVersion.AutoSize = True
        Me.lblVersion.Location = New System.Drawing.Point(220, 63)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.Size = New System.Drawing.Size(42, 13)
        Me.lblVersion.TabIndex = 11
        Me.lblVersion.Text = "Version"
        '
        'butMail
        '
        Me.butMail.Location = New System.Drawing.Point(1565, 19)
        Me.butMail.Name = "butMail"
        Me.butMail.Size = New System.Drawing.Size(42, 26)
        Me.butMail.TabIndex = 12
        Me.butMail.Text = "&Mail"
        Me.butMail.UseVisualStyleBackColor = True
        '
        'butDblDebis
        '
        Me.butDblDebis.Location = New System.Drawing.Point(12, 85)
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
        Me.GroupBox2.Location = New System.Drawing.Point(1419, 18)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(140, 133)
        Me.GroupBox2.TabIndex = 17
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Optionen"
        '
        'frmImportMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(1672, 637)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.butDblDebis)
        Me.Controls.Add(Me.butMail)
        Me.Controls.Add(Me.lblVersion)
        Me.Controls.Add(Me.cmbPerioden)
        Me.Controls.Add(Me.butImportK)
        Me.Controls.Add(Me.txtNumber)
        Me.Controls.Add(Me.dgvInfo)
        Me.Controls.Add(Me.dgvBookingSub)
        Me.Controls.Add(Me.butImport)
        Me.Controls.Add(Me.dgvBookings)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cmbBuha)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmImportMain"
        Me.Text = "Sage200 - Importer"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.dgvBookings, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvBookingSub, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvInfo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cmbBuha As ComboBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents butKreditoren As Button
    Friend WithEvents butDebitoren As Button
    Friend WithEvents dgvBookings As DataGridView
    Friend WithEvents butImport As Button
    Friend WithEvents dgvBookingSub As DataGridView
    Friend WithEvents dgvInfo As DataGridView
    Friend WithEvents txtNumber As TextBox
    Friend WithEvents butImportK As Button
    Friend WithEvents cmbPerioden As ComboBox
    Friend WithEvents lblVersion As Label
    Friend WithEvents butMail As Button
    Friend WithEvents butDblDebis As Button
    Friend WithEvents butDblKredis As Button
    Friend WithEvents chkValutaCorrect As CheckBox
    Friend WithEvents dtpValutaCorrect As DateTimePicker
    Friend WithEvents GroupBox2 As GroupBox
End Class
