<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class RunSql
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
		Me.dgResult = New System.Windows.Forms.DataGridView()
		Me.txtSql = New System.Windows.Forms.TextBox()
		Me.btnExecute = New System.Windows.Forms.Button()
		Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
		Me.txtLog = New System.Windows.Forms.TextBox()
		Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
		Me.txtPassword = New System.Windows.Forms.TextBox()
		Me.txtFolder = New System.Windows.Forms.TextBox()
		Me.lblStatus = New System.Windows.Forms.Label()
		Me.Splitter1 = New System.Windows.Forms.Splitter()
		CType(Me.dgResult, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SplitContainer1.Panel1.SuspendLayout()
		Me.SplitContainer1.Panel2.SuspendLayout()
		Me.SplitContainer1.SuspendLayout()
		Me.SuspendLayout()
		'
		'dgResult
		'
		Me.dgResult.AllowUserToAddRows = False
		Me.dgResult.AllowUserToDeleteRows = False
		Me.dgResult.AllowUserToOrderColumns = True
		Me.dgResult.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
			Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.dgResult.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.dgResult.Location = New System.Drawing.Point(3, 3)
		Me.dgResult.Name = "dgResult"
		Me.dgResult.Size = New System.Drawing.Size(719, 223)
		Me.dgResult.TabIndex = 21
		'
		'txtSql
		'
		Me.txtSql.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
			Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtSql.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtSql.Location = New System.Drawing.Point(3, 3)
		Me.txtSql.Multiline = True
		Me.txtSql.Name = "txtSql"
		Me.txtSql.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.txtSql.Size = New System.Drawing.Size(719, 157)
		Me.txtSql.TabIndex = 25
		'
		'btnExecute
		'
		Me.btnExecute.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnExecute.Location = New System.Drawing.Point(647, 221)
		Me.btnExecute.Name = "btnExecute"
		Me.btnExecute.Size = New System.Drawing.Size(75, 23)
		Me.btnExecute.TabIndex = 26
		Me.btnExecute.Text = "Execute"
		Me.btnExecute.UseVisualStyleBackColor = True
		'
		'txtLog
		'
		Me.txtLog.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtLog.Location = New System.Drawing.Point(3, 196)
		Me.txtLog.Multiline = True
		Me.txtLog.Name = "txtLog"
		Me.txtLog.Size = New System.Drawing.Size(638, 48)
		Me.txtLog.TabIndex = 27
		Me.txtLog.Visible = False
		'
		'SplitContainer1
		'
		Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
			Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.SplitContainer1.Location = New System.Drawing.Point(2, 0)
		Me.SplitContainer1.Name = "SplitContainer1"
		Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
		'
		'SplitContainer1.Panel1
		'
		Me.SplitContainer1.Panel1.Controls.Add(Me.txtPassword)
		Me.SplitContainer1.Panel1.Controls.Add(Me.txtFolder)
		Me.SplitContainer1.Panel1.Controls.Add(Me.lblStatus)
		Me.SplitContainer1.Panel1.Controls.Add(Me.txtSql)
		Me.SplitContainer1.Panel1.Controls.Add(Me.txtLog)
		Me.SplitContainer1.Panel1.Controls.Add(Me.btnExecute)
		'
		'SplitContainer1.Panel2
		'
		Me.SplitContainer1.Panel2.Controls.Add(Me.dgResult)
		Me.SplitContainer1.Size = New System.Drawing.Size(725, 482)
		Me.SplitContainer1.SplitterDistance = 247
		Me.SplitContainer1.TabIndex = 28
		'
		'txtPassword
		'
		Me.txtPassword.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtPassword.Location = New System.Drawing.Point(622, 170)
		Me.txtPassword.Name = "txtPassword"
		Me.txtPassword.Size = New System.Drawing.Size(100, 20)
		Me.txtPassword.TabIndex = 30
		'
		'txtFolder
		'
		Me.txtFolder.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtFolder.Location = New System.Drawing.Point(3, 170)
		Me.txtFolder.Name = "txtFolder"
		Me.txtFolder.Size = New System.Drawing.Size(613, 20)
		Me.txtFolder.TabIndex = 29
		'
		'lblStatus
		'
		Me.lblStatus.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.lblStatus.AutoSize = True
		Me.lblStatus.Location = New System.Drawing.Point(648, 202)
		Me.lblStatus.Name = "lblStatus"
		Me.lblStatus.Size = New System.Drawing.Size(0, 13)
		Me.lblStatus.TabIndex = 28
		'
		'Splitter1
		'
		Me.Splitter1.Dock = System.Windows.Forms.DockStyle.Bottom
		Me.Splitter1.Location = New System.Drawing.Point(0, 483)
		Me.Splitter1.Name = "Splitter1"
		Me.Splitter1.Size = New System.Drawing.Size(727, 11)
		Me.Splitter1.TabIndex = 29
		Me.Splitter1.TabStop = False
		'
		'RunSql
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(727, 494)
		Me.Controls.Add(Me.Splitter1)
		Me.Controls.Add(Me.SplitContainer1)
		Me.Name = "RunSql"
		Me.Text = "RunSql"
		CType(Me.dgResult, System.ComponentModel.ISupportInitialize).EndInit()
		Me.SplitContainer1.Panel1.ResumeLayout(False)
		Me.SplitContainer1.Panel1.PerformLayout()
		Me.SplitContainer1.Panel2.ResumeLayout(False)
		Me.SplitContainer1.ResumeLayout(False)
		Me.ResumeLayout(False)

	End Sub
	Friend WithEvents dgResult As System.Windows.Forms.DataGridView
	Friend WithEvents txtSql As System.Windows.Forms.TextBox
	Friend WithEvents btnExecute As System.Windows.Forms.Button
	Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
	Friend WithEvents txtLog As System.Windows.Forms.TextBox
	Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
	Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
	Friend WithEvents lblStatus As System.Windows.Forms.Label
	Friend WithEvents txtFolder As System.Windows.Forms.TextBox
	Friend WithEvents txtPassword As System.Windows.Forms.TextBox
End Class
