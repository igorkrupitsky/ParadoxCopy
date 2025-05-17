Imports System.Data.OleDb
Imports System.Text.RegularExpressions

Public Class Form1
	Inherits System.Windows.Forms.Form

	Private bStop As Boolean = False

	Dim oAppSetting As New AppSetting
	Friend WithEvents chkHideNotSelected As System.Windows.Forms.CheckBox
	Dim oSelTables As New Hashtable

#Region " Windows Form Designer generated code "

	Public Sub New()
		MyBase.New()

		'This call is required by the Windows Form Designer.
		InitializeComponent()

		'Add any initialization after the InitializeComponent() call

	End Sub

	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
		If disposing Then
			If Not (components Is Nothing) Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(disposing)
	End Sub

	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer

	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.  
	'Do not modify it using the code editor.
	Friend WithEvents btnExport As System.Windows.Forms.Button
	Friend WithEvents btnConnect1 As System.Windows.Forms.Button
	Friend WithEvents txtFolderPath As System.Windows.Forms.TextBox
	Friend WithEvents btnCancel As System.Windows.Forms.Button
	Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
	Friend WithEvents Label1 As System.Windows.Forms.Label
	Friend WithEvents chkCreateTable As System.Windows.Forms.CheckBox
	Friend WithEvents txtLog As System.Windows.Forms.TextBox
	Friend WithEvents chkDeleteData As System.Windows.Forms.CheckBox
	Friend WithEvents Label2 As System.Windows.Forms.Label
	Friend WithEvents txtConnectTo As System.Windows.Forms.TextBox
	Friend WithEvents btnConnect2 As System.Windows.Forms.Button
	Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
	Friend WithEvents dgTables As System.Windows.Forms.DataGridView
	Friend WithEvents btnSQL As System.Windows.Forms.Button
	Friend WithEvents Label3 As System.Windows.Forms.Label
	Friend WithEvents txtPassword As System.Windows.Forms.TextBox
	Friend WithEvents chkDropTable As System.Windows.Forms.CheckBox
	Friend WithEvents chkSQL2008 As System.Windows.Forms.CheckBox
	Friend WithEvents btnCheckAll As System.Windows.Forms.Button
	Friend WithEvents btnUncheckAll As System.Windows.Forms.Button
	Friend WithEvents chkCopyFiles As System.Windows.Forms.CheckBox
	Friend WithEvents ProgressBar2 As System.Windows.Forms.ProgressBar
	Friend WithEvents Panel1 As System.Windows.Forms.Panel
	Friend WithEvents btnCheckNew As System.Windows.Forms.Button
	Friend WithEvents btnCheckDiff As System.Windows.Forms.Button
	Friend WithEvents btnStop As System.Windows.Forms.LinkLabel
	Friend WithEvents btnOpenTempFolder As System.Windows.Forms.Button
	Friend WithEvents lbCount As System.Windows.Forms.Label
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.btnExport = New System.Windows.Forms.Button()
		Me.txtFolderPath = New System.Windows.Forms.TextBox()
		Me.btnConnect1 = New System.Windows.Forms.Button()
		Me.btnCancel = New System.Windows.Forms.Button()
		Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
		Me.Label1 = New System.Windows.Forms.Label()
		Me.chkCreateTable = New System.Windows.Forms.CheckBox()
		Me.txtLog = New System.Windows.Forms.TextBox()
		Me.chkDeleteData = New System.Windows.Forms.CheckBox()
		Me.lbCount = New System.Windows.Forms.Label()
		Me.Label2 = New System.Windows.Forms.Label()
		Me.txtConnectTo = New System.Windows.Forms.TextBox()
		Me.btnConnect2 = New System.Windows.Forms.Button()
		Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
		Me.dgTables = New System.Windows.Forms.DataGridView()
		Me.btnSQL = New System.Windows.Forms.Button()
		Me.Label3 = New System.Windows.Forms.Label()
		Me.txtPassword = New System.Windows.Forms.TextBox()
		Me.chkDropTable = New System.Windows.Forms.CheckBox()
		Me.chkSQL2008 = New System.Windows.Forms.CheckBox()
		Me.btnCheckAll = New System.Windows.Forms.Button()
		Me.btnUncheckAll = New System.Windows.Forms.Button()
		Me.chkCopyFiles = New System.Windows.Forms.CheckBox()
		Me.ProgressBar2 = New System.Windows.Forms.ProgressBar()
		Me.Panel1 = New System.Windows.Forms.Panel()
		Me.btnCheckDiff = New System.Windows.Forms.Button()
		Me.btnCheckNew = New System.Windows.Forms.Button()
		Me.btnStop = New System.Windows.Forms.LinkLabel()
		Me.btnOpenTempFolder = New System.Windows.Forms.Button()
		Me.chkHideNotSelected = New System.Windows.Forms.CheckBox()
		CType(Me.dgTables, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.Panel1.SuspendLayout()
		Me.SuspendLayout()
		'
		'btnExport
		'
		Me.btnExport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnExport.Location = New System.Drawing.Point(566, 446)
		Me.btnExport.Name = "btnExport"
		Me.btnExport.Size = New System.Drawing.Size(100, 23)
		Me.btnExport.TabIndex = 0
		Me.btnExport.Text = "Copy tables"
		'
		'txtFolderPath
		'
		Me.txtFolderPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtFolderPath.BackColor = System.Drawing.SystemColors.Window
		Me.txtFolderPath.Location = New System.Drawing.Point(73, 8)
		Me.txtFolderPath.Name = "txtFolderPath"
		Me.txtFolderPath.Size = New System.Drawing.Size(469, 20)
		Me.txtFolderPath.TabIndex = 1
		'
		'btnConnect1
		'
		Me.btnConnect1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnConnect1.Location = New System.Drawing.Point(708, 9)
		Me.btnConnect1.Name = "btnConnect1"
		Me.btnConnect1.Size = New System.Drawing.Size(75, 23)
		Me.btnConnect1.TabIndex = 3
		Me.btnConnect1.Text = "Connect"
		'
		'btnCancel
		'
		Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnCancel.Location = New System.Drawing.Point(672, 445)
		Me.btnCancel.Name = "btnCancel"
		Me.btnCancel.Size = New System.Drawing.Size(104, 24)
		Me.btnCancel.TabIndex = 6
		Me.btnCancel.Text = "Cancel"
		'
		'ProgressBar1
		'
		Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.ProgressBar1.Location = New System.Drawing.Point(224, 609)
		Me.ProgressBar1.Name = "ProgressBar1"
		Me.ProgressBar1.Size = New System.Drawing.Size(474, 11)
		Me.ProgressBar1.TabIndex = 7
		'
		'Label1
		'
		Me.Label1.AutoSize = True
		Me.Label1.Location = New System.Drawing.Point(5, 13)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(36, 13)
		Me.Label1.TabIndex = 8
		Me.Label1.Text = "Folder"
		'
		'chkCreateTable
		'
		Me.chkCreateTable.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.chkCreateTable.AutoSize = True
		Me.chkCreateTable.Checked = True
		Me.chkCreateTable.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkCreateTable.Location = New System.Drawing.Point(8, 445)
		Me.chkCreateTable.Name = "chkCreateTable"
		Me.chkCreateTable.Size = New System.Drawing.Size(113, 17)
		Me.chkCreateTable.TabIndex = 11
		Me.chkCreateTable.Text = "Create target table"
		Me.chkCreateTable.UseVisualStyleBackColor = True
		'
		'txtLog
		'
		Me.txtLog.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtLog.Location = New System.Drawing.Point(8, 474)
		Me.txtLog.Multiline = True
		Me.txtLog.Name = "txtLog"
		Me.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.txtLog.Size = New System.Drawing.Size(786, 133)
		Me.txtLog.TabIndex = 12
		'
		'chkDeleteData
		'
		Me.chkDeleteData.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.chkDeleteData.AutoSize = True
		Me.chkDeleteData.Checked = True
		Me.chkDeleteData.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkDeleteData.Location = New System.Drawing.Point(8, 415)
		Me.chkDeleteData.Name = "chkDeleteData"
		Me.chkDeleteData.Size = New System.Drawing.Size(118, 17)
		Me.chkDeleteData.TabIndex = 13
		Me.chkDeleteData.Text = "Delete before insert"
		Me.chkDeleteData.UseVisualStyleBackColor = True
		'
		'lbCount
		'
		Me.lbCount.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.lbCount.AutoSize = True
		Me.lbCount.Location = New System.Drawing.Point(739, 610)
		Me.lbCount.Name = "lbCount"
		Me.lbCount.Size = New System.Drawing.Size(25, 13)
		Me.lbCount.TabIndex = 14
		Me.lbCount.Text = "000"
		Me.lbCount.Visible = False
		'
		'Label2
		'
		Me.Label2.AutoSize = True
		Me.Label2.Location = New System.Drawing.Point(5, 41)
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(62, 13)
		Me.Label2.TabIndex = 15
		Me.Label2.Text = "SQL Server"
		'
		'txtConnectTo
		'
		Me.txtConnectTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtConnectTo.Location = New System.Drawing.Point(73, 41)
		Me.txtConnectTo.Name = "txtConnectTo"
		Me.txtConnectTo.Size = New System.Drawing.Size(629, 20)
		Me.txtConnectTo.TabIndex = 16
		'
		'btnConnect2
		'
		Me.btnConnect2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnConnect2.Location = New System.Drawing.Point(708, 39)
		Me.btnConnect2.Name = "btnConnect2"
		Me.btnConnect2.Size = New System.Drawing.Size(75, 23)
		Me.btnConnect2.TabIndex = 17
		Me.btnConnect2.Text = "Connect"
		Me.btnConnect2.UseVisualStyleBackColor = True
		'
		'dgTables
		'
		Me.dgTables.AllowUserToAddRows = False
		Me.dgTables.AllowUserToDeleteRows = False
		Me.dgTables.AllowUserToOrderColumns = True
		Me.dgTables.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
			Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.dgTables.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.dgTables.Location = New System.Drawing.Point(5, 68)
		Me.dgTables.Name = "dgTables"
		Me.dgTables.Size = New System.Drawing.Size(778, 342)
		Me.dgTables.TabIndex = 20
		'
		'btnSQL
		'
		Me.btnSQL.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnSQL.Location = New System.Drawing.Point(459, 446)
		Me.btnSQL.Name = "btnSQL"
		Me.btnSQL.Size = New System.Drawing.Size(102, 23)
		Me.btnSQL.TabIndex = 21
		Me.btnSQL.Text = "SQL"
		Me.btnSQL.UseVisualStyleBackColor = True
		'
		'Label3
		'
		Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.Label3.AutoSize = True
		Me.Label3.Location = New System.Drawing.Point(548, 11)
		Me.Label3.Name = "Label3"
		Me.Label3.Size = New System.Drawing.Size(53, 13)
		Me.Label3.TabIndex = 31
		Me.Label3.Text = "Password"
		'
		'txtPassword
		'
		Me.txtPassword.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtPassword.Location = New System.Drawing.Point(607, 8)
		Me.txtPassword.Name = "txtPassword"
		Me.txtPassword.Size = New System.Drawing.Size(95, 20)
		Me.txtPassword.TabIndex = 30
		'
		'chkDropTable
		'
		Me.chkDropTable.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.chkDropTable.AutoSize = True
		Me.chkDropTable.Location = New System.Drawing.Point(132, 445)
		Me.chkDropTable.Name = "chkDropTable"
		Me.chkDropTable.Size = New System.Drawing.Size(112, 17)
		Me.chkDropTable.TabIndex = 33
		Me.chkDropTable.Text = "Drop table if exists"
		Me.chkDropTable.UseVisualStyleBackColor = True
		'
		'chkSQL2008
		'
		Me.chkSQL2008.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.chkSQL2008.AutoSize = True
		Me.chkSQL2008.Checked = True
		Me.chkSQL2008.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkSQL2008.Location = New System.Drawing.Point(250, 445)
		Me.chkSQL2008.Name = "chkSQL2008"
		Me.chkSQL2008.Size = New System.Drawing.Size(93, 17)
		Me.chkSQL2008.TabIndex = 34
		Me.chkSQL2008.Text = "Sql Ser 2008+"
		Me.chkSQL2008.UseVisualStyleBackColor = True
		'
		'btnCheckAll
		'
		Me.btnCheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnCheckAll.Location = New System.Drawing.Point(208, 4)
		Me.btnCheckAll.Name = "btnCheckAll"
		Me.btnCheckAll.Size = New System.Drawing.Size(100, 23)
		Me.btnCheckAll.TabIndex = 35
		Me.btnCheckAll.Text = "Check All"
		Me.btnCheckAll.UseVisualStyleBackColor = True
		'
		'btnUncheckAll
		'
		Me.btnUncheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnUncheckAll.Location = New System.Drawing.Point(314, 4)
		Me.btnUncheckAll.Name = "btnUncheckAll"
		Me.btnUncheckAll.Size = New System.Drawing.Size(103, 23)
		Me.btnUncheckAll.TabIndex = 36
		Me.btnUncheckAll.Text = "Uncheck All"
		Me.btnUncheckAll.UseVisualStyleBackColor = True
		'
		'chkCopyFiles
		'
		Me.chkCopyFiles.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.chkCopyFiles.AutoSize = True
		Me.chkCopyFiles.Checked = True
		Me.chkCopyFiles.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkCopyFiles.Location = New System.Drawing.Point(132, 415)
		Me.chkCopyFiles.Name = "chkCopyFiles"
		Me.chkCopyFiles.Size = New System.Drawing.Size(103, 17)
		Me.chkCopyFiles.TabIndex = 37
		Me.chkCopyFiles.Text = "Copy files locally"
		Me.chkCopyFiles.UseVisualStyleBackColor = True
		'
		'ProgressBar2
		'
		Me.ProgressBar2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
		Me.ProgressBar2.Location = New System.Drawing.Point(5, 609)
		Me.ProgressBar2.Name = "ProgressBar2"
		Me.ProgressBar2.Size = New System.Drawing.Size(213, 11)
		Me.ProgressBar2.TabIndex = 38
		'
		'Panel1
		'
		Me.Panel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.Panel1.BackColor = System.Drawing.SystemColors.ControlLight
		Me.Panel1.Controls.Add(Me.btnCheckDiff)
		Me.Panel1.Controls.Add(Me.btnCheckNew)
		Me.Panel1.Controls.Add(Me.btnCheckAll)
		Me.Panel1.Controls.Add(Me.btnUncheckAll)
		Me.Panel1.Location = New System.Drawing.Point(359, 408)
		Me.Panel1.Name = "Panel1"
		Me.Panel1.Size = New System.Drawing.Size(424, 31)
		Me.Panel1.TabIndex = 39
		'
		'btnCheckDiff
		'
		Me.btnCheckDiff.Location = New System.Drawing.Point(5, 4)
		Me.btnCheckDiff.Name = "btnCheckDiff"
		Me.btnCheckDiff.Size = New System.Drawing.Size(89, 23)
		Me.btnCheckDiff.TabIndex = 38
		Me.btnCheckDiff.Text = "Check Diff Files"
		Me.btnCheckDiff.UseVisualStyleBackColor = True
		'
		'btnCheckNew
		'
		Me.btnCheckNew.Location = New System.Drawing.Point(100, 4)
		Me.btnCheckNew.Name = "btnCheckNew"
		Me.btnCheckNew.Size = New System.Drawing.Size(102, 23)
		Me.btnCheckNew.TabIndex = 37
		Me.btnCheckNew.Text = "Check New Rec"
		Me.btnCheckNew.UseVisualStyleBackColor = True
		'
		'btnStop
		'
		Me.btnStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnStop.AutoSize = True
		Me.btnStop.Location = New System.Drawing.Point(704, 609)
		Me.btnStop.Name = "btnStop"
		Me.btnStop.Size = New System.Drawing.Size(29, 13)
		Me.btnStop.TabIndex = 40
		Me.btnStop.TabStop = True
		Me.btnStop.Text = "Stop"
		Me.btnStop.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.btnStop.Visible = False
		'
		'btnOpenTempFolder
		'
		Me.btnOpenTempFolder.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnOpenTempFolder.Location = New System.Drawing.Point(364, 445)
		Me.btnOpenTempFolder.Name = "btnOpenTempFolder"
		Me.btnOpenTempFolder.Size = New System.Drawing.Size(89, 23)
		Me.btnOpenTempFolder.TabIndex = 41
		Me.btnOpenTempFolder.Text = "Open Folder"
		Me.btnOpenTempFolder.UseVisualStyleBackColor = True
		'
		'chkHideNotSelected
		'
		Me.chkHideNotSelected.AutoSize = True
		Me.chkHideNotSelected.Location = New System.Drawing.Point(250, 415)
		Me.chkHideNotSelected.Name = "chkHideNotSelected"
		Me.chkHideNotSelected.Size = New System.Drawing.Size(109, 17)
		Me.chkHideNotSelected.TabIndex = 42
		Me.chkHideNotSelected.Text = "Hide not selected"
		Me.chkHideNotSelected.UseVisualStyleBackColor = True
		'
		'Form1
		'
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.ClientSize = New System.Drawing.Size(802, 625)
		Me.Controls.Add(Me.chkHideNotSelected)
		Me.Controls.Add(Me.dgTables)
		Me.Controls.Add(Me.btnOpenTempFolder)
		Me.Controls.Add(Me.btnStop)
		Me.Controls.Add(Me.Panel1)
		Me.Controls.Add(Me.ProgressBar2)
		Me.Controls.Add(Me.chkCopyFiles)
		Me.Controls.Add(Me.chkSQL2008)
		Me.Controls.Add(Me.chkDropTable)
		Me.Controls.Add(Me.Label3)
		Me.Controls.Add(Me.txtPassword)
		Me.Controls.Add(Me.btnSQL)
		Me.Controls.Add(Me.btnConnect2)
		Me.Controls.Add(Me.txtConnectTo)
		Me.Controls.Add(Me.Label2)
		Me.Controls.Add(Me.lbCount)
		Me.Controls.Add(Me.chkDeleteData)
		Me.Controls.Add(Me.txtLog)
		Me.Controls.Add(Me.chkCreateTable)
		Me.Controls.Add(Me.Label1)
		Me.Controls.Add(Me.ProgressBar1)
		Me.Controls.Add(Me.btnCancel)
		Me.Controls.Add(Me.btnConnect1)
		Me.Controls.Add(Me.txtFolderPath)
		Me.Controls.Add(Me.btnExport)
		Me.Name = "Form1"
		Me.Text = "Copy tables from Paradox to SQL Server"
		CType(Me.dgTables, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Panel1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

#End Region

	Private Sub frmExport_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

		txtFolderPath.Text = oAppSetting.GetSetting("FolderPath")
		txtPassword.Text = oAppSetting.GetSetting("Password")
		txtConnectTo.Text = oAppSetting.GetSetting("ConnectTo")
		chkCopyFiles.Checked = oAppSetting.GetSetting("CopyFiles") = "1"
		chkHideNotSelected.Checked = oAppSetting.GetSetting("HideNotSelected") = "1"
		btnOpenTempFolder.Visible = chkCopyFiles.Checked
		btnCheckDiff.Visible = chkCopyFiles.Checked

		oSelTables = GetSelectedTablesFromReg()
	End Sub

	Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

		Dim oHash As New Hashtable
		oHash("FolderPath") = txtFolderPath.Text
		oHash("Password") = txtPassword.Text
		oHash("ConnectTo") = txtConnectTo.Text
		oHash("CopyFiles") = IIf(chkCopyFiles.Checked, "1", "0").ToString()
		oHash("HideNotSelected") = IIf(chkHideNotSelected.Checked, "1", "0").ToString()

		If Not dgTables.SortedColumn Is Nothing Then
			oHash("SortedColumn") = dgTables.SortedColumn.Name
			oHash("SortOrder") = dgTables.SortOrder.ToString()
		End If

		Dim oTables As List(Of String) = GetSelectedTables()
		Dim sTables As String = String.Join(",", oTables.ToArray)

		oHash("SelectedTables") = sTables

		oAppSetting.SaveSettings(oHash)

	End Sub

	Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
		If System.IO.Directory.Exists(txtFolderPath.Text) = False Then
			txtFolderPath.Text = ""
		End If

		Windows.Forms.Application.DoEvents()
		SetTableGrid(False)

		Dim sSortedColumn As String = oAppSetting.GetSetting("SortedColumn")
		If sSortedColumn <> "" Then
			Dim sSortOrder As String = oAppSetting.GetSetting("SortOrder")
			If sSortOrder = "Ascending" Then
				dgTables.Sort(dgTables.Columns(sSortedColumn), System.ComponentModel.ListSortDirection.Ascending)
			Else
				dgTables.Sort(dgTables.Columns(sSortedColumn), System.ComponentModel.ListSortDirection.Descending)
			End If
		End If
	End Sub

	Private Function GetSelectedTablesFromReg() As Hashtable

		Dim oSelTables As New Hashtable
		Dim sSelectedTables As String = oAppSetting.GetSetting("SelectedTables")
		If sSelectedTables <> "" Then
			Dim oSelectedTables As String() = Split(sSelectedTables, ",")
			For i As Integer = 0 To oSelectedTables.Length - 1
				Dim sTable As String = oSelectedTables(i)
				oSelTables(sTable) = True
			Next
		End If
		Return oSelTables
	End Function

	Private Sub btnConnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConnect1.Click

		FolderBrowserDialog1.SelectedPath = txtFolderPath.Text
		FolderBrowserDialog1.ShowDialog()

		If FolderBrowserDialog1.SelectedPath = "" Then
			dgTables.Rows.Clear()
			Exit Sub
		End If

		If FolderBrowserDialog1.SelectedPath = "" Then
			Exit Sub
		End If

		txtFolderPath.Text = FolderBrowserDialog1.SelectedPath

		SetTableGrid(True)
	End Sub

	Sub SetTableGrid(ByVal bRefresh As Boolean)

		If txtFolderPath.Text = "" Then
			Exit Sub
		End If

		Dim sSortOrder As String = ""
		Dim sSortedColumn As String = ""
		If bRefresh Then
			'Update oSelTables by selected tables from grid
			UpdateSelectedTables()

			If Not dgTables.SortedColumn Is Nothing Then
				sSortedColumn = dgTables.SortedColumn.Name
				sSortOrder = dgTables.SortOrder.ToString()
			End If
		End If

		Dim oSqlServerTables As System.Data.DataTable = GetSqlServerTables()
		Dim oTable As Data.DataTable = GetDbFiles(txtFolderPath.Text)

		'Update Checked, LocalDateModified, DestRowCount
		For iRow As Integer = 0 To oTable.Rows.Count - 1
			Dim sTableName As String = oTable.Rows(iRow)("Name").ToString()

			oTable.Rows(iRow)("Checked") = oSelTables.ContainsKey(sTableName)

			If chkCopyFiles.Checked Then
				Dim sToFilePath As String = System.IO.Path.Combine(GetTempFolderPath(), sTableName & ".db")
				If IO.File.Exists(sToFilePath) Then
					Dim oLocalFileInfo As New IO.FileInfo(sToFilePath)
					oTable.Rows(iRow)("LocalDateModified") = oLocalFileInfo.LastWriteTime
				End If
			End If

			If Not oSqlServerTables Is Nothing Then
				Dim oRows As Data.DataRow() = oSqlServerTables.Select("Name='" & sTableName & "'")
				If oRows.Length > 0 Then
					oTable.Rows(iRow)("DestRowCount") = oRows(0)("Rows")
				End If
			End If
		Next

		dgTables.DataSource = oTable
		dgTables.Update()

		Dim oCol As DataGridViewCheckBoxColumn = DirectCast(dgTables.Columns("Checked"), DataGridViewCheckBoxColumn)
		oCol.TrueValue = True
		oCol.SortMode = DataGridViewColumnSortMode.Automatic
		oCol.Width = 35
		oCol.HeaderText = ""

		dgTables.Columns("LocalDateModified").Visible = chkCopyFiles.Checked
		dgTables.Columns("DestRowCount").Visible = Not oSqlServerTables Is Nothing

		UpdateDataColumn("DateModified", "", "Date Modified")
		UpdateDataColumn("LocalDateModified", "", "Local Date")
		UpdateDataColumn("Size", "#,#", "Size")
		UpdateDataColumn("RowCount", "#,#", "Src Row Count")
		UpdateDataColumn("DestRowCount", "#,#", "Dest Row Count")

		SetupBackground()

		If sSortedColumn <> "" Then
			If sSortOrder = "Ascending" Then
				dgTables.Sort(dgTables.Columns(sSortedColumn), System.ComponentModel.ListSortDirection.Ascending)
			Else
				dgTables.Sort(dgTables.Columns(sSortedColumn), System.ComponentModel.ListSortDirection.Descending)
			End If
		End If

	End Sub

	Private Sub UpdateSelectedTables()

		oSelTables = New Hashtable
		Dim oTables As List(Of String) = GetSelectedTables()

		For i As Integer = 0 To oTables.Count - 1
			Dim sTable As String = oTables(i).ToString
			oSelTables(sTable) = True
		Next
	End Sub

	Private Sub SetupBackground()
		For iRow = 0 To dgTables.RowCount - 1
			Dim sSrcCount As String = dgTables.Rows(iRow).Cells("RowCount").Value.ToString()
			Dim sDstCount As String = dgTables.Rows(iRow).Cells("DestRowCount").Value.ToString()
			If sSrcCount <> "" AndAlso sDstCount <> "" Then
				If CInt(sSrcCount) = CInt(sDstCount) Then
					dgTables.Rows(iRow).Cells("DestRowCount").Style.BackColor = Color.LightBlue
				Else
					dgTables.Rows(iRow).Cells("DestRowCount").Style.BackColor = Color.LightPink
				End If
			Else
				dgTables.Rows(iRow).Cells("DestRowCount").Style.BackColor = Color.White
			End If

			If chkCopyFiles.Checked Then
				Dim sRemoteDate As String = dgTables.Rows(iRow).Cells("DateModified").Value.ToString()
				Dim sLocalDate As String = dgTables.Rows(iRow).Cells("LocalDateModified").Value.ToString()
				If sRemoteDate <> "" AndAlso sLocalDate <> "" Then
					If sRemoteDate = sLocalDate Then
						dgTables.Rows(iRow).Cells("LocalDateModified").Style.BackColor = Color.LightBlue
					Else
						dgTables.Rows(iRow).Cells("LocalDateModified").Style.BackColor = Color.LightPink
					End If
				End If
			End If
		Next
	End Sub

	Private Sub CheckCompare(sCol1 As String, sCol2 As String)
		For iRow = 0 To dgTables.RowCount - 1
			Dim sColVal1 As String = dgTables.Rows(iRow).Cells(sCol1).Value.ToString()
			Dim sColVal2 As String = dgTables.Rows(iRow).Cells(sCol2).Value.ToString()
			If sColVal1 <> "" AndAlso sColVal2 <> "" AndAlso sColVal1 <> sColVal2 Then
				dgTables.Rows(iRow).Cells("Checked").Value = True
			Else
				dgTables.Rows(iRow).Cells("Checked").Value = False
			End If
		Next
	End Sub

	Private Function GetDbFiles(ByVal sFolderPath As String) As Data.DataTable
		If chkHideNotSelected.Checked = False Then
			Return GetDbFiles2(sFolderPath)
		End If

		Dim oTable As Data.DataTable = GetDbFiles2(sFolderPath)
		For i As Integer = oTable.Rows.Count - 1 To 0 Step -1
			Dim sTableName As String = oTable.Rows(i)("Name").ToString()
			If oSelTables.ContainsKey(sTableName) = False Then
				oTable.Rows(i).Delete()
			End If
		Next

		Return oTable
	End Function

	Private Function GetDbFiles2(ByVal sFolderPath As String) As Data.DataTable

		'Try to get list if files from cache
		Dim sTempFilePath As String = GetTempFileName(sFolderPath, "xml")
		Dim ds As New System.Data.DataSet()
		If IO.File.Exists(sTempFilePath) Then
			Dim oFileInfo As New IO.FileInfo(sTempFilePath)
			If DateTime.Now.Subtract(oFileInfo.LastWriteTime).Hours > 2 Then
				'File is 2 hours old - delete
				System.IO.File.Delete(sTempFilePath)
			Else
				ds.ReadXml(sTempFilePath)
				Return ds.Tables(0)
			End If
		End If

		Dim dStart As DateTime = DateTime.Now

		Dim oTable As New Data.DataTable
		oTable.Columns.Add(New Data.DataColumn("Checked", System.Type.GetType("System.Boolean"))) '<--
		oTable.Columns.Add(New Data.DataColumn("Name"))
		oTable.Columns.Add(New Data.DataColumn("DateModified", System.Type.GetType("System.DateTime")))
		oTable.Columns.Add(New Data.DataColumn("LocalDateModified", System.Type.GetType("System.DateTime"))) '<--
		oTable.Columns.Add(New Data.DataColumn("Size", System.Type.GetType("System.Int64")))
		oTable.Columns.Add(New Data.DataColumn("RowCount", System.Type.GetType("System.Int64")))
		oTable.Columns.Add(New Data.DataColumn("DestRowCount", System.Type.GetType("System.Int64"))) '<--

		Dim oFiles As String() = System.IO.Directory.GetFiles(sFolderPath)
		For i As Integer = 0 To oFiles.Length - 1
			Dim sFilePath As String = oFiles(i)
			Dim oFileInfo As New IO.FileInfo(sFilePath)
			Dim sTableName As String = IO.Path.GetFileNameWithoutExtension(sFilePath)
			If oFileInfo.Extension.ToLower() = ".db" Then
				Dim iRowCount As Integer = GetParadoxRecCount(sFolderPath, sTableName)
				If iRowCount > 0 Then
					Dim oDataRow As DataRow = oTable.NewRow()
					oDataRow("Name") = sTableName
					oDataRow("DateModified") = oFileInfo.LastWriteTime
					oDataRow("Size") = oFileInfo.Length
					oDataRow("RowCount") = iRowCount
					oTable.Rows.Add(oDataRow)
				End If
			End If
		Next

		Dim oDuration As TimeSpan = DateTime.Now.Subtract(dStart)
		If oDuration.Seconds > 5 Then
			'Save to Cache if it took more than 5 seconds to get the list of tables
			ds.Tables.Add(oTable)
			ds.WriteXml(sTempFilePath, XmlWriteMode.WriteSchema)
		End If

		Return oTable
	End Function

	Private Function GetTempFileName(ByVal sKey As String, sExt As String) As String
		Dim oRegex As New Regex(String.Format("[{0}]", Regex.Escape(New String(IO.Path.GetInvalidFileNameChars()))), RegexOptions.Compiled)
		Dim sFileName As String = oRegex.Replace(sKey, "-") & "." & sExt
		Return IO.Path.Combine(GetTempFolderPath(), sFileName)
	End Function


	Private Function GetSqlServerTables() As Data.DataTable

		If txtConnectTo.Text = "" Then
			Return Nothing
		End If

		Dim cn As OleDbConnection = New OleDbConnection(txtConnectTo.Text)
		cn.Open()

		Dim sSql As String = "SELECT TABLE_SCHEMA, TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'"
		Dim oTable As Data.DataTable = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, Nothing})

		Dim oRetTable As New Data.DataTable
		oRetTable.Columns.Add(New Data.DataColumn("Name"))
		oRetTable.Columns.Add("Rows", System.Type.GetType("System.Int64"))

		For i As Long = 0 To oTable.Rows.Count - 1
			Dim sSchema As String = oTable.Rows(i)("TABLE_SCHEMA") & ""
			Dim sTName As String = oTable.Rows(i)("TABLE_NAME") & ""
			Dim sKey As String = sSchema & "." & sTName

			If sSchema <> "sys" Then
				If sSchema = "" Or sSchema = "dbo" Then
					sKey = sTName
				End If

				Try
					Dim cmd As New OleDbCommand("sp_MStablespace '" & sKey & "'", cn)
					Dim dr As OleDbDataReader = cmd.ExecuteReader()
					If dr.Read Then
						Dim iRowCount As Integer = CInt(dr.GetValue(dr.GetOrdinal("Rows")))
						If iRowCount > 0 Then
							Dim oDataRow As DataRow = oRetTable.NewRow()
							oDataRow("Name") = sKey
							oDataRow("Rows") = iRowCount
							oRetTable.Rows.Add(oDataRow)
						End If
					End If
					dr.Close()
				Catch ex As Exception
					'Do Nothing
				End Try
			End If

		Next

		cn.Close()
		Return oRetTable
	End Function

	Function GetParadoxConnectionString(ByVal sFolderPath As String, ByVal sPassword As String) As String
		If sFolderPath = "" Then
			Return ""
		End If

		If sPassword <> "" Then
			Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sFolderPath & ";Extended Properties=Paradox 5.x;Jet OLEDB:Database Password=" & sPassword & ";"
		Else
			Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sFolderPath & ";Extended Properties=Paradox 5.x;"
		End If
	End Function

	Private Function GetFolderPath() As String
		If chkCopyFiles.Checked Then
			Return GetTempFolderPath()
		Else
			Return txtFolderPath.Text
		End If
	End Function

	Private Sub CopyFileLocally(ByVal sFileName As String)
		If chkCopyFiles.Checked = False OrElse txtFolderPath.Text = "" OrElse IO.Directory.Exists(txtFolderPath.Text) = False Then
			Exit Sub
		End If

		Dim sFromFilePath As String = System.IO.Path.Combine(txtFolderPath.Text, sFileName)
		Dim oFromFileInfo As New IO.FileInfo(sFromFilePath)
		If oFromFileInfo.Exists = False Then
			'Source file does not exist
			Exit Sub
		End If

		Dim sTempFolderPath As String = GetTempFolderPath()
		Dim sToFilePath As String = System.IO.Path.Combine(sTempFolderPath, sFileName)
		Dim oToFileInfo As New IO.FileInfo(sToFilePath)

		If oToFileInfo.Exists AndAlso oFromFileInfo.Length = oToFileInfo.Length AndAlso oFromFileInfo.LastWriteTime = oToFileInfo.LastWriteTime Then
			'Assume files are the same
			Exit Sub
		End If

		Dim dStart2 As DateTime = DateTime.Now
		System.IO.File.Copy(sFromFilePath, sToFilePath, True)
		System.IO.File.SetLastWriteTime(sToFilePath, System.IO.File.GetLastWriteTime(sFromFilePath))
		Log("Copied " & sFromFilePath & " to " & sToFilePath & " in " & GetDuration(dStart2))
	End Sub


	Private Sub UpdateDataColumn(sColName As String, sFormat As String, sHeaderText As String)
		Dim oCol As DataGridViewColumn = dgTables.Columns(sColName)
		If sFormat <> "" Then oCol.DefaultCellStyle.Format = sFormat
		If sHeaderText <> "" Then oCol.HeaderText = sHeaderText
	End Sub

	Protected Function EditConnectionString(ByVal sConnectionString As String) As String
		Try
			Dim oDataLinks As Object = CreateObject("DataLinks")
			Dim cn As Object = CreateObject("ADODB.Connection")

			cn.ConnectionString = sConnectionString
			oDataLinks.hWnd = Me.Handle

			If Not oDataLinks.PromptEdit(cn) Then
				'User pressed cancel button
				Return ""
			End If

			cn.Open()

			Return cn.ConnectionString

		Catch ex As Exception
			MsgBox(ex.Message)
			Return ""
		End Try
	End Function

	Function GetSelectedTables() As List(Of String)
		Dim oRet As New List(Of String)

		For Each oRow As DataGridViewRow In dgTables.Rows
			Dim oCheckbox As DataGridViewCheckBoxCell = DirectCast(oRow.Cells.Item(0), DataGridViewCheckBoxCell)

			If oCheckbox.Value.ToString = oCheckbox.TrueValue.ToString() Then
				Dim sName As String = oRow.Cells(1).Value.ToString()
				oRet.Add(sName)
			End If
		Next

		Return oRet
	End Function


	Private Sub btnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport.Click

		If dgTables.Rows.Count = 0 Then
			MsgBox("Please connect to the source database.")
			Exit Sub
		End If

		Dim oTables As List(Of String) = GetSelectedTables()
		If oTables.Count = 0 Then
			MsgBox("Please select tables to copy.")
			Exit Sub
		End If

		txtLog.Clear()

		If oTables.Count = 0 Then
			Exit Sub
		End If

		Dim cnDst As OleDbConnection = New OleDbConnection(txtConnectTo.Text)
		cnDst.Open()

		ProgressBar2.Maximum = oTables.Count
		bStop = False

		For i As Integer = 0 To oTables.Count - 1
			ProgressBar2.Value = i

			Dim sTable As String = oTables(i).ToString
			If Len(sTable) > 26 Then
				Log("File name " & sTable & " is " & Len(sTable) & " characters long. It exceeds 26 characters allowed by Microsoft.Jet.OLEDB provider.")
			Else
				ExportTable(sTable, cnDst)

				If bStop Then
					Exit For
				End If
			End If
		Next

		ProgressBar2.Value = 0
		cnDst.Close()
		SetTableGrid(True)
	End Sub

	Private Sub OpenConnections(ByRef cn As OleDbConnection)
		If cn.State <> ConnectionState.Open Then
			cn.Open()
		End If
	End Sub


	Private Sub ExportTable(ByVal sTableName As String, ByRef cnDst As OleDbConnection)

		Dim dStart As DateTime = DateTime.Now
		Dim bDestTableExists As Boolean = False
		Dim iDestRecCount As Integer = 0

		Try
			Dim oSrcCmd As New OleDbCommand("SELECT Count(*) FROM " & PadSqlColumnName(sTableName), cnDst)
			iDestRecCount = Integer.Parse(oSrcCmd.ExecuteScalar().ToString())
			bDestTableExists = True
		Catch ex As Exception
			'Ignore - asume table dos not exist
		End Try

		If chkDeleteData.Checked And iDestRecCount > 0 Then
			Log("Deleteting data from table: " & sTableName)

			OpenConnections(cnDst)
			Dim sSql1 As String = "DELETE FROM " & PadSqlColumnName(sTableName)
			Dim oCmd1 As New OleDbCommand(sSql1, cnDst)

			Try
				oCmd1.ExecuteNonQuery()
			Catch ex As Exception
				Log(ex.Message & vbTab & "SQL: " & sSql1)
			End Try
		End If

		CopyFileLocally(sTableName & ".db")
		CopyFileLocally(sTableName & ".PX")
		CopyFileLocally(sTableName & ".MB")

		'Open Src connection
		Dim sFolderPath As String = GetFolderPath()
		Dim sParadoxConnectionString As String = GetParadoxConnectionString(sFolderPath, txtPassword.Text)
		Dim cnSrc As New OleDbConnection(sParadoxConnectionString)

		Try
			cnSrc.Open()
		Catch ex As Exception
			Log(ex.Message & ", ConnectionString: " & sParadoxConnectionString)
			Exit Sub
		End Try

		Dim sSql As String = "SELECT * FROM " & PadParaColumnName(sTableName & ".db")
		Dim cmd As New OleDbCommand(sSql, cnSrc)
		Dim dr As OleDbDataReader

		Try
			dr = cmd.ExecuteReader()
		Catch ex As Exception
			Log(ex.Message & ", SQL: " & sSql)
			Exit Sub
		End Try

		If dr.IsClosed Then
			Log("Recordset is closed , SQL: " & sSql)
			Exit Sub
		End If

		If chkCreateTable.Checked Then

			If bDestTableExists And chkDropTable.Checked Then
				Log("Drop table: " & sTableName)

				Dim oCmdDrop As New OleDbCommand("DROP TABLE " & PadSqlColumnName(sTableName), cnDst)

				Try
					oCmdDrop.ExecuteNonQuery()
					bDestTableExists = False
				Catch ex As Exception
					Log("Could not drop table: " & sTableName & ", " & ex.Message & vbTab)
				End Try

			End If

			If bDestTableExists = False Then
				Log("Create table: " & sTableName)

				'Make create table statement
				Dim sSql1 As String = ""

				Try
					sSql1 = GetCreateTableSqlFromParadox(sTableName, dr)
				Catch ex As Exception
					Log("GetCreateTableSqlFromParadox Error: " & ex.Message)
					Exit Sub
				End Try

				OpenConnections(cnDst)
				Dim oCmd1 As New OleDbCommand(sSql1, cnDst)

				Try
					oCmd1.ExecuteNonQuery()
					bDestTableExists = True
				Catch ex As Exception
					Log(ex.Message & vbTab & "SQL: " & sSql1)
				End Try
			End If
		End If

		If bDestTableExists = False Then
			Log("Destination table does not exist: " & sTableName)
			Exit Sub
		End If

		Dim iCount As Integer = GetParadoxRecCount(sFolderPath, sTableName)
		If iCount = 0 Then
			'Nothing to copy - Exit
			Exit Sub
		End If

		ProgressBar1.Maximum = iCount
		lbCount.Visible = True
		btnStop.Visible = True

		Log("Copying " & iCount & " rows from table: " & sTableName)
		CopyTableJet(sTableName, dr, cnDst)

		ProgressBar1.Value = 0
		lbCount.Visible = False
		lbCount.Text = ""
		btnStop.Visible = False

		Log("Copied table " & sTableName & vbTab & " in " & GetDuration(dStart))

		dr.Close()
		cnSrc.Close()
	End Sub

	Private Function GetDuration(ByVal dStart As DateTime) As String
		Dim oDuration As TimeSpan = DateTime.Now.Subtract(dStart)
		Return (New DateTime(oDuration.Ticks)).ToString("HH 'hrs' mm 'mins' ss 'secs'").Replace("00 hrs", "").Replace("00 mins", "").Trim()
	End Function

	Private Sub CopyTableJet(ByVal sTableName As String, dr As OleDbDataReader, ByRef cnDst As OleDbConnection)

		Dim oSchemaRows As Data.DataRowCollection = dr.GetSchemaTable.Rows
		Dim sRow As String
		Dim i As Integer
		Dim iRow As Integer = 0
		Dim iRowCount As Integer = 0

		'Get Header
		Dim sHeader As String = ""
		For i = 0 To oSchemaRows.Count - 1
			Dim sColumn As String = oSchemaRows(i)("ColumnName")
			If i <> 0 Then
				sHeader += ", "
			End If
			sHeader += PadSqlColumnName(sColumn)
		Next

		Dim sValues As String = ""

		While dr.Read()
			iRowCount += 1
			sRow = ""

			For i = 0 To oSchemaRows.Count - 1
				If sRow <> "" Then
					sRow += ", "
				End If

				sRow += GetValueString(dr.GetValue(i))
			Next

			If chkSQL2008.Checked Then
				If sValues <> "" Then sValues += ", "
				sValues += "(" & sRow & ")"

				If iRowCount >= 1000 Then
					Dim sSql1 As String = "INSERT INTO " & PadSqlColumnName(sTableName) & " (" & sHeader & ") VALUES " & sValues
					OpenConnections(cnDst)
					ExecuteSql(sSql1, cnDst)
					iRowCount = 0
					sValues = ""
				End If
			Else
				Dim sSql1 As String = "INSERT INTO " & PadSqlColumnName(sTableName) & " (" & sHeader & ") VALUES (" & sRow & ")"
				OpenConnections(cnDst)
				ExecuteSql(sSql1, cnDst)
			End If

			iRow += 1
			ProgressBar1.Value = Math.Min(ProgressBar1.Maximum, iRow)
			lbCount.Text = iRow.ToString()
			lbCount.Refresh()

			'Listen for the user to press Cancel button
			Windows.Forms.Application.DoEvents()
			If bStop Then
				Log("Copied table " & sTableName & " stopped. ")
				Exit While
			End If

		End While


		If chkSQL2008.Checked And sValues <> "" Then
			Dim sSql1 As String = "INSERT INTO " & PadSqlColumnName(sTableName) & " (" & sHeader & ") VALUES " & sValues
			ExecuteSql(sSql1, cnDst)
		End If

	End Sub

	Private Sub ExecuteSql(ByVal sSql As String, ByRef cnDst As OleDbConnection)
		Dim oCmd1 As New OleDbCommand(sSql, cnDst)
		Dim bError As Boolean = False
		Try
			oCmd1.ExecuteNonQuery()
		Catch ex As Exception
			bError = True
		End Try

		If bError = False Then
			Exit Sub
		End If

		'Remove Control Charachters and try again
		oCmd1.CommandText = Regex.Replace(sSql, "\p{C}+", "")
		Try
			oCmd1.ExecuteNonQuery()
		Catch ex As Exception
			Log(ex.Message & vbTab & "SQL: " & sSql)
		End Try

	End Sub

	Private Function GetParadoxRecCount(ByVal sFolderPath As String, ByVal sTableName As String) As Integer

		If IO.File.Exists(IO.Path.Combine(sFolderPath, sTableName & ".db")) = False Then
			Return 0
		End If

		Try
			Dim oTable = New ParadoxReader.ParadoxTable(sFolderPath, sTableName)
			Return oTable.RecordCount
		Catch ex As Exception
			'Log("GetParadoxRecCount Error for: " & sTableName)
		End Try

		Return 0
	End Function

	Private Sub Log(s As String)

		If txtLog.Text = "" Then
			txtLog.Text = s
		Else
			txtLog.AppendText(vbCrLf & s)
		End If

		txtLog.Visible = True
		txtLog.ScrollToCaret()
		txtLog.Refresh()
	End Sub

	Private Function GetValueString(ByVal obj As Object) As String
		If (IsDBNull(obj)) Then Return "NULL"

		Select Case obj.GetType.FullName

			Case "System.Boolean"
				If (obj = True) Then
					Return "1"
				Else
					Return "0"
				End If

			Case "System.String"
				Dim str As String = obj
				Return "'" + str.Replace("'", "''") + "'"

			Case "System.DateTime"
				Dim dDate As Date = obj
				If dDate.Year > 1960 AndAlso dDate.Year < 2060 Then
					Return "CONVERT(datetime, '" + obj.ToString() + "',101)"
				Else
					Return "NULL"
				End If

			Case "System.Drawing.Image"
				Return "NULL"

			Case "System.Drawing.Bitmap"
				Return "NULL"

			Case "System.Byte[]"
				Return "0x" + GetHexString(obj)

			Case "System.Double"
				If obj.ToString() = "NaN" Then
					Return "NULL"
				ElseIf Format(obj, "0").Length > 20 Then
					Return "NULL"
				Else
					Return obj.ToString()
				End If

			Case Else
				Return obj.ToString()

		End Select
	End Function

	Private Function GetHexString(ByRef bytes() As Byte) As String
		Dim sb As New System.Text.StringBuilder
		Dim b As Byte
		Dim i As Integer = 0

		For Each b In bytes
			i += 1
			sb.Append(b.ToString("X2"))
			If i > 10 Then
				Return sb.ToString()
			End If
		Next

		Return sb.ToString()
	End Function

	Private Function GetCreateTableSqlFromParadox(ByVal sTableName As String, dr As OleDbDataReader) As String

		Dim sb As New System.Text.StringBuilder()
		Dim oSchemaRows As Data.DataRowCollection = dr.GetSchemaTable.Rows
		Dim sKeyColumns As String = ""
		Dim i As Integer = 0

		sb.Append("CREATE TABLE " & PadSqlColumnName(sTableName) & " (" & vbCrLf)

		For iCol As Integer = 0 To oSchemaRows.Count - 1
			Dim sColumn As String = oSchemaRows(iCol).Item("ColumnName").ToString() & ""
			Dim sColumnSize As String = oSchemaRows(iCol).Item("ColumnSize").ToString() & ""
			Dim sDataType As String = oSchemaRows(iCol).Item("DATATYPE").FullName.ToString()
			Dim bAllowDBNull As Boolean = oSchemaRows(iCol).Item("AllowDBNull")	'Does not always work

			If i > 0 Then
				sb.Append(",")
				sb.Append(vbCrLf)
			End If

			sb.Append(PadSqlColumnName(sColumn))
			sb.Append(" " & PadAccessDataType(sDataType, sColumnSize))

			If bAllowDBNull Then
				sb.Append(" NULL")
			Else
				sb.Append(" NOT NULL")
			End If

			i += 1
		Next

		sb.Append(")")

		If i = 0 Then
			Return ""
		Else
			Return sb.ToString()
		End If

	End Function

	Private Function PadAccessDataType(ByVal sDataType As String, ByVal sColumnSize As String) As String

		Dim bUseSize As Boolean = True
		sDataType = Replace(sDataType, "System.", "")

		Select Case LCase(sDataType)
			Case "string" : sDataType = "VARCHAR"
			Case "int16" : sDataType = "SMALLINT" : bUseSize = False
			Case "int32" : sDataType = "INT" : bUseSize = False
			Case "int64" : sDataType = "BIGINT" : bUseSize = False
			Case "double" : sDataType = "FLOAT"
			Case "datetime" : bUseSize = False
		End Select

		If sColumnSize <> "" And bUseSize Then
			If sDataType = "VARCHAR" AndAlso CInt(sColumnSize) >= 8000 Then
				Return "VARCHAR(MAX)"
			Else
				Return sDataType & "(" & sColumnSize & ")"
			End If
		Else
			Return sDataType
		End If

	End Function

	Public Function PadParaColumnName(ByVal sTable As String, Optional ByVal bAlwaysPad As Boolean = True) As String
		If bAlwaysPad _
			OrElse sTable.IndexOf(" ") <> -1 Then
			Return "`" & sTable & "`"
		Else
			Return sTable
		End If
	End Function

	Public Function PadSqlColumnName(ByVal sTable As String, Optional ByVal bAlwaysPad As Boolean = True) As String
		If bAlwaysPad _
			OrElse sTable.IndexOf(" ") <> -1 _
			OrElse sTable.IndexOf("#") <> -1 _
			OrElse IsNumeric(sTable.Substring(0, 1)) Then
			Return "[" & sTable & "]"
		Else
			Return sTable
		End If
	End Function


	Private Sub btnSQL_Click(sender As Object, e As EventArgs) Handles btnSQL.Click

		Dim sTable As String = ""
		For Each oRow As DataGridViewRow In dgTables.Rows
			For Each oCell As DataGridViewCell In oRow.Cells
				If oCell.Selected Then
					Dim sName As String = oRow.Cells(1).Value.ToString()
					sTable = sName
					Exit For
				End If
			Next
		Next

		RunSql.Show()

		If sTable <> "" Then
			CopyFileLocally(sTable & ".db")
			CopyFileLocally(sTable & ".PX")
			CopyFileLocally(sTable & ".MB")

			RunSql.txtPassword.Text = txtPassword.Text
			RunSql.txtFolder.Text = GetFolderPath()
			RunSql.txtSql.Text = "SELECT * FROM " & PadParaColumnName(sTable & ".db")
		End If

	End Sub

	Private Function GetTempFolderPath() As String
		Dim sFolder As String = System.IO.Path.GetTempPath()
		Dim sRetFolder As String = System.IO.Path.Combine(sFolder, "ParadoxCopy")


		If Not System.IO.Directory.Exists(sRetFolder) Then
			System.IO.Directory.CreateDirectory(sRetFolder)
		End If

		Return sRetFolder
	End Function

	Private Sub btnConnect2_Click(sender As Object, e As EventArgs) Handles btnConnect2.Click
		Dim sConnectionString As String = txtConnectTo.Text

		If sConnectionString = "" Then
			sConnectionString = "Provider=SQLOLEDB.1"
		End If

		sConnectionString = EditConnectionString(sConnectionString)
		If sConnectionString = "" Then
			Exit Sub
		End If

		txtConnectTo.Text = sConnectionString

		SetTableGrid(True)
	End Sub

	Private Sub btnCheckAll_Click(sender As Object, e As EventArgs) Handles btnCheckAll.Click
		For iRow = 0 To dgTables.RowCount - 1
			dgTables.Rows(iRow).Cells("Checked").Value = True
		Next
	End Sub

	Private Sub btnUncheckAll_Click(sender As Object, e As EventArgs) Handles btnUncheckAll.Click
		For iRow = dgTables.RowCount - 1 To 0 Step -1
			dgTables.Rows(iRow).Cells("Checked").Value = False
		Next
	End Sub

	Private Sub btnCheckNew_Click(sender As Object, e As EventArgs) Handles btnCheckNew.Click
		CheckCompare("DestRowCount", "RowCount")
	End Sub

	Private Sub btnCheckDiff_Click(sender As Object, e As EventArgs) Handles btnCheckDiff.Click
		CheckCompare("DateModified", "LocalDateModified")
	End Sub

	Private Sub dgTables_Sorted(sender As Object, e As EventArgs) Handles dgTables.Sorted
		SetupBackground()
	End Sub

	Private Sub txtFolderPath_KeyUp(sender As Object, e As KeyEventArgs) Handles txtFolderPath.KeyUp
		If e.KeyCode = Keys.Enter Then
			SetTableGrid(True)
		End If
	End Sub

	Private Sub txtConnectTo_KeyUp(sender As Object, e As KeyEventArgs) Handles txtConnectTo.KeyUp
		If e.KeyCode = Keys.Enter Then
			SetTableGrid(True)
		End If
	End Sub

	Private Sub chkCreateTable_CheckedChanged(sender As Object, e As EventArgs) Handles chkCreateTable.CheckedChanged
		chkDropTable.Visible = chkCreateTable.Checked
	End Sub

	Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
		Me.Close()
	End Sub

	Private Sub btnStop_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles btnStop.LinkClicked
		bStop = True
	End Sub

	Private Sub chkCopyFiles_CheckedChanged(sender As Object, e As EventArgs) Handles chkCopyFiles.CheckedChanged
		btnOpenTempFolder.Visible = chkCopyFiles.Checked
		btnCheckDiff.Visible = chkCopyFiles.Checked
		SetTableGrid(True)
	End Sub

	Private Sub btnOpenTempFolder_Click(sender As Object, e As EventArgs) Handles btnOpenTempFolder.Click
		Dim sTempFolderPath As String = GetTempFolderPath()
		Process.Start("explorer.exe", sTempFolderPath)
	End Sub

	Private Sub chkHideSelected_CheckedChanged(sender As Object, e As EventArgs) Handles chkHideNotSelected.CheckedChanged
		SetTableGrid(True)
	End Sub
End Class
