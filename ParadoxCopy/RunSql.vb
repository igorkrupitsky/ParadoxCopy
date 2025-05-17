Imports System.Data.OleDb
Imports System.Data.Odbc

Public Class RunSql

	'Private Sub btnExecute_Click(sender As Object, e As EventArgs) Handles btnExecute.Click

	'	txtLog.Text = ""
	'	lblStatus.Text = "Running..."
	'	Windows.Forms.Application.DoEvents()
	'	Me.Refresh()

	'	Dim sConnectionString As String = GetFromConnectionString()
	'	Dim cn As New OdbcConnection(sConnectionString)
	'	cn.Open()

	'	If txtSql.Text.IndexOf(";" & vbCrLf) <> -1 Then
	'		ExecuteBatchCommand(cn)
	'		Exit Sub
	'	End If

	'	Try
	'		Dim ad As OdbcDataAdapter = New OdbcDataAdapter(txtSql.Text, cn)
	'		Dim ds As DataSet = New DataSet
	'		ad.Fill(ds)
	'		If ds.Tables.Count > 0 Then
	'			dgResult.DataSource = ds.Tables(0)
	'		End If
	'	Catch ex As Exception
	'		Log(ex.Message)
	'	End Try

	'	dgResult.Update()
	'	cn.Close()

	'	lblStatus.Text = ""
	'End Sub

	'Sub ExecuteBatchCommand(cn As OdbcConnection)
	'	Dim bSuccess As Boolean = True
	'	Dim iCount As Integer = 0
	'	Dim oList As String() = System.Text.RegularExpressions.Regex.Split(txtSql.Text, ";" & vbCrLf)
	'	For i As Integer = 0 To oList.Length - 1
	'		Dim sSql As String = Trim(oList(i))
	'		If sSql <> "" Then
	'			Try
	'				Dim cmd As New OdbcCommand(sSql, cn)
	'				cmd.ExecuteNonQuery()
	'				iCount += 1
	'			Catch ex As Exception
	'				Log(ex.Message & "; SQL: " & sSql)
	'				bSuccess = False
	'			End Try
	'		End If
	'	Next

	'	If bSuccess Then
	'		Log("Successfully executed " & iCount & " commands.")
	'	End If

	'	cn.Close()
	'End Sub

	Private Sub btnExecute_Click(sender As Object, e As EventArgs) Handles btnExecute.Click

		txtLog.Text = ""
		lblStatus.Text = "Running..."
		Windows.Forms.Application.DoEvents()
		Me.Refresh()

		Dim sConnectionString As String = GetFromConnectionString()
		Dim cn As New OleDbConnection(sConnectionString)

		Try
			cn.Open()
		Catch ex As Exception
			Log(ex.Message & "; ConnectionString: " & sConnectionString)
		End Try

		If txtSql.Text.IndexOf(";" & vbCrLf) <> -1 Then
			ExecuteBatchCommand(cn)
			Exit Sub
		End If

		Try
			Dim ad As OleDbDataAdapter = New OleDbDataAdapter(txtSql.Text, cn)
			Dim ds As DataSet = New DataSet
			ad.Fill(ds)
			If ds.Tables.Count > 0 Then
				dgResult.DataSource = ds.Tables(0)
			End If
		Catch ex As Exception
			Log(ex.Message)
		End Try

		dgResult.Update()
		cn.Close()

		lblStatus.Text = ""
	End Sub

	Sub ExecuteBatchCommand(cn As OleDbConnection)
		Dim bSuccess As Boolean = True
		Dim iCount As Integer = 0
		Dim oList As String() = System.Text.RegularExpressions.Regex.Split(txtSql.Text, ";" & vbCrLf)
		For i As Integer = 0 To oList.Length - 1
			Dim sSql As String = Trim(oList(i))
			If sSql <> "" Then
				Try
					Dim cmd As New OleDbCommand(sSql, cn)
					cmd.ExecuteNonQuery()
					iCount += 1
				Catch ex As Exception
					Log(ex.Message & "; SQL: " & sSql)
					bSuccess = False
				End Try
			End If
		Next

		If bSuccess Then
			Log("Successfully executed " & iCount & " commands.")
		End If

		cn.Close()
	End Sub

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

	Function GetFromConnectionString() As String
		If txtFolder.Text = "" Then
			Return ""
		End If

		'http://www.reportizer.net/documentation/reportizer/opening-paradox-files.htm
		'https://docs.microsoft.com/en-us/sql/odbc/microsoft/sqlconfigdatasource-paradox-driver

		'Provider=MSDASQL.1;Extended Properties="DefaultDir=C:\MyParadoxFolder;Driver={Microsoft Paradox Driver (*.db )};DriverId=26;"
		'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\MyParadoxFolder;Extended Properties=dBASE IV;
		'Driver={Microsoft Paradox Driver (*.db )};DriverID=538;Fil=Paradox 5.X;DefaultDir=c:\pathToDb\;Dbq=c:\pathToDb\;CollatingSequence=ASCII;
		'Return "Provider=MSDASQL.1;Extended Properties=""DefaultDir=" & txtFolder.Text & "\;Dbq=" & txtFolder.Text & "\;Driver={Microsoft Paradox Driver (*.db )};DriverId=538;Fil=Paradox 5.X"";PWD=" & txtPassword.Text & ";"

		'If txtPassword.Text <> "" Then
		'	Return "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & txtFolder.Text & "';Extended Properties=dBASE;Jet OLEDB:Database Password=" & txtPassword.Text & ";"
		'Else
		'	Return "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & txtFolder.Text & "';Extended Properties=dBASE;"
		'End If


		If txtPassword.Text <> "" Then
			Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtFolder.Text & ";Extended Properties=Paradox 7.x;Jet OLEDB:Database Password=" & txtPassword.Text & ";"
		Else
			Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtFolder.Text & ";Extended Properties=Paradox 7.x;"
		End If

	End Function

End Class