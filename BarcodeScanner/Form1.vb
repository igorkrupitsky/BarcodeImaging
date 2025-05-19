Imports System.IO

Public Class Form1
	Inherits System.Windows.Forms.Form

	Declare Unicode Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameW" (ByVal longPath As String, ByVal ShortPath As System.Text.StringBuilder, ByVal bufferSize As Integer) As Integer

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
	Friend WithEvents Label1 As System.Windows.Forms.Label
	Friend WithEvents txtFrom As System.Windows.Forms.TextBox
	Friend WithEvents btnFrom As System.Windows.Forms.Button
	Friend WithEvents fldFrom As System.Windows.Forms.FolderBrowserDialog
	Friend WithEvents btnProcess As System.Windows.Forms.Button
	Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
	Friend WithEvents chkUseBarcodeZones As System.Windows.Forms.CheckBox
	Friend WithEvents cbBarcodeType As System.Windows.Forms.ComboBox
	Friend WithEvents lblType As System.Windows.Forms.Label
	Friend WithEvents Label2 As System.Windows.Forms.Label
	Friend WithEvents cbDirection As System.Windows.Forms.ComboBox
	Friend WithEvents chkShowTime As System.Windows.Forms.CheckBox
	Friend WithEvents chkCSV As System.Windows.Forms.CheckBox
	Friend WithEvents chkLog As System.Windows.Forms.CheckBox
	Friend WithEvents chkBytescout As System.Windows.Forms.CheckBox
	Friend WithEvents pnCS As System.Windows.Forms.Panel
	Friend WithEvents pnByteScout As System.Windows.Forms.Panel
	Friend WithEvents txtRegKey As System.Windows.Forms.TextBox
	Friend WithEvents txtRegName As System.Windows.Forms.TextBox
	Friend WithEvents Label4 As System.Windows.Forms.Label
	Friend WithEvents Label3 As System.Windows.Forms.Label
	Friend WithEvents chkFast As System.Windows.Forms.CheckBox
    Friend WithEvents Label5 As Label
	Friend WithEvents txtStartFile As TextBox
	Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
	Friend WithEvents txtOutput As System.Windows.Forms.TextBox
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtFrom = New System.Windows.Forms.TextBox()
        Me.btnFrom = New System.Windows.Forms.Button()
        Me.txtOutput = New System.Windows.Forms.TextBox()
        Me.fldFrom = New System.Windows.Forms.FolderBrowserDialog()
        Me.btnProcess = New System.Windows.Forms.Button()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.chkUseBarcodeZones = New System.Windows.Forms.CheckBox()
        Me.cbBarcodeType = New System.Windows.Forms.ComboBox()
        Me.lblType = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbDirection = New System.Windows.Forms.ComboBox()
        Me.chkShowTime = New System.Windows.Forms.CheckBox()
        Me.chkCSV = New System.Windows.Forms.CheckBox()
        Me.chkLog = New System.Windows.Forms.CheckBox()
        Me.chkBytescout = New System.Windows.Forms.CheckBox()
        Me.pnCS = New System.Windows.Forms.Panel()
        Me.pnByteScout = New System.Windows.Forms.Panel()
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel()
        Me.chkFast = New System.Windows.Forms.CheckBox()
        Me.txtRegKey = New System.Windows.Forms.TextBox()
        Me.txtRegName = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtStartFile = New System.Windows.Forms.TextBox()
        Me.pnCS.SuspendLayout()
        Me.pnByteScout.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(19, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(54, 20)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Folder"
        '
        'txtFrom
        '
        Me.txtFrom.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFrom.Location = New System.Drawing.Point(86, 15)
        Me.txtFrom.Name = "txtFrom"
        Me.txtFrom.Size = New System.Drawing.Size(1049, 26)
        Me.txtFrom.TabIndex = 5
        '
        'btnFrom
        '
        Me.btnFrom.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFrom.Location = New System.Drawing.Point(1145, 12)
        Me.btnFrom.Name = "btnFrom"
        Me.btnFrom.Size = New System.Drawing.Size(45, 33)
        Me.btnFrom.TabIndex = 7
        Me.btnFrom.Text = "..."
        Me.btnFrom.UseVisualStyleBackColor = True
        '
        'txtOutput
        '
        Me.txtOutput.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtOutput.Location = New System.Drawing.Point(26, 346)
        Me.txtOutput.Multiline = True
        Me.txtOutput.Name = "txtOutput"
        Me.txtOutput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtOutput.Size = New System.Drawing.Size(1164, 121)
        Me.txtOutput.TabIndex = 13
        '
        'btnProcess
        '
        Me.btnProcess.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnProcess.Location = New System.Drawing.Point(24, 304)
        Me.btnProcess.Name = "btnProcess"
        Me.btnProcess.Size = New System.Drawing.Size(1164, 34)
        Me.btnProcess.TabIndex = 14
        Me.btnProcess.Text = "Process"
        Me.btnProcess.UseVisualStyleBackColor = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(26, 475)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(1164, 28)
        Me.ProgressBar1.TabIndex = 15
        '
        'chkUseBarcodeZones
        '
        Me.chkUseBarcodeZones.AutoSize = True
        Me.chkUseBarcodeZones.Location = New System.Drawing.Point(582, 7)
        Me.chkUseBarcodeZones.Name = "chkUseBarcodeZones"
        Me.chkUseBarcodeZones.Size = New System.Drawing.Size(177, 24)
        Me.chkUseBarcodeZones.TabIndex = 16
        Me.chkUseBarcodeZones.Text = "Use Barcode Zones"
        Me.chkUseBarcodeZones.UseVisualStyleBackColor = True
        '
        'cbBarcodeType
        '
        Me.cbBarcodeType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbBarcodeType.FormattingEnabled = True
        Me.cbBarcodeType.Items.AddRange(New Object() {"All", "Code39", "Code128", "EAN"})
        Me.cbBarcodeType.Location = New System.Drawing.Point(61, 4)
        Me.cbBarcodeType.Name = "cbBarcodeType"
        Me.cbBarcodeType.Size = New System.Drawing.Size(205, 28)
        Me.cbBarcodeType.TabIndex = 17
        '
        'lblType
        '
        Me.lblType.AutoSize = True
        Me.lblType.Location = New System.Drawing.Point(5, 10)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(43, 20)
        Me.lblType.TabIndex = 18
        Me.lblType.Text = "Type"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(291, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 20)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "Direction"
        '
        'cbDirection
        '
        Me.cbDirection.FormattingEnabled = True
        Me.cbDirection.Items.AddRange(New Object() {"All", "Vertical", "Horizontal"})
        Me.cbDirection.Location = New System.Drawing.Point(379, 4)
        Me.cbDirection.Name = "cbDirection"
        Me.cbDirection.Size = New System.Drawing.Size(194, 28)
        Me.cbDirection.TabIndex = 20
        '
        'chkShowTime
        '
        Me.chkShowTime.AutoSize = True
        Me.chkShowTime.Checked = True
        Me.chkShowTime.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowTime.Location = New System.Drawing.Point(971, 57)
        Me.chkShowTime.Name = "chkShowTime"
        Me.chkShowTime.Size = New System.Drawing.Size(113, 24)
        Me.chkShowTime.TabIndex = 21
        Me.chkShowTime.Text = "Show Time"
        Me.chkShowTime.UseVisualStyleBackColor = True
        '
        'chkCSV
        '
        Me.chkCSV.AutoSize = True
        Me.chkCSV.Checked = True
        Me.chkCSV.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCSV.Location = New System.Drawing.Point(971, 94)
        Me.chkCSV.Name = "chkCSV"
        Me.chkCSV.Size = New System.Drawing.Size(111, 24)
        Me.chkCSV.TabIndex = 22
        Me.chkCSV.Text = "Make CSV"
        Me.chkCSV.UseVisualStyleBackColor = True
        '
        'chkLog
        '
        Me.chkLog.AutoSize = True
        Me.chkLog.Checked = True
        Me.chkLog.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkLog.Location = New System.Drawing.Point(843, 57)
        Me.chkLog.Name = "chkLog"
        Me.chkLog.Size = New System.Drawing.Size(106, 24)
        Me.chkLog.TabIndex = 23
        Me.chkLog.Text = "Show Log"
        Me.chkLog.UseVisualStyleBackColor = True
        '
        'chkBytescout
        '
        Me.chkBytescout.AutoSize = True
        Me.chkBytescout.Location = New System.Drawing.Point(843, 91)
        Me.chkBytescout.Name = "chkBytescout"
        Me.chkBytescout.Size = New System.Drawing.Size(113, 24)
        Me.chkBytescout.TabIndex = 24
        Me.chkBytescout.Text = "Byte Scout"
        Me.chkBytescout.UseVisualStyleBackColor = True
        '
        'pnCS
        '
        Me.pnCS.Controls.Add(Me.cbDirection)
        Me.pnCS.Controls.Add(Me.cbBarcodeType)
        Me.pnCS.Controls.Add(Me.lblType)
        Me.pnCS.Controls.Add(Me.Label2)
        Me.pnCS.Controls.Add(Me.chkUseBarcodeZones)
        Me.pnCS.Location = New System.Drawing.Point(26, 57)
        Me.pnCS.Name = "pnCS"
        Me.pnCS.Size = New System.Drawing.Size(785, 41)
        Me.pnCS.TabIndex = 25
        Me.pnCS.Visible = False
        '
        'pnByteScout
        '
        Me.pnByteScout.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnByteScout.Controls.Add(Me.LinkLabel1)
        Me.pnByteScout.Controls.Add(Me.chkFast)
        Me.pnByteScout.Controls.Add(Me.txtRegKey)
        Me.pnByteScout.Controls.Add(Me.txtRegName)
        Me.pnByteScout.Controls.Add(Me.Label4)
        Me.pnByteScout.Controls.Add(Me.Label3)
        Me.pnByteScout.Location = New System.Drawing.Point(26, 124)
        Me.pnByteScout.Name = "pnByteScout"
        Me.pnByteScout.Size = New System.Drawing.Size(1140, 122)
        Me.pnByteScout.TabIndex = 26
        '
        'LinkLabel1
        '
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(174, 92)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(162, 20)
        Me.LinkLabel1.TabIndex = 5
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Download Byte Scout"
        '
        'chkFast
        '
        Me.chkFast.AutoSize = True
        Me.chkFast.Location = New System.Drawing.Point(10, 92)
        Me.chkFast.Name = "chkFast"
        Me.chkFast.Size = New System.Drawing.Size(111, 24)
        Me.chkFast.TabIndex = 4
        Me.chkFast.Text = "Fast Mode"
        Me.chkFast.UseVisualStyleBackColor = True
        '
        'txtRegKey
        '
        Me.txtRegKey.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRegKey.Location = New System.Drawing.Point(62, 50)
        Me.txtRegKey.Name = "txtRegKey"
        Me.txtRegKey.Size = New System.Drawing.Size(1069, 26)
        Me.txtRegKey.TabIndex = 3
        Me.txtRegKey.Text = "demo"
        '
        'txtRegName
        '
        Me.txtRegName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRegName.Location = New System.Drawing.Point(62, 7)
        Me.txtRegName.Name = "txtRegName"
        Me.txtRegName.Size = New System.Drawing.Size(1069, 26)
        Me.txtRegName.TabIndex = 2
        Me.txtRegName.Text = "demo"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 50)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(35, 20)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "Key"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(51, 20)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Name"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(35, 256)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(44, 20)
        Me.Label5.TabIndex = 27
        Me.Label5.Text = "Start"
        '
        'txtStartFile
        '
        Me.txtStartFile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtStartFile.Location = New System.Drawing.Point(88, 256)
        Me.txtStartFile.Name = "txtStartFile"
        Me.txtStartFile.Size = New System.Drawing.Size(1068, 26)
        Me.txtStartFile.TabIndex = 28
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(8, 19)
        Me.ClientSize = New System.Drawing.Size(1209, 510)
        Me.Controls.Add(Me.txtStartFile)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.pnByteScout)
        Me.Controls.Add(Me.pnCS)
        Me.Controls.Add(Me.chkBytescout)
        Me.Controls.Add(Me.chkLog)
        Me.Controls.Add(Me.chkCSV)
        Me.Controls.Add(Me.chkShowTime)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.btnProcess)
        Me.Controls.Add(Me.txtOutput)
        Me.Controls.Add(Me.btnFrom)
        Me.Controls.Add(Me.txtFrom)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Form1"
        Me.Text = "Barcode Scanner"
        Me.pnCS.ResumeLayout(False)
        Me.pnCS.PerformLayout()
        Me.pnByteScout.ResumeLayout(False)
        Me.pnByteScout.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Dim oAppSetting As New AppSetting()
	Dim oLogFile As System.IO.StreamWriter
	Dim bStartFile As Boolean = True

	Private Sub Form1_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		oAppSetting.SetValue("FromPath", txtFrom.Text)
		oAppSetting.SetValue("RegName", txtRegName.Text)
		oAppSetting.SetValue("RegKey", txtRegKey.Text)
		oAppSetting.SetValue("StartFile", txtStartFile.Text)

		oAppSetting.SetValue("BarcodeType", cbBarcodeType.SelectedIndex)
		oAppSetting.SetValue("Direction", cbDirection.SelectedIndex)

		oAppSetting.SetValue("UseBarcodeZones", IIf(chkUseBarcodeZones.Checked, "1", "0"))
		oAppSetting.SetValue("Log", IIf(chkLog.Checked, "1", "0"))
		oAppSetting.SetValue("ShowTime", IIf(chkShowTime.Checked, "1", "0"))
		oAppSetting.SetValue("CSV", IIf(chkCSV.Checked, "1", "0"))

		oAppSetting.SetValue("Bytescout", IIf(chkBytescout.Checked, "1", "0"))
		oAppSetting.SetValue("Fast", IIf(chkFast.Checked, "1", "0"))

		oAppSetting.SaveData()
	End Sub

	Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

		oAppSetting.LoadData()

		txtFrom.Text = oAppSetting.GetValue("FromPath")
		txtRegName.Text = oAppSetting.GetValueDef("RegName", txtRegName.Text)
		txtRegKey.Text = oAppSetting.GetValueDef("RegKey", txtRegKey.Text)
		txtStartFile.Text = oAppSetting.GetValue("StartFile")

		Dim sBarcodeType As String = oAppSetting.GetValue("BarcodeType")
		If sBarcodeType <> "" Then
			cbBarcodeType.SelectedIndex = CInt(sBarcodeType)
		Else
			cbBarcodeType.SelectedIndex = 0
		End If

		Dim sDirection As String = oAppSetting.GetValue("Direction")
		If sDirection <> "" Then
			cbDirection.SelectedIndex = CInt(sDirection)
		Else
			cbDirection.SelectedIndex = 0
		End If

		chkUseBarcodeZones.Checked = oAppSetting.GetValue("UseBarcodeZones") = "1"
		chkLog.Checked = oAppSetting.GetValue("Log") <> "0"
		chkShowTime.Checked = oAppSetting.GetValue("ShowTime") <> "0"
		chkCSV.Checked = oAppSetting.GetValue("CSV") <> "0"

		chkBytescout.Checked = oAppSetting.GetValue("Bytescout") = "1"
		chkFast.Checked = oAppSetting.GetValue("Fast") = "1"

		BytescoutChecked()
	End Sub

	Public Shared Function GetShortName(ByVal sLongFileName As String) As String

		If sLongFileName.Length < 250 Then
			Return sLongFileName
		End If

		Dim lRetVal As Long
		Dim iLen As Integer = 1024
		Dim sShortPathName As System.Text.StringBuilder = New System.Text.StringBuilder(iLen)
		lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
		Dim sRet As String = sShortPathName.ToString()

		If sRet <> "" Then
			Return sRet
		Else
			Return sLongFileName
		End If

	End Function

	Private Sub btnFrom_Click(sender As System.Object, e As System.EventArgs) Handles btnFrom.Click
		fldFrom.SelectedPath = txtFrom.Text
		fldFrom.ShowDialog()
		txtFrom.Text = fldFrom.SelectedPath
	End Sub

	Private Sub btnProcess_Click(sender As System.Object, e As System.EventArgs) Handles btnProcess.Click
		btnProcess.Enabled = False

		Dim sFromPath As String = txtFrom.Text
		If Not Directory.Exists(sFromPath) Then
			btnProcess.Enabled = True
			MsgBox("Folder does not exist")
			Exit Sub
		End If

		If chkCSV.Checked Then
			Dim sLogFileName As String = Now.Month.ToString() & "-" & _
			 Now.Day.ToString() & "-" & _
			 Now.Year.ToString() & "_" & _
			 Now.Hour.ToString() & "-" & _
			 Now.Minute.ToString() & "-" & _
			 Now.Second.ToString() & ".csv"
			Dim sLogFilePath As String = IO.Path.Combine(sFromPath, sLogFileName)
			oLogFile = New System.IO.StreamWriter(sLogFilePath)
		End If

		If txtStartFile.Text <> "" Then
			bStartFile = False
		End If

		txtOutput.Text = ""
		txtOutput.AppendText("Starting..." & vbCrLf)
		ProccessFolder(sFromPath)
		txtOutput.AppendText("Done!")

		btnProcess.Enabled = True

		If chkCSV.Checked Then
			oLogFile.Close()
		End If

	End Sub

	Sub ProccessFolder(ByVal sFolderPath As String)
		Dim sFromPath As String = txtFrom.Text
		Dim oFiles As String() = Directory.GetFiles(sFolderPath)
		ProgressBar1.Maximum = oFiles.Length
		For i As Integer = 0 To oFiles.Length - 1
			Dim sFromFilePath As String = oFiles(i)

			If txtStartFile.Text <> "" Then
				If Trim(LCase(txtStartFile.Text)) = LCase(sFromFilePath) Then
					bStartFile = True
				End If
			End If

			If bStartFile Then
				Dim oFileInfo As New FileInfo(GetShortName(sFromFilePath))
				Dim sExt As String = PadExt(oFileInfo.Extension)

				If sExt = "JPG" Or sExt = "GIF" Or sExt = "PNG" Or sExt = "BMP" Or sExt = "TIF" Then

					If chkBytescout.Checked Then
						Bytescout_ReadBarcode(sFromFilePath)
					Else
						ReadBarcode(sFromFilePath)
					End If

				End If
			End If

			ProgressBar1.Value = i
			Application.DoEvents()
		Next

		ProgressBar1.Value = 0

		Dim oFolders As String() = Directory.GetDirectories(sFolderPath)
		For i As Integer = 0 To oFolders.Length - 1
			Dim sChildFolder As String = oFolders(i)
			Dim iPos As Integer = sChildFolder.LastIndexOf("\")
			Dim sFolderName As String = sChildFolder.Substring(iPos + 1)
			ProccessFolder(sChildFolder)
		Next

	End Sub

	Private Sub Bytescout_ReadBarcode(ByVal sFromFilePath As String)
		'C:\Program Files\Bytescout BarCode Reader SDK\Redistributable\net2.00

		Dim sFromPath As String = txtFrom.Text
		Dim sFileName As String = sFromFilePath.Replace(sFromPath & "\", "")
		Dim dtStart As DateTime = DateTime.Now

		Dim oReader As New Bytescout.BarCodeReader.Reader()
		oReader.RegistrationName = txtRegName.Text
		oReader.RegistrationKey = txtRegKey.Text

		'oReader.BarcodeTypesToFind.All1D = True
		oReader.BarcodeTypesToFind.Code39 = True
		oReader.FastMode = chkFast.Checked
		'oReader.HeuristicMode = False
		'oReader.Orientation = Bytescout.BarCodeReader.OrientationType.HorizontalFromLeftToRight

		Dim oCodes As Bytescout.BarCodeReader.FoundBarcode()
		Try
			oCodes = oReader.ReadFrom(sFromFilePath)
		Catch ex As Exception
			WriteLog(sFromFilePath & "," & "Could not open")
			Exit Sub
		End Try

		Dim sSec As String = ""
		If chkShowTime.Checked Then
			sSec = vbTab & DateTime.Now.Subtract(dtStart).TotalSeconds.ToString("#0.00")
		End If

		For Each oCode As Bytescout.BarCodeReader.FoundBarcode In oCodes
			'barcode.Type, barcode.Value, barcode.Page, barcode.Rect

			If chkLog.Checked Then
				txtOutput.AppendText(sFileName & vbTab & oCode.Value & sSec & vbCrLf)
			End If

			WriteLog(sFileName & Chr(9) & oCode.Value)
		Next

	End Sub

	Private Sub ReadBarcode(sFromFilePath As String)

		Dim sFromPath As String = txtFrom.Text
		Dim sFileName As String = sFromFilePath.Replace(sFromPath & "\", "")

		Dim oImage As System.Drawing.Image = Nothing

		Try
			oImage = System.Drawing.Image.FromFile(sFromFilePath)
		Catch ex As Exception
			If chkLog.Checked Then
				txtOutput.AppendText(sFileName & vbTab & "Could not open" & vbCrLf)
			End If

			WriteLog(sFileName & "," & "Could not open")
			Exit Sub
		End Try

		Dim barcodes As New System.Collections.ArrayList
		Dim iScans As Integer = 100
		Dim dtStart As DateTime = DateTime.Now

		BarcodeImaging.UseBarcodeZones = chkUseBarcodeZones.Checked

		Dim oBarcodeType As BarcodeImaging.BarcodeType = BarcodeImaging.BarcodeType.All
		Select Case cbBarcodeType.Text
			Case "Code39"
				oBarcodeType = BarcodeImaging.BarcodeType.Code39
			Case "Code128"
				oBarcodeType = BarcodeImaging.BarcodeType.Code128
			Case "EAN"
				oBarcodeType = BarcodeImaging.BarcodeType.EAN
		End Select

		Select Case cbDirection.Text
			Case "All"
				BarcodeImaging.FullScanBarcodeTypes = oBarcodeType
				BarcodeImaging.FullScanPage(barcodes, oImage, iScans)
			Case "Vertical"
				BarcodeImaging.ScanPage(barcodes, oImage, iScans, BarcodeImaging.ScanDirection.Horizontal, oBarcodeType)
			Case "Horizontal"
				BarcodeImaging.ScanPage(barcodes, oImage, iScans, BarcodeImaging.ScanDirection.Vertical, oBarcodeType)
		End Select

		Dim sSec As String = ""
		If chkShowTime.Checked Then
			sSec = vbTab & DateTime.Now.Subtract(dtStart).TotalSeconds.ToString("#0.00")
		End If

		If barcodes.Count = 0 Then

			If chkLog.Checked Then
				txtOutput.AppendText(sFileName & vbTab & "Failed" & sSec & vbCrLf)
			End If

			WriteLog(sFileName & "," & "Failed")
		Else
			For Each bc As Object In barcodes

				If chkLog.Checked Then
					txtOutput.AppendText(sFileName & vbTab & bc & sSec & vbCrLf)
				End If

				WriteLog(sFileName & "," & bc)
			Next
		End If

		oImage.Dispose()
	End Sub



	Private Sub WriteLog(sLine As String)
		If chkCSV.Checked Then
			oLogFile.WriteLine(sLine)
		End If
	End Sub

	Private Function PadExt(ByVal s As String) As String
		s = UCase(s)
		If s.Length > 3 Then
			s = s.Substring(1, 3)
		End If
		Return s
	End Function

	Private Sub chkBytescout_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkBytescout.CheckedChanged
		BytescoutChecked()
	End Sub

	Private Sub BytescoutChecked()
		pnCS.Visible = chkBytescout.Checked = False
		pnByteScout.Visible = chkBytescout.Checked
	End Sub

	Private Sub LinkLabel1_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
		Process.Start("https://bytescout.com/download/web-installer")
	End Sub
End Class
