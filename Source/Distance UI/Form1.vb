' ---------------------------------------------------------------------------------------
' Safety System Calculator
' Author: Tamara Myers
' Description: Windows Forms application to calculate minimum safe distances for
' automation robotics based on ISO 13855 standard. Reads actuator and safety device
' specifications from an Excel file. Designed for use on shared network drive environments.
' ---------------------------------------------------------------------------------------

Imports System
Imports System.IO
Imports System.Windows.Forms
Imports OfficeOpenXml
Imports System.Diagnostics

Public Class Form1
    Inherits System.Windows.Forms.Form

    ' Path to the Excel data file (placed alongside EXE)
    Private ReadOnly ExcelFilePath As String = Path.Combine(Application.StartupPath, "Safe_Distance_Data.xlsx")

    ' UI Elements
    Private lightCurtainSelection1, lightCurtainSelection2, lightCurtainSelection3 As ComboBox
    Private syncCheckbox, useSafetyControllerCheckbox As CheckBox
    Private safetyPLCSelection As ComboBox
    Private actuatorList As CheckedListBox
    Private btnCalculate, btnOpenExcel As Button
    Private resultsListBox As ListBox
    Private lblLightCurtainC, lblLightCurtainT, lblSafetyControllerT, lblSafetyPLCT, lblTotalT, lblTotalC, lblKValue, lblEquation As Label

    ' Constants
    Private Const K_VALUE As Integer = 2000 ' mm/s hand approach speed
    Private WithEvents refreshTimer As New Timer With {.Interval = 2000} ' Refresh every 2 seconds

    ' Form Initialization
    Public Sub New()
        InitializeComponent()
        Me.StartPosition = FormStartPosition.CenterScreen
        LoadPartsData()
        refreshTimer.Start()
    End Sub

    ' Setup UI Components
    Private Sub InitializeComponent()
        Me.ClientSize = New System.Drawing.Size(450, 950)
        Me.Text = "Safety System Calculator"

        ' Open Excel Button
        btnOpenExcel = CreateButton("Open Excel File", 10, 10, 420, 30, AddressOf OpenExcelFile)
        Me.Controls.Add(btnOpenExcel)

        ' Light Curtain Selection Group
        Dim lightCurtainGroup = CreateGroupBox("Light Curtain Selection", 10, 50, 420, 210)
        lightCurtainSelection1 = CreateComboBox(10, 20, 390)
        lightCurtainSelection2 = CreateComboBox(10, 50, 390)
        lightCurtainSelection3 = CreateComboBox(10, 80, 390)
        syncCheckbox = New CheckBox With {.Location = New Point(10, 110), .Size = New Size(250, 20), .Text = "Using Synchronized Wire? (Subtracts 3ms)"}
        lblLightCurtainC = CreateLabel("C: 0 mm", 10, 140)
        lblLightCurtainT = CreateLabel("T: 0 ms", 10, 165) ' moved down slightly
        AddHandler syncCheckbox.CheckedChanged, AddressOf UpdateOverviewOnly
        lightCurtainGroup.Controls.AddRange({lightCurtainSelection1, lightCurtainSelection2, lightCurtainSelection3, syncCheckbox, lblLightCurtainC, lblLightCurtainT})
        Me.Controls.Add(lightCurtainGroup)

        ' Safety Controller Group
        Dim safetyControllerGroup = CreateGroupBox("Safety Controller", 10, 270, 420, 80)
        useSafetyControllerCheckbox = New CheckBox With {.Location = New Point(10, 20), .Size = New Size(200, 20), .Text = "Use Safety Controller?"}
        lblSafetyControllerT = CreateLabel("T: 0 ms", 220, 20)
        AddHandler useSafetyControllerCheckbox.CheckedChanged, AddressOf UpdateOverviewOnly
        safetyControllerGroup.Controls.AddRange({useSafetyControllerCheckbox, lblSafetyControllerT})
        Me.Controls.Add(safetyControllerGroup)

        ' Safety PLC Group
        Dim safetyPLCGroup = CreateGroupBox("Safety PLC Selection", 10, 360, 420, 80)
        safetyPLCSelection = CreateComboBox(10, 20, 390)
        lblSafetyPLCT = CreateLabel("T: 0 ms", 220, 50)
        safetyPLCGroup.Controls.AddRange({safetyPLCSelection, lblSafetyPLCT})
        Me.Controls.Add(safetyPLCGroup)

        ' Actuator Selection Group
        Dim actuatorGroup = CreateGroupBox("Actuator Selection", 10, 450, 420, 150)
        actuatorList = New CheckedListBox With {.Location = New Point(10, 20), .Size = New Size(390, 100), .CheckOnClick = True}
        actuatorGroup.Controls.Add(actuatorList)
        Me.Controls.Add(actuatorGroup)

        ' Final Calculation Overview Group
        Dim finalCalcGroup = CreateGroupBox("Final Calculation Overview", 10, 610, 420, 120)
        lblTotalT = CreateLabel("Total T: 0 ms", 10, 20)
        lblTotalC = CreateLabel("Total C: 0 mm", 10, 40)
        lblKValue = CreateLabel($"K Value: {K_VALUE} mm/s", 10, 60)
        lblEquation = CreateLabel("Equation: S = (K × T) + C", 10, 80)
        finalCalcGroup.Controls.AddRange({lblTotalT, lblTotalC, lblKValue, lblEquation})
        Me.Controls.Add(finalCalcGroup)

        ' Calculate Safe Distance Button
        btnCalculate = CreateButton("Calculate Safe Distance", 10, 740, 420, 30, AddressOf CalculateSafeDistances)
        Me.Controls.Add(btnCalculate)

        ' Results List
        Dim lblResults = New Label With {.Location = New Point(10, 780), .Size = New Size(420, 20), .Text = "Safe Distances:"}
        Me.Controls.Add(lblResults)

        resultsListBox = New ListBox With {.Location = New Point(10, 810), .Size = New Size(420, 120)}
        resultsListBox.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        Me.Controls.Add(resultsListBox)
    End Sub


    ' Create a new GroupBox UI element
    Private Function CreateGroupBox(title As String, x As Integer, y As Integer, width As Integer, height As Integer) As GroupBox
        Return New GroupBox With {.Text = title, .Location = New Point(x, y), .Size = New Size(width, height)}
    End Function

    ' Create a new ComboBox UI element
    Private Function CreateComboBox(x As Integer, y As Integer, width As Integer) As ComboBox
        Return New ComboBox With {.Location = New Point(x, y), .Size = New Size(width, 21)}
    End Function

    ' Create a new Label UI element
    Private Function CreateLabel(text As String, x As Integer, y As Integer) As Label
        Return New Label With {.Text = text, .Location = New Point(x, y), .Size = New Size(250, 20)}
    End Function

    ' Create a new Button UI element
    Private Function CreateButton(text As String, x As Integer, y As Integer, width As Integer, height As Integer, handler As EventHandler) As Button
        Dim btn = New Button With {.Text = text, .Location = New Point(x, y), .Size = New Size(width, height)}
        AddHandler btn.Click, handler
        Return btn
    End Function

    ' Load all parts data from Excel file
    Private Sub LoadPartsData()
        If Not File.Exists(ExcelFilePath) Then Exit Sub
        Try
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial
            Using package As New ExcelPackage(New FileInfo(ExcelFilePath))
                PopulateComboBox(package, "Safety PLCs", safetyPLCSelection)
                PopulateComboBox(package, "Light Curtains", lightCurtainSelection1, lightCurtainSelection2, lightCurtainSelection3)
                PopulateCheckedListBox(package, "Actuators", actuatorList)
            End Using
        Catch ex As Exception
            MessageBox.Show($"Failed to load Excel data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Populate ComboBox options from Excel sheet
    Private Sub PopulateComboBox(package As ExcelPackage, sheetName As String, ParamArray comboBoxes() As ComboBox)
        Dim worksheet = GetSheet(package, sheetName)
        If worksheet Is Nothing Then Exit Sub

        Dim partNames = New List(Of String)
        For row = 2 To worksheet.Dimension.End.Row
            Dim partName = worksheet.Cells(row, 1).Value?.ToString()
            If Not String.IsNullOrEmpty(partName) Then partNames.Add(partName)
        Next

        For Each comboBox In comboBoxes
            Dim selected = comboBox.SelectedItem
            comboBox.Items.Clear()
            comboBox.Items.AddRange(partNames.ToArray())
            comboBox.SelectedItem = selected
        Next
    End Sub

    ' Populate CheckedListBox options from Excel sheet
    Private Sub PopulateCheckedListBox(package As ExcelPackage, sheetName As String, checkedListBox As CheckedListBox)
        Dim worksheet = GetSheet(package, sheetName)
        If worksheet Is Nothing Then Exit Sub

        Dim checked = checkedListBox.CheckedItems.Cast(Of String).ToList()
        checkedListBox.Items.Clear()
        For row = 2 To worksheet.Dimension.End.Row
            Dim actuatorName = worksheet.Cells(row, 1).Value?.ToString()
            If Not String.IsNullOrEmpty(actuatorName) Then
                checkedListBox.Items.Add(actuatorName, checked.Contains(actuatorName))
            End If
        Next
    End Sub

    ' Retrieve a specific worksheet by name
    Private Function GetSheet(package As ExcelPackage, name As String) As ExcelWorksheet
        Return package.Workbook.Worksheets.FirstOrDefault(Function(ws) ws.Name.ToLower().Contains(name.ToLower()))
    End Function

    ' Find the row number by value in column 1
    Private Function FindRowByValue(worksheet As ExcelWorksheet, value As String) As Integer
        Return Enumerable.Range(2, worksheet.Dimension.End.Row - 1).FirstOrDefault(Function(r) worksheet.Cells(r, 1).Value?.ToString() = value)
    End Function

    ' Update display values for overview without running calculations
    Private Sub UpdateOverviewOnly()
        Dim totalLightCurtainT As Double = 0
        Dim totalC As Double = 0
        Dim safetyControllerT As Double = 0
        Dim safetyPLCT As Double = 0

        ' Reset labels early
        lblLightCurtainC.Text = "C: 0 mm"
        lblLightCurtainT.Text = "T: 0 ms"
        lblSafetyControllerT.Text = "T: 0 ms"
        lblSafetyPLCT.Text = "T: 0 ms"
        lblTotalT.Text = "Total T: 0 ms"
        lblTotalC.Text = "Total C: 0 mm"

        If Not File.Exists(ExcelFilePath) Then Exit Sub
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Using package As New ExcelPackage(New FileInfo(ExcelFilePath))
            Dim lightCurtainSheet = GetSheet(package, "Light Curtains")
            Dim safetyControllerSheet = GetSheet(package, "Safety Controllers")
            Dim safetyPLCSheet = GetSheet(package, "Safety PLCs")

            ' Light Curtain Calculation
            If lightCurtainSheet IsNot Nothing Then
                Dim selectedCurtains = {lightCurtainSelection1, lightCurtainSelection2, lightCurtainSelection3}.
                Where(Function(cb) cb.SelectedIndex <> -1).Select(Function(cb) cb.SelectedItem.ToString()).ToList()

                ' If curtains are selected, sum them
                If selectedCurtains.Count > 0 Then
                    totalC = selectedCurtains.Sum(Function(lc) Convert.ToDouble(lightCurtainSheet.Cells(FindRowByValue(lightCurtainSheet, lc), 2).Value))
                    totalLightCurtainT = selectedCurtains.Sum(Function(lc) Convert.ToDouble(lightCurtainSheet.Cells(FindRowByValue(lightCurtainSheet, lc), 3).Value))

                    ' Only apply sync wire subtraction if there were valid light curtains selected
                    If syncCheckbox.Checked Then
                        totalLightCurtainT = Math.Max(0, totalLightCurtainT - 3)
                    End If
                End If

                ' Always update these even if 0
                lblLightCurtainC.Text = $"C: {totalC} mm"
                lblLightCurtainT.Text = $"T: {totalLightCurtainT} ms"
            End If

            ' Safety Controller T
            If useSafetyControllerCheckbox.Checked AndAlso safetyControllerSheet IsNot Nothing Then
                safetyControllerT = Enumerable.Range(2, safetyControllerSheet.Dimension.End.Row - 1).
                Sum(Function(row) Convert.ToDouble(safetyControllerSheet.Cells(row, 2).Value))
            End If
            lblSafetyControllerT.Text = $"T: {safetyControllerT} ms"

            ' Safety PLC T
            If safetyPLCSheet IsNot Nothing AndAlso safetyPLCSelection.SelectedIndex <> -1 Then
                Dim row = FindRowByValue(safetyPLCSheet, safetyPLCSelection.SelectedItem.ToString())
                If row <> 0 Then
                    safetyPLCT = Convert.ToDouble(safetyPLCSheet.Cells(row, 2).Value)
                End If
            End If
            lblSafetyPLCT.Text = $"T: {safetyPLCT} ms"

            ' Final Overview
            Dim totalT = totalLightCurtainT + safetyControllerT + safetyPLCT
            lblTotalT.Text = $"Total T: {totalT} ms"
            lblTotalC.Text = $"Total C: {totalC} mm"
        End Using
    End Sub



    ' Calculate final safe distances when button pressed
    Private Sub CalculateSafeDistances(sender As Object, e As EventArgs)
        resultsListBox.Items.Clear()

        Dim totalLightCurtainT As Double = 0
        Dim totalC As Double = 0
        Dim safetyControllerT As Double = 0
        Dim safetyPLCT As Double = 0

        If Not File.Exists(ExcelFilePath) Then Exit Sub
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Using package As New ExcelPackage(New FileInfo(ExcelFilePath))
            Dim lightCurtainSheet = GetSheet(package, "Light Curtains")
            Dim safetyControllerSheet = GetSheet(package, "Safety Controllers")
            Dim safetyPLCSheet = GetSheet(package, "Safety PLCs")
            Dim actuatorSheet = GetSheet(package, "Actuators")

            ' Light Curtain Calculation
            If lightCurtainSheet IsNot Nothing Then
                Dim selectedCurtains = {lightCurtainSelection1, lightCurtainSelection2, lightCurtainSelection3}.
                Where(Function(cb) cb.SelectedIndex <> -1).Select(Function(cb) cb.SelectedItem.ToString()).ToList()

                totalC = selectedCurtains.Sum(Function(lc) Convert.ToDouble(lightCurtainSheet.Cells(FindRowByValue(lightCurtainSheet, lc), 2).Value))
                totalLightCurtainT = selectedCurtains.Sum(Function(lc) Convert.ToDouble(lightCurtainSheet.Cells(FindRowByValue(lightCurtainSheet, lc), 3).Value))
                If syncCheckbox.Checked Then
                    totalLightCurtainT = Math.Max(0, totalLightCurtainT - 3)
                End If
            End If

            ' Safety Controller T
            If useSafetyControllerCheckbox.Checked AndAlso safetyControllerSheet IsNot Nothing Then
                safetyControllerT = Enumerable.Range(2, safetyControllerSheet.Dimension.End.Row - 1).
                Sum(Function(row) Convert.ToDouble(safetyControllerSheet.Cells(row, 2).Value))
            End If

            ' Safety PLC T
            If safetyPLCSheet IsNot Nothing AndAlso safetyPLCSelection.SelectedIndex <> -1 Then
                Dim row = FindRowByValue(safetyPLCSheet, safetyPLCSelection.SelectedItem.ToString())
                If row <> 0 Then safetyPLCT = Convert.ToDouble(safetyPLCSheet.Cells(row, 2).Value)
            End If

            ' Calculate distances per actuator
            If actuatorSheet IsNot Nothing Then
                For Each act In actuatorList.CheckedItems.Cast(Of String)()
                    Dim row = FindRowByValue(actuatorSheet, act)
                    If row <> 0 Then
                        Dim actuatorT = Convert.ToDouble(actuatorSheet.Cells(row, 2).Value)
                        Dim totalT = totalLightCurtainT + safetyControllerT + safetyPLCT + actuatorT

                        ' Calculate safe distance
                        Dim distanceMM = (K_VALUE * (totalT / 1000)) + totalC
                        Dim distanceIN = Math.Ceiling((distanceMM / 25.4) * 4) / 4

                        ' Enforce 4-inch minimum
                        If distanceIN < 4 Then distanceIN = 4

                        resultsListBox.Items.Add($"{act} ➤ {Math.Round(distanceMM)} mm / {distanceIN} in")
                    End If
                Next
            End If
        End Using
    End Sub


    ' Handle "Open Excel File" button click safely
    Private Sub OpenExcelFile(sender As Object, e As EventArgs)
        If File.Exists(ExcelFilePath) Then
            Process.Start(New ProcessStartInfo(ExcelFilePath) With {.UseShellExecute = True})
        Else
            MessageBox.Show("Could not find the Excel file at:" & vbCrLf & ExcelFilePath, "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    ' Refresh parts data every few seconds
    Private Sub refreshTimer_Tick(sender As Object, e As EventArgs) Handles refreshTimer.Tick
        LoadPartsData()
        UpdateOverviewOnly()
    End Sub

    ' Warn immediately if parts_data.xlsx missing on launch
    Protected Overrides Sub OnLoad(e As EventArgs)
        MyBase.OnLoad(e)
        If Not File.Exists(ExcelFilePath) Then
            MessageBox.Show("Warning: The Safe_Distance_Data.xlsx file could not be found." & vbCrLf & "Some functionality may not work until this is fixed." & vbCrLf & "Expected path: " & ExcelFilePath, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

End Class
