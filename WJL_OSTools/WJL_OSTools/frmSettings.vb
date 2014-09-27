Imports System.Windows.Forms
Imports System.IO

Public Class frmSettings
    Const ExportExtension As String = "mps"

    Dim Initialising As Boolean = False
    Dim UnsavedChanges As Boolean = False
    Dim myStatus As String = ""

    Dim StringDelimiter As Char

    Dim strSearchPathsMM500 As String
    Dim strFreezeLayersMM500 As String
    Dim intColourIndexMM500 As Integer

    Dim strSearchPathsMM1000 As String
    Dim strFreezeLayersMM1000 As String
    Dim intColourIndexMM1000 As Integer

    Dim booUseMM500 As Boolean
    Dim booUse50kNew As Boolean

    Dim strSearchPathsSV As String
    Dim strSearchPaths50kNew As String
    Dim strSearchPaths50kOld As String

    ReadOnly Property Status As String
        Get
            Return myStatus
        End Get
    End Property

    Sub New()

        '' This call is required by the designer.
        InitializeComponent()
        GetSettings()
    End Sub

    Private Sub frmSettings_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Initialise()
    End Sub

    Private Sub Initialise()
        Initialising = True
        AddListItemsFromString(lstSearchPathsMM500, strSearchPathsMM500, StringDelimiter)
        AddListItemsFromString(lstLayerFreezeMM500, strFreezeLayersMM500, StringDelimiter)
        numColourMM500.Value = intColourIndexMM500

        If booUseMM500 Then
            AddListItemsFromString(lstSearchPathsMM1000, strSearchPathsMM500, StringDelimiter)
            AddListItemsFromString(lstLayerFreezeMM1000, strFreezeLayersMM500, StringDelimiter)
            numColourMM500.Value = intColourIndexMM500
        Else
            AddListItemsFromString(lstSearchPathsMM1000, strSearchPathsMM1000, StringDelimiter)
            AddListItemsFromString(lstLayerFreezeMM1000, strFreezeLayersMM1000, StringDelimiter)
            numColourMM500.Value = intColourIndexMM1000
        End If

        AddListItemsFromString(lstSearchPathsSV, strSearchPathsSV, StringDelimiter)
        AddListItemsFromString(lstSearchPaths50kNew, strSearchPaths50kNew, StringDelimiter)

        If booUse50kNew Then
            AddListItemsFromString(lstSearchPaths50kOld, strSearchPaths50kNew, StringDelimiter)
        Else
            AddListItemsFromString(lstSearchPaths50kOld, strSearchPaths50kOld, StringDelimiter)
        End If

        Initialising = False
    End Sub

    Sub GetSettings()

        StringDelimiter = My.Settings.StringDelimiter
        booUseMM500 = My.Settings.UseMM500
        booUse50kNew = My.Settings.Use50kNew

        strSearchPathsMM500 = My.Settings.SearchPathsMM500
        strFreezeLayersMM500 = My.Settings.LayerFreezeListMM500
        intColourIndexMM500 = My.Settings.LayerColourIndexMM500

        strSearchPathsMM1000 = My.Settings.SearchPathsMM1000
        strFreezeLayersMM1000 = My.Settings.LayerFreezeListMM1000
        intColourIndexMM1000 = My.Settings.LayerColourIndexMM1000

        strSearchPathsSV = My.Settings.SearchPathsSV
        strSearchPaths50kNew = My.Settings.SearchPaths50kNew
        strSearchPaths50kOld = My.Settings.SearchPaths50kOld

    End Sub

    Sub SetSettings()

        strSearchPathsMM500 = GetListAsString(lstSearchPathsMM500, StringDelimiter)
        My.Settings.SearchPathsMM500 = strSearchPathsMM500
        My.Settings.LayerFreezeListMM500 = strFreezeLayersMM500
        My.Settings.LayerColourIndexMM500 = intColourIndexMM500

        strSearchPathsMM1000 = GetListAsString(lstSearchPathsMM1000, StringDelimiter)
        My.Settings.SearchPathsMM1000 = strSearchPathsMM1000
        My.Settings.LayerFreezeListMM1000 = strFreezeLayersMM1000
        My.Settings.LayerColourIndexMM1000 = intColourIndexMM1000
        
        strSearchPathsSV = GetListAsString(lstSearchPathsSV, StringDelimiter)
        My.Settings.SearchPathsSV = strSearchPathsSV

        strSearchPaths50kNew = GetListAsString(lstSearchPaths50kNew, StringDelimiter)
        My.Settings.SearchPaths50kNew = strSearchPaths50kNew

        strSearchPaths50kOld = GetListAsString(lstSearchPaths50kOld, StringDelimiter)
        My.Settings.SearchPaths50kOld = strSearchPaths50kOld
        
        My.Settings.Save()
    End Sub

#Region "Form Controls"

#Region "Ok, Cancel, Apply"
    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        If UnsavedChanges Then SetSettings()
        myStatus = "OK"
        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        My.Settings.Reload()
        myStatus = "Cancel"
        Me.Close()
    End Sub

    Private Sub cmdApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdApply.Click
        SetSettings()
        UnsavedChanges = False
        cmdApply.Enabled = False
    End Sub

#End Region

#Region "Search Paths"

#Region "MultiMap500"
    Private Sub cmdAddSearchPathMM500_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddSearchPathMM500.Click
        Dim FolderBrowser As FolderBrowserDialog = New FolderBrowserDialog
        FolderBrowser.RootFolder = Environment.SpecialFolder.MyComputer
        FolderBrowser.ShowNewFolderButton = False

        Dim SelectedPath As Object
        If FolderBrowser.ShowDialog() = DialogResult.OK Then
            SelectedPath = FolderBrowser.SelectedPath
            InsertItem(lstSearchPathsMM500, SelectedPath)
            UnsavedChanges = True
            cmdApply.Enabled = True

        End If
    End Sub

    Private Sub cmdRemoveSearchPathMM500_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemoveSearchPathMM500.Click
        Dim SelectedItems As New List(Of Object)
        For Each Item As Object In lstSearchPathsMM500.SelectedItems
            SelectedItems.Add(Item)
        Next
        For Each SelectedItem As Object In SelectedItems
            lstSearchPathsMM500.Items.Remove(SelectedItem)
            UnsavedChanges = True
            cmdApply.Enabled = True

        Next

    End Sub

    Private Sub cmdMoveSearchPathUpMM500_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMoveSearchPathUpMM500.Click
        MoveSelectedItemUp(lstSearchPathsMM500)
        UnsavedChanges = True
    End Sub

    Private Sub cmdMoveSearchPathDownMM500_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMoveSearchPathDownMM500.Click
        MoveSelectedItemDown(lstSearchPathsMM500)
        UnsavedChanges = True
    End Sub
#End Region

#Region "MultiMap1000"
    Private Sub cmdAddSearchPathMM1000_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddSearchPathMM1000.Click
        Dim FolderBrowser As FolderBrowserDialog = New FolderBrowserDialog
        FolderBrowser.RootFolder = Environment.SpecialFolder.MyComputer
        FolderBrowser.ShowNewFolderButton = False

        Dim SelectedPath As Object
        If FolderBrowser.ShowDialog() = DialogResult.OK Then
            SelectedPath = FolderBrowser.SelectedPath
            InsertItem(lstSearchPathsMM1000, SelectedPath)
            UnsavedChanges = True
            cmdApply.Enabled = True
        End If
    End Sub

    Private Sub cmdRemoveSearchPathMM1000_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemoveSearchPathMM1000.Click
        Dim SelectedItems As New List(Of Object)
        For Each Item As Object In lstSearchPathsMM1000.SelectedItems
            SelectedItems.Add(Item)
        Next
        For Each SelectedItem As Object In SelectedItems
            lstSearchPathsMM1000.Items.Remove(SelectedItem)
            UnsavedChanges = True
            cmdApply.Enabled = True
        Next

    End Sub

    Private Sub cmdMoveSearchPathUpMM1000_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMoveSearchPathUpMM1000.Click
        MoveSelectedItemUp(lstSearchPathsMM1000)
        UnsavedChanges = True
    End Sub

    Private Sub cmdMoveSearchPathDownMM1000_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMoveSearchPathDownMM1000.Click
        MoveSelectedItemDown(lstSearchPathsMM1000)
        UnsavedChanges = True
    End Sub
#End Region

#Region "StreetView"
    Private Sub cmdAddSearchPathSV_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddSearchPathSV.Click
        Dim FolderBrowser As FolderBrowserDialog = New FolderBrowserDialog
        FolderBrowser.RootFolder = Environment.SpecialFolder.MyComputer
        FolderBrowser.ShowNewFolderButton = False

        Dim SelectedPath As Object
        If FolderBrowser.ShowDialog() = DialogResult.OK Then
            SelectedPath = FolderBrowser.SelectedPath
            InsertItem(lstSearchPathsSV, SelectedPath)
            UnsavedChanges = True
            cmdApply.Enabled = True
        End If
    End Sub

    Private Sub cmdRemoveSearchPathSV_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemoveSearchPathSV.Click
        Dim SelectedItems As New List(Of Object)
        For Each Item As Object In lstSearchPathsSV.SelectedItems
            SelectedItems.Add(Item)
        Next
        For Each SelectedItem As Object In SelectedItems
            lstSearchPathsSV.Items.Remove(SelectedItem)
            UnsavedChanges = True
            cmdApply.Enabled = True
        Next

    End Sub

    Private Sub cmdMoveSearchPathUpSV_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMoveSearchPathUpSV.Click
        MoveSelectedItemUp(lstSearchPathsSV)
        UnsavedChanges = True
    End Sub

    Private Sub cmdMoveSearchPathDownSV_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMoveSearchPathDownSV.Click
        MoveSelectedItemDown(lstSearchPathsSV)
        UnsavedChanges = True
    End Sub
#End Region

#Region "50kNew"
    Private Sub cmdAddSearchPath50kNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddSearchPath50kNew.Click
        Dim FolderBrowser As FolderBrowserDialog = New FolderBrowserDialog
        FolderBrowser.RootFolder = Environment.SpecialFolder.MyComputer
        FolderBrowser.ShowNewFolderButton = False

        Dim SelectedPath As Object
        If FolderBrowser.ShowDialog() = DialogResult.OK Then
            SelectedPath = FolderBrowser.SelectedPath
            InsertItem(lstSearchPaths50kNew, SelectedPath)
            UnsavedChanges = True
            cmdApply.Enabled = True
        End If
    End Sub

    Private Sub cmdRemoveSearchPath50kNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemoveSearchPath50kNew.Click
        Dim SelectedItems As New List(Of Object)
        For Each Item As Object In lstSearchPaths50kNew.SelectedItems
            SelectedItems.Add(Item)
        Next
        For Each SelectedItem As Object In SelectedItems
            lstSearchPaths50kNew.Items.Remove(SelectedItem)
            UnsavedChanges = True
            cmdApply.Enabled = True
        Next

    End Sub

    Private Sub cmdMoveSearchPathUp50kNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMoveSearchPathUp50kNew.Click
        MoveSelectedItemUp(lstSearchPaths50kNew)
        UnsavedChanges = True
    End Sub

    Private Sub cmdMoveSearchPathDown50kNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMoveSearchPathDown50kNew.Click
        MoveSelectedItemDown(lstSearchPaths50kNew)
        UnsavedChanges = True
    End Sub
#End Region

#Region "50kOld"
    Private Sub cmdAddSearchPath50kOld_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddSearchPath50kOld.Click
        Dim FolderBrowser As FolderBrowserDialog = New FolderBrowserDialog
        FolderBrowser.RootFolder = Environment.SpecialFolder.MyComputer
        FolderBrowser.ShowNewFolderButton = False

        Dim SelectedPath As Object
        If FolderBrowser.ShowDialog() = DialogResult.OK Then
            SelectedPath = FolderBrowser.SelectedPath
            InsertItem(lstSearchPaths50kOld, SelectedPath)
            UnsavedChanges = True
            cmdApply.Enabled = True
        End If
    End Sub

    Private Sub cmdRemoveSearchPath50kOld_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemoveSearchPath50kOld.Click
        Dim SelectedItems As New List(Of Object)
        For Each Item As Object In lstSearchPaths50kOld.SelectedItems
            SelectedItems.Add(Item)
        Next
        For Each SelectedItem As Object In SelectedItems
            lstSearchPaths50kOld.Items.Remove(SelectedItem)
            UnsavedChanges = True
            cmdApply.Enabled = True
        Next

    End Sub

    Private Sub cmdMoveSearchPathUp50kOld_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMoveSearchPathUp50kOld.Click
        MoveSelectedItemUp(lstSearchPaths50kOld)
        UnsavedChanges = True
    End Sub

    Private Sub cmdMoveSearchPathDown50kOld_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMoveSearchPathDown50kOld.Click
        MoveSelectedItemDown(lstSearchPaths50kOld)
        UnsavedChanges = True
    End Sub
#End Region

#End Region

#Region "Freeze Layers"

#Region "MultiMap500"
    Private Sub cmdAddLayerFreezeMM500_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddLayerFreezeMM500.Click
        Dim LayerName As String = txtLayerFreezeMM500.Text 'This needs validating
        If LayerName <> "" Then
            InsertItem(lstLayerFreezeMM500, LayerName)
            UnsavedChanges = True
            cmdApply.Enabled = True
        End If
    End Sub

    Private Sub cmdRemoveLayerFreezeMM500_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemoveLayerFreezeMM500.Click
        Dim SelectedItems As New List(Of Object)
        For Each Item As Object In lstLayerFreezeMM500.SelectedItems
            SelectedItems.Add(Item)
        Next
        For Each SelectedItem As Object In SelectedItems
            lstLayerFreezeMM500.Items.Remove(SelectedItem)
            UnsavedChanges = True
            cmdApply.Enabled = True
        Next
    End Sub

    Private Sub cmdMoveLayerFreezeUpMM500_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMoveLayerFreezeUpMM500.Click
        MoveSelectedItemUp(lstLayerFreezeMM500)
        UnsavedChanges = True
    End Sub

    Private Sub cmdMoveLayerFreezeDownMM500_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMoveLayerFreezeDownMM500.Click
        MoveSelectedItemDown(lstLayerFreezeMM500)
        UnsavedChanges = True
    End Sub
#End Region

#Region "MultiMap1000"
    Private Sub cmdAddLayerFreezeMM1000_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddLayerFreezeMM1000.Click
        Dim LayerName As String = txtLayerFreezeMM1000.Text 'This needs validating
        If LayerName <> "" Then
            InsertItem(lstLayerFreezeMM1000, LayerName)
            UnsavedChanges = True
            cmdApply.Enabled = True
        End If
    End Sub

    Private Sub cmdRemoveLayerFreezeMM1000_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemoveLayerFreezeMM1000.Click
        Dim SelectedItems As New List(Of Object)
        For Each Item As Object In lstLayerFreezeMM1000.SelectedItems
            SelectedItems.Add(Item)
        Next
        For Each SelectedItem As Object In SelectedItems
            lstLayerFreezeMM1000.Items.Remove(SelectedItem)
            UnsavedChanges = True
            cmdApply.Enabled = True
        Next
    End Sub

    Private Sub cmdMoveLayerFreezeUpMM1000_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMoveLayerFreezeUpMM1000.Click
        MoveSelectedItemUp(lstLayerFreezeMM1000)
        UnsavedChanges = True
    End Sub

    Private Sub cmdMoveLayerFreezeDownMM1000_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMoveLayerFreezeDownMM1000.Click
        MoveSelectedItemDown(lstLayerFreezeMM1000)
        UnsavedChanges = True
    End Sub
#End Region

#End Region

#Region "Map Colour"

#Region "MultiMap500"
    Private Sub cmdColourPickMM500_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdColourPickMM500.Click
        Dim ColorDialog As New Autodesk.AutoCAD.Windows.ColorDialog()
        ColorDialog.SetDialogTabs(Autodesk.AutoCAD.Windows.ColorDialog.ColorTabs.ACITab)
        ColorDialog.IncludeByBlockByLayer = False

        Dim PickedColourIndex As Integer
        If ColorDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim ColourString As String = ColorDialog.Color.ToString
            Select Case ColourString
                Case "red"
                    PickedColourIndex = 1
                Case "yellow"
                    PickedColourIndex = 2
                Case "green"
                    PickedColourIndex = 3
                Case "cyan"
                    PickedColourIndex = 4
                Case "blue"
                    PickedColourIndex = 5
                Case "magenta"
                    PickedColourIndex = 6
                Case "white"
                    PickedColourIndex = 7
                Case Else
                    Try
                        PickedColourIndex = CInt(ColorDialog.Color.ToString)
                    Catch ex As Exception

                    End Try
            End Select

            If Not PickedColourIndex = Nothing Then
                ''@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                ''needs to convert named colours to index
                numColourMM500.Value = PickedColourIndex
                ''@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
            End If
            UnsavedChanges = True
        End If
    End Sub

    Private Sub numColourMM500_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numColourMM500.ValueChanged
        If intColourIndexMM500 <> numColourMM500.Value Then
            intColourIndexMM500 = numColourMM500.Value
            UnsavedChanges = True
            cmdApply.Enabled = True
        End If

    End Sub
#End Region

#Region "MultiMap1000"
    Private Sub cmdColourPickMM1000_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdColourPickMM1000.Click
        Dim ColorDialog As New Autodesk.AutoCAD.Windows.ColorDialog()
        ColorDialog.SetDialogTabs(Autodesk.AutoCAD.Windows.ColorDialog.ColorTabs.ACITab)
        ColorDialog.IncludeByBlockByLayer = False

        Dim PickedColourIndex As Integer
        If ColorDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim ColourString As String = ColorDialog.Color.ToString
            Select Case ColourString
                Case "red"
                    PickedColourIndex = 1
                Case "yellow"
                    PickedColourIndex = 2
                Case "green"
                    PickedColourIndex = 3
                Case "cyan"
                    PickedColourIndex = 4
                Case "blue"
                    PickedColourIndex = 5
                Case "magenta"
                    PickedColourIndex = 6
                Case "white"
                    PickedColourIndex = 7
                Case Else
                    Try
                        PickedColourIndex = CInt(ColorDialog.Color.ToString)
                    Catch ex As Exception

                    End Try
            End Select

            If Not PickedColourIndex = Nothing Then
                ''@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                ''needs to convert named colours to index
                numColourMM1000.Value = PickedColourIndex
                ''@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
            End If
            UnsavedChanges = True
        End If
    End Sub

    Private Sub numColourMM1000_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numColourMM1000.ValueChanged
        If intColourIndexMM1000 <> numColourMM1000.Value Then
            intColourIndexMM1000 = numColourMM1000.Value
            UnsavedChanges = True
            cmdApply.Enabled = True
        End If

    End Sub
#End Region

#End Region

#Region "Import, Export, Reset"
    
    Private Sub cmdImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdImport.Click

        Dim msgTitle = "Import Settings"
        Dim msgPrompt = "Current settings will be lost, export current settings first?"
        Dim msgStyle = MsgBoxStyle.YesNoCancel Or MsgBoxStyle.Critical Or MsgBoxStyle.DefaultButton1
        Dim msgResult = MsgBox(msgPrompt, msgStyle, msgTitle)

        Select Case msgResult
            Case MsgBoxResult.Yes
                ExportSettings()
                ImportSettings()
            Case MsgBoxResult.No
                ImportSettings()
        End Select
        UnsavedChanges = True
    End Sub

    Private Sub cmdExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExport.Click
        If UnsavedChanges Then
            MsgBox("Please apply changes before exporting settings.")
        Else
            ExportSettings()
        End If
    End Sub
    
    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click

        Dim msgTitle = "Reset to Default Settings"
        Dim msgPrompt = "Reset to defaults?"
        Dim msgStyle = MsgBoxStyle.YesNo Or MsgBoxStyle.Critical Or MsgBoxStyle.DefaultButton1
        Dim msgResult = MsgBox(msgPrompt, msgStyle, msgTitle)

        Select Case msgResult
            Case MsgBoxResult.Yes
                booUseMM500 = True
                booUse50kNew = False

                strSearchPathsMM500 = ""
                strFreezeLayersMM500 = ""
                intColourIndexMM500 = 0

                strSearchPathsMM1000 = ""
                strFreezeLayersMM1000 = ""
                intColourIndexMM1000 = 0

                strSearchPathsSV = ""
                strSearchPaths50kNew = ""
                strSearchPaths50kOld = ""
                Initialise()
                UnsavedChanges = True
            Case Else
        End Select

    End Sub

#End Region

    Sub MoveSelectedItemUp(ByVal ListBox As ListBox)
        Dim Index As Integer = ListBox.SelectedIndex    'Index of selected item
        Dim Swap As Object = ListBox.SelectedItem       'Selected Item
        If (Index <> -1) AndAlso (Index - 1 > -1) Then
            ListBox.Items.RemoveAt(Index)                   'Remove it
            ListBox.Items.Insert(Index - 1, Swap)           'Add it back in one spot up
            ListBox.SelectedItem = Swap                     'Keep this item selected
            UnsavedChanges = True
            cmdApply.Enabled = True
        End If
    End Sub

    Sub MoveSelectedItemDown(ByVal ListBox As ListBox)
        Dim Index As Integer = ListBox.SelectedIndex    'Index of selected item
        Dim Swap As Object = ListBox.SelectedItem       'Selected Item
        If (Index <> -1) AndAlso (Index + 1 < ListBox.Items.Count) Then
            ListBox.Items.RemoveAt(Index)                   'Remove it
            ListBox.Items.Insert(Index + 1, Swap)           'Add it back in one spot up
            ListBox.SelectedItem = Swap                     'Keep this item selected
            UnsavedChanges = True
            cmdApply.Enabled = True
        End If
    End Sub

    Sub InsertItem(ByVal ListBox As ListBox, ByVal item As Object)
        If ListBox.Items.Contains(item) Then
            MsgBox(item & " is already in the list.")
            ListBox.SelectedItem = item
        Else
            Dim Index As Integer = ListBox.SelectedIndex    'Index of selected item
            ListBox.Items.Insert(Index + 1, item)           'Add it back in one spot up
            ListBox.SelectedItem = item                     'Keep this item selected
        End If
    End Sub

    Public Sub AddListItemsFromString(ByRef ListBox As Windows.Forms.ListBox, ByVal strItems As String, Optional ByVal StringDelimiter As Char = ";")
        ListBox.Items.Clear()
        Dim Items() As String = Split(strItems, StringDelimiter)
        For Each Item As String In Items
            If Item <> "" Then
                ListBox.Items.Add(Item)
            End If
        Next
    End Sub

    Public Function GetListAsString(ByVal Listbox As ListBox, Optional ByVal StringDelimiter As Char = ";") As String

        Dim SearchPaths(0 To (Listbox.Items.Count - 1))
        For i = 0 To (Listbox.Items.Count - 1)
            If Listbox.Items.Item(i) <> "" Then
                SearchPaths(i) = Listbox.Items.Item(i)
            End If
        Next

        Dim strSearchPaths As String = ""
        For Each SearchPath As String In SearchPaths
            If SearchPath <> "" Then
                'If the item's blank, don't add it - otherwise add the item and a delimeter to the end of the return string
                strSearchPaths = strSearchPaths & SearchPath & StringDelimiter
            End If
        Next
        If strSearchPaths <> "" Then
            'Remove the last delimiter
            strSearchPaths = Strings.Left(strSearchPaths, Len(strSearchPaths) - 1)
        End If

        Return strSearchPaths

    End Function

    Private Sub ImportSettings()
        Dim OpenDialog As New OpenFileDialog
        OpenDialog.Title = "Import Settings"
        OpenDialog.Filter = "OS Tile Picker Settings (*." & ExportExtension & ")|*" & ExportExtension
        If OpenDialog.ShowDialog() = DialogResult.OK Then
            Using StreamReader As New StreamReader(OpenDialog.FileName)
                While StreamReader.Peek() > 0
                    Dim ReadLine = StreamReader.ReadLine()
                    ' Split comma delimited data ( SettingName,SettingValue )  
                    Dim SplitLine = ReadLine.Split(CChar(","))

                    Dim SettingName As String = SplitLine(0)
                    Dim SettingValue As String = SplitLine(1)

                    ''###############################################################
                    ''This is a bodge...
                    Select Case SettingName
                        Case "MapDetach"
                            My.Settings(SettingName) = CType(SettingValue, Boolean)
                        Case "AskIfAttached"
                            My.Settings(SettingName) = CType(SettingValue, Boolean)
                        Case "StringDelimiter"
                            My.Settings(SettingName) = CType(SettingValue, Char)
                        Case "LayerColourIndexMM500"
                            My.Settings(SettingName) = CType(SettingValue, Integer)
                        Case "LayerColourIndexMM1000"
                            My.Settings(SettingName) = CType(SettingValue, Integer)
                        Case "CurrentMapType"
                            My.Settings(SettingName) = CType(SettingValue, Integer)
                        Case "UseMM500"
                            My.Settings(SettingName) = CType(SettingValue, Boolean)
                        Case "Use50kNew"
                            My.Settings(SettingName) = CType(SettingValue, Boolean)
                        Case "FirstRun"

                        Case Else
                            My.Settings(SettingName) = SettingValue
                    End Select
                    ''###############################################################


                End While

            End Using

            GetSettings()

            Initialise()
        End If
    End Sub

    Private Sub ExportSettings()
        Dim sDialog As New SaveFileDialog()
        sDialog.DefaultExt = "." & ExportExtension
        sDialog.Filter = "OS Tile Picker Settings (*." & ExportExtension & ")|*" & ExportExtension
        sDialog.Title = "Export Settings"

        If sDialog.ShowDialog() = DialogResult.OK Then
            Using sWriter As New StreamWriter(sDialog.FileName)
                For Each setting As Configuration.SettingsPropertyValue In My.Settings.PropertyValues
                    sWriter.WriteLine(setting.Name & "," & setting.PropertyValue.ToString())
                Next
            End Using
            My.Settings.Save()
            'MessageBox.Show("Settings saved")
        End If
    End Sub

#End Region


    Private Sub chkUseMasterMap500_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUseMasterMap500.CheckedChanged
        booUseMM500 = chkUseMasterMap500.Checked

        grpSearchPathsMM1000.Enabled = Not chkUseMasterMap500.Checked
        grpLayerFreezeMM1000.Enabled = Not chkUseMasterMap500.Checked
        grpColourMM1000.Enabled = Not chkUseMasterMap500.Checked

        If booUseMM500 Then
            strSearchPathsMM1000 = GetListAsString(lstSearchPathsMM1000, StringDelimiter)
            AddListItemsFromString(lstSearchPathsMM1000, strSearchPathsMM500, StringDelimiter)

            strFreezeLayersMM1000 = GetListAsString(lstLayerFreezeMM1000, StringDelimiter)
            AddListItemsFromString(lstLayerFreezeMM1000, strFreezeLayersMM500, StringDelimiter)

            numColourMM1000.Value = intColourIndexMM500
        Else
            AddListItemsFromString(lstSearchPathsMM1000, strSearchPathsMM1000, StringDelimiter)
            AddListItemsFromString(lstLayerFreezeMM1000, strFreezeLayersMM1000, StringDelimiter)
            numColourMM1000.Value = intColourIndexMM1000
        End If
    End Sub

    Private Sub chkUse50kNew_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUse50kNew.CheckedChanged
        booUse50kNew = chkUse50kNew.Checked

        grpSearchPaths50kOld.Enabled = Not chkUse50kNew.Checked

        If booUse50kNew Then
            MsgBox("Using the same search path for 10km and 20km tiles may cause conflicts.", MsgBoxStyle.Exclamation)

            AddListItemsFromString(lstSearchPaths50kOld, strSearchPaths50kNew, StringDelimiter)
        Else
            AddListItemsFromString(lstSearchPaths50kOld, strSearchPaths50kOld, StringDelimiter)
        End If
    End Sub

    Private Sub lstLayerFreezeMM500_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstLayerFreezeMM500.TextChanged
        If booUseMM500 Then AddListItemsFromString(lstSearchPathsMM1000, strSearchPathsMM1000, StringDelimiter)
    End Sub
End Class


