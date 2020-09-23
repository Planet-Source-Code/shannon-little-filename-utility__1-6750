VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "ntFile Utility v1.0"
   ClientHeight    =   3000
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7125
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   3000
   ScaleWidth      =   7125
   Begin VB.Timer tmrCheckForNewVersionOfProg 
      Interval        =   1000
      Left            =   6000
      Top             =   120
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   6120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   6120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   6480
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowseForPath 
      Caption         =   "Browse..."
      Height          =   285
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtCurrentPath 
      Height          =   285
      Left            =   600
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      ToolTipText     =   "Path in which all operation will be performed."
      Top             =   120
      Width           =   3975
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   255
      Left            =   2040
      TabIndex        =   27
      Top             =   2760
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   28
      Top             =   2685
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1182
            MinWidth        =   1182
            Object.ToolTipText     =   "Progress percent"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8440
            Object.ToolTipText     =   "Progress bar"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   582
            MinWidth        =   88
            Picture         =   "frmMain.frx":1272
            Object.ToolTipText     =   "NeoTrix"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "11:08 AM"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraFrame 
      BorderStyle     =   0  'None
      Caption         =   "Renaming"
      Height          =   1695
      Index           =   2
      Left            =   120
      TabIndex        =   35
      Top             =   840
      Width           =   6855
      Begin VB.Frame fraReplacePresets 
         Caption         =   "&Presets"
         Height          =   855
         Left            =   120
         TabIndex        =   46
         Top             =   0
         Width           =   1215
         Begin VB.OptionButton optReplacePresetNone 
            Caption         =   "None"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optReplacePresetMP3 
            Caption         =   "MP3"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   490
            Width           =   735
         End
      End
      Begin VB.TextBox txtReplaceTo 
         Height          =   285
         Left            =   2160
         TabIndex        =   45
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtReplaceFrom 
         Height          =   285
         Left            =   2160
         TabIndex        =   44
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdReplaceGo 
         Caption         =   "&Go!"
         Default         =   -1  'True
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   960
         Width           =   1215
      End
      Begin VB.Frame fraReplaceType 
         Caption         =   "&Type"
         Height          =   1095
         Left            =   3600
         TabIndex        =   37
         Top             =   240
         Width           =   1455
         Begin VB.OptionButton optReplaceFile 
            Caption         =   "F&iles"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optReplaceBoth 
            Caption         =   "&Both"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton optReplaceFolder 
            Caption         =   "Fol&ders"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdReplaceCancel 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblReplaceInfoText 
         Caption         =   "This will replace every instance of the source string with the target string. Extensions will not be changed."
         Height          =   1455
         Left            =   5160
         TabIndex        =   49
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label lblReplaceWith 
         Caption         =   "&With:"
         Height          =   255
         Left            =   1680
         TabIndex        =   43
         Top             =   885
         Width           =   495
      End
      Begin VB.Label lblReplaceEvery 
         Caption         =   "&Replace:"
         Height          =   255
         Left            =   1440
         TabIndex        =   42
         Top             =   525
         Width           =   735
      End
   End
   Begin VB.Frame fraFrame 
      BorderStyle     =   0  'None
      Caption         =   "Extensions"
      Height          =   1695
      Index           =   1
      Left            =   120
      TabIndex        =   30
      Top             =   840
      Width           =   6855
      Begin VB.CommandButton cmdExtensionsCancel 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Frame fraExtensionsSpecific 
         Caption         =   "Specific Extensions"
         Height          =   1335
         Left            =   1920
         TabIndex        =   7
         Top             =   120
         Width           =   2175
         Begin VB.TextBox txtSpecificExtFrom 
            Height          =   285
            Left            =   720
            TabIndex        =   9
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtSpecificExtTo 
            Height          =   285
            Left            =   720
            TabIndex        =   11
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblExtensionsTo 
            Caption         =   "To:"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lblExtensionsFrom 
            Caption         =   "From:"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame fraAllExtensions 
         Caption         =   "All &Extensions"
         Height          =   1335
         Left            =   1920
         TabIndex        =   12
         Top             =   120
         Width           =   2175
         Begin VB.TextBox txtExtensionsChangeAllTo 
            Height          =   285
            Left            =   1200
            TabIndex        =   14
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lblExtensionsChangeAllTo 
            Caption         =   "&Change all to:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdExtensionsGo 
         Caption         =   "&Go!"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optSpecificExtensions 
         Caption         =   "&Specific extensions"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optAllExtensions 
         Caption         =   "&All extensions"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblExtensionsRange 
         Caption         =   "Range:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblExtensionsInfo 
         Caption         =   $"frmMain.frx":2F7E
         Height          =   1575
         Left            =   4200
         TabIndex        =   16
         Top             =   120
         Width           =   2655
      End
   End
   Begin VB.Frame fraFrame 
      BorderStyle     =   0  'None
      Caption         =   "Renaming"
      Height          =   1695
      Index           =   0
      Left            =   120
      TabIndex        =   29
      Top             =   840
      Width           =   6855
      Begin VB.CommandButton cmdRenameCancel 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox chkCapAfterEverySpace 
         Caption         =   "Capitalize every letter after a space"
         Height          =   255
         Left            =   1920
         TabIndex        =   32
         ToolTipText     =   "Caps every letter after a space"
         Top             =   40
         Width           =   2775
      End
      Begin VB.Frame fraRenamingType 
         Caption         =   "&Type"
         Height          =   1095
         Left            =   3360
         TabIndex        =   22
         Top             =   360
         Width           =   1455
         Begin VB.OptionButton optRenameFolder 
            Caption         =   "Fol&ders"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton optRenameAll 
            Caption         =   "&Both"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton optRenameFile 
            Caption         =   "F&iles"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame fraRenamingCase 
         Caption         =   "Ca&se"
         Height          =   1095
         Left            =   1440
         TabIndex        =   18
         Top             =   360
         Width           =   1695
         Begin VB.OptionButton optNoChange 
            Caption         =   "&No Change"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optUppercase 
            Caption         =   "&Uppercase"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton optLowercase 
            Caption         =   "&Lowercase"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdRenameGo 
         Caption         =   "&Go!"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkCapFirstLetter 
         Caption         =   "C&aptialize first letter"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Caps the first letter"
         Top             =   40
         Width           =   1935
      End
      Begin VB.Label lblWindowsUppercaseNote 
         Caption         =   $"frmMain.frx":309E
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   4920
         TabIndex        =   31
         Top             =   240
         Width           =   1935
      End
   End
   Begin MSComctlLib.TabStrip tbsTab 
      Height          =   2175
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3836
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Format Case"
            Key             =   "renaming"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Change Extension"
            Key             =   "extensions"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Replace Text"
            Key             =   "Replace"
            ImageVarType    =   2
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Line MenuLineLight 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   7080
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Label lblPath 
      Caption         =   "Path:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   495
   End
   Begin VB.Line MenuLineDark 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   0
      X2              =   7080
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOnTop 
         Caption         =   "&Always on top"
      End
      Begin VB.Menu mnuHyphen2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowResultsWindow 
         Caption         =   "&Show Results Window"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuClearResultsWindow 
         Caption         =   "&Clear Results Window"
      End
      Begin VB.Menu mnuFileHyphen1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuTotalFileCount 
         Caption         =   "&Total file count..."
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuHelpHyphen2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpGotoNTWebsite 
         Caption         =   "&Goto NeoTrix Website..."
      End
      Begin VB.Menu mnuHelpHyphen1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About ntFile Utility..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intCurFrame As Integer ' Current Frame visible
Dim lngTotalFileCount As Long
Dim bCancel As Boolean      'Set to TRUE when a cancel button is hit


Private Sub cmdBrowseForPath_Click()
    Dim strTemp As String
    strTemp = BrowseForFolder(Me.hWnd)
    If strTemp <> "" Then txtCurrentPath.Text = strTemp
End Sub




'**********                                              **********
'********** All functions for Renaming(Formatting) Files and Folders **********
'**********                                              **********

Private Sub cmdRenameGo_Click()
    On Error GoTo lError
    Dim strTemp As String, N As Integer
    Dim doFolder As Boolean, doFile As Boolean
    Dim doLower As Boolean, doUpper As Boolean
    Dim doCapAfterSpace As Boolean, doCapFirstLetter As Boolean
    Dim strFront As String, strMid As String, strEnd As String 'Used by doCapFirstLetter
    Dim fso, f, fc, sfc, fTemp
    'FileSystem Object, File, FileCollection, SubFolderCollection, FileFolderTemp
    Dim intTemp As Integer    'Used by progressbar
    
    DisableEverything
        
    'Make sure at least a c:\ is there for the path
    If Len(txtCurrentPath.Text) < 3 Then
        TypError "Please insert a path. I'm not a mind reader."
    Else    'Got a path
        Set fso = CreateObject("Scripting.FileSystemObject")    'Obj
        Set f = fso.GetFolder(txtCurrentPath.Text)              'File obj
        Set sfc = f.SubFolders
        Set fc = f.Files
        
        'Read each option
        'Assign a T/F value to each bool var that represents each option
        
        'Do Files/Folders/Both options
        intTemp = 0
        If optRenameFolder.Value Then      'Only do folders
            doFolder = True
            doFile = False
            intTemp = sfc.Count 'Subfolder count, used for progress MAX
        Else
            If optRenameFile.Value Then   'Only do files
                doFolder = False
                doFile = True
                intTemp = fc.Count  'File count, used for progress MAX
            Else                                'Do both files/folders
                doFolder = True
                doFile = True
                intTemp = intTemp + sfc.Count
                intTemp = intTemp + fc.Count
            End If
        End If
        
        'No files or folders, show message
        If intTemp = 0 Then
            TypError "There are no files or folders in the currently selected folder. Please select a folder with something in it!"
            GoTo lEnd
        End If
        
        'Set-up progress bar
        SetProgressLimits 1, intTemp     '1 amount for each file
        SetProgressValue 1
        
        'Check which case we are converting to Lower/Upper/Don'tChange
        If optLowercase.Value = True Then   'Only do lower case
            doLower = True
            doUpper = False
        Else
            If optUppercase.Value = True Then   'Only do upper case
                doLower = False
                doUpper = True
            Else            'Do not change the current case
                doLower = False
                doUpper = False
            End If
        End If
        
        'Cap first letter option
        doCapFirstLetter = IIf(chkCapFirstLetter.Value = vbChecked, True, False)
        
        'Cap after every space option
        doCapAfterSpace = IIf(chkCapAfterEverySpace.Value = vbChecked, True, False)
        
        '**** Now down to the real work ****
        '**** If folders is selected ****
        If doFolder Then
            For Each fTemp In sfc
                'strTemp will be operated on by all of the functions, not fTemp.name
                'at the end it will rename the file
                strTemp = fTemp.Name
                                        
                DoEvents
                StepProgress 'Increase progress
                'Adds to list of results
                With frmResults.lstResults.ListItems
                    .Add 1, , , "Folder", "Folder"      'Icon
                    .Item(1).SubItems(1) = strTemp      'Original filename
                End With
                
                'Formatting branch, this is done for each folder
                If doLower Then        'Lowercase the foldername
                    strTemp = LCase(strTemp)
                Else
                    If doUpper Then    'Uppercase the foldername
                        strTemp = UCase(strTemp)
                    End If  'ELSE No change selected
                End If
                
                If doCapFirstLetter Then
                    strTemp = CapFirstLetter(strTemp)
                End If
                    
                If doCapAfterSpace Then
                    strTemp = StrConv(strTemp, vbProperCase)
                End If
                
                'FSO wouldn't allow the name to be the same, even if the case
                'was different, so add a "[" (very uncommon file extension )
                'to the end while re-casing
                'then remove it after the case has been changed
                fTemp.Name = strTemp & "["
                fTemp.Name = Left(fTemp.Name, Len(fTemp.Name) - 1)
                
                
                lngTotalFileCount = lngTotalFileCount + 1
                'Finish the list item off with the final file name
                With frmResults.lstResults.ListItems
                    .Item(1).SubItems(2) = fTemp.Name   'Change filename
                End With
                frmResults.IncrementCounter    'Sub on frmResults to increment file counter
                If bCancel Then GoTo lEnd   'Cancel was pressed
            Next
        End If
        '**** If files is selected ****
        If doFile Then
            For Each fTemp In fc
                strTemp = fTemp.Name
                DoEvents
                StepProgress 'Increase progress
                'Adds to list of results
                With frmResults.lstResults.ListItems
                    .Add 1, , , "File", "File" 'Icon
                    .Item(1).SubItems(1) = fTemp.Name   'Original filename
                End With

                'Formatting branch, this is done for each file
                If doLower Then        'Lowercase the foldername
                    strTemp = LCase(strTemp)
                Else
                    If doUpper Then    'Uppercase the foldername
                        strTemp = UCase(strTemp)
                    End If  'ELSE No change selected
                End If
                
                If doCapFirstLetter Then
                    strTemp = CapFirstLetter(strTemp)
                End If
                    
                If doCapAfterSpace Then
                    strTemp = StrConv(strTemp, vbProperCase)
                End If

                'FSO wouldn't allow the name to be the same, even if the case
                'was different, so add a "[" to the end while re-casing
                'then remove it after the case has been changed
                fTemp.Name = strTemp & "["
                fTemp.Name = Left(fTemp.Name, Len(fTemp.Name) - 1)

                lngTotalFileCount = lngTotalFileCount + 1
                'Finish the list item off with the final file name
                With frmResults.lstResults.ListItems
                    .Item(1).SubItems(2) = fTemp.Name   'Change filename
                End With
                frmResults.IncrementCounter    'Sub on frmResults to increment file counter
                If bCancel Then GoTo lEnd   'Cancel was pressed
            Next
        End If
    End If
    
    EnableEverything
    Exit Sub
lError:
    'If intTemp = 0 Then Err.Clear   'It was already handled above, clear err
    
    Select Case Err.Number
        Case 5: 'Illegal chars in path
                TypError "The path contains illegal character(s)" & vbNewLine & " \ / : * ? " & Chr(34) & " < > |"
        Case 76: 'Path does not exist error
                TypError "Path does not exist"
        Case Default: 'For anything I haven't yet programmed a response for
                GenError
    End Select
    
lEnd:
    EnableEverything
End Sub

Private Sub cmdRenameCancel_Click()
    bCancel = True
End Sub

'******************************************************************
'******************************************************************
'******************************************************************





'**********                                              **********
'********** All functions for Renaming Extensions       **********
'**********                                              **********

Private Sub cmdExtensionsGo_Click()
    On Error GoTo lError
    Dim fso, f, fc, sfc, fTemp
    'FileSystem Object, File, FileCollection, SubFolderCollection, FileFolderTemp
    Dim strTemp As String, strExt1 As String, strExt2 As String, strTempExt As String
    Dim strExt3 As String   'Used by All Extensions
    Dim N As Integer
    
    DisableEverything
    
    'Make sure at least a c:\ is there for the path
    If Len(txtCurrentPath.Text) < 3 Then
        TypError "Please insert a path. I'm not a mind reader."
    Else    'Got a path
        
        Set fso = CreateObject("Scripting.FileSystemObject")    'Obj
        Set f = fso.GetFolder(txtCurrentPath.Text)              'File obj
        Set fc = f.Files
        
        'Set-up progress bar
        SetProgressLimits 1, fc.Count     '1 amount for each file
        SetProgressValue 1
        
        If optSpecificExtensions.Value Then     'Specific extensions
            strExt1 = txtSpecificExtFrom.Text
            strExt2 = txtSpecificExtTo.Text
            
            
            'Check to make sure something was entered
            If Len(strExt1) = 0 Then
                TypError "You need to enter an extension for the From field"
                GoTo lEnd
            End If
            If Len(strExt2) = 0 Then
                TypError "You need to enter an extension for the To field"
                GoTo lEnd
            End If
            
            'Check for a "." and remove it and anything that comes before it
            strExt1 = StrReturnRight(strExt1, ".")
            'Check for a "." and remove it and anything that comes before it
            strExt2 = StrReturnRight(strExt2, ".")
            'Make sure there are not any illegal characters
            ' \ / : * ? " < > |
            'char 34 = "
            'Raise a custom error message for illegal extensions
            '513 for the FROM field
            '514 for the TO field
            If DoesCharExist(strExt1, "/") Then Err.Raise (513)
            If DoesCharExist(strExt1, "\") Then Err.Raise (513)
            If DoesCharExist(strExt1, ":") Then Err.Raise (513)
            If DoesCharExist(strExt1, "*") Then Err.Raise (513)
            If DoesCharExist(strExt1, "?") Then Err.Raise (513)
            If DoesCharExist(strExt1, Chr(34)) Then Err.Raise (513)
            If DoesCharExist(strExt1, "<") Then Err.Raise (513)
            If DoesCharExist(strExt1, ">") Then Err.Raise (513)
            If DoesCharExist(strExt1, "|") Then Err.Raise (513)
            
            If DoesCharExist(strExt2, "/") Then Err.Raise (514)
            If DoesCharExist(strExt2, "\") Then Err.Raise (514)
            If DoesCharExist(strExt2, ":") Then Err.Raise (514)
            If DoesCharExist(strExt2, "*") Then Err.Raise (514)
            If DoesCharExist(strExt2, "?") Then Err.Raise (514)
            If DoesCharExist(strExt2, Chr(34)) Then Err.Raise (514)
            If DoesCharExist(strExt2, "<") Then Err.Raise (514)
            If DoesCharExist(strExt2, ">") Then Err.Raise (514)
            If DoesCharExist(strExt2, "|") Then Err.Raise (514)
            
            'Now that we have 2 formatted extensions, we can compare with
            'each file in the file collection to see if they are a match
            'then replace if they are
        
        
            For Each fTemp In fc
                strTemp = fTemp.Name
                DoEvents
                StepProgress 'Increase progress
                'Compare extenstions to see if they match
                strTempExt = StrReturnRight(strTemp, ".")   'Returns the extension
                If UCase(strTempExt) = UCase(strExt1) Then
                    'Adds to list of results
                    With frmResults.lstResults.ListItems
                        .Add 1, , , "File", "File" 'Icon
                        .Item(1).SubItems(1) = strTemp      'Original filename
                    End With

                    'Mash it all together
                    strTemp = StrReturnLeft(strTemp, ".") & "." & strExt2
                    
                    fTemp.Name = strTemp & "["
                    
                    'This checks for duplicate files
                    'Returns the name of the file if it exists
                    If Dir(txtCurrentPath.Text & "\" & Left(fTemp.Name, Len(fTemp.Name) - 1)) <> "" Then
                        GoTo lCSkipDuplicateFile1
                    End If
                    
                    fTemp.Name = Left(fTemp.Name, Len(fTemp.Name) - 1)
                    
                    'Finish the list item off with the final file name
                    With frmResults.lstResults.ListItems
                        .Item(1).SubItems(2) = strTemp   'Change filename
                    End With
                    frmResults.IncrementCounter    'Sub on frmResults to increment file counter
                    GoTo lJumpOver1                 'Just to skip over below
lCSkipDuplicateFile1:
                    With frmResults.lstResults.ListItems
                        'Need to restore old filename
                        fTemp.Name = .Item(1).SubItems(1)   'SubItem1 holds old file name
                        .Item(1).SubItems(2) = "Duplicate filename"
                        .Item(1).ListSubItems.Item(1).ForeColor = vbRed
                        .Item(1).ListSubItems.Item(2).ForeColor = vbRed
                        '.Remove (1) 'Remove the previous entered file
                        'Since it should be skipped and not processed
                    End With
lJumpOver1:
                End If
                If bCancel Then GoTo lEnd   'Cancel was pressed
            Next
            
        Else      'All extensions
        
            'Same as above but with no comparison
            
            strExt3 = txtExtensionsChangeAllTo.Text
            
            If Len(strExt3) = 0 Then
                TypError "You need to enter an extension for the Change All To field"
                GoTo lEnd
            End If
            
            'Check for a "." and remove it and anything that comes before it
            strExt3 = StrReturnRight(strExt3, ".")
            'Make sure there are not any illegal characters
            'Raise a custom error message for illegal extensions
            If DoesCharExist(strExt3, "/") Then Err.Raise (515)
            If DoesCharExist(strExt3, "\") Then Err.Raise (515)
            If DoesCharExist(strExt3, ":") Then Err.Raise (515)
            If DoesCharExist(strExt3, "*") Then Err.Raise (515)
            If DoesCharExist(strExt3, "?") Then Err.Raise (515)
            If DoesCharExist(strExt3, Chr(34)) Then Err.Raise (515)
            If DoesCharExist(strExt3, "<") Then Err.Raise (515)
            If DoesCharExist(strExt3, ">") Then Err.Raise (515)
            If DoesCharExist(strExt3, "|") Then Err.Raise (515)
            
            For Each fTemp In fc
                strTemp = fTemp.Name
                DoEvents
                StepProgress 'Increase progress
                
                'Adds to list of results
                With frmResults.lstResults.ListItems
                    .Add 1, , , "File", "File" 'Icon
                    .Item(1).SubItems(1) = strTemp      'Original filename
                End With

                'Mash it all together with the new extension
                strTemp = StrReturnLeft(strTemp, ".") & "." & strExt3
                
                fTemp.Name = strTemp & "["
                
                'This checks for duplicate files
                'Returns the name of the file if it exists
                If Dir(txtCurrentPath.Text & "\" & Left(fTemp.Name, Len(fTemp.Name) - 1)) <> "" Then
                    GoTo lCSkipDuplicateFile2
                End If
                    
                'Duplicate file error occurs here (below)
                fTemp.Name = Left(fTemp.Name, Len(fTemp.Name) - 1)
                
                'Finish the list item off with the final file name
                With frmResults.lstResults.ListItems
                    .Item(1).SubItems(2) = strTemp   'Change filename
                End With
                frmResults.IncrementCounter    'Sub on frmResults to increment file counter
                GoTo lJumpOver2                 'Just to skip over below
lCSkipDuplicateFile2:
                With frmResults.lstResults.ListItems
                    'Need to restore old filename
                    fTemp.Name = .Item(1).SubItems(1)   'SubItem1 holds old file name
                    .Item(1).SubItems(2) = "Duplicate filename"
                    .Item(1).ListSubItems.Item(1).ForeColor = vbRed
                    .Item(1).ListSubItems.Item(2).ForeColor = vbRed
                End With
lJumpOver2:
                If bCancel Then GoTo lEnd   'Cancel was pressed
            Next
        End If  'End extension type All or Specific
    End If  'End if a path exists
    EnableEverything
    Exit Sub
    
lError:
     
     Select Case Err.Number
        Case 5: 'Illegal chars in path
                TypError "The path contains illegal character(s)" & vbNewLine & " \ / : * ? " & Chr(34) & " < > |"
        Case 58: 'Duplicate file error
                TypError "Unexpected duplicate file"
        Case 76: 'Path does not exist error
                TypError "Path does not exist"
        'Errors 513 and above are user-defined
        Case 513:   'Illegal chars in From field
                TypError "The extension From field contains illegal character(s)" & vbNewLine & " \ / : * ? " & Chr(34) & " < > |"
        Case 514:   'Illegal chars in TO field
                TypError "The extension To field contains illegal character(s)" & vbNewLine & " \ / : * ? " & Chr(34) & " < > |"
        Case 515:   'Illegal chars in All extension field
                TypError "The Change All To field contains illegal character(s)" & vbNewLine & " \ / : * ? " & Chr(34) & " < > |"
        Case Default: 'For anything I haven't yet programmed a response for
                GenError
    End Select
    Err.Clear
lEnd:
    EnableEverything
End Sub

Private Sub cmdExtensionsCancel_Click()
    bCancel = True
End Sub

'******************************************************************
'******************************************************************
'******************************************************************





'******************************************************************
'******** Functions to replace text strings in file names *********
'******************************************************************


Private Sub cmdReplaceGo_Click()
    On Error GoTo lError
    Dim strTemp As String, N As Integer
    Dim doFolder As Boolean, doFile As Boolean
    Dim strFrom As String, strTo As String, strExt As String
    Dim fso, f, fc, sfc, fTemp
    'FileSystem Object, File, FileCollection, SubFolderCollection, FileFolderTemp
    Dim intTemp As Integer    'Used by progressbar
    
    DisableEverything
        
    'Make sure at least a c:\ is there for the path
    If Len(txtCurrentPath.Text) < 3 Then
        TypError "Please insert a path. I'm not a mind reader."
    Else    'Got a path
        Set fso = CreateObject("Scripting.FileSystemObject")    'Obj
        Set f = fso.GetFolder(txtCurrentPath.Text)              'File obj
        Set sfc = f.SubFolders
        Set fc = f.Files
        
        'Read each option
        'Assign a T/F value to each bool var that represents each option
        
        'Do Files/Folders/Both options
        intTemp = 0
        If optReplaceFolder.Value Then      'Only do folders
            doFolder = True
            doFile = False
            intTemp = sfc.Count 'Subfolder count, used for progress MAX
        Else
            If optReplaceFile.Value Then   'Only do files
                doFolder = False
                doFile = True
                intTemp = fc.Count  'File count, used for progress MAX
            Else                                'Do both files/folders
                doFolder = True
                doFile = True
                intTemp = intTemp + sfc.Count
                intTemp = intTemp + fc.Count
            End If
        End If
        
        'Make sure there is something in the Replace from textbox
        'The Replace To box is not check because they could be wanting
        ' to remove a character, so it would be left blank
        If Len(txtReplaceFrom.Text) = 0 Then
            TypError "There is nothing entered for Replace"
            GoTo lEnd
        End If
        
        'No files or folders, show message
        If intTemp = 0 Then
            TypError "There are no files or folders in the currently selected folder. Please select a folder with something in it!"
            GoTo lEnd
        End If
        
        strFrom = txtReplaceFrom.Text
        strTo = txtReplaceTo.Text
        
        'Make sure there are not any illegal characters
        'Raise a custom error message for illegal extensions
        If DoesCharExist(strTo, "/") Then Err.Raise (515)
        If DoesCharExist(strTo, "\") Then Err.Raise (515)
        If DoesCharExist(strTo, ":") Then Err.Raise (515)
        If DoesCharExist(strTo, "*") Then Err.Raise (515)
        If DoesCharExist(strTo, "?") Then Err.Raise (515)
        If DoesCharExist(strTo, Chr(34)) Then Err.Raise (515)
        If DoesCharExist(strTo, "<") Then Err.Raise (515)
        If DoesCharExist(strTo, ">") Then Err.Raise (515)
        If DoesCharExist(strTo, "|") Then Err.Raise (515)
        
        'Set-up progress bar
        SetProgressLimits 1, intTemp     '1 amount for each file
        SetProgressValue 1
        
        '**** Now down to the real work ****
        '**** If folders is selected ****
        If doFolder Then
            For Each fTemp In sfc
                'strTemp will be operated on by all of the functions, not fTemp.name
                'at the end it will rename the file
                strTemp = fTemp.Name
                                        
                DoEvents
                StepProgress 'Increase progress
                'Adds to list of results
                With frmResults.lstResults.ListItems
                    .Add 1, , , "Folder", "Folder"      'Icon
                    .Item(1).SubItems(1) = strTemp      'Original filename
                End With
                
                'Replaces every strFrom with strTo
                strTemp = Replace(strTemp, strFrom, strTo, , , vbBinaryCompare)
                
                'FSO wouldn't allow the name to be the same, even if the case
                'was different, so add a "[" (very uncommon file extension )
                'to the end while re-casing
                'then remove it after the case has been changed
                fTemp.Name = strTemp & "["
                fTemp.Name = Left(fTemp.Name, Len(fTemp.Name) - 1)
                
                lngTotalFileCount = lngTotalFileCount + 1
                'Finish the list item off with the final file name
                With frmResults.lstResults.ListItems
                    .Item(1).SubItems(2) = fTemp.Name   'Change filename
                End With
                frmResults.IncrementCounter    'Sub on frmResults to increment file counter
                If bCancel Then GoTo lEnd   'Cancel was pressed
            Next
        End If
        '**** If files is selected ****
        If doFile Then
            For Each fTemp In fc
                strTemp = fTemp.Name
                DoEvents
                StepProgress 'Increase progress
                'Adds to list of results
                With frmResults.lstResults.ListItems
                    .Add 1, , , "File", "File" 'Icon
                    .Item(1).SubItems(1) = fTemp.Name   'Original filename
                End With

                'Replaces every strFrom with strTo
                strTemp = Replace(strTemp, strFrom, strTo, , , vbBinaryCompare)

                'FSO wouldn't allow the name to be the same, even if the case
                'was different, so add a "[" to the end while re-casing
                'then remove it after the case has been changed
                fTemp.Name = strTemp & "["
                fTemp.Name = Left(fTemp.Name, Len(fTemp.Name) - 1)

                lngTotalFileCount = lngTotalFileCount + 1
                'Finish the list item off with the final file name
                With frmResults.lstResults.ListItems
                    .Item(1).SubItems(2) = fTemp.Name   'Change filename
                End With
                frmResults.IncrementCounter    'Sub on frmResults to increment file counter
                If bCancel Then GoTo lEnd   'Cancel was pressed
            Next
        End If
    End If
    
    EnableEverything
    Exit Sub
lError:
    'If intTemp = 0 Then Err.Clear   'It was already handled above, clear err
    
    Select Case Err.Number
        Case 5: 'Illegal chars in replace to
                TypError "The path contains illegal character(s)" & vbNewLine & " \ / : * ? " & Chr(34) & " < > |"
        Case 76: 'Path does not exist error
                TypError "Path does not exist"
        Case 515:   'Illegal chars in the Replace To textbox
                TypError "The To field contains illegal character(s)" & vbNewLine & " \ / : * ? " & Chr(34) & " < > |"
        Case Default: 'For anything I haven't yet programmed a response for
                GenError
    End Select
    
lEnd:
    EnableEverything
End Sub

Private Sub cmdReplaceCancel_Click()
    bCancel = True
End Sub



'******************************************************************
'******************************************************************
'******************************************************************


Private Sub EnableEverything()
    tbsTab.Enabled = True
    cmdBrowseForPath.Enabled = True
    txtCurrentPath.Enabled = True
    SetProgressLimits 0, 2
    SetProgressValue 0
    bCancel = False
    
    '** Renaming files/folders
    cmdRenameGo.Caption = "&Go!"
    cmdRenameGo.Enabled = True
    cmdRenameCancel.Enabled = False
    fraRenamingType.Enabled = True
    fraRenamingCase.Enabled = True
    optRenameFolder.Enabled = True
    optRenameFile.Enabled = True
    optRenameAll.Enabled = True
    optLowercase.Enabled = True
    optUppercase.Enabled = True
    optNoChange.Enabled = True
    chkCapFirstLetter.Enabled = True
    chkCapAfterEverySpace.Enabled = True
    '** Renaming extensions
    optSpecificExtensions.Enabled = True
    optAllExtensions.Enabled = True
    cmdExtensionsGo.Caption = "&Go!"
    cmdExtensionsGo.Enabled = True
    cmdExtensionsCancel.Enabled = False
    fraExtensionsSpecific.Enabled = True
    txtSpecificExtFrom.Enabled = True
    txtSpecificExtTo.Enabled = True
    fraAllExtensions.Enabled = True
    txtExtensionsChangeAllTo.Enabled = True
    'Replacing
    cmdReplaceGo.Enabled = True
    cmdReplaceCancel.Enabled = False
    fraReplaceType.Enabled = True
    optReplaceFile.Enabled = True
    optReplaceFolder.Enabled = True
    optReplaceBoth.Enabled = True
    fraReplacePresets.Enabled = True
    optReplacePresetMP3.Enabled = True
    optReplacePresetNone.Enabled = True
    txtReplaceFrom.Enabled = True
    txtReplaceTo.Enabled = True
End Sub

Private Sub DisableEverything()
    txtCurrentPath.Enabled = False
    cmdBrowseForPath.Enabled = False
    tbsTab.Enabled = False
    bCancel = False
    
   '** Renaming files/folders
    cmdRenameGo.Enabled = False
    cmdRenameGo.Caption = "Processing..."
    cmdRenameCancel.Enabled = True
    fraRenamingCase.Enabled = False
    fraRenamingType.Enabled = False
    optRenameFolder.Enabled = False
    optRenameFile.Enabled = False
    optRenameAll.Enabled = False
    optLowercase.Enabled = False
    optUppercase.Enabled = False
    optNoChange.Enabled = False
    chkCapFirstLetter.Enabled = False
    chkCapAfterEverySpace.Enabled = False
    '** Renaming extensions
    optSpecificExtensions.Enabled = False
    optAllExtensions.Enabled = False
    cmdExtensionsGo.Caption = "Processing..."
    cmdExtensionsGo.Enabled = False
    cmdExtensionsCancel.Enabled = True
    fraExtensionsSpecific.Enabled = False
    txtSpecificExtFrom.Enabled = False
    txtSpecificExtTo.Enabled = False
    fraAllExtensions.Enabled = False
    txtExtensionsChangeAllTo.Enabled = False
    'Replacing
    cmdReplaceGo.Enabled = False
    cmdReplaceCancel.Enabled = True
    fraReplaceType.Enabled = False
    optReplaceFile.Enabled = False
    optReplaceFolder.Enabled = False
    optReplaceBoth.Enabled = False
    fraReplacePresets.Enabled = False
    optReplacePresetMP3.Enabled = False
    optReplacePresetNone.Enabled = False
    txtReplaceFrom.Enabled = False
    txtReplaceTo.Enabled = False
End Sub

Private Sub Form_Load()
    modWinProc.Initialize frmMain
    gHW = Me.hWnd
    modWinProc.Hook
    LoadSettings
    strProgramName = "ntFile Utility"
    intCurFrame = 0     'Used for tab control
    If txtCurrentPath.Text = "" Then txtCurrentPath.Text = App.Path
    Dim intTemp
    'Set all frames as not visible, they are positioned when the resize event fires
    For intTemp = 0 To 2
        With fraFrame(intTemp)
            .Visible = False
        End With
    Next intTemp
    fraFrame(0).Visible = True
    SetProgressVisible True
    PositionProgressBar
    SetProgressLimits 0, 2
    SetProgressValue 0
    Load frmResults
    frmResults.lngFileCount = 0
    bCancel = False
End Sub

'******************************************
'********* Progress bar Functions *********
'******************************************

Private Sub PositionProgressBar()
    'Position the progress bar over the status bar panel
    With ProgressBar
        .Top = StatusBar.Top + 50
        .Left = StatusBar.Panels(2).Left + 25
        .Width = StatusBar.Panels(2).Width - 45
        .Height = StatusBar.Height - 65
    End With
End Sub

Private Sub SetProgressVisible(ByVal blnVisible As Boolean)
    ProgressBar.Visible = blnVisible
    StatusBar.Panels(1).Visible = blnVisible
End Sub

Private Sub SetProgressLimits(ByVal intMin As Integer, ByVal intMax As Integer)
    ProgressBar.Min = intMin
    ProgressBar.Max = intMax
End Sub

Private Sub SetProgressValue(ByVal intValue As Integer)
    ProgressBar.Value = intValue
    StatusBar.Panels(1).Text = FormatPercent(intValue / ProgressBar.Max)
End Sub

Private Sub StepProgress()
    If ProgressBar.Value + 1 <= ProgressBar.Max Then    'Keep from going over max and erroring
        ProgressBar.Value = ProgressBar.Value + 1
        StatusBar.Panels(1).Text = FormatPercent(ProgressBar.Value / ProgressBar.Max)
    End If
End Sub



'******************************************
'******************************************
'******************************************

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    modWinProc.Unhook
End Sub

Private Sub Form_Resize()
    Dim intTemp
    If WindowState <> vbMinimized Then
        If Height < 4095 Then Height = 4095
        PositionProgressBar
        tbsTab.Width = ScaleWidth - (2 * tbsTab.Left)
        tbsTab.Height = ScaleHeight - StatusBar.Height - tbsTab.Top
        'Position all frames so they are centered
        For intTemp = 0 To 2
            With fraFrame(intTemp)
                .Left = tbsTab.ClientLeft
                .Top = tbsTab.ClientTop + tbsTab.TabFixedHeight
                .Height = tbsTab.ClientHeight
                .Width = tbsTab.ClientWidth
            End With
        Next intTemp
        MenuLineLight.X2 = Width
        MenuLineDark.X2 = Width
        If mnuShowResultsWindow.Checked = True Then
            If frmResults.Visible = False Then frmResults.Visible = True
        End If
    Else
        frmResults.Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSettings
    frmResults.UnloadForm
    Unload frmAbout
    Unhook
End Sub

Private Sub mnuClearResultsWindow_Click()
    frmResults.lstResults.ListItems.Clear
End Sub

Private Sub mnuFileExit_Click()
    Unload frmMain
End Sub

Private Sub mnuFileOnTop_Click()
    mnuFileOnTop.Checked = mnuFileOnTop.Checked Xor True
    If mnuFileOnTop.Checked = True Then
        SetTopWindow Me.hWnd, True
    Else
        SetTopWindow Me.hWnd, False
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    SetTopWindow Me.hWnd, False
    frmAbout.Show vbModal, Me
    SetTopWindow Me.hWnd, True
End Sub

Private Sub mnuHelpGotoNTWebsite_Click()
    GotoMyWebSite
End Sub

Private Sub mnuShowResultsWindow_Click()
    'Toggle state then show or hide
    mnuShowResultsWindow.Checked = mnuShowResultsWindow.Checked Xor True
    If mnuShowResultsWindow.Checked = True Then
        frmResults.Show
    Else
        frmResults.Hide
    End If
End Sub

Private Sub mnuTotalFileCount_Click()
    MsgBox "Total file count is " & lngTotalFileCount, vbOKOnly + vbInformation, "Total files processed"
End Sub

Private Sub optAllExtensions_Click()
    fraAllExtensions.Visible = True
    fraExtensionsSpecific.Visible = False
End Sub

Private Sub optReplacePresetMP3_Click()
    txtReplaceFrom.Text = "_"
    txtReplaceTo.Text = " "
End Sub

Private Sub optReplacePresetNone_Click()
    txtReplaceFrom.Text = ""
    txtReplaceTo.Text = ""
End Sub

Private Sub optSpecificExtensions_Click()
    fraAllExtensions.Visible = False
    fraExtensionsSpecific.Visible = True
End Sub

Private Sub tbsTab_Click()
    ' No need to change frame.
    If tbsTab.SelectedItem.Index - 1 = intCurFrame Then Exit Sub
    ' Otherwise, hide old frame, show new.
    fraFrame(tbsTab.SelectedItem.Index - 1).Visible = True
    fraFrame(intCurFrame).Visible = False
    ' Set mintCurFrame to new value.
    intCurFrame = tbsTab.SelectedItem.Index - 1
    
    'Makes sure the correct "Go!" button is the default
    Select Case intCurFrame
        Case 0: cmdRenameGo.Default = True
        Case 1: cmdExtensionsGo.Default = True
        Case 2: cmdReplaceGo.Default = True
    End Select
End Sub

Private Sub LoadSettings()
    '**** Load stored values from registry ****
    'On top ****
    If TestKey("AlwaysOnTop", "False") = "True" Then
        mnuFileOnTop.Checked = True
        SetTopWindow Me.hWnd, True
    Else                        'Normal
        mnuFileOnTop.Checked = False
        SetTopWindow Me.hWnd, False
    End If
    'Show results window ****
    If TestKey("ShowResults", "False") = "True" Then
        mnuShowResultsWindow.Checked = True
        frmResults.Show
    Else                        'Not displayed
        mnuShowResultsWindow.Checked = False
    End If
    'Current folder ****
    txtCurrentPath.Text = TestKey("CurrentPath", App.Path)
    ' Window Position ****
    TestKey "WindowMTop", 100
    TestKey "WindowMLeft", 100
    TestKey "WindowMWidth", 7245
    TestKey "WindowMHeight", 3690
    Top = GetKey("WindowMTop")
    Left = GetKey("WindowMLeft")
    Width = GetKey("WindowMWidth")
    Height = GetKey("WindowMHeight")
    'Counter for grand total of all files processed
    lngTotalFileCount = Val(TestKey("TotalFileCount", 0))
End Sub

Private Sub SaveSettings()
    SetKey "AlwaysOnTop", mnuFileOnTop.Checked
    SetKey "ShowResults", mnuShowResultsWindow.Checked
    SetKey "CurrentPath", txtCurrentPath.Text
    SetKey "WindowMTop", Top
    SetKey "WindowMLeft", Left
    SetKey "WindowMWidth", Width
    SetKey "WindowMHeight", Height
    SetKey "TotalFileCount", Str(lngTotalFileCount)
End Sub

Private Sub tmrCheckForNewVersionOfProg_Timer()
    'Checks for a new version of the program on the net
    'This would have been called in Load() but I hangs the program for 5 secs
    If InternetConnectionPresent(Winsock) Then
        CheckForNewVersionOfProgram Inet, "ntFileUtility", 1, 0, True
    End If
    tmrCheckForNewVersionOfProg.Enabled = False
End Sub

'This will let the user just drag the path to use on the program instead
'of hitting browse
Private Sub txtCurrentPath_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim strTemp As String
    If Data.GetFormat(vbCFFiles) Then  'If the data being dropped is a file or folder then we can go ahead
        Dim temp
        For Each temp In Data.Files
            'Get the first file/folder, the extract the path from it
            'C:\Folder\File.tmp
            If Mid(temp, Len(temp) - 3, 1) = "." Then
                strTemp = StrReturnLeftFromEnd(temp, "\")
                'If len = 2 then strTemp = C:, so we have to add an "\" to the end
                'if its longer than 2 then there is a folder after C: and no "\"
                'is needed
                If Len(strTemp) = 2 Then
                    txtCurrentPath.Text = strTemp & "\"
                Else
                    txtCurrentPath.Text = strTemp
                End If
            Else    'Its a folder, just add it unmodified to the path
                txtCurrentPath.Text = temp
            End If
            Exit For
        Next temp
    End If
End Sub

Private Sub txtCurrentPath_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    If Data.GetFormat(vbCFFiles) Then
        'If the data is in the proper format, inform the source of the action to be taken
        Effect = vbDropEffectCopy And Effect
        Exit Sub
    End If
    'If the data is not desired format, no drop
    Effect = vbDropEffectNone
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
      txtCurrentPath_OLEDragDrop Data, Effect, Button, Shift, x, y
End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    txtCurrentPath_OLEDragOver Data, Effect, Button, Shift, x, y, State
End Sub

