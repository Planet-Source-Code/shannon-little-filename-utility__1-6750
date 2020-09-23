VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmResults 
   Caption         =   "Results"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmResults.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList iglListViewImages 
      Left            =   3840
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResults.frx":27A2
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResults.frx":2D3E
            Key             =   "File"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   2880
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   556
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin MSComctlLib.ListView lstResults 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      Icons           =   "iglListViewImages"
      SmallIcons      =   "iglListViewImages"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Type"
         Text            =   "Type"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "From"
         Text            =   "From"
         Object.Width           =   5115
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "To"
         Text            =   "To"
         Object.Width           =   5115
      EndProperty
   End
   Begin VB.Menu mnuResults 
      Caption         =   "Results"
      Visible         =   0   'False
      Begin VB.Menu mnuResultsClearList 
         Caption         =   "&Clear list"
      End
      Begin VB.Menu mnuResultsHyphen1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResultsCloseWindow 
         Caption         =   "Close &Results Window"
      End
   End
End
Attribute VB_Name = "frmResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lngFileCount As Long
Dim bNoReset As Boolean



Private Sub Form_Load()
    ' Window Position ****
    TestKey "WindowRTop", 100
    TestKey "WindowRLeft", 100
    TestKey "WindowRWidth", 6555
    TestKey "WindowRHeight", 3150
    Top = GetKey("WindowRTop")
    Left = GetKey("WindowRLeft")
    Width = GetKey("WindowRWidth")
    Height = GetKey("WindowRHeight")
    StatusBar.SimpleText = "Total files processed this instance: " & lngFileCount
    bNoReset = False
End Sub

Private Sub Form_Resize()
    DoEvents
    If WindowState <> vbMinimized Then
        If Height < 1500 Then Height = 1500
        If Width < 2000 Then Width = 2000
        lstResults.Height = ScaleHeight - StatusBar.Height
        lstResults.Width = ScaleWidth
    End If
End Sub

Public Sub UnloadForm()
    bNoReset = True 'Used to tell the form not to change then check mark when unloaded
    'This is used for when the program shuts down, its doesn't remove the checkmark
    'Because then it is saved to the registry as not being open when the prog was closed
    Unload frmResults
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If bNoReset = False Then
        'The X was clicked, don't unload, just hide
        'Only unload when given the command from frmMain "UnloadForm()" as above
        frmMain.mnuShowResultsWindow.Checked = False
        frmResults.Hide
        Cancel = True   'Keep the window from being unloaded
        Exit Sub
    End If
        
    SetKey "WindowRTop", Top
    SetKey "WindowRLeft", Left
    SetKey "WindowRWidth", Width
    SetKey "WindowRHeight", Height
End Sub

Public Sub IncrementCounter()
    'Adds one to the total file count the results window keep
    lngFileCount = lngFileCount + 1
    StatusBar.SimpleText = "Total files processed this instance: " & lngFileCount
End Sub

Private Sub lstResults_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        'Displays the ClearList menu when the results list is right-clicked on
        Me.PopupMenu mnuResults, , , , mnuResultsClearList
    End If
End Sub

Private Sub mnuResultsClearList_Click()
    lstResults.ListItems.Clear
End Sub

Private Sub mnuResultsCloseWindow_Click()
    Unload Me
End Sub
