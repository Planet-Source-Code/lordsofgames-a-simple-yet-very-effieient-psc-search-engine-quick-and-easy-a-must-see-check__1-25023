VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   Caption         =   "Planet-Source-Code Search for Visual Basic Code              Status: HOLD"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8895
   Icon            =   "pscsearch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Stop"
      Height          =   375
      Left            =   2280
      TabIndex        =   22
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   255
      Left            =   2880
      TabIndex        =   21
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Reload"
      Height          =   255
      Left            =   2160
      TabIndex        =   20
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Foward"
      Height          =   255
      Left            =   1440
      TabIndex        =   19
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   255
      Left            =   720
      TabIndex        =   18
      Top             =   0
      Width           =   735
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Timer ohyeah 
      Enabled         =   0   'False
      Interval        =   7
      Left            =   8160
      Top             =   1800
   End
   Begin MSComctlLib.StatusBar yoyo 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   7635
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser results 
      Height          =   4695
      Left            =   0
      TabIndex        =   5
      Top             =   2880
      Width           =   8895
      ExtentX         =   15690
      ExtentY         =   8281
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   3975
      Begin VB.TextBox search 
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Text            =   "Fight"
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search!"
         Height          =   375
         Left            =   2280
         TabIndex        =   6
         Top             =   1440
         Width           =   1575
      End
      Begin VB.ListBox List1 
         Height          =   450
         ItemData        =   "pscsearch.frx":030A
         Left            =   120
         List            =   "pscsearch.frx":0314
         TabIndex        =   4
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox rpp 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Text            =   "10"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "QuickSearch:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Search Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Results per Page:"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "DateDescending = NewCode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   4800
      TabIndex        =   25
      Top             =   2520
      Width           =   6015
   End
   Begin VB.Label Label6 
      Caption         =   "Alphabetical = ABC "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   24
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "search type. Then SEARCH!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   600
      TabIndex        =   17
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "search type. Then SEARCH!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4800
      TabIndex        =   15
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "want per page. Choose a "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4800
      TabIndex        =   14
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "Then how many results you"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   13
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "Type in your search key word."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   12
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "Created by: Darwin Yu"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   11
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "PSC Search Engine"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4800
      TabIndex        =   10
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public StartingAddress As String
Dim mbDontNavigateNow As Boolean

Private Sub Command1_Click()
StartingAddress = ("http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?txtCriteria=" & search.Text & "&blnWorldDropDownUsed=FALSE&txtMaxNumberOfEntriesPerPage=" & rpp.Text & "&blnResetAllVariables=TRUE&lngWId=1&B1=Quick+Search&optSort=" & List1.Text)
results.Navigate StartingAddress
Form1.Caption = "Planet-Source-Code Search for Visual Basic Code/\Status: Logged on. Please be patient."
ohyeah.Enabled = True
End Sub

Private Sub Command2_Click()
results.GoBack
End Sub

Private Sub Command3_Click()
results.gofoward
End Sub

Private Sub Command4_Click()
results.Refresh
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Command6_Click()
results.Stop
End Sub

Private Sub ohyeah_Timer()
If results.Busy = False Then
        ohyeah.Enabled = False
        yoyo.SimpleText = results.LocationName
        Form1.Caption = "Planet-Source-Code Search for Visual Basic Code/\Status: Done."
    Else
        yoyo.SimpleText = "Working..."
    End If
End Sub
