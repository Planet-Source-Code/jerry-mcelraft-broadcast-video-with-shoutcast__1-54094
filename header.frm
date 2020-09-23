VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form header 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Header"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleMode       =   0  'User
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7223
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Basic"
      TabPicture(0)   =   "header.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Advanced"
      TabPicture(1)   =   "header.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "advanced"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "Other"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   20
         Top             =   2760
         Width           =   4215
         Begin VB.TextBox editother 
            Height          =   285
            Left            =   720
            TabIndex        =   23
            Top             =   900
            Width           =   3375
         End
         Begin VB.ListBox other 
            Height          =   645
            ItemData        =   "header.frx":0038
            Left            =   120
            List            =   "header.frx":003A
            TabIndex        =   8
            ToolTipText     =   "To edit any of these, use Advanced."
            Top             =   240
            Width           =   3975
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Edit:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   960
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Optional"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   4215
         Begin VB.TextBox aim 
            Height          =   285
            Left            =   840
            TabIndex        =   7
            Text            =   "anon"
            Top             =   960
            Width           =   3255
         End
         Begin VB.TextBox icq 
            Height          =   285
            Left            =   840
            TabIndex        =   6
            Text            =   "1234567"
            Top             =   720
            Width           =   3255
         End
         Begin VB.TextBox irc 
            Height          =   285
            Left            =   840
            TabIndex        =   5
            Text            =   "#random"
            Top             =   480
            Width           =   3255
         End
         Begin VB.TextBox url 
            Height          =   285
            Left            =   840
            TabIndex        =   4
            Text            =   "http:"
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "AIM:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "ICQ:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "IRC:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "URL:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Required"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   4215
         Begin VB.CheckBox pub 
            Caption         =   "Public"
            Height          =   195
            Left            =   3120
            TabIndex        =   3
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox genre 
            Height          =   285
            Left            =   960
            TabIndex        =   2
            Text            =   "video"
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox streamname 
            Height          =   255
            Left            =   960
            TabIndex        =   1
            Text            =   "nsvgui stream"
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Genre:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   14
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox advanced 
         Height          =   3615
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   11
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.CommandButton dontsave 
      Caption         =   "Don't Save"
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton save 
      Caption         =   "Save"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "nsvgui v0.6"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1800
      TabIndex        =   21
      Top             =   4320
      Width           =   1095
   End
End
Attribute VB_Name = "header"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dontsave_Click()
Unload Me
End Sub

Private Sub editother_Change()
On Error Resume Next

other.List(other.ListIndex) = editother.Text
End Sub

Private Sub Form_Load()
Dim fso As New FileSystemObject
Dim exist As Boolean, linein As String
Dim header
Dim headerarray As Variant
Dim headericy As String
Dim headervalue As String

exist = fso.FileExists(App.Path + "\header.txt")

On Error Resume Next

If exist = True Then
    Set header = fso.OpenTextFile(App.Path + "\header.txt")
    advanced.Text = header.ReadAll
    header.Close
Else
End If
    
If exist = True Then
    Set header = fso.OpenTextFile(App.Path + "\header.txt")
    Do Until header.AtEndOfStream
        linein = header.ReadLine
        headerarray = Split(linein, ":")
        headericy = headerarray(LBound(headerarray))
        headervalue = headerarray(UBound(headerarray))
        
        If headericy = "icy-name" Then
            streamname.Text = headervalue
        Else
        If headericy = "icy-genre" Then
            genre.Text = headervalue
        Else
        If headericy = "icy-pub" Then
            pub.Value = headervalue
        Else
        If headericy = "icy-url" Then
            url.Text = url.Text + headervalue
        Else
        If headericy = "icy-irc" Then
            irc.Text = headervalue
        Else
        If headericy = "icy-icq" Then
            icq.Text = headervalue
        Else
        If headericy = "icy-aim" Then
            aim.Text = headervalue
        Else
            other.AddItem (linein)
        End If
        End If
        End If
        End If
        End If
        End If
        End If
    Loop
header.Close
Else
End If
Set fso = Nothing
End Sub

Private Sub other_Click()
editother.Text = other.List(other.ListIndex)
End Sub

Private Sub save_Click()
Dim fso As New FileSystemObject
Dim header
Dim othernumber As Integer, current As Integer

Set header = fso.CreateTextFile(App.Path + "\header.txt", True)

If SSTab1.Tab = 1 Then
header.Write (advanced.Text)
Else
othernumber = other.ListCount
current = 0
header.WriteLine ("icy-name:" + streamname.Text)
header.WriteLine ("icy-genre:" + genre.Text)
If pub.Value = 1 Then
header.WriteLine ("icy-pub:1")
Else
header.WriteLine ("icy-pub:0")
End If
header.WriteLine ("icy-url:" + url.Text)
header.WriteLine ("icy-irc:" + irc.Text)
header.WriteLine ("icy-icq:" + icq.Text)
header.WriteLine ("icy-aim:" + aim.Text)
Do While othernumber > current
header.WriteLine (other.List(current))
current = current + 1
Loop
End If

header.Close
Set fso = Nothing
Unload Me
End Sub
