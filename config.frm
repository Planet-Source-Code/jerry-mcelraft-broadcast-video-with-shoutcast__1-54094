VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form config 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuration"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton nevermind 
      Caption         =   "Don't Save"
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
      Left            =   3360
      TabIndex        =   19
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton save 
      Caption         =   "Save"
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
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin TabDlg.SSTab tabs 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4683
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "config.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "closegui"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "startwindows"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Shoutcast"
      TabPicture(1)   =   "config.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "ip"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "port"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "password"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Encoding"
      TabPicture(2)   =   "config.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label11"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label12"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label10"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label13"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "vp3"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "vp6"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "bitrate"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "custompro"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "delete"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "location"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "framerate"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "About"
      TabPicture(3)   =   "config.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label5"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label6"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label7"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label8"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label9"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      Begin VB.TextBox framerate 
         Height          =   285
         Left            =   -71760
         TabIndex        =   25
         ToolTipText     =   "Needed for some files, such as xvids. Leave blank for none."
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox location 
         Height          =   285
         Left            =   -73320
         TabIndex        =   23
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CheckBox delete 
         Caption         =   "Delete encoded files after they are used."
         Height          =   375
         Left            =   -74880
         TabIndex        =   21
         Top             =   1680
         Width           =   3255
      End
      Begin VB.CheckBox custompro 
         Caption         =   "Use a custom profile for encoding. This will be set when you click the Stream button."
         Height          =   615
         Left            =   -74880
         TabIndex        =   20
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CheckBox startwindows 
         Caption         =   "Start stream with Windows. This is good if this machine is a dedicated server and you would like it to always be streaming."
         Height          =   975
         Left            =   -74760
         TabIndex        =   18
         Top             =   480
         Width           =   3975
      End
      Begin VB.CheckBox closegui 
         Caption         =   "Close program on start of stream"
         Height          =   735
         Left            =   -74760
         TabIndex        =   17
         Top             =   1560
         Width           =   3975
      End
      Begin VB.ComboBox bitrate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "config.frx":0070
         Left            =   -74040
         List            =   "config.frx":008B
         TabIndex        =   16
         Text            =   "Select..."
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton vp6 
         Caption         =   "VP6/AAC"
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
         Left            =   -73080
         TabIndex        =   14
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton vp3 
         Caption         =   "VP3/MP3"
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
         Left            =   -74280
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox password 
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
         IMEMode         =   3  'DISABLE
         Left            =   -73680
         PasswordChar    =   "*"
         TabIndex        =   7
         Text            =   "changeme"
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox port 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -73680
         TabIndex        =   5
         Text            =   "8000"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox ip 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -73680
         TabIndex        =   3
         Text            =   "localhost"
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label9 
         Caption         =   "jr_mcelraft@yahoo.com"
         Height          =   255
         Left            =   960
         TabIndex        =   27
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label8 
         Caption         =   "SHOUTCast Video Broadcasting. This isnt like that lame ass AVPHONE!"
         Height          =   495
         Left            =   600
         TabIndex        =   26
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Force Framerate:"
         Height          =   375
         Left            =   -71760
         TabIndex        =   24
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Save encoded files to:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74880
         TabIndex        =   22
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Bitrate:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   15
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Type:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "by Jerry McElraft"
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "version 0.613"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   960
         TabIndex        =   10
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "nsvgui"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1200
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   8
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Port:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   6
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Host IP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "nsvgui v0.6 "
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   2880
      Width           =   2055
   End
End
Attribute VB_Name = "config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub custompro_Click()
If custompro.Value = 1 Then
    vp3.Enabled = False
    vp6.Enabled = False 'if custom profiles is enabled, turn all these off
    bitrate.Enabled = False
Else
    vp3.Enabled = True
    vp6.Enabled = True 'if custom profiles is disabled, turn all these on
    bitrate.Enabled = True
End If
End Sub

Private Sub Form_Load()
Dim fso As New FileSystemObject 'lets me work with files
Dim exist As Boolean 'this will remember if config existed
Dim configfile 'if it exists, this will read it
Dim codec As String

On Error Resume Next 'the config file went corrupt?! this will keep everything at the defaults

location.Text = App.Path + "\Videos" 'this will set the default

exist = fso.FileExists(App.Path + "\nsvgui.cfg") 'actually checks if it exists

If exist = True Then 'it exists? great, lets load it
    Set configfile = fso.OpenTextFile(App.Path + "\nsvgui.cfg", ForReading) 'prepare config for reading
    ip.Text = configfile.ReadLine
    port.Text = configfile.ReadLine
    password.Text = configfile.ReadLine
    codec = configfile.ReadLine
        If codec = "vp6aac" Then
            vp6.Value = True
        Else
        End If
        If codec = "vp3mp3" Then
            vp3.Value = True
        Else
        End If
    bitrate.Text = configfile.ReadLine
    framerate.Text = configfile.ReadLine
    custompro.Value = configfile.ReadLine
    delete.Value = configfile.ReadLine
    location.Text = configfile.ReadLine
    
Else 'doesn't exist? well lets just keep everything at defaults then
End If

configfile.Close 'don't need this anymore
Set fso = Nothing 'we read file already, no need to have this open

End Sub

Private Sub url_Click()
Shell "explorer.exe http://pissant.stinkyhands.com/nsvgui" 'opens the link
End Sub

Private Sub nevermind_Click()
Unload Me 'closes window
End Sub

Private Sub save_Click()
Dim fso As New FileSystemObject 'we need to work with files again
Dim configfile
Dim profile, profilesave
Dim codec As String, bitrte As String

Set configfile = fso.CreateTextFile(App.Path + "\nsvgui.cfg", True) 'prep file for writing

configfile.WriteLine (ip.Text)
configfile.WriteLine (port.Text)
configfile.WriteLine (password.Text)

If vp6.Value = True Then
configfile.WriteLine ("vp6aac")
configfile.WriteLine (bitrate.Text)
codec = "vp6aac"
Else
End If
If vp3.Value = True Then
configfile.WriteLine ("vp3mp3")
configfile.WriteLine (bitrate.Text)
codec = "vp3mp3"
Else
End If

configfile.WriteLine (framerate.Text)
configfile.WriteLine (custompro.Value)
configfile.WriteLine (delete.Value)
configfile.WriteLine (location.Text)

configfile.Close 'finished with file

If bitrate.Text = "64 kbps" Then
    bitrte = "064"
Else
End If
If bitrate.Text = "96 kbps" Then
    bitrte = "096"
Else
End If
If bitrate.Text = "128 kbps" Then
    bitrte = "128"
Else
End If
If bitrate.Text = "192 kbps" Then
    bitrte = "192"
Else
End If
If bitrate.Text = "320 kbps" Then
    bitrte = "320"
Else
End If

If custompro.Value = 0 Then
Set profile = fso.OpenTextFile(App.Path + "\profiles\" + bitrte + codec + ".config", ForReading)
Set profilesave = fso.CreateTextFile(App.Path + "\nsvenc.exe.config", True)
profilesave.Write (profile.ReadAll) 'copys appropiate profile as the nsvenc config
profilesave.Close
profile.Close
Else
End If

Set fso = Nothing 'finished with all files for now

Unload Me 'don't need this window anymore
End Sub
