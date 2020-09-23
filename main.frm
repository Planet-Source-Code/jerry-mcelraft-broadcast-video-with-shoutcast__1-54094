VERSION 5.00
Begin VB.Form main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "nsvgui"
   ClientHeight    =   6840
   ClientLeft      =   1365
   ClientTop       =   1080
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox anotherlist 
      Height          =   375
      Left            =   4440
      TabIndex        =   28
      Text            =   "1"
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox oldnumber 
      Height          =   375
      Left            =   4440
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton updateplaylist 
      Caption         =   "Update Playlist"
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
      Left            =   240
      TabIndex        =   24
      Top             =   6240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton editheader 
      Caption         =   "Edit Header"
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
      Left            =   3840
      TabIndex        =   16
      Top             =   6240
      Width           =   1695
   End
   Begin VB.TextBox m3ulisttest 
      Height          =   375
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   5640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton addall 
      Caption         =   " >> ALL"
      Height          =   855
      Left            =   4440
      TabIndex        =   5
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton add 
      Caption         =   ">>"
      Height          =   855
      Left            =   4440
      TabIndex        =   4
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton exit 
      Caption         =   "Exit"
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
      Left            =   7440
      TabIndex        =   18
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton helpme 
      Caption         =   "Newbie Help"
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
      Left            =   5640
      TabIndex        =   17
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton configuration 
      Caption         =   "Configuration"
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
      Left            =   2040
      TabIndex        =   15
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton stream 
      Caption         =   "Stream"
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
      Left            =   240
      TabIndex        =   14
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Playlist"
      Height          =   6015
      Left            =   5040
      TabIndex        =   19
      Top             =   120
      Width           =   4335
      Begin VB.Frame Frame3 
         Caption         =   "Current contents"
         Height          =   975
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   4095
         Begin VB.ListBox old 
            Height          =   645
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   3855
         End
      End
      Begin VB.TextBox numberenc 
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Text            =   "0"
         Top             =   4800
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton listdel 
         Caption         =   "X"
         Height          =   735
         Left            =   3840
         TabIndex        =   8
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton listdown 
         Caption         =   "\/"
         Height          =   735
         Left            =   3840
         TabIndex        =   9
         Top             =   3240
         Width           =   375
      End
      Begin VB.CommandButton listup 
         Caption         =   "/\"
         Height          =   735
         Left            =   3840
         TabIndex        =   7
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton Clearlist 
         Caption         =   "Clear playlist"
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   5640
         Width           =   1215
      End
      Begin VB.CheckBox randomlist 
         Caption         =   "Random"
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
         Left            =   960
         TabIndex        =   12
         Top             =   5640
         Width           =   1455
      End
      Begin VB.TextBox playlistname 
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
         Left            =   2160
         TabIndex        =   10
         Text            =   "playlist"
         Top             =   5160
         Width           =   2055
      End
      Begin VB.CommandButton saveplaylist 
         Caption         =   "Save Playlist"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   5160
         Width           =   1095
      End
      Begin VB.ListBox playlist 
         Height          =   4545
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label labelenc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 file(s) need to be encoded"
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
         Left            =   120
         TabIndex        =   22
         Top             =   4800
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
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
         Left            =   1440
         TabIndex        =   20
         Top             =   5160
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Source Files"
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.FileListBox File 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2730
         Left            =   120
         Pattern         =   "*.nsv;*.m3u;*.avi;*.mpg;*.mpeg;*.mov;*.m2v;*.m1v;*.vob"
         TabIndex        =   3
         Top             =   3240
         Width           =   4095
      End
      Begin VB.DirListBox Folder 
         Height          =   2565
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   4095
      End
      Begin VB.DriveListBox Drive 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4095
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub add_Click()
Dim fso As New FileSystemObject 'we might need to mess with some files
Dim m3u 'this'll be the file to access
Dim filenamearray As Variant 'this well help determine extension
Dim fileextension As String 'this will store extension
Dim filenameonly As String
Dim linein As String 'this will help in reading playlists
Dim playlistlength As Integer, playlistline As Integer 'so we can read lines correctly
Dim needenc As Boolean

On Error Resume Next 'keep the program running if something happens

filenamearray = Split(File.filename, ".") 'split up filename by the .s
fileextension = filenamearray(UBound(filenamearray)) 'take everything after the last ., which will be the extension
filenameonly = filenamearray(LBound(filenamearray))

If fileextension = "nsv" Then 'if its already a nsv, lets just add to list
    playlist.AddItem (File.Path + "\" + File.filename) 'add to list
End If
If fileextension = "m3u" Then 'if it's a playlist, lets load it
    Set m3u = fso.OpenTextFile(File.Path + "\" + File.filename) 'sets list up for reading
    Do Until m3u.AtEndOfStream 'we read entire file
        linein = m3u.ReadLine 'reads a line
        m3ulisttest.Text = linein 'this will check to see if it is an extended m3u line
            If linein = "" Then 'if its blank, we don't want it
            Else
                If m3ulisttest.Text = "#" Then 'if its an extended comment, ignore
                Else
                filenamearray = Split(linein, ".")
                fileextension = filenamearray(UBound(filenamearray)) 'find extension, again
                If fileextension = "nsv" Then
                playlist.AddItem (linein) 'finally add it
                Else
                If fileextension = "mpg" Or "mpeg" Or "avi" Or "mov" Or "m2v" Or "m1v" Then
                playlist.AddItem (linein)
                numberenc.Text = numberenc.Text + 1
                labelenc.Caption = numberenc.Text + " file(s) need to be encoded"
                End If
                End If
                End If
            End If
    Loop 'loops our do until function
    m3u.Close 'closes the playlist
    fileextension = "nsv"
End If
If fileextension = "nsv" Then
Else
    numberenc.Text = numberenc.Text + 1
    labelenc.Caption = numberenc.Text + " file(s) need to be encoded"
    playlist.AddItem (File.Path + "\" + File.filename)
End If

Set fso = Nothing 'don't need this anymore
End Sub

Private Sub addall_Click()
Dim fso As New FileSystemObject
Dim filecount As Integer, filecur As Integer
Dim filenamearray As Variant 'this well help determine extension
Dim fileextension As String 'this will store extension
Dim linein As String, playlistlength As Integer, playlistline As Integer


filecount = File.ListCount
filecur = 0

Do While filecount > filecur
    filenamearray = Split(File.List(filecur), ".") 'split up filename by the .s
    fileextension = filenamearray(UBound(filenamearray)) 'take everything after the last ., which will be the extension

If fileextension = "nsv" Then 'if its already a nsv, lets just add to list
    playlist.AddItem (Folder.Path + "\" + File.List(filecur)) 'add to list
End If
If fileextension = "nsv" Or fileextension = "m3u" Then
Else 'if it ain't a nsv or m3u list, then lets do the drill with it
    numberenc.Text = numberenc.Text + 1
    labelenc.Caption = numberenc.Text + " file(s) need to be encoded"
    playlist.AddItem (Folder.Path + "\" + File.List(filecur))
End If
    
    filecur = filecur + 1
Loop

Set fso = Nothing
End Sub

Private Sub Clearlist_Click()
playlist.Clear
numberenc.Text = 0
labelenc.Caption = "0 file(s) need to be encoded"
End Sub

Private Sub configuration_Click()
config.Visible = True 'opens the config window
End Sub

Private Sub Drive_Change()
Folder.Path = Drive.Drive 'changes the folder box to match the drive
End Sub

Private Sub editheader_Click()
header.Visible = True
End Sub

Private Sub exit_Click()
Dim fso As New FileSystemObject
Dim configfile
Dim config1 As String, config2 As String, config3 As String, config4 As String
Dim config5 As String, config6 As String, config7 As String, config8 As String, config9 As String

On Error Resume Next 'we came this far, don't show an error now

Set configfile = fso.OpenTextFile(App.Path + "\nsvgui.cfg") 'opens current config
config1 = configfile.ReadLine 'reads config
config2 = configfile.ReadLine
config3 = configfile.ReadLine
config4 = configfile.ReadLine
config5 = configfile.ReadLine
config6 = configfile.ReadLine
config7 = configfile.ReadLine
config8 = configfile.ReadLine
config9 = configfile.ReadLine
configfile.Close 'closes so we can reopen and write to it
Set configfile = fso.OpenTextFile(App.Path + "\nsvgui.cfg", ForWriting) 'opens so we can write to it now
configfile.WriteLine (config1) 'rewrites config
configfile.WriteLine (config2)
configfile.WriteLine (config3)
configfile.WriteLine (config4)
configfile.WriteLine (config5)
configfile.WriteLine (config6)
configfile.WriteLine (config7)
configfile.WriteLine (config8)
configfile.WriteLine (config9)
configfile.WriteLine (Drive.Drive) 'adds drive we last used
configfile.WriteLine (Folder.Path) 'adds folder last used
configfile.Close 'closes file

End  'closes the app
End Sub

Private Sub Folder_Change()
File.Path = Folder.Path 'changes the file box to match the folder
End Sub

Private Sub Form_Load()
Dim fso As New FileSystemObject
Dim configfile
Dim exist As Boolean, wanthelp As Boolean

On Error Resume Next

exist = fso.FolderExists(App.Path + "\temp") 'does the temp folder exist?
If exist = False Then
    fso.CreateFolder (App.Path + "\temp") 'creates temp folder
Else
End If

exist = fso.FileExists(App.Path + "\nsvgui.cfg") 'does config exist?
If exist = False Then
    wanthelp = MsgBox("No config file was found, meaning this is probably your first time running. Would you like help with setting up your config?", vbYesNo, "No config was found")
    If wanthelp = True Then
        Shell "explorer.exe help\config.html"
        Load config
    Else
        Load config
    End If
Else
    Set configfile = fso.OpenTextFile(App.Path + "\nsvgui.cfg")
    configfile.SkipLine
    configfile.SkipLine
    configfile.SkipLine 'first 9 lines are useless right now
    configfile.SkipLine
    configfile.SkipLine
    configfile.SkipLine
    configfile.SkipLine
    configfile.SkipLine
    configfile.SkipLine
    Drive.Drive = configfile.ReadLine 'go back to last used drive
    Folder.Path = configfile.ReadLine 'go back to last used folder
    configfile.Close
End If

Set fso = Nothing

Randomize

End Sub

Private Sub helpme_Click()
On Error Resume Next
Shell "explorer help\index.html"
End Sub

Private Sub listdel_Click()
Dim sel As Integer
Dim filenamearray As Variant
Dim fileextension As String

On Error Resume Next

sel = playlist.ListIndex

filenamearray = Split(playlist.List(sel), ".")
fileextension = filenamearray(UBound(filenamearray))

playlist.RemoveItem (sel)

playlist.Selected(sel) = True

If fileextension = "nsv" Then
Else
numberenc.Text = numberenc.Text - 1
labelenc.Caption = numberenc.Text + " file(s) need to be encoded"
End If

End Sub

Private Sub listdown_Click()
Dim sel As Integer
Dim linedata As String

On Error Resume Next

sel = playlist.ListIndex 'this remembers which one u picked
linedata = playlist.List(sel + 1) 'this remembers what one under it was
playlist.List(sel + 1) = playlist.List(sel) 'this sets one under it as what u picked
playlist.List(sel) = linedata 'this makes one u picked the one under it
playlist.Selected(sel + 1) = True

End Sub

Private Sub listup_Click()
Dim sel As Integer
Dim linedata As String

On Error Resume Next

sel = playlist.ListIndex
linedata = playlist.List(sel - 1)
playlist.List(sel - 1) = playlist.List(sel)
playlist.List(sel) = linedata
playlist.Selected(sel - 1) = True

End Sub

Private Sub saveplaylist_Click()
Dim fso As New FileSystemObject
Dim m3uexport
Dim itemnumber As Integer

On Error Resume Next

Set m3uexport = fso.CreateTextFile(Folder.Path + playlistname.Text + ".m3u", True) 'makes a playlist in current folder
itemnumber = 0 'sets the first file in list up for processing

Do Until itemnumber > playlist.ListCount 'do this until we reach end of list
    m3uexport.WriteLine (playlist.List(itemnumber)) 'write it
    itemnumber = itemnumber + 1 'set next file for saving
Loop

m3uexport.Close
Set fso = Nothing

File.Refresh 'make it so list shows up immediately

End Sub

Private Sub stream_Click()
Dim fso As New FileSystemObject
Dim streambat
Dim config
Dim listlength As Integer, newnumber As Integer, current As Integer, rando As Integer
Dim ip As String, port As String, password As String, framerate As String, custom As String, delete As String, location As String
Dim filenamearray As Variant
Dim fileextension As String, filename As String, filenamewoext As String
Dim exist As Boolean, needenc As Boolean

On Error Resume Next

Set config = fso.OpenTextFile(App.Path + "\nsvgui.cfg") 'opens config
ip = config.ReadLine
port = config.ReadLine 'reads data from config
password = config.ReadLine
config.SkipLine
config.SkipLine
framerate = config.ReadLine
custom = config.ReadLine
delete = config.ReadLine
location = config.ReadLine
config.Close 'closes config

exist = fso.FolderExists(location)
If exist = False Then
    fso.CreateFolder (location)
Else
End If

Set streambat = fso.CreateTextFile(App.Path + "\stream.bat", True) 'create the file

streambat.WriteLine ("cd " + Chr(34) + App.Path + Chr(34)) 'this is needed to get the header to work correctly
streambat.WriteLine ("del *.nsv") 'deletes a playlist that may exist

If custom = "1" Then 'if custom profile is enabled, call up nsvenc's config
streambat.WriteLine (Chr(34) + App.Path + "\nsvenc" + Chr(34) + " /config")
Else
End If

listlength = playlist.ListCount 'find out how long playlist is
newnumber = 1000 'new names for files, so they arrange in the right order
current = 0

Do While listlength > 0 'do this until we get all files in playlist processed
        filenamearray = Split(playlist.List(current), "\")
        filename = filenamearray(UBound(filenamearray))
        filenamearray = Split(filename, ".")
        fileextension = filenamearray(UBound(filenamearray))
        filenamewoext = filenamearray(LBound(filenamearray))
    If randomlist.Value = 1 Then
        rando = Rnd * 10000 'gets a random number between 0 and 10000
            If fileextension = "nsv" Then
                streambat.WriteLine ("copy " + Chr(34) + playlist.List(current) + Chr(34) + " " + Chr(34) + App.Path + "\temp\" & rando & ".nsv" + Chr(34)) 'copy the file
            Else
            needenc = fso.FileExists(location + "\" + filenamewoext + ".nsv")
            If needenc = False Then
                If framerate = "" Then
                streambat.WriteLine (Chr(34) + App.Path + "\nsvenc" + Chr(34) + " " + Chr(34) + playlist.List(current) + Chr(34) + " " + Chr(34) + location + "\" + filenamewoext + ".nsv" + Chr(34))
                Else
                streambat.WriteLine (Chr(34) + App.Path + "\nsvenc" + Chr(34) + " /fr=" + framerate + " " + Chr(34) + playlist.List(current) + Chr(34) + " " + Chr(34) + location + "\" + filenamewoext + ".nsv" + Chr(34))
                End If
                streambat.WriteLine ("copy " + Chr(34) + location + "\" + filenamewoext + ".nsv" + Chr(34) + " " + Chr(34) + App.Path + "\temp\" & rando & ".nsv" + Chr(34)) 'copy the file
            Else
                streambat.WriteLine ("copy " + Chr(34) + location + "\" + filenamewoext + ".nsv" + Chr(34) + " " + Chr(34) + App.Path + "\temp\" & rando & ".nsv" + Chr(34)) 'copy the file
            End If
            End If
    Else
            If fileextension = "nsv" Then
                streambat.WriteLine ("copy " + Chr(34) + playlist.List(current) + Chr(34) + " " + Chr(34) + App.Path + "\temp\" & newnumber & ".nsv" + Chr(34)) 'copy the file
            Else
            needenc = fso.FileExists(location + "\" + filenamewoext + ".nsv")
            If needenc = False Then
                If framerate = "" Then
                streambat.WriteLine (Chr(34) + App.Path + "\nsvenc" + Chr(34) + " " + Chr(34) + playlist.List(current) + Chr(34) + " " + Chr(34) + location + "\" + filenamewoext + ".nsv" + Chr(34))
                Else
                streambat.WriteLine (Chr(34) + App.Path + "\nsvenc" + Chr(34) + " /fr=" + framerate + " " + Chr(34) + playlist.List(current) + Chr(34) + " " + Chr(34) + location + "\" + filenamewoext + ".nsv" + Chr(34))
                End If
                streambat.WriteLine ("copy " + Chr(34) + location + "\" + filenamewoext + ".nsv" + Chr(34) + " " + Chr(34) + App.Path + "\temp\" & newnumber & ".nsv" + Chr(34)) 'copy the file
            Else
                streambat.WriteLine ("copy " + Chr(34) + location + "\" + filenamewoext + ".nsv" + Chr(34) + " " + Chr(34) + App.Path + "\temp\" & newnumber & ".nsv" + Chr(34)) 'copy the file
            End If
            End If
    End If
    listlength = listlength - 1
    current = current + 1 'adds one so that it processes the new in the list
    newnumber = newnumber + 1 'adds one so that the playlist is ordered correctly
Loop
oldnumber.Text = newnumber
streambat.WriteLine (Chr(34) + App.Path + "\rar32" + Chr(34) + " m playlist.nsv -m0 -msnsv -en -ep temp\*.nsv") 'zips all the files
If delete = "1" Then
streambat.WriteLine ("del " + Chr(34) + location + "\*.nsv" + Chr(34)) 'deletes the nsvs copied during the making
Else
End If
streambat.WriteLine (Chr(34) + App.Path + "\nsvscsrc" + Chr(34) + " /sc " + ip + ":" + port + ":" + password + ":header.txt " + Chr(34) + App.Path + Chr(34)) 'starts streaming
streambat.Close

Shell (App.Path + "\stream.bat")

updateplaylist.Visible = True
stream.Visible = False
Frame3.Visible = True
playlist.Height = 3375
playlist.Top = 1320
current = 0
listlength = playlist.ListCount
Do While listlength > 0
    old.AddItem (playlist.List(current))
    current = current + 1
    listlength = listlength - 1
Loop
playlist.Clear

Set fso = Nothing
End Sub

Private Sub updateplaylist_Click()
Dim fso As New FileSystemObject
Dim streambat
Dim filenamearray As Variant
Dim fileextension As String, filename As String, filenamewoext As String
Dim exist As Boolean, needenc As Boolean
Dim listlength As Integer, newnumber As Integer, current As Integer, rando As Integer

Set streambat = fso.CreateTextFile(App.Path + "\update.bat", True)
streambat.WriteLine ("cd " + Chr(34) + App.Path + Chr(34)) 'this is needed to get the header to work correctly

listlength = playlist.ListCount 'find out how long playlist is
newnumber = oldnumber.Text 'new names for files, so they arrange in the right order
current = 0

Do While listlength > 0 'do this until we get all files in playlist processed
        filenamearray = Split(playlist.List(current), "\")
        filename = filenamearray(UBound(filenamearray))
        filenamearray = Split(filename, ".")
        fileextension = filenamearray(UBound(filenamearray))
        filenamewoext = filenamearray(LBound(filenamearray))
    If randomlist.Value = 1 Then
        rando = Rnd * 10000 'gets a random number between 0 and 10000
            If fileextension = "nsv" Then
                streambat.WriteLine ("copy " + Chr(34) + playlist.List(current) + Chr(34) + " " + Chr(34) + App.Path + "\temp\" & rando & ".nsv" + Chr(34)) 'copy the file
            Else
            needenc = fso.FileExists(location + "\" + filenamewoext + ".nsv")
            If needenc = False Then
                If framerate = "" Then
                streambat.WriteLine (Chr(34) + App.Path + "\nsvenc" + Chr(34) + " " + Chr(34) + playlist.List(current) + Chr(34) + " " + Chr(34) + location + "\" + filenamewoext + ".nsv" + Chr(34))
                Else
                streambat.WriteLine (Chr(34) + App.Path + "\nsvenc" + Chr(34) + " /fr=" + framerate + " " + Chr(34) + playlist.List(current) + Chr(34) + " " + Chr(34) + location + "\" + filenamewoext + ".nsv" + Chr(34))
                End If
                streambat.WriteLine ("copy " + Chr(34) + location + "\" + filenamewoext + ".nsv" + Chr(34) + " " + Chr(34) + App.Path + "\temp\" & rando & ".nsv" + Chr(34)) 'copy the file
            Else
                streambat.WriteLine ("copy " + Chr(34) + location + "\" + filenamewoext + ".nsv" + Chr(34) + " " + Chr(34) + App.Path + "\temp\" & rando & ".nsv" + Chr(34)) 'copy the file
            End If
            End If
    Else
            If fileextension = "nsv" Then
                streambat.WriteLine ("copy " + Chr(34) + playlist.List(current) + Chr(34) + " " + Chr(34) + App.Path + "\temp\" & newnumber & ".nsv" + Chr(34)) 'copy the file
            Else
            needenc = fso.FileExists(location + "\" + filenamewoext + ".nsv")
            If needenc = False Then
                If framerate = "" Then
                streambat.WriteLine (Chr(34) + App.Path + "\nsvenc" + Chr(34) + " " + Chr(34) + playlist.List(current) + Chr(34) + " " + Chr(34) + location + "\" + filenamewoext + ".nsv" + Chr(34))
                Else
                streambat.WriteLine (Chr(34) + App.Path + "\nsvenc" + Chr(34) + " /fr=" + framerate + " " + Chr(34) + playlist.List(current) + Chr(34) + " " + Chr(34) + location + "\" + filenamewoext + ".nsv" + Chr(34))
                End If
                streambat.WriteLine ("copy " + Chr(34) + location + "\" + filenamewoext + ".nsv" + Chr(34) + " " + Chr(34) + App.Path + "\temp\" & newnumber & ".nsv" + Chr(34)) 'copy the file
            Else
                streambat.WriteLine ("copy " + Chr(34) + location + "\" + filenamewoext + ".nsv" + Chr(34) + " " + Chr(34) + App.Path + "\temp\" & newnumber & ".nsv" + Chr(34)) 'copy the file
            End If
            End If
    End If
    listlength = listlength - 1
    current = current + 1 'adds one so that it processes the new in the list
    newnumber = newnumber + 1 'adds one so that the playlist is ordered correctly
Loop
streambat.WriteLine (Chr(34) + App.Path + "\rar32" + Chr(34) + " m playlist" + anotherlist.Text + ".nsv -m0 -msnsv -en -ep temp\*.nsv") 'zips all the files
streambat.WriteLine ("del " + Chr(34) + location + "\*.nsv" + Chr(34))
streambat.WriteLine ("exit")
streambat.Close

anotherlist.Text = anotherlist.Text + 1

Shell (App.Path + "\update.bat")

Set fso = Nothing

current = 0
listlength = playlist.ListCount
Do While listlength > 0
    old.AddItem (playlist.List(current))
    current = current + 1
    listlength = listlength - 1
Loop
playlist.Clear

End Sub
