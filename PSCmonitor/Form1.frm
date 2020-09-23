VERSION 5.00
Object = "{683364A1-B37D-11D1-ADC5-006008A5848C}#1.0#0"; "DHTMLED.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "PSC submission monitor"
   ClientHeight    =   2475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   165
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   528
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   270
      Left            =   9240
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      _Version        =   393216
      Min             =   60
      Max             =   600
      SelStart        =   60
      TickStyle       =   3
      Value           =   60
   End
   Begin DHTMLEDLibCtl.DHTMLEdit DHTMLEdit1 
      Height          =   135
      Left            =   9960
      TabIndex        =   7
      Top             =   600
      Width           =   150
      ActivateApplets =   0   'False
      ActivateActiveXControls=   0   'False
      ActivateDTCs    =   -1  'True
      ShowDetails     =   0   'False
      ShowBorders     =   0   'False
      Appearance      =   1
      Scrollbars      =   0   'False
      ScrollbarAppearance=   1
      SourceCodePreservation=   -1  'True
      AbsoluteDropMode=   0   'False
      SnapToGrid      =   0   'False
      SnapToGridX     =   50
      SnapToGridY     =   50
      BrowseMode      =   0   'False
      UseDivOnCarriageReturn=   0   'False
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   7860
      TabIndex        =   4
      Top             =   1875
      Width           =   7920
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00404040&
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   7860
      TabIndex        =   3
      Top             =   2175
      Width           =   7920
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7680
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "b0"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0354
            Key             =   "b.5"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":06A8
            Key             =   "b1"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":09FC
            Key             =   "b1.5"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0D50
            Key             =   "b2"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10A4
            Key             =   "b2.5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":13F8
            Key             =   "b3"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":174C
            Key             =   "b3.5"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1AA0
            Key             =   "b4"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1DF4
            Key             =   "b4.5"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2148
            Key             =   "b5"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":249C
            Key             =   "prog"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2608
            Key             =   "vote"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":295C
            Key             =   "novote"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5520
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   9737
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Name"
         Object.Tag             =   "Name"
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Compat"
         Text            =   "Compatability"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Level"
         Text            =   "Level"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "views"
         Text            =   "Views/Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Rating"
         Text            =   "Rating"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "RatingDetails"
         Text            =   "RatingDetails"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8280
      Top             =   1680
   End
   Begin VB.TextBox txtauthorurl 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   10215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Monitor"
      Height          =   270
      Left            =   8280
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lbldelay 
      Caption         =   "60 second refresh"
      Height          =   255
      Left            =   4680
      TabIndex        =   9
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter URL for PSC (all submissions by this author) URL"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim regEx1, regEx2, regEx3, regEx4, regEx5, regEx6 ' Create variable.
Dim Matches1, Matches2, Matches3, Matches4, Match1, Matches6
Dim mycount, go, votecount, totalviews As Long
Dim strvotes, strprevvotes, rating, votes, views, rateicon, voteicon, tmpstr As String

Sub saveSettings()
On Error Resume Next
PutIni "PSCmon.ini", "settings", "txtauthorurl", Me.txtauthorurl
PutIni "PSCmon.ini", "settings", "Slider1", Me.Slider1.Value
' save settings for form size
PutIni "PSCmon.ini", "settings", "formwidth", Me.Width
PutIni "PSCmon.ini", "settings", "formheight", Me.Height
PutIni "PSCmon.ini", "settings", "windowstate", Me.WindowState

' Save listview settings
Dim colhead As ColumnHeader
For Each colhead In Me.ListView1.ColumnHeaders
    'Debug.Print colhead.Width & colhead.Key & colhead.position
    PutIni "PSCmon.ini", "ListView1-" & colhead.Key, "Width", colhead.Width
    PutIni "PSCmon.ini", "ListView1-" & colhead.Key, "Position", colhead.Position
Next
End Sub

Sub loadSettings()
On Error Resume Next
Me.txtauthorurl = GetIni("PSCmon.ini", "settings", "txtauthorurl", "")
Me.Slider1.Value = GetIni("PSCmon.ini", "settings", "Slider1", "60")
lbldelay.Caption = Slider1.Value & " second refresh"
' get setting for form size
Me.Width = GetIni("PSCmon.ini", "settings", "formwidth", Me.Width)
Me.Height = GetIni("PSCmon.ini", "settings", "formheight", Me.Height)
Me.WindowState = GetIni("PSCmon.ini", "settings", "windowstate", Me.WindowState)

'load list view settings
Dim colhead As ColumnHeader
For Each colhead In Me.ListView1.ColumnHeaders
   colhead.Width = GetIni("PSCmon.ini", "ListView1-" & colhead.Key, "Width", "90")
   colhead.Position = GetIni("PSCmon.ini", "ListView1-" & colhead.Key, "Position", colhead.Position)
Next
End Sub
Private Sub Command1_Click()
go = 1
download
End Sub

Private Sub Command2_Click()
Select Case Command2.Caption
    Case "Pause"
        go = 0
        Timer1.Enabled = False
        Command2.Caption = "Continue"
    Case "Continue"
        go = 1
        Timer1.Enabled = True
        Command2.Caption = "Pause"
End Select
End Sub

Sub download()
On Error GoTo Bail
    Command2.Enabled = False
    If Me.txtauthorurl = "" Then
        MsgBox "You must enter a URL"
        Exit Sub
    End If
    status "Downloading document"
    status2 ""
    DHTMLEdit1.LoadURL Me.txtauthorurl
    'WebBrowser1.Navigate Me.txtauthorurl
    go = 1
Exit Sub
Bail:
status "An error occured downloading the page"
go = 0
Timer1.Enabled = False

End Sub

Private Sub DHTMLEdit1_DocumentComplete()
status "Download complete"
If go = 1 Then parsepage
End Sub

Private Sub Form_Load()
loadSettings
'----------------------------------------------
'Define reg exp for page content this retrieves a list of HTML for each program
'----------------------------------------------
    Set regEx1 = New RegExp ' get program list
    regEx1.Pattern = "<!--descrip-->[\w\W]*?<!description>"
    regEx1.IgnoreCase = True
    regEx1.Global = True
'----------------------------------------------
'Define reg exp to extract info this gets the info for each program
'----------------------------------------------
    Set regEx2 = New RegExp ' get program content
    regEx2.Pattern = "<!--descrip-->[\w\W]*?alt=""(.*?)\""[\w\W]*?<!--code compat-->([^<]*)[\w\W]*?<!--level-->(\w*)[\w\W]*?<!--views/date submitted--><TD><!i><FONT Size=1 >([\w\W]*?)<BR>([\w\W]*?)</TD><!--user rating-->[\w\W]*<!i><center>([\w\W]*?)<BR></center>"
    '<!--views/date submitted--><TD><!i><FONT Size=1 >333 since<BR>10/22/2001 6:09:47 AM</TD><!--user rating-->
    regEx2.IgnoreCase = True
    regEx2.Global = True

'----------------------------------------------
'Define reg exp for full rating
'----------------------------------------------
    Set regEx3 = New RegExp ' get program list
    regEx3.Pattern = "RatingSmall.jpg"
    regEx3.IgnoreCase = True
    regEx3.Global = True
'----------------------------------------------
'Define reg exp for half rating
'----------------------------------------------
    Set regEx4 = New RegExp ' get program list
    regEx4.Pattern = "RatingHalfSmall.jpg"
    regEx4.IgnoreCase = True
    regEx4.Global = True
'----------------------------------------------
'Define reg exp to remove tags from votes
'----------------------------------------------
  Set regEx5 = New RegExp
  regEx5.Pattern = "(<[^>]+>)|([^a-z0-9:/ ])|[ ]{2}"
  regEx5.IgnoreCase = True
  regEx5.Global = True
'----------------------------------------------
'Define reg exp to extract first number from string
'----------------------------------------------
  Set regEx6 = New RegExp
  regEx6.Pattern = "\d*"
  regEx6.IgnoreCase = False
  regEx6.Global = False
End Sub

Sub parsepage()
    On Error Resume Next
    ListView1.ListItems.Clear
    Set Matches1 = regEx1.Execute(DHTMLEdit1.DocumentHTML)    ' Execute search.
    If Matches1.Count > 0 Then
        Timer1.Enabled = True ' only start timer if page parses
        Command2.Enabled = True
    Else
        status2 "Couldn't Parse HTML. Are you sure you entered a correct (Planet-source-code.com / all submisions by this author) URL?"
    Exit Sub
    End If
        
        For Each Match1 In Matches1
            Set Matches2 = regEx2.Execute(Match1.Value)
            Set Matches3 = regEx3.Execute(Match1.Value)
            Set Matches4 = regEx4.Execute(Match1.Value)
            rating = (Matches3.Count) + (Matches4.Count * 0.5)
            rateicon = "b" & rating
            votes = Matches2(0).SubMatches(5)
            votes = regEx5.Replace(votes, "")
            votes = Replace(votes, "Users", "Users ")
            views = Matches2(0).SubMatches(3) & " " & Matches2(0).SubMatches(4)
            views = regEx5.Replace(views, "")
            If votes = "Unrated" Then voteicon = "novote" Else: voteicon = "vote"
                        ListView1.ListItems.Add , Matches2(0).SubMatches(0), Matches2(0).SubMatches(0), "prog", "prog"
                        ListView1.ListItems(Matches2(0).SubMatches(0)).ListSubItems.Add , "Compat", Matches2(0).SubMatches(1)
                        ListView1.ListItems(Matches2(0).SubMatches(0)).ListSubItems.Add , "Level", Matches2(0).SubMatches(2)
                        ListView1.ListItems(Matches2(0).SubMatches(0)).ListSubItems.Add , "Views", views
                        ListView1.ListItems(Matches2(0).SubMatches(0)).ListSubItems.Add , "Rating", rating, rateicon
                        ListView1.ListItems(Matches2(0).SubMatches(0)).ListSubItems.Add , "RatingDetails", votes, voteicon
        strvotes = strvotes & votes
        'tally totalviews
         Set Matches6 = regEx6.Execute(views)
         totalviews = totalviews + Matches6(0).Value
        Next
        Me.Caption = totalviews & " totalviews - PSCmonitor"
        totalviews = 0
    If strvotes <> strprevvotes Then ' first time or somebody voted notify user
        strprevvotes = strvotes
        If votecount > 0 Then 'Wow, somebody voted
            Beep
            status2 "You recieved a new vote"
        Else
            Beep
            votecount = 1
            status2 "From now on I will only Beep when you get a new vote."
        End If
    Else
            status2 "No votes recieved for any of your progies in the last " & Me.Slider1.Value & " seconds."
    End If
    strvotes = ""
End Sub

Sub status(strmsg As String)
    Picture1.Cls
    Picture1.Print strmsg
End Sub
Sub status2(strmsg As String)
    Picture2.Cls
    Picture2.Print strmsg
End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Or 0 Then Exit Sub
If Me.Width < 8500 Then Me.Width = 8500
If Me.Height < 3000 Then Me.Height = 3000
Me.ListView1.Left = 0
Me.ListView1.Height = Me.ScaleHeight - Me.ListView1.Top - Me.Picture1.Height - Me.Picture2.Height
Me.ListView1.Width = Me.ScaleWidth
Me.txtauthorurl.Left = 0
Me.txtauthorurl.Width = ScaleWidth
Me.Command2.Left = ScaleWidth - Command2.Width
Me.Command1.Left = Command2.Left - Command1.Width - 5
Me.DHTMLEdit1.Left = Command2.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
saveSettings
End Sub

Private Sub Slider1_Scroll()
lbldelay.Caption = Slider1.Value & " second refresh"
End Sub

Private Sub Timer1_Timer()
 mycount = mycount + 1
 If mycount > Slider1.Value Then ' do it all over again
    Timer1.Enabled = False
    mycount = 0
    download
 Else
    status "Pausing " & Slider1.Value & " seconds so we don't hurt the site ..." & mycount
 End If
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
If go = 1 Then
    status "Parsing HTML"
    parsepage
    status "Done"
End If
End Sub

