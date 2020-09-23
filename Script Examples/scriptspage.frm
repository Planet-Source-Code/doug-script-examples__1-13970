VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form scriptspage 
   Caption         =   "Scripts"
   ClientHeight    =   7080
   ClientLeft      =   1245
   ClientTop       =   675
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   9465
   Begin VB.ComboBox cboDhtml 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   360
      Width           =   1575
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4935
      Left            =   1680
      TabIndex        =   1
      Top             =   2040
      Width           =   7575
      ExtentX         =   13361
      ExtentY         =   8705
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1695
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   2990
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"scriptspage.frx":0000
   End
End
Attribute VB_Name = "scriptspage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboDhtml_Click()
Select Case cboDhtml.Text
    Case "Slide In Menu I" 'Loads Pre-Selected Text Into The RichTextBox And Shows A Preview In The Browser Window
        RichTextBox1.Text = "A cool menu bar that automatically slides open from the left edge of the screen as the surfer moves the mouse over it. Moving the mouse out will cause it the bar to slide back in. Browsers other than NS 4+ and IE 4+ will simply see nothing."
        WebBrowser1.SetFocus 'Sets The Browser Window Active
        WebBrowser1.Navigate App.Path & "\SlideInMenuI.html"
    Case "Slide In Menu II" 'Loads Pre-Selected Text Into The RichTextBox And Shows A Preview In The Browser Window
        RichTextBox1.Text = " A keyboard-controlled menu bar that slides open/contracts with the press of a key. x is the key that will expand the menu, whereas z is the key that contracts it. Browsers other than NS 4+ and IE 4+ will simply see nothing."
        WebBrowser1.SetFocus 'Sets The Browser Window Active
        WebBrowser1.Navigate App.Path & "\SlideInMenuII.html"
    Case "Slide In Menu III" 'Loads Pre-Selected Text Into The RichTextBox And Shows A Preview In The Browser Window
        RichTextBox1.Text = "A manually controlled menu bar. Drag the bar to expand or contract it. Browsers other than NS 4+ and IE 4+ will simply see nothing."
        WebBrowser1.SetFocus 'Sets The Browser Window Active
        WebBrowser1.Navigate App.Path & "\SlideInMenuIII.html"
End Select


End Sub

Private Sub Form_Load()
MsgBox "All Scripts Used In This Demo Were Made At http://www.dynamicdrive.com", vbOKOnly
'Loads The Web Browser Window With A Blank HTML Page
WebBrowser1.Navigate App.Path & "\Blank.html"
'Fills The Combo Box
cboDhtml.AddItem "Slide In Menu I"
cboDhtml.AddItem "Slide In Menu II"
cboDhtml.AddItem "Slide In Menu III"
cboDhtml.Text = "DHTML Scripts"
End Sub

