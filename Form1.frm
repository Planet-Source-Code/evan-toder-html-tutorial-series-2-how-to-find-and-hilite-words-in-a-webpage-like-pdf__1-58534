VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   ScaleHeight     =   8835
   ScaleWidth      =   11745
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5625
      Top             =   4185
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnColor 
      Caption         =   "&sel color"
      Height          =   285
      Left            =   9990
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   45
      Width           =   870
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   8010
      TabIndex        =   3
      Text            =   "Visual Basic"
      Top             =   45
      Width           =   1950
   End
   Begin VB.CommandButton btnFind 
      Caption         =   "&find && hilite..."
      Height          =   285
      Left            =   6885
      TabIndex        =   2
      Top             =   45
      Width           =   1095
   End
   Begin SHDocVwCtl.WebBrowser WB1 
      Height          =   8430
      Left            =   0
      TabIndex        =   1
      Top             =   405
      Width           =   11760
      ExtentX         =   20743
      ExtentY         =   14870
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
      Location        =   ""
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Text            =   "http://www.planetsourcecode.com/vb"
      ToolTipText     =   "ENTER KEY TO GO"
      Top             =   45
      Width           =   6000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Dim oDoc     As MSHTML.HTMLDocument
  Dim oBody    As MSHTML.HTMLBody
  
  
 
Private Sub btnColor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  '
  'select a hilite color
  '
  CD1.ShowColor
  btnColor.BackColor = CD1.color
  
End Sub

Private Sub btnFind_Click()

 Call find_and_hilite(txtFind)

End Sub

Private Sub Form_Load()

  Show
  WB1.Navigate Text1
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  Set oBody = Nothing
  Set oDoc = Nothing
  
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 '
 'if enter key, kill the boooop  sound and navigate
 '
 If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    WB1.Navigate Text1
 End If
 
End Sub

Private Sub WB1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

  btnFind.Enabled = False
  btnColor.Enabled = False
  
End Sub

Private Sub WB1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
  
  'set ref to the browsers document
   Set oDoc = WB1.Document
   DoEvents
  'set ref to the documents body
   Set oBody = oDoc.body
   
   btnFind.Enabled = True
   btnColor.Enabled = True
   
End Sub

Private Sub find_and_hilite(str_to_find As String)

  Dim oRange   As MSHTML.IHTMLTxtRange
  Dim bfound   As Boolean
  
  'start the range to encompass ALL the pages text
   Set oRange = oBody.createTextRange
   
   Do 'tell the orange object to find the text (str_to_find (txtFind))
     bfound = oRange.findText(str_to_find)
     
     'If its found, select it, change its backcolor
     If bfound Then
       oRange.Select
       oDoc.execCommand "backcolor", False, btnColor.BackColor
       'this tells the orange object to resume the search with
       'the start point being the end of the word just found
       oRange.collapse False
     End If

     DoEvents
      'keep going til we dont find the word(s) anymore
   Loop Until Not (bfound)
   
    'scroll the page back to the top
   oDoc.parentWindow.Scroll 0, 0
   Set oRange = Nothing
   
End Sub
