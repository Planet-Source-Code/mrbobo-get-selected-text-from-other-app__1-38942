VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get Selected Text"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   4455
      Begin VB.TextBox Text1 
         Height          =   1935
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   720
         Width           =   4215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Get Selected Text"
         Height          =   375
         Left            =   2880
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.TextBox txtHwnd 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Text            =   "Untitled - Notepad"
         Top             =   600
         Width           =   3255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Find"
         Height          =   375
         Left            =   3480
         TabIndex        =   1
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblEdit 
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblParent 
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Enter Window Caption"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3135
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********Copyright PSST Software 2002**********************
'Submitted to Planet Source Code - September 2002
'If you got it elsewhere - they stole it from PSC.

'Written by MrBobo - enjoy
'Please visit our website - www.psst.com.au

'How to get the selected text from another application is a
'question that is often asked in the Discussion Forum.
'Here's my rather simple attempt. It will not suit
'all situations but at least it gives a pointer as to
'how the task might be achieved
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Const GW_NEXT = 2
Private Const GW_CHILD = 5
Private Const WM_GETTEXT = &HD
Private Const WM_USER = &H400
Private Const EM_GETSEL = &HB0
Private Const EM_GETSELTEXT = (WM_USER + 62)
Private Const WM_GETTEXTLENGTH = &HE
Dim TargetHwnd As Long
Dim TargetParentHwnd As Long
Private Sub Command1_Click()
    Text1.Text = GetSelectedText(TargetHwnd)
End Sub

Private Sub Command2_Click()
    Dim EditParent As Long
    lblParent.Caption = ""
    lblEdit.Caption = ""
    'Search for a window with the desired caption
    TargetParentHwnd = FindWindow(vbNullString, txtHwnd.Text)
    If TargetParentHwnd > 0 Then
        lblParent.Caption = "Parent Window: " & TargetParentHwnd
        'Find the Edit window
        TargetHwnd = GetEditWindow(TargetParentHwnd)
        If TargetParentHwnd > 0 Then
            lblEdit.Caption = "Edit Window: " & TargetHwnd
        Else
            MsgBox "Failed to locate edit window"
        End If
    Else
        MsgBox "Failed to locate parent window"
    End If
End Sub
Public Function GetEditWindow(mHwnd As Long) As Long
    Dim hwnda As Long
    Dim ClWind As String * 5
    hwnda = GetWindow(mHwnd, GW_CHILD)
    'Loop through all child windows
    'to find the Edit window
    Do While hwnda <> 0
        GetClassName hwnda, ClWind, 5
        If Left(ClWind, 4) = "Edit" Then
            Exit Do
        End If
        'Didn't find it - try again
        hwnda = GetWindow(hwnda, GW_NEXT)
    Loop
    GetEditWindow = hwnda
End Function

Public Function GetSelectedText(mEditWindow As Long) As String
    Dim SelLocation As Long, SelStart As Long, SelEnd As Long
    Dim FullText As String, FullTextlength As Long
    'Get the selected text location
    SelLocation = SendMessage(mEditWindow, EM_GETSEL, 0, 0&)
    SelStart = SelLocation And &H7FFF 'LOWORD
    SelEnd = SelLocation \ &H10000 'HIWORD
    'Get all the text
    FullTextlength = SendMessage(mEditWindow, WM_GETTEXTLENGTH, 0&, 0&)
    FullText = String(FullTextlength, 0&)
    SendMessageByString mEditWindow, WM_GETTEXT, FullTextlength + 1&, FullText
    'Return only the selected text
    GetSelectedText = Mid(FullText, SelStart + 1, SelEnd - SelStart)
End Function
