VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "System Treeview Theft"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   ScaleHeight     =   379
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   286
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Set Path to App.Path"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   60
      Width           =   3555
   End
   Begin VB.PictureBox PicBrowse 
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   360
      ScaleHeight     =   337
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********Copyright PSST Software 2002**********************
'Submitted to Planet Source Code - October 2002
'If you got it elsewhere - they stole it from PSC.

'Written by MrBobo - enjoy
'Please visit our website - www.psst.com.au

'Prividing a Treeview for navigating local hard drives can be tricky.
'Getting it to also include the LAN is even trickier. The task when
'using a VB Treeview, whilst not impossible by any means, involves
'quite a lot of code and is never quite as fast as Windows.
'The job becomes even harder when trying to obtain all
'the System Icons.

'This is a quick demo of thievery. It demonstrates how to steal
'the Treeview from the BrowseForFolder dialog for use by
'us poor VB coders.

'Advantages:
'Speed
'System error handling
'Small amount of code
'Access to LAN(Network Neighborhood)
'System Icons without the overhead

'Disadvantages:
'There's only really 3 useful functions
'1.Set path
'2.Get Path
'3.Resize

'Issues:
'Project must load through Sub Main
'Project MUST close correctly


Option Explicit

Private Sub Command1_Click()
    'Example of how to set a new path
    ChangePath App.Path
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Return the Treeview to BrowseForFolder dialog
    'and close the hidden BrowseForFolder dialog
    CloseUp
End Sub

Public Sub PathChange()
    'Recieve path change from the Treeview
    Me.Caption = m_CurrentDirectory
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    PicBrowse.Move 25, 25, Me.ScaleWidth - 50, Me.ScaleHeight - 50
End Sub

Private Sub PicBrowse_Resize()
    'resize the Treeview as needed
    On Error Resume Next
    SizeTV 0, 0, PicBrowse.ScaleWidth, PicBrowse.ScaleHeight
End Sub
