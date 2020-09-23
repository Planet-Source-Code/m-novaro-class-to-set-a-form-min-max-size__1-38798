VERSION 5.00
Begin VB.Form frmMinMax 
   Caption         =   "Sub-Classing Demo 1"
   ClientHeight    =   3930
   ClientLeft      =   5265
   ClientTop       =   2010
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   4710
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Try to resize this form to see the effect.....                           NO FLICKERING!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmMinMax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private a As New cFormSize

Private Sub Form_Load()
    
    ' Initialize the class
    a.Init Me.hwnd
    
    ' Set max. and min. sizes, in twips !
    a.MaxWidth = 6000
    a.MaxHeight = 6000
    
    a.MinHeight = 3000
    a.MinWidth = 3000

    ' Resize the form to its max size. This is needed to avoid
    ' starting with a wrong size...
    a.ResizeToMax

End Sub
