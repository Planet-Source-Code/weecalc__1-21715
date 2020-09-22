VERSION 5.00
Begin VB.Form WC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WeeCalc"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1575
   Icon            =   "WC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   1575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdClose 
      Caption         =   "Close"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton CmdClr 
      Caption         =   "Clear"
      Default         =   -1  'True
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox TxtIn 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton CmdFc 
      Caption         =   "F° ->  C°"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton CmdCf 
      Caption         =   "C° ->  F°"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   1440
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   120
      Y1              =   480
      Y2              =   1080
   End
   Begin VB.Line Line1 
      X1              =   1440
      X2              =   1440
      Y1              =   480
      Y2              =   1080
   End
   Begin VB.Label LblA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "WC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdClose_Click()
Unload Me 'Closes the program
End Sub
Private Sub CmdClr_Click()
TxtIn.Text = "" 'Clears the text field
LblA.Caption = "" 'Clears the label field
Beep ' I think this is farely obvious
End Sub
Private Sub CmdCf_Click()
If TxtIn.Text = "" Then 'Makes sure that the input field is filled in
MsgBox "Please insert a temperature value" & vbNewLine & "in the provided text box"
Else
LblA.Caption = ((TxtIn.Text * 9) / 5) + 32 & "°F"
Beep
End If
End Sub
Private Sub CmdFc_Click()
If TxtIn.Text = "" Then 'Makes sure that the input field is filled in
MsgBox "Please insert a temperature value" & vbNewLine & "in the provided text box"
Else
LblA.Caption = ((TxtIn.Text - 32) * 5) / 9 & "°C"
Beep
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
MsgBox "Thank You for using WeeCalc" & vbNewLine & "                 Steph" 'Closing statement in a message box
End Sub
Private Sub TxtIn_Click()
TxtIn.Text = "" 'Clears the Text box on mouse click
End Sub
'I made this as an example on how to
'calculate within set paramaters
'Stephane Lessard
