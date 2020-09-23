VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Sparq - Effects"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   3300
      TabIndex        =   11
      Top             =   2220
      Width           =   1815
      Begin VB.Label lblSubmit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Submit User Info"
         Height          =   195
         Left            =   0
         TabIndex        =   12
         Top             =   180
         Width           =   1815
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   660
      Top             =   4200
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   4
      Left            =   2340
      TabIndex        =   4
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   3
      Left            =   1500
      MaxLength       =   14
      TabIndex        =   3
      Top             =   1080
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   1
      Top             =   300
      Width           =   2355
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   2355
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "User Input Form"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   60
      TabIndex        =   15
      Top             =   2580
      Width           =   2295
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   -180
      Top             =   2520
      Width           =   6195
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   420
      TabIndex        =   14
      Top             =   1935
      Width           =   585
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "This info. is not really submitted anywhere - so relax :)"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1080
      TabIndex        =   13
      Top             =   1935
      Width           =   3735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "                                                Pot Hedds Inc.                                           "
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1560
      Width           =   5235
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   60
      X2              =   5160
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   60
      X2              =   5160
      Y1              =   1810
      Y2              =   1810
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   60
      X2              =   5160
      Y1              =   1510
      Y2              =   1510
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   60
      X2              =   5160
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail Address"
      Height          =   195
      Left            =   2280
      TabIndex        =   9
      Top             =   810
      Width           =   1050
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ext."
      Height          =   195
      Left            =   1440
      TabIndex        =   8
      Top             =   810
      Width           =   270
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number"
      Height          =   195
      Left            =   60
      TabIndex        =   7
      Top             =   810
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      Height          =   195
      Left            =   2700
      TabIndex        =   6
      Top             =   30
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   30
      Width           =   750
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   4
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   1020
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   3
      Left            =   1440
      Shape           =   4  'Rounded Rectangle
      Top             =   1020
      Width           =   675
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   2
      Left            =   60
      Shape           =   4  'Rounded Rectangle
      Top             =   1020
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   1
      Left            =   2700
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   2475
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   60
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   2475
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function ResetAllControls()
    Dim X As Integer
    
    For X = 0 To Text1.Count - 1
            Text1(X).BackColor = &HC00000
            Text1(X).ForeColor = vbWhite
            Shape1(X).FillColor = &HC00000
    Next X
End Function

Private Sub Frame1_Click()
    lblSubmit_Click
End Sub

Private Sub lblSubmit_Click()
  Dim Extension As String
    If Len(Text1(3)) < 1 Then
        Extension = ""
    Else
        Extension = "  Ext. " & Text1(3)
    End If
    MsgBox "Name: " & StrConv(Text1(0), vbProperCase) & " " & StrConv(Text1(1), vbProperCase) & vbCrLf & _
           "Phone: " & Text1(2) & " " & Extension & vbCrLf & _
           "E-Mail Address: " & Text1(4), vbInformation, "User Info"
           
    MsgBox "If you like this code, dont be shy - vote for me." & vbCrLf & vbCrLf & _
             "If you want to make any changes to this code, feel free!"
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Dim X As Integer
    
    For X = 0 To Text1.Count - 1
        If X = Index Then
            Text1(X).BackColor = vbWhite
            Text1(X).ForeColor = vbBlack
            Shape1(X).FillColor = vbWhite
        Else
            Text1(X).BackColor = &HC00000
            Text1(X).ForeColor = vbWhite
            Shape1(X).FillColor = &HC00000
        End If
    Next X
    
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index))
End Sub

Private Sub Timer1_Timer()
    Dim String1 As String
    Dim String2 As String
    
    
    String1 = Left(Label8, 1)
    String2 = Right(Label8, Len(Label8) - 1)
    
    Label8 = String2 & String1
End Sub
