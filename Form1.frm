VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mengkonversi Bilangan Binary ke Desimal"
   ClientHeight    =   1140
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   6150
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox textAkhir 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox textAwal 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton cmKonversi 
      Caption         =   "Konversi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   ":"
      Height          =   195
      Left            =   1200
      TabIndex        =   7
      Top             =   600
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   ":"
      Height          =   195
      Left            =   1200
      TabIndex        =   6
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Decimal"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Binary"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   435
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim S As String

Private Function BinaryKeDesimal(ByVal BinValue As String) As Long
    Dim lngValue As Long
    Dim x As Long
    Dim k As Long
        
        k = Len(BinValue)
        
        For x = k To 1 Step -1
            If Mid$(BinValue, x, 1) = "1" Then
                If k - x > 30 Then
                    lngValue = lngValue Or -2147483648#
                Else
                    lngValue = lngValue + 2 ^ (k - x)
                End If
            End If
        Next x
        
    BinaryKeDesimal = lngValue
End Function

Private Sub cmKonversi_Click()
    textAkhir.Text = BinaryKeDesimal(Val(textAwal.Text))
End Sub

Private Sub cmReset_Click()
    Dim Objek As Control
    For Each Objek In Me
        If TypeName(Objek) = "TextBox" Then
            Objek.Text = ""
        End If
        textAwal.SetFocus
    Next
End Sub

Private Sub Form_Load()
    S = "Isikan hanya angka 1 dan 0 ..."
    
    With Me
        .textAwal.Text = ""
        .textAkhir.Text = ""
        .textAkhir.Locked = True
        .textAkhir.BackColor = Form1.BackColor
    End With
    
    textAwal_LostFocus
End Sub

Private Sub textAwal_GotFocus()
    If textAwal.Text = S Then
        With textAwal
            .Text = ""
            .ForeColor = &H80000008
        End With
    End If
End Sub

Private Sub textAwal_LostFocus()
    If textAwal.Text = "" Then
        With textAwal
            .Text = S
            .ForeColor = &H80000000
        End With
    End If
End Sub
