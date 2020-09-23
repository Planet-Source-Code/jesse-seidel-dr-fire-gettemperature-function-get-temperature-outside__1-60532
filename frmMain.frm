VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmMain 
   Caption         =   "Get Temperature"
   ClientHeight    =   375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4320
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   375
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   3360
      Top             =   2520
   End
   Begin VB.TextBox Text1 
      Height          =   195
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   150
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3840
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   45
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'By Jesse Seidel (drfire@gmail.com)

Private Sub Form_Load()
Timer1_Timer
End Sub

Private Sub Timer1_Timer()
Label1 = "Current temperature for 90210: " & GetTemperature(90210)
End Sub


Public Function GetTemperature(Zipcode As Long) As String

    Dim Expression As String, First As String, Last As String
    
    GetTemperature = Inet1.OpenURL("http://www.weather.com/weather/local/" & Zipcode & "?lswe=" & Zipcode & "&lwsa=WeatherLocalUndeclared")
    
    Expression = GetTemperature
    First = "<B CLASS=obsTempTextA>"
    Last = "</B>"
    
    IFirst = InStr(Expression, First$)
    ILast = InStr(Expression, Last$)


    If IFirst And ILast Then
        Expression = Mid(Expression, IFirst + Len(First$))
        ILast = InStr(Expression$, Last$)
        Expression = Left(Expression$, ILast - 1)
    End If
    
    GetTemperature = Expression
    GetTemperature = Replace(GetTemperature, "&deg;", "°")
    
    If Len(GetTemperature) >= 5 Then
        GetTemperature = "Error"
    End If

End Function
