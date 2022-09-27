VERSION 5.00
Begin VB.Form FHackMDB 
   Caption         =   "Hack MDB"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2985
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TStartString 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2055
      TabIndex        =   1
      Top             =   1170
      Width           =   3075
   End
   Begin VB.TextBox TNoOfChar 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2055
      TabIndex        =   2
      Top             =   1620
      Width           =   3075
   End
   Begin VB.TextBox TDSN 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2055
      TabIndex        =   0
      Top             =   705
      Width           =   3075
   End
   Begin VB.CommandButton CStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3315
      TabIndex        =   3
      Top             =   2175
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2445
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label2 
      Caption         =   "Start String"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   405
      TabIndex        =   6
      Top             =   1170
      Width           =   1515
   End
   Begin VB.Label Label4 
      Caption         =   "No of Characters"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   405
      TabIndex        =   5
      Top             =   1620
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "DSN"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   405
      TabIndex        =   4
      Top             =   765
      Width           =   1350
   End
End
Attribute VB_Name = "FHackMDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cn As ADODB.Connection

Private Sub CStart_Click()
Dim sPassword As String, iI As Long, i As Long, lNumberOfChar As Long
Dim lCharCodeArray(11) As String, iPropagate As Long, sString As String
    
    sString = StrReverse(TStartString.Text)
    lNumberOfChar = Val(TNoOfChar.Text)
    i = 0
    While i < Len(sString)
        lCharCodeArray(i) = Mid(sString, i + 1, 1)
        i = i + 1
    Wend
    
    iI = i - 1
    While (iI <= lNumberOfChar)
        i = 0
        iPropagate = 0
        sPassword = ""
        While i <= iI
            sPassword = lCharCodeArray(i) & sPassword
            If i = 0 Or iPropagate = 1 Then
                If lCharCodeArray(i) = "z" Then
                    lCharCodeArray(i) = incrementCode(lCharCodeArray(i))
                    iPropagate = 1
                    If iI = i Then
                        iPropagate = 0
                        iI = iI + 1
                    End If
                Else
                    lCharCodeArray(i) = incrementCode(lCharCodeArray(i))
                    iPropagate = 0
                End If
            End If
            i = i + 1
        Wend
        If openDatabaseWith(sPassword) Then
            Exit Sub
        End If
        Me.Caption = sPassword
    Wend
End Sub

Private Function incrementCode(sCode As String) As String
        Select Case sCode
        Case "0"
            sCode = "1"
            GoTo GoOut
        Case "1"
            sCode = "2"
            GoTo GoOut
        Case "2"
            sCode = "3"
            GoTo GoOut
        Case "3"
            sCode = "4"
            GoTo GoOut
        Case "4"
            sCode = "5"
            GoTo GoOut
        Case "5"
            sCode = "6"
            GoTo GoOut
        Case "6"
            sCode = "7"
            GoTo GoOut
        Case "7"
            sCode = "8"
            GoTo GoOut
        Case "8"
            sCode = "9"
            GoTo GoOut
        Case "9"
            sCode = "A"
            GoTo GoOut
        Case "A"
            sCode = "B"
            GoTo GoOut
        Case "B"
            sCode = "C"
            GoTo GoOut
        Case "C"
            sCode = "D"
            GoTo GoOut
        Case "D"
            sCode = "E"
            GoTo GoOut
        Case "E"
            sCode = "F"
            GoTo GoOut
        Case "F"
            sCode = "G"
            GoTo GoOut
        Case "G"
            sCode = "H"
            GoTo GoOut
        Case "H"
            sCode = "I"
            GoTo GoOut
        Case "I"
            sCode = "J"
            GoTo GoOut
        Case "J"
            sCode = "K"
            GoTo GoOut
        Case "K"
            sCode = "L"
            GoTo GoOut
        Case "L"
            sCode = "M"
            GoTo GoOut
        Case "M"
            sCode = "N"
            GoTo GoOut
        Case "N"
            sCode = "O"
            GoTo GoOut
        Case "O"
            sCode = "P"
            GoTo GoOut
        Case "P"
            sCode = "Q"
            GoTo GoOut
        Case "Q"
            sCode = "R"
            GoTo GoOut
        Case "R"
            sCode = "S"
            GoTo GoOut
        Case "S"
            sCode = "T"
            GoTo GoOut
        Case "T"
            sCode = "U"
            GoTo GoOut
        Case "U"
            sCode = "V"
            GoTo GoOut
        Case "V"
            sCode = "W"
            GoTo GoOut
        Case "W"
            sCode = "X"
            GoTo GoOut
        Case "X"
            sCode = "Y"
            GoTo GoOut
        Case "Y"
            sCode = "Z"
            GoTo GoOut
        Case "Z"
            sCode = "a"
        Case "a"
            sCode = "b"
            GoTo GoOut
        Case "b"
            sCode = "c"
            GoTo GoOut
        Case "c"
            sCode = "d"
            GoTo GoOut
        Case "d"
            sCode = "e"
            GoTo GoOut
        Case "e"
            sCode = "f"
            GoTo GoOut
        Case "f"
            sCode = "g"
            GoTo GoOut
        Case "g"
            sCode = "h"
            GoTo GoOut
        Case "h"
            sCode = "i"
            GoTo GoOut
        Case "i"
            sCode = "j"
            GoTo GoOut
        Case "j"
            sCode = "k"
            GoTo GoOut
        Case "k"
            sCode = "l"
            GoTo GoOut
        Case "l"
            sCode = "m"
            GoTo GoOut
        Case "m"
            sCode = "n"
            GoTo GoOut
        Case "n"
            sCode = "o"
            GoTo GoOut
        Case "o"
            sCode = "p"
            GoTo GoOut
        Case "p"
            sCode = "q"
            GoTo GoOut
        Case "q"
            sCode = "r"
            GoTo GoOut
        Case "r"
            sCode = "s"
            GoTo GoOut
        Case "s"
            sCode = "t"
            GoTo GoOut
        Case "t"
            sCode = "u"
            GoTo GoOut
        Case "u"
            sCode = "v"
            GoTo GoOut
        Case "v"
            sCode = "w"
            GoTo GoOut
        Case "w"
            sCode = "x"
            GoTo GoOut
        Case "x"
            sCode = "y"
            GoTo GoOut
        Case "y"
            sCode = "z"
            GoTo GoOut
        Case "z"
            sCode = "0"
            GoTo GoOut
    End Select
GoOut:
    incrementCode = sCode
End Function

Private Function openDatabaseWith(sPassword As String) As Boolean
On Error GoTo GoOut
        Set cn = New ADODB.Connection
        cn.ConnectionString = "DSN=" & TDSN.Text & ";Uid=;Pwd=" & sPassword & ";"
        cn.Open
        MsgBox sPassword
        openDatabaseWith = True
        Exit Function
GoOut:
    openDatabaseWith = False
End Function

