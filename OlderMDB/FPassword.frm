VERSION 5.00
Begin VB.Form FHackMDB 
   Caption         =   "Hack MDB"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3690
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TStart 
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
      Top             =   1200
      Width           =   3075
   End
   Begin VB.TextBox TEnd 
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
      Top             =   1680
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
      TabIndex        =   3
      Top             =   2175
      Width           =   3075
   End
   Begin VB.TextBox TDatabase 
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
      TabIndex        =   4
      Top             =   2730
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   1140
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
      TabIndex        =   8
      Top             =   2175
      Width           =   1515
   End
   Begin VB.Label Label3 
      Caption         =   "End Character"
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
      TabIndex        =   7
      Top             =   1695
      Width           =   1350
   End
   Begin VB.Label Label2 
      Caption         =   "Start Character"
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
      Top             =   1230
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "Database"
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
      Top             =   765
      Width           =   1350
   End
End
Attribute VB_Name = "FHackMDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function getPasswordString(lChar As Long) As String
    getPasswordString = Chr$(lChar)
End Function

Private Sub CStart_Click()
Dim db As Database, sPassword As String, iI As Long, iK As Long, lNumberOfChar As Long
Dim lCharCodeArray(11) As Long, lPropogate As Long, lStartChar As Long, lEndChar As Long
    
    lStartChar = Val(TStart.Text)
    lEndChar = Val(TEnd.Text)
    lNumberOfChar = Val(TNoOfChar.Text)
    
    iI = 1
    iK = 0
    While iK <= 11
        lCharCodeArray(iK) = lStartChar - 1
        iK = iK + 1
    Wend
    While (iI <= lNumberOfChar)
        
        iK = 1
        sPassword = ""
        lPropogate = 0
        
        While iK <= iI
            sPassword = getPasswordString(lCharCodeArray(iK)) & sPassword
            If iK = 1 Or lPropogate = 1 Then
                If lCharCodeArray(iK) = lEndChar Then
                    lCharCodeArray(iK) = lStartChar
                    lPropogate = 1
                    If iI = iK Then
                        iI = iI + 1
                    End If
                Else
                    lCharCodeArray(iK) = lCharCodeArray(iK) + 1
                    lPropogate = 0
                End If
            End If
            iK = iK + 1
        Wend
        If openDatabaseWith(sPassword) Then
            Exit Sub
        End If
        Me.Caption = sPassword

    Wend
End Sub

Private Function openDatabaseWith(sPassword As String) As Boolean
On Error GoTo GoOut
        Set db = OpenDatabase(TDatabase.Text & ".mdb", False, False, "MS Access;PWD=" & sPassword)
        MsgBox sPassword
        openDatabaseWith = True
        Exit Function
GoOut:
    openDatabaseWith = False
End Function
