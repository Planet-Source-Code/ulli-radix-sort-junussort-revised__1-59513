VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "SORT DEMO"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   9180
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox txNum 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1545
      MaxLength       =   7
      TabIndex        =   11
      Text            =   "1000000"
      Top             =   720
      Width           =   1440
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create now"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1545
      TabIndex        =   10
      Top             =   1305
      Width           =   1440
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "Verify"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4365
      TabIndex        =   1
      Top             =   2610
      Width           =   1665
   End
   Begin VB.Frame Frame3 
      Caption         =   " Performance test "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   4290
      TabIndex        =   2
      Top             =   810
      Width           =   4590
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Start tick: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   300
         TabIndex        =   8
         Top             =   375
         Width           =   1515
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Stop tick: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   300
         TabIndex        =   7
         Top             =   750
         Width           =   1515
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Elapsed time: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   300
         TabIndex        =   6
         Top             =   1125
         Width           =   1515
      End
      Begin VB.Label lblR 
         Alignment       =   1  'Rechts
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   1875
         TabIndex        =   5
         Top             =   375
         Width           =   2490
      End
      Begin VB.Label lblR 
         Alignment       =   1  'Rechts
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         Left            =   1875
         TabIndex        =   4
         Top             =   750
         Width           =   2490
      End
      Begin VB.Label lblR 
         Alignment       =   1  'Rechts
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   2
         Left            =   1875
         TabIndex        =   3
         Top             =   1125
         Width           =   2490
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Sorting "
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4365
      TabIndex        =   0
      Top             =   285
      Width           =   1515
   End
   Begin VB.Label lblR 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   6255
      TabIndex        =   12
      Top             =   2640
      Width           =   75
   End
   Begin VB.Label Label2 
      Caption         =   "Number of array elements to sort"
      Height          =   390
      Left            =   165
      TabIndex        =   9
      Top             =   690
      Width           =   1245
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Based on code by Junus (see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=59491&lngWId=1)

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Type Element
    Key         As Integer
    'Payload    As [whatever] 'payloads will slow down the sort of course
End Type

Private UnsortedElements()      As Element
Private SortedElements()        As Element

Private Sub cmdCreate_Click()

  Dim i         As Long

  'create unsorted

    cmdStart.Enabled = False
    ReDim UnsortedElements(1 To Val(txNum))
    For i = LBound(UnsortedElements) To UBound(UnsortedElements)
        UnsortedElements(i).Key = Int(Rnd * 65536) - 32768
    Next i
    cmdStart.Enabled = True

End Sub

Private Sub cmdStart_Click()

  Dim Tick      As Long

  Dim Counts(-32768 To 32767) As Long
  Dim Temp1     As Long
  Dim Temp2     As Long
  Dim i         As Long

    cmdVerify.Enabled = False
    lblR(0) = ""
    lblR(1) = ""
    lblR(2) = ""
    DoEvents
    ReDim SortedElements(LBound(UnsortedElements) To UBound(UnsortedElements))
    Tick = GetTickCount

    '''''
    'SORT
    '''''

    'count occurances
    For i = LBound(UnsortedElements) To UBound(UnsortedElements)
        Temp1 = UnsortedElements(i).Key
        Counts(Temp1) = Counts(Temp1) + 1
    Next i

    'convert occurences to output pointers
    Temp1 = LBound(SortedElements)
    For i = LBound(Counts) To UBound(Counts)
        Temp2 = Counts(i)
        Counts(i) = Temp1
        Temp1 = Temp1 + Temp2
    Next i

    'output to sorted
    For i = LBound(UnsortedElements) To UBound(UnsortedElements)
        Temp1 = UnsortedElements(i).Key
        SortedElements(Counts(Temp1)) = UnsortedElements(i)
        Counts(Temp1) = Counts(Temp1) + 1
    Next i

    '''''
    'DONE
    '''''

    lblR(0) = Tick
    lblR(1) = GetTickCount
    lblR(2) = lblR(1) - Tick & " msec"
    cmdVerify.Enabled = True

End Sub

Private Sub cmdVerify_Click()

  Dim i         As Long

    For i = LBound(SortedElements) To UBound(SortedElements) - 1
        If SortedElements(i).Key > SortedElements(i + 1).Key Then
            Exit For 'loop varying i
        End If
    Next i
    If i = UBound(SortedElements) Then
        lblR(3) = "Hurray - properly sorted :-)"
      Else 'NOT I...
        lblR(3) = "Seems you have a bug :-("
    End If

End Sub

':) Ulli's VB Code Formatter V2.18.3 (2005-Mrz-17 09:21)  Decl: 13  Code: 88  Total: 101 Lines
':) CommentOnly: 12 (11,9%)  Commented: 1 (1%)  Empty: 24 (23,8%)  Max Logic Depth: 3
