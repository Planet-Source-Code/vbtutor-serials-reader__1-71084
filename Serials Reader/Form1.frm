VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Serials Reader"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7755
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   410
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   517
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Operating system"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   270
      TabIndex        =   11
      Top             =   3960
      Width           =   7215
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   10
         Left            =   1800
         TabIndex        =   22
         Top             =   1080
         Width           =   4980
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   9
         Left            =   1800
         TabIndex        =   13
         Top             =   720
         Width           =   4980
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   8
         Left            =   1800
         TabIndex        =   12
         Top             =   360
         Width           =   4980
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "HDD Model"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   330
         TabIndex        =   23
         Top             =   1080
         Width           =   1290
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1005
         TabIndex        =   15
         Top             =   720
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "SN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   360
         Width           =   345
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "MotherBoard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   270
      TabIndex        =   6
      Top             =   1920
      Width           =   7215
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   7
         Left            =   1845
         TabIndex        =   20
         Top             =   1440
         Width           =   4980
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   6
         Left            =   1845
         TabIndex        =   18
         Top             =   1080
         Width           =   4980
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   1845
         TabIndex        =   8
         Top             =   360
         Width           =   4980
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   5
         Left            =   1845
         TabIndex        =   7
         Top             =   720
         Width           =   4980
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "BIOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1050
         TabIndex        =   21
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   945
         TabIndex        =   19
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "SN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   360
         Width           =   345
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Manufacturer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   150
         TabIndex        =   9
         Top             =   720
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "CPU"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   270
      TabIndex        =   1
      Top             =   120
      Width           =   7215
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   1875
         TabIndex        =   16
         Top             =   720
         Width           =   4980
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   1875
         TabIndex        =   3
         Top             =   360
         Width           =   4980
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   1875
         TabIndex        =   2
         Top             =   1080
         Width           =   4980
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Manufacturer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   165
         TabIndex        =   17
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1020
         TabIndex        =   5
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "SN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   1080
         Width           =   345
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3150
      TabIndex        =   0
      Top             =   5640
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Const Arr = 10

Private Sub Form_load()
  Dim SWbemSet(Arr) As SWbemObjectSet
  Dim SWbemObj As SWbemObject
  Dim varObjectToId(Arr) As String
  Dim varSerial(Arr) As String
  Dim i, j As Integer
  On Error Resume Next
  varObjectToId(1) = "Win32_Processor,Name"
  varObjectToId(2) = "Win32_Processor,Manufacturer"
  varObjectToId(3) = "Win32_Processor,ProcessorId"
  varObjectToId(4) = "Win32_BaseBoard,SerialNumber"
  varObjectToId(5) = "Win32_BaseBoard,manufacturer"
  varObjectToId(6) = "Win32_Baseboard,product"
  varObjectToId(7) = "Win32_BIOS,Manufacturer"
  varObjectToId(8) = "Win32_OperatingSystem,SerialNumber"
  varObjectToId(9) = "Win32_OperatingSystem,Caption"
  varObjectToId(10) = "Win32_DiskDrive,Model"
  For i = 1 To Arr
    Set SWbemSet(i) = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf(Split(varObjectToId(i), ",")(0))
    varSerial(i) = ""
    For Each SWbemObj In SWbemSet(i)
      varSerial(i) = SWbemObj.Properties_(Split(varObjectToId(i), ",")(1)) 'Property value
      varSerial(i) = Trim(varSerial(i))
      If Len(varSerial(i)) < 1 Then varSerial(i) = "Unknown value"
    Next
    Text1(i) = varSerial(i)
  Next
End Sub

Private Sub Command1_Click()
  Unload Me
End Sub
