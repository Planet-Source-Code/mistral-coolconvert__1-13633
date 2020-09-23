VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Cool Converter -- Main Screen"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   13260
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   13260
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Index           =   5
      Left            =   5160
      TabIndex        =   53
      Top             =   1920
      Width           =   4335
      Begin VB.Frame Frame11 
         Caption         =   "Convert"
         Height          =   700
         Left            =   840
         TabIndex        =   56
         Top             =   480
         Width           =   2895
         Begin VB.TextBox Text12 
            Height          =   315
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   1815
         End
         Begin VB.ComboBox cboAngle1 
            Height          =   315
            Left            =   2040
            TabIndex        =   57
            Text            =   "Deg"
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   1920
         Width           =   2535
      End
      Begin VB.ComboBox cboAngle2 
         Height          =   315
         Left            =   3240
         TabIndex        =   54
         Text            =   "Deg"
         Top             =   2000
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   60
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label11 
         Height          =   135
         Left            =   5400
         TabIndex        =   59
         Top             =   2880
         Width           =   255
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Index           =   4
      Left            =   11640
      TabIndex        =   49
      Top             =   3720
      Width           =   4335
      Begin VB.OptionButton Option2 
         Caption         =   "Oct"
         Height          =   255
         Index           =   3
         Left            =   2880
         TabIndex        =   9
         Top             =   2880
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Hex"
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   8
         Top             =   2640
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Dec"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   7
         Top             =   2400
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Bin"
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   6
         Top             =   2160
         Width           =   615
      End
      Begin VB.Frame Frame13 
         Caption         =   "Convert"
         Height          =   1335
         Left            =   1080
         TabIndex        =   50
         Top             =   120
         Width           =   2055
         Begin VB.OptionButton Option1 
            Caption         =   "Oct"
            Height          =   255
            Index           =   3
            Left            =   1320
            TabIndex        =   4
            Top             =   960
            Width           =   615
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Hex"
            Height          =   255
            Index           =   2
            Left            =   1320
            TabIndex        =   3
            Top             =   720
            Width           =   615
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Dec"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   2
            Top             =   480
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Bin"
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   1
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   120
            MaxLength       =   12
            TabIndex        =   0
            Top             =   540
            Width           =   1095
         End
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   52
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label9 
         Height          =   135
         Left            =   5400
         TabIndex        =   51
         Top             =   2880
         Width           =   255
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Index           =   3
      Left            =   9840
      TabIndex        =   45
      Top             =   720
      Width           =   4335
      Begin VB.ComboBox cboTemp2 
         Height          =   315
         Left            =   3000
         TabIndex        =   14
         Text            =   "°C"
         Top             =   2000
         Width           =   495
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Frame Frame10 
         Caption         =   "Convert"
         Height          =   700
         Left            =   1200
         TabIndex        =   46
         Top             =   480
         Width           =   1935
         Begin VB.ComboBox cboTemp1 
            Height          =   315
            Left            =   1320
            TabIndex        =   12
            Text            =   "°C"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox Text7 
            Height          =   315
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Label Label8 
         Height          =   135
         Left            =   5400
         TabIndex        =   48
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   47
         Top             =   1320
         Width           =   375
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Index           =   2
      Left            =   7560
      TabIndex        =   35
      Top             =   6720
      Width           =   4335
      Begin VB.ListBox lstWei2 
         Height          =   2205
         Left            =   2760
         TabIndex        =   44
         Top             =   1080
         Width           =   855
      End
      Begin VB.ListBox lstWei1 
         Height          =   2205
         Left            =   600
         TabIndex        =   43
         Top             =   1080
         Width           =   855
      End
      Begin VB.Frame Frame9 
         Caption         =   "To"
         Height          =   2655
         Left            =   2520
         TabIndex        =   40
         Top             =   840
         Width           =   1335
      End
      Begin VB.Frame Frame8 
         Caption         =   "From"
         Height          =   2655
         Left            =   360
         TabIndex        =   39
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   220
         Width           =   1815
      End
      Begin VB.Frame Frame7 
         Caption         =   "Convert"
         Height          =   615
         Left            =   360
         TabIndex        =   36
         Top             =   120
         Width           =   1335
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   120
            MaxLength       =   15
            TabIndex        =   37
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Label Label6 
         Height          =   135
         Left            =   5400
         TabIndex        =   42
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1850
         TabIndex        =   41
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Index           =   1
      Left            =   4560
      TabIndex        =   25
      Top             =   4920
      Width           =   4335
      Begin VB.ListBox lstCapa2 
         Height          =   2205
         Left            =   2760
         TabIndex        =   34
         Top             =   1080
         Width           =   855
      End
      Begin VB.ListBox lstCapa1 
         Height          =   2205
         Left            =   600
         TabIndex        =   33
         Top             =   1080
         Width           =   855
      End
      Begin VB.Frame Frame6 
         Caption         =   "Convert"
         Height          =   615
         Left            =   360
         TabIndex        =   29
         Top             =   120
         Width           =   1335
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   120
            MaxLength       =   15
            TabIndex        =   30
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   220
         Width           =   1815
      End
      Begin VB.Frame Frame5 
         Caption         =   "From"
         Height          =   2655
         Left            =   360
         TabIndex        =   27
         Top             =   840
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "To"
         Height          =   2655
         Left            =   2520
         TabIndex        =   26
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1850
         TabIndex        =   32
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label3 
         Height          =   135
         Left            =   5400
         TabIndex        =   31
         Top             =   2880
         Width           =   255
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   4920
      Width           =   4335
      Begin VB.ListBox lstDist2 
         Height          =   2205
         Left            =   2760
         TabIndex        =   22
         Top             =   1080
         Width           =   855
      End
      Begin VB.Frame Frame4 
         Caption         =   "To"
         Height          =   2655
         Left            =   2520
         TabIndex        =   21
         Top             =   840
         Width           =   1335
      End
      Begin VB.ListBox lstDist1 
         Height          =   2205
         Left            =   600
         TabIndex        =   20
         Top             =   1080
         Width           =   855
      End
      Begin VB.Frame Frame3 
         Caption         =   "From"
         Height          =   2655
         Left            =   360
         TabIndex        =   19
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   220
         Width           =   1815
      End
      Begin VB.Frame Frame2 
         Caption         =   "Convert"
         Height          =   615
         Left            =   360
         TabIndex        =   16
         Top             =   120
         Width           =   1335
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            MaxLength       =   15
            TabIndex        =   17
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Label Label2 
         Height          =   135
         Left            =   5400
         TabIndex        =   24
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1850
         TabIndex        =   23
         Top             =   120
         Width           =   375
      End
   End
   Begin MSComctlLib.TabStrip tabRtf 
      Height          =   4335
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7646
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      TabMinWidth     =   0
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Distance"
            Key             =   "tabDistance"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Capacity (Gal / l)"
            Key             =   "tabCap"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Weight"
            Key             =   "tabWeight"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Temperature"
            Key             =   "tabTemp"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Numeric (HEX)"
            Key             =   "tabNum"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Angles"
            Key             =   "tabAngles"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InterMeter As Double
Function BinToDec(bite) As Integer
    Y = Len(bite) - 1

    For X = 1 To Len(bite)
        BinToDec = BinToDec + ((2 ^ Y)) * (Mid$(bite, X, 1))
        Y = Y - 1
    Next X

End Function


Function DecToBin(dec As Double)
Dim Bin As String
Dim Count, nline As Integer

Bin$ = "00000000"

For Count = 1 To 8
    If dec Mod 2 ^ Count <> 0 Then Mid$(Bin$, Count, 1) = "1": dec = dec - (dec Mod 2 ^ Count)
Next

For nline = 8 To 1 Step -1
    DecToBin = DecToBin & Mid$(Bin$, nline, 1)
Next nline

End Function

Private Sub cboAngle1_Change()

If cboAngle1.Text = "Deg" Then
InterMeter = Val(Text12.Text) / 360
ElseIf cboAngle1.Text = "Rad" Then
InterMeter = Val(Text12.Text) / 6.28318530718
ElseIf cboAngle1.Text = "Grad" Then
InterMeter = Val(Text12.Text) / 400
End If

If cboAngle2.Text = "Deg" Then
Text11.Text = InterMeter * 360
ElseIf cboAngle2.Text = "Rad" Then
Text11.Text = InterMeter * 6.28318530718
ElseIf cboAngle2.Text = "Grad" Then
Text11.Text = InterMeter * 400
End If
End Sub

Private Sub cboAngle1_Click()
If cboAngle1.Text = "Deg" Then
InterMeter = Val(Text12.Text) / 360
ElseIf cboAngle1.Text = "Rad" Then
InterMeter = Val(Text12.Text) / 6.28318530718
ElseIf cboAngle1.Text = "Grad" Then
InterMeter = Val(Text12.Text) / 400
End If

If cboAngle2.Text = "Deg" Then
Text11.Text = InterMeter * 360
ElseIf cboAngle2.Text = "Rad" Then
Text11.Text = InterMeter * 6.28318530718
ElseIf cboAngle2.Text = "Grad" Then
Text11.Text = InterMeter * 400
End If
End Sub

Private Sub cboAngle2_Change()
If cboAngle1.Text = "Deg" Then
InterMeter = Val(Text12.Text) / 360
ElseIf cboAngle1.Text = "Rad" Then
InterMeter = Val(Text12.Text) / 6.28318530718
ElseIf cboAngle1.Text = "Grad" Then
InterMeter = Val(Text12.Text) / 400
End If

If cboAngle2.Text = "Deg" Then
Text11.Text = InterMeter * 360
ElseIf cboAngle2.Text = "Rad" Then
Text11.Text = InterMeter * 6.28318530718
ElseIf cboAngle2.Text = "Grad" Then
Text11.Text = InterMeter * 400
End If
End Sub

Private Sub cboAngle2_Click()
If cboAngle1.Text = "Deg" Then
InterMeter = Val(Text12.Text) / 360
ElseIf cboAngle1.Text = "Rad" Then
InterMeter = Val(Text12.Text) / 6.28318530718
ElseIf cboAngle1.Text = "Grad" Then
InterMeter = Val(Text12.Text) / 400
End If

If cboAngle2.Text = "Deg" Then
Text11.Text = InterMeter * 360
ElseIf cboAngle2.Text = "Rad" Then
Text11.Text = InterMeter * 6.28318530718
ElseIf cboAngle2.Text = "Grad" Then
Text11.Text = InterMeter * 400
End If
End Sub

Private Sub cboTemp1_Change()

If cboTemp1.Text = "°C" Then
InterMeter = 32 + 9 / 5 * Val(Text7.Text)
ElseIf cboTemp1.Text = "°F" Then
InterMeter = Val(Text7.Text)
End If

If cboTemp2.Text = "°C" Then
Text8.Text = 5 / 9 * (InterMeter - 32)
ElseIf cboTemp2.Text = "°F" Then
Text8.Text = InterMeter
End If

End Sub

Private Sub cboTemp1_Click()
If cboTemp1.Text = "°C" Then
InterMeter = 32 + 9 / 5 * Val(Text7.Text)
ElseIf cboTemp1.Text = "°F" Then
InterMeter = Val(Text7.Text)
End If

If cboTemp2.Text = "°C" Then
Text8.Text = 5 / 9 * (InterMeter - 32)
ElseIf cboTemp2.Text = "°F" Then
Text8.Text = InterMeter
End If
End Sub

Private Sub cboTemp2_Change()
If cboTemp1.Text = "°C" Then
InterMeter = 32 + 9 / 5 * Val(Text7.Text)
ElseIf cboTemp1.Text = "°F" Then
InterMeter = Val(Text7.Text)
End If

If cboTemp2.Text = "°C" Then
Text8.Text = 5 / 9 * (InterMeter - 32)
ElseIf cboTemp2.Text = "°F" Then
Text8.Text = InterMeter
End If
End Sub

Private Sub cboTemp2_Click()
If cboTemp1.Text = "°C" Then
InterMeter = 32 + 9 / 5 * Val(Text7.Text)
Else
InterMeter = Val(Text7.Text)
End If

If cboTemp2.Text = "°C" Then
Text8.Text = 5 / 9 * (InterMeter - 32)
Else
Text8.Text = InterMeter
End If
End Sub

Private Sub Form_Load()
'***************************************
'Définit la grandeur de la fenetre     *

Me.Width = 4815
Me.Height = 5025


'***************************************
'Déclarations des objets des listes    *

lstDist1.AddItem "Miles"
lstDist1.AddItem "Km"
lstDist1.AddItem "Hm"
lstDist1.AddItem "Dam"
lstDist1.AddItem "M"
lstDist1.AddItem "Yards"
lstDist1.AddItem "Feets"
lstDist1.AddItem "Dm"
lstDist1.AddItem "Inches"
lstDist1.AddItem "Cm"
lstDist1.AddItem "Mm"


lstDist2.AddItem "Miles"
lstDist2.AddItem "Km"
lstDist2.AddItem "Hm"
lstDist2.AddItem "Dam"
lstDist2.AddItem "M"
lstDist2.AddItem "Yards"
lstDist2.AddItem "Feets"
lstDist2.AddItem "Dm"
lstDist2.AddItem "Inches"
lstDist2.AddItem "Cm"
lstDist2.AddItem "Mm"
'------------------------

lstCapa1.AddItem "Gal"
lstCapa1.AddItem "Kl"
lstCapa1.AddItem "Hl"
lstCapa1.AddItem "Dal"
lstCapa1.AddItem "Liters"
lstCapa1.AddItem "Dl"
lstCapa1.AddItem "Cl"
lstCapa1.AddItem "Ml"

lstCapa2.AddItem "Gal"
lstCapa2.AddItem "Kl"
lstCapa2.AddItem "Hl"
lstCapa2.AddItem "Dal"
lstCapa2.AddItem "Liters"
lstCapa2.AddItem "Dl"
lstCapa2.AddItem "Cl"
lstCapa2.AddItem "Ml"


'------------------------

lstWei1.AddItem "Tons"
lstWei1.AddItem "Kg"
lstWei1.AddItem "Pounds"
lstWei1.AddItem "Hg"
lstWei1.AddItem "Ounce"
lstWei1.AddItem "Dag"
lstWei1.AddItem "Grams"
lstWei1.AddItem "Dg"
lstWei1.AddItem "Cg"
lstWei1.AddItem "Mg"

lstWei2.AddItem "Tons"
lstWei2.AddItem "Kg"
lstWei2.AddItem "Pounds"
lstWei2.AddItem "Hg"
lstWei2.AddItem "Ounce"
lstWei2.AddItem "Dag"
lstWei2.AddItem "Grams"
lstWei2.AddItem "Dg"

lstWei2.AddItem "Cg"
lstWei2.AddItem "Mg"
'------------------------

cboTemp1.AddItem "°C"
cboTemp1.AddItem "°F"

cboTemp2.AddItem "°C"
cboTemp2.AddItem "°F"

'------------------------

cboAngle1.AddItem "Deg"
cboAngle1.AddItem "Rad"
cboAngle1.AddItem "Grad"

cboAngle2.AddItem "Deg"
cboAngle2.AddItem "Rad"
cboAngle2.AddItem "Grad"

'***************************************
'Routine pour le placement des cadres  *

For i = 0 To fraTab.Count - 1
   With fraTab(i)
      .Move tabRtf.ClientLeft, _
      tabRtf.ClientTop, _
      tabRtf.ClientWidth, _
      tabRtf.ClientHeight
      End With
      Next i
   fraTab(0).ZOrder 0
End Sub

Private Sub lstCapa1_Click()
'Vérifie si l'autre liste a qqch de séléctionné

If lstCapa2.Text <> "" Then
'Opère avec l'opération de conversion
Select Case lstCapa1.Text
Case "Kl"
InterMeter = Val(Text4.Text) * 1000
Case "Hl"
InterMeter = Val(Text4.Text) * 100
Case "Dl"
InterMeter = Val(Text4.Text) * 10
Case "Liters"
InterMeter = Val(Text4.Text)
Case "Dl"
InterMeter = Val(Text4.Text) / 10
Case "Cl"
InterMeter = Val(Text4.Text) / 100
Case "Ml"
InterMeter = Val(Text4.Text) / 1000
Case "Gal"
InterMeter = Val(Text4.Text) / 3.785411784
End Select

Select Case lstCapa2.Text
Case "Kl"
Text3.Text = InterMeter / 1000
Case "Hl"
Text3.Text = InterMeter / 100
Case "Dl"
Text3.Text = InterMeter / 10
Case "Liters"
Text3.Text = InterMeter
Case "Dl"
Text3.Text = InterMeter * 10
Case "Cl"
Text3.Text = InterMeter * 100
Case "Ml"
Text3.Text = InterMeter * 1000
Case "Gal"
Text3.Text = InterMeter * 3.785411784
End Select
End If
End Sub

Private Sub lstCapa2_Click()

If lstCapa1.Text <> "" Then
'Opère avec l'opération de conversion
Select Case lstCapa1.Text
Case "Kl"
InterMeter = Val(Text4.Text) * 1000
Case "Hl"
InterMeter = Val(Text4.Text) * 100
Case "Dl"
InterMeter = Val(Text4.Text) * 10
Case "Liters"
InterMeter = Val(Text4.Text)
Case "Dl"
InterMeter = Val(Text4.Text) / 10
Case "Cl"
InterMeter = Val(Text4.Text) / 100
Case "Ml"
InterMeter = Val(Text4.Text) / 1000
Case "Gal"
InterMeter = Val(Text4.Text) / 0.264172052

End Select

Select Case lstCapa2.Text
Case "Kl"
Text3.Text = InterMeter / 1000
Case "Hl"
Text3.Text = InterMeter / 100
Case "Dl"
Text3.Text = InterMeter / 10
Case "Liters"
Text3.Text = InterMeter
Case "Dl"
Text3.Text = InterMeter * 10
Case "Cl"
Text3.Text = InterMeter * 100
Case "Ml"
Text3.Text = InterMeter * 1000
Case "Gal"
Text3.Text = InterMeter * 0.264172052

End Select
End If
End Sub

Private Sub lstDist1_Click()
'Vérifie si l'autre liste a qqch de séléctionné

If lstDist2.Text <> "" Then
'Opère avec l'opération de conversion
Select Case lstDist1.Text
Case "Km"
InterMeter = Val(Text1.Text) * 1000
Case "Hm"
InterMeter = Val(Text1.Text) * 100
Case "Dam"
InterMeter = Val(Text1.Text) * 10
Case "M"
InterMeter = Val(Text1.Text)
Case "Dm"
InterMeter = Val(Text1.Text) / 10
Case "Cm"
InterMeter = Val(Text1.Text) / 100
Case "Mm"
InterMeter = Val(Text1.Text) / 1000
Case "Miles"
InterMeter = Val(Text1.Text) * 1610
Case "Yards"
InterMeter = Val(Text1.Text) / 1.09361329833771
Case "Inches"
InterMeter = Val(Text1.Text) * 0.025399
Case "Feets"
InterMeter = Val(Text1.Text) * 0.3048
End Select
Select Case lstDist2.Text
Case "Km"
Text2.Text = InterMeter / 1000
Case "Hm"
Text2.Text = InterMeter / 100
Case "Dam"
Text2.Text = InterMeter / 10
Case "M"
Text2.Text = InterMeter
Case "Dm"
Text2.Text = InterMeter * 10
Case "Cm"
Text2.Text = InterMeter * 100
Case "Mm"
Text2.Text = InterMeter * 1000
Case "Miles"
Text2.Text = InterMeter / 1610
Case "Yards"
Text2.Text = InterMeter * 1.09361329833771
Case "Inches"
Text2.Text = InterMeter / 0.025399
Case "Feets"
Text2.Text = InterMeter / 0.3048
End Select
End If
End Sub

Private Sub lstDist2_Click()
'Vérifie si l'autre liste a qqch de séléctionné

If lstDist1.Text <> "" Then
'Opère avec l'opération de conversion
Select Case lstDist1.Text
Case "Km"
InterMeter = Val(Text1.Text) * 1000
Case "Hm"
InterMeter = Val(Text1.Text) * 100
Case "Dam"
InterMeter = Val(Text1.Text) * 10
Case "M"
InterMeter = Val(Text1.Text)
Case "Dm"
InterMeter = Val(Text1.Text) / 10
Case "Cm"
InterMeter = Val(Text1.Text) / 100
Case "Mm"
InterMeter = Val(Text1.Text) / 1000
Case "Miles"
InterMeter = Val(Text1.Text) * 1610
Case "Yards"
InterMeter = Val(Text1.Text) / 1.09361329833771
Case "Inches"
InterMeter = Val(Text1.Text) * 0.025399
Case "Feets"
InterMeter = Val(Text1.Text) * 0.3048
End Select
Select Case lstDist2.Text
Case "Km"
Text2.Text = InterMeter / 1000
Case "Hm"
Text2.Text = InterMeter / 100
Case "Dam"
Text2.Text = InterMeter / 10
Case "M"
Text2.Text = InterMeter
Case "Dm"
Text2.Text = InterMeter * 10
Case "Cm"
Text2.Text = InterMeter * 100
Case "Mm"
Text2.Text = InterMeter * 1000
Case "Miles"
Text2.Text = InterMeter / 1610
Case "Yards"
Text2.Text = InterMeter * 1.09361329833771
Case "Inches"
Text2.Text = InterMeter / 0.025399
Case "Feets"
Text2.Text = InterMeter / 0.3048
End Select
End If
End Sub



Private Sub lstWei1_Click()
If lstWei2.Text <> "" Then
'Opère avec l'opération de conversion
Select Case lstWei1.Text
Case "Tons"
InterMeter = Val(Text5.Text) * 1000000
Case "Kg"
InterMeter = Val(Text5.Text) * 1000
Case "Pounds"
InterMeter = Val(Text5.Text) * 453.59237
Case "Hg"
InterMeter = Val(Text5.Text) * 100
Case "Ounce"
InterMeter = Val(Text5.Text) * 28.3495
Case "Dag"
InterMeter = Val(Text5.Text) * 10
Case "Grams"
InterMeter = Val(Text5.Text)
Case "Dg"
InterMeter = Val(Text5.Text) / 10
Case "Cg"
InterMeter = Val(Text5.Text) / 100
Case "Mg"
InterMeter = Val(Text5.Text) / 1000
End Select

Select Case lstWei2.Text
Case "Tons"
Text6.Text = InterMeter / 1000000
Case "Kg"
Text6.Text = InterMeter / 1000
Case "Pounds"
Text6.Text = InterMeter / 453.59237
Case "Hg"
Text6.Text = InterMeter / 100
Case "Ounce"
Text6.Text = InterMeter / 28.3495
Case "Dag"
Text6.Text = InterMeter / 10
Case "Grams"
Text6.Text = InterMeter
Case "Dg"
Text6.Text = InterMeter * 10
Case "Cg"
Text6.Text = InterMeter * 100
Case "Mg"
Text6.Text = InterMeter * 1000
End Select
End If
End Sub

Private Sub lstWei2_Click()

If lstWei1.Text <> "" Then
Select Case lstWei1.Text
Case "Tons"
InterMeter = Val(Text5.Text) * 1000000
Case "Kg"
InterMeter = Val(Text5.Text) * 1000
Case "Pounds"
InterMeter = Val(Text5.Text) * 453.59237
Case "Hg"
InterMeter = Val(Text5.Text) * 100
Case "Ounce"
InterMeter = Val(Text5.Text) * 28.3495
Case "Dag"
InterMeter = Val(Text5.Text) * 10
Case "Grams"
InterMeter = Val(Text5.Text)
Case "Dg"
InterMeter = Val(Text5.Text) / 10
Case "Cg"
InterMeter = Val(Text5.Text) / 100
Case "Mg"
InterMeter = Val(Text5.Text) / 1000
End Select

Select Case lstWei2.Text
Case "Tons"
Text6.Text = InterMeter / 1000000
Case "Kg"
Text6.Text = InterMeter / 1000
Case "Pounds"
Text6.Text = InterMeter / 453.59237
Case "Hg"
Text6.Text = InterMeter / 100
Case "Ounce"
Text6.Text = InterMeter / 28.3495
Case "Dag"
Text6.Text = InterMeter / 10
Case "Grams"
Text6.Text = InterMeter
Case "Dg"
Text6.Text = InterMeter * 10
Case "Cg"
Text6.Text = InterMeter * 100
Case "Mg"
Text6.Text = InterMeter * 1000
End Select
End If
End Sub

Private Sub Option1_Click(Index As Integer)
Dim Strtem As String
Dim Ind2 As Integer
Select Case Index
Case 0
InterMeter = BinToDec(Text10.Text)
Case 1
InterMeter = Val(Text10.Text)
Case 2
Strtem = "&H" + Text10.Text
InterMeter = Val(Strtem)
Case 3
Strtem = "&o" + Text10.Text
InterMeter = Val(Strtem)
End Select

If Option2(0).Value = True Then
Ind2 = 0
ElseIf Option2(1).Value = True Then
Ind2 = 1
ElseIf Option2(2).Value = True Then
Ind2 = 2
ElseIf Option2(3).Value = True Then
Ind2 = 3
End If

Select Case Ind2
Case 0
Text9.Text = DecToBin(InterMeter)
Case 1
Text9.Text = InterMeter
Case 2
Text9.Text = Hex(InterMeter)
Case 3
Text9.Text = Oct(InterMeter)
End Select

End Sub

Private Sub Option2_Click(Index As Integer)
Dim Ind1 As Integer
If Option1(0).Value = True Then
Ind1 = 0
ElseIf Option1(1).Value = True Then
Ind1 = 1
ElseIf Option1(2).Value = True Then
Ind1 = 2
ElseIf Option1(3).Value = True Then
Ind1 = 3
End If

Dim Strtem As String
Select Case Ind1
Case 0
InterMeter = BinToDec(Text10.Text)
Case 1
InterMeter = Val(Text10.Text)
Case 2
Strtem = "&H" + Text10.Text
InterMeter = Strtem
Case 3
Strtem = "&o" + Text10.Text
InterMeter = Val(Strtem)
End Select


Select Case Index
Case 0
Text9.Text = DecToBin(InterMeter)
Case 1
Text9.Text = InterMeter
Case 2
Text9.Text = Hex(InterMeter)
Case 3
Text9.Text = Oct(InterMeter)
End Select
End Sub

Private Sub tabRtf_Click()
fraTab(tabRtf.SelectedItem.Index - 1).ZOrder 0
Select Case (tabRtf.SelectedItem.Index - 1)
Case 0
Text1.SetFocus
Case 1
Text4.SetFocus
Case 2
Text5.SetFocus
Case 3
Text7.SetFocus
Case 4
Text10.SetFocus
Case 5
Case 6
End Select
End Sub

Private Sub Text1_Change()
If lstDist2.Text <> "" And lstDist1.Text <> "" Then
'Opère avec l'opération de conversion
Select Case lstDist1.Text
Case "Km"
InterMeter = Val(Text1.Text) * 1000
Case "Hm"
InterMeter = Val(Text1.Text) * 100
Case "Dam"
InterMeter = Val(Text1.Text) * 10
Case "M"
InterMeter = Val(Text1.Text)
Case "Dm"
InterMeter = Val(Text1.Text) / 10
Case "Cm"
InterMeter = Val(Text1.Text) / 100
Case "Mm"
InterMeter = Val(Text1.Text) / 1000
Case "Miles"
InterMeter = Val(Text1.Text) * 1610
Case "Yards"
InterMeter = Val(Text1.Text) / 1.09361329833771
Case "Inches"
InterMeter = Val(Text1.Text) * 0.025399
Case "Feets"
InterMeter = Val(Text1.Text) * 0.3048
End Select
Select Case lstDist2.Text
Case "Km"
Text2.Text = InterMeter / 1000
Case "Hm"
Text2.Text = InterMeter / 100
Case "Dam"
Text2.Text = InterMeter / 10
Case "M"
Text2.Text = InterMeter
Case "Dm"
Text2.Text = InterMeter * 10
Case "Cm"
Text2.Text = InterMeter * 100
Case "Mm"
Text2.Text = InterMeter * 1000
Case "Miles"
Text2.Text = InterMeter / 1610
Case "Yards"
Text2.Text = InterMeter * 1.09361329833771
Case "Inches"
Text2.Text = InterMeter / 0.025399
Case "Feets"
Text2.Text = InterMeter / 0.3048
End Select
End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case vbKey0
Case vbKey1
Case vbKey2
Case vbKey3
Case vbKey4
Case vbKey5
Case vbKey6
Case vbKey7
Case vbKey8
Case vbKey9
Case vbKeyHome
Case vbKeyEnd
Case vbKeyInsert
Case vbKeyBack
Case vbKeyDelete
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub Text10_Change()
Dim Strtem As String
Dim Ind2 As Integer
Select Case Index
Case 0
InterMeter = BinToDec(Text10.Text)
Case 1
InterMeter = Val(Text10.Text)
Case 2
Strtem = "&H" + Text10.Text
InterMeter = Strtem
Case 3
Strtem = "&o" + Text10.Text
InterMeter = Val(Strtem)
End Select

If Option2(0).Value = True Then
Ind2 = 0
ElseIf Option2(1).Value = True Then
Ind2 = 1
ElseIf Option2(2).Value = True Then
Ind2 = 2
ElseIf Option2(3).Value = True Then
Ind2 = 3
End If

Select Case Ind2
Case 0
Text9.Text = DecToBin(InterMeter)
Case 1
Text9.Text = InterMeter
Case 2
Text9.Text = Hex(InterMeter)
Case 3
Text9.Text = Oct(InterMeter)
End Select
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If Option1(2).Value = False Then
Select Case KeyAscii
Case vbKey0
Case vbKey1
Case vbKey2
Case vbKey3
Case vbKey4
Case vbKey5
Case vbKey6
Case vbKey7
Case vbKey8
Case vbKey9
Case vbKeyHome
Case vbKeyEnd
Case vbKeyInsert
Case vbKeyBack
Case vbKeyDelete
Case Else
KeyAscii = 0
End Select
Else
Select Case KeyAscii
Case vbKey0
Case vbKey1
Case vbKey2
Case vbKey3
Case vbKey4
Case vbKey5
Case vbKey6
Case vbKey7
Case vbKey8
Case vbKey9
Case vbKeyA
Case vbKeyB
Case vbKeyC
Case vbKeyD
Case vbKeyE
Case vbKeyF
Case vbKeyHome
Case vbKeyEnd
Case vbKeyInsert
Case vbKeyBack
Case vbKeyDelete
Case Else
KeyAscii = 0
End Select
End If
End Sub
Private Sub Text12_Change()
If cboAngle1.Text = "Deg" Then
InterMeter = Val(Text12.Text) / 360
ElseIf cboAngle1.Text = "Rad" Then
InterMeter = Val(Text12.Text) / 6.28318530718
ElseIf cboAngle1.Text = "Grad" Then
InterMeter = Val(Text12.Text) / 400
End If

If cboAngle2.Text = "Deg" Then
Text11.Text = InterMeter * 360
ElseIf cboAngle2.Text = "Rad" Then
Text11.Text = InterMeter * 6.28318530718
ElseIf cboAngle2.Text = "Grad" Then
Text11.Text = InterMeter * 400
End If
End Sub

Private Sub Text4_Change()
If lstCapa1.Text <> "" And lstCapa2.Text <> "" Then

'Opère avec l'opération de conversion
Select Case lstCapa1.Text
Case "Kl"
InterMeter = Val(Text4.Text) * 1000
Case "Hl"
InterMeter = Val(Text4.Text) * 100
Case "Dl"
InterMeter = Val(Text4.Text) * 10
Case "Liters"
InterMeter = Val(Text4.Text)
Case "Dl"
InterMeter = Val(Text4.Text) / 10
Case "Cl"
InterMeter = Val(Text4.Text) / 100
Case "Ml"
InterMeter = Val(Text4.Text) / 1000
Case "Gal"
InterMeter = Val(Text4.Text) / 0.264172052

End Select

Select Case lstCapa2.Text
Case "Kl"
Text3.Text = InterMeter / 1000
Case "Hl"
Text3.Text = InterMeter / 100
Case "Dl"
Text3.Text = InterMeter / 10
Case "Liters"
Text3.Text = InterMeter
Case "Dl"
Text3.Text = InterMeter * 10
Case "Cl"
Text3.Text = InterMeter * 100
Case "Ml"
Text3.Text = InterMeter * 1000
Case "Gal"
Text3.Text = InterMeter * 0.264172052

End Select
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case vbKey0
Case vbKey1
Case vbKey2
Case vbKey3
Case vbKey4
Case vbKey5
Case vbKey6
Case vbKey7
Case vbKey8
Case vbKey9
Case vbKeyHome
Case vbKeyEnd
Case vbKeyInsert
Case vbKeyBack
Case vbKeyDelete
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub Text5_Change()
If lstWei1.Text <> "" And lstWei2.Text <> "" Then
Select Case lstWei1.Text
Case "Tons"
InterMeter = Val(Text5.Text) * 1000000
Case "Kg"
InterMeter = Val(Text5.Text) * 1000
Case "Pounds"
InterMeter = Val(Text5.Text) * 453.59237
Case "Hg"
InterMeter = Val(Text5.Text) * 100
Case "Ounce"
InterMeter = Val(Text5.Text) * 28.3495
Case "Dag"
InterMeter = Val(Text5.Text) * 10
Case "Grams"
InterMeter = Val(Text5.Text)
Case "Dg"
InterMeter = Val(Text5.Text) / 10
Case "Cg"
InterMeter = Val(Text5.Text) / 100
Case "Mg"
InterMeter = Val(Text5.Text) / 1000
End Select

Select Case lstWei2.Text
Case "Tons"
Text6.Text = InterMeter / 1000000
Case "Kg"
Text6.Text = InterMeter / 1000
Case "Pounds"
Text6.Text = InterMeter / 453.59237
Case "Hg"
Text6.Text = InterMeter / 100
Case "Ounce"
Text6.Text = InterMeter / 28.3495
Case "Dag"
Text6.Text = InterMeter / 10
Case "Grams"
Text6.Text = InterMeter
Case "Dg"
Text6.Text = InterMeter * 10
Case "Cg"
Text6.Text = InterMeter * 100
Case "Mg"
Text6.Text = InterMeter * 1000
End Select
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case vbKey0
Case vbKey1
Case vbKey2
Case vbKey3
Case vbKey4
Case vbKey5
Case vbKey6
Case vbKey7
Case vbKey8
Case vbKey9
Case vbKeyHome
Case vbKeyEnd
Case vbKeyInsert
Case vbKeyBack
Case vbKeyDelete
Case Else
KeyAscii = 0
End Select
End Sub

Private Sub Text6_Change()
If lstWei1.Text <> "" And lstWei2.Text <> "" Then
Select Case lstWei1.Text
Case "Tons"
InterMeter = Val(Text5.Text) * 1000000
Case "Kg"
InterMeter = Val(Text5.Text) * 1000
Case "Pounds"
InterMeter = Val(Text5.Text) * 453.59237
Case "Hg"
InterMeter = Val(Text5.Text) * 100
Case "Ounce"
InterMeter = Val(Text5.Text) * 28.3495
Case "Dag"
InterMeter = Val(Text5.Text) * 10
Case "Grams"
InterMeter = Val(Text5.Text)
Case "Dg"
InterMeter = Val(Text5.Text) / 10
Case "Cg"
InterMeter = Val(Text5.Text) / 100
Case "Mg"
InterMeter = Val(Text5.Text) / 1000
End Select

Select Case lstWei2.Text
Case "Tons"
Text6.Text = InterMeter / 1000000
Case "Kg"
Text6.Text = InterMeter / 1000
Case "Pounds"
Text6.Text = InterMeter / 453.59237
Case "Hg"
Text6.Text = InterMeter / 100
Case "Ounce"
Text6.Text = InterMeter / 28.3495
Case "Dag"
Text6.Text = InterMeter / 10
Case "Grams"
Text6.Text = InterMeter
Case "Dg"
Text6.Text = InterMeter * 10
Case "Cg"
Text6.Text = InterMeter * 100
Case "Mg"
Text6.Text = InterMeter * 1000
End Select
End If
End Sub

Private Sub Text7_Change()
If cboTemp1.Text = "°C" Then
InterMeter = Val(Text7.Text)
ElseIf cboTemp1.Text = "°F" Then
InterMeter = Val(Text7.Text) * (-17.22222)
End If

If cboTemp2.Text = "°C" Then
Text8.Text = InterMeter
ElseIf cboTemp2.Text = "°F" Then
Text8.Text = InterMeter / (-17.22222)
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case vbKey0
Case vbKey1
Case vbKey2
Case vbKey3
Case vbKey4
Case vbKey5
Case vbKey6
Case vbKey7
Case vbKey8
Case vbKey9
Case vbKeyHome
Case vbKeyEnd
Case vbKeyInsert
Case vbKeyBack
Case vbKeyDelete
Case Else
KeyAscii = 0
End Select
End Sub
