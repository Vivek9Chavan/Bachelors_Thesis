VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000016&
   Caption         =   "Form1"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   20160
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20160
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "OUTPUTS"
      Height          =   7935
      Left            =   8280
      TabIndex        =   59
      Top             =   960
      Width           =   7335
      Begin VB.Frame Frame5 
         Caption         =   "PRESSURE DROP"
         Height          =   3135
         Left            =   360
         TabIndex        =   62
         Top             =   4320
         Width           =   6735
      End
      Begin VB.Frame Frame4 
         Caption         =   "TEMP PROFILE"
         Height          =   3375
         Left            =   120
         TabIndex        =   61
         Top             =   720
         Width           =   6855
         Begin VB.TextBox Text30 
            Height          =   285
            Left            =   5040
            TabIndex        =   78
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox Text29 
            Height          =   375
            Left            =   4920
            TabIndex        =   77
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox Text28 
            Height          =   375
            Left            =   4920
            TabIndex        =   76
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox Text27 
            Height          =   285
            Left            =   1320
            TabIndex        =   72
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox Text26 
            Height          =   285
            Left            =   1200
            TabIndex        =   71
            Top             =   1920
            Width           =   855
         End
         Begin VB.TextBox Text25 
            Height          =   285
            Left            =   1200
            TabIndex        =   70
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox Text24 
            Height          =   285
            Left            =   1200
            TabIndex        =   69
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label27 
            Caption         =   "TF3"
            Height          =   255
            Left            =   3240
            TabIndex        =   75
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label Label26 
            Caption         =   "TF2"
            Height          =   375
            Left            =   3240
            TabIndex        =   74
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label25 
            Caption         =   "TF1"
            Height          =   255
            Left            =   3240
            TabIndex        =   73
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label24 
            Caption         =   "TG4"
            Height          =   255
            Left            =   240
            TabIndex        =   68
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label Label23 
            Caption         =   "TG3"
            Height          =   375
            Left            =   240
            TabIndex        =   67
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label Label22 
            Caption         =   "TG2"
            Height          =   375
            Left            =   240
            TabIndex        =   66
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "TG1"
            Height          =   375
            Left            =   240
            TabIndex        =   65
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label20 
            Caption         =   "THERMIC FLUID"
            Height          =   375
            Left            =   3840
            TabIndex        =   64
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label19 
            Caption         =   "FLUE GAS"
            Height          =   375
            Left            =   360
            TabIndex        =   63
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Label Label18 
         Caption         =   "Label18"
         Height          =   375
         Left            =   4560
         TabIndex        =   60
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000016&
      Caption         =   "COIL DETAILS"
      Height          =   8895
      Left            =   4440
      TabIndex        =   24
      Top             =   120
      Width           =   3615
      Begin VB.TextBox Text23 
         Height          =   285
         Left            =   2040
         TabIndex        =   58
         Top             =   8400
         Width           =   1095
      End
      Begin VB.TextBox Text22 
         Height          =   285
         Left            =   2040
         TabIndex        =   57
         Top             =   7920
         Width           =   1095
      End
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   2040
         TabIndex        =   53
         Top             =   6960
         Width           =   1095
      End
      Begin VB.TextBox Text20 
         Height          =   375
         Left            =   2040
         TabIndex        =   52
         Top             =   6360
         Width           =   1095
      End
      Begin VB.TextBox Text19 
         Height          =   375
         Left            =   2040
         TabIndex        =   51
         Top             =   5760
         Width           =   1335
      End
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   2040
         TabIndex        =   36
         Top             =   5280
         Width           =   1215
      End
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   2040
         TabIndex        =   35
         Top             =   4680
         Width           =   1215
      End
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   1920
         TabIndex        =   34
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   1920
         TabIndex        =   33
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   1920
         TabIndex        =   32
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   1920
         TabIndex        =   31
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   1920
         TabIndex        =   30
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "THK"
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   8400
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "DIA"
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   7920
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "OUTER JACKET"
         Height          =   255
         Left            =   1080
         TabIndex        =   54
         Top             =   7440
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "NO OF TURNS"
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   6960
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "NO. OF STARTS"
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   6360
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "TUBE THK"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   5760
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "TUBE DIA"
         Height          =   255
         Left            =   360
         TabIndex        =   47
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "PCD"
         Height          =   255
         Left            =   360
         TabIndex        =   46
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "INNER COILS"
         Height          =   255
         Left            =   1200
         TabIndex        =   40
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "OUTER COILS"
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   37
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "NO OF TURNS"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   29
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "NO. OF STARTS"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   28
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "TUBE DIA"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   27
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "TUBE THK"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   26
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "PCD"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   25
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Caption         =   "FUEL DETAILS"
      Height          =   8895
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   4095
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   1680
         TabIndex        =   23
         Text            =   "20"
         Top             =   6960
         Width           =   1215
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   1680
         TabIndex        =   22
         Text            =   "300"
         Top             =   6360
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1680
         TabIndex        =   21
         Text            =   "200"
         Top             =   5760
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   1680
         TabIndex        =   20
         Text            =   "9650"
         Top             =   5040
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1680
         TabIndex        =   19
         Text            =   "0"
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1680
         TabIndex        =   18
         Text            =   "1.5"
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1680
         TabIndex        =   17
         Text            =   "3.5"
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1680
         TabIndex        =   16
         Text            =   "0"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Text            =   "0"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Text            =   "11"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Text            =   "84"
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "H2O"
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   12
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "ASH"
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   11
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "CAL. VAL."
         Height          =   375
         Index           =   7
         Left            =   0
         TabIndex        =   10
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "FIRING RATE"
         Height          =   375
         Index           =   6
         Left            =   0
         TabIndex        =   9
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "ATMOSHERIC TEMP"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   6360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "EXCESS AIR"
         Height          =   375
         Index           =   4
         Left            =   0
         TabIndex        =   7
         Top             =   6840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "0"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "N"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "S"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "H"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "C"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000016&
      Caption         =   "CALCULATE"
      Height          =   495
      Left            =   8280
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label7 
      Caption         =   "NO OF TURNS"
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   45
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "NO. OF STARTS"
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   44
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "TUBE DIA"
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   43
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TUBE THK"
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   42
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "PCD"
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   41
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "OUTER COILS"
      Height          =   375
      Index           =   2
      Left            =   5520
      TabIndex        =   39
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "OUTER COILS"
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   38
      Top             =   3840
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim m(1 To 5, 1 To 5) As Variant
Dim cp(1 To 5, 1 To 5) As Variant

Dim id(1 To 5, 1 To 5) As Variant
Dim va(1 To 5, 1) As Variant
Dim sol(1 To 5, 1) As Variant


'RADIATION :

mc = val(Text1.Text) * 20 / 12 'input
mh = val(Text2.Text) * 10 / 2 'input
mo = val(Text3.Text) * 10 / 16 'input
mn = val(Text4.Text) * 0 'input
mso2 = val(Text5.Text) * 10 * 2 / 32 'input
mw = val(Text6.Text) * 10 / 18 'input
d = 1 'input
h = 4 'input

mt = mc + mh - mo + mso2
mo2 = mt / 2
mn2 = mo2 * 79 / 21

EA = val(Text11.Text)

nco2 = mc / 2
nh2o = mh + mw
no2 = EA * mo2 * 0.01
nn2 = mn2 * (1 + EA / 100)
nso2 = mso2 / 2

ntot = nco2 + nh2o + no2 + nn2 + nso2
xc = nco2 / ntot
xw = nh2o / ntot
xo2 = no2 / ntot
xso2 = nso2 / ntot
xn2 = nn2 / ntot
gf = (nco2 * 44 + nh2o * 18 + nso2 * 64 + no2 * 32 + nn2 * 28) / 1000
T1 = 1500 'Assumption



dT = 20
iter:
If dT > 0.1 Then

rar = (22 / 7) / 4 * d ^ 2
har = (22 / 7) * d * h
vol = (22 / 7) / 4 * d ^ 2 * h
lb = 3.6 * vol / (rar + har)
b = 1
k = ((0.8 + 1.6 * xw) * (1 - (0.38 * (T1) / 1000)) * (xc + xw)) / ((xc + xw) * lb) ^ 0.5
Eg = b * (1 - Exp(-k * lb))


x = T1 - 273

'Specific heat

sh2o = (-1) * 10 ^ (-4) * x ^ 2 + 0.7691 * x + 1788.7
so2 = (-7) * 10 ^ (-5) * x ^ 2 + 0.2838 * x + 912.86
sn2 = (-6) * 10 ^ (-5) * x ^ 2 + 0.2571 * x + 1005.2
sco2 = (-0.0002) * x ^ 2 + 0.5929 * x + 882.71
sso2 = 0.0000001 * x ^ 3 - 0# * x ^ 2 + 0.608 * x + 608
smix = (sh2o * xw * 18 + so2 * 32 * xo2 + sn2 * 28 * xn2 + sso2 * xso2 * 64 + sco2 * 44 * xco2) / (xw * 18 + xso2 * 64 + 32 * xo2 + 28 * xn2 + 44 * xco2)

'Thermal conductivity

kn2 = 0.00000003 * x ^ 2 + 0.00002 * x + 0.032
kco2 = -0.00000002 * x ^ 2 + 0.00009 * x + 0.012
ko2 = -0.000000006 * x ^ 2 + 0.00007 * x + 0.026
kh2o = 0.00000002 * x ^ 2 + 0# * x + 0.011
kso2 = -0.00000001 * x ^ 2 + 0.00006 * x + 0.007

kmix = (kh2o * xw * (18) ^ (1 / 3) + ko2 * 32 * (xo2) ^ (1 / 3) + kn2 * 28 * (xn2) ^ (1 / 3) + kso2 * 64 * (xso2) ^ (1 / 3) + kco2 * 44 * (xco2) ^ (1 / 3)) / ((xw * (18) ^ (1 / 3) + 32 * (xo2) ^ (1 / 3) + 64 * (xso2) ^ (1 / 3) + 28 * (xn2) ^ (1 / 3) + 44 * (xco2) ^ (1 / 3)))

'Dynamic Viscosity
muco2 = -0.000000000005 * x ^ 2 + 0.00000004 * x + 0.00002
mun2 = 0.0000000000002 * x ^ 3 - 0.0000000003 * x ^ 2 + 0.0000001 * x + 0.000008
muo2 = -0.000000000005 * x ^ 2 + 0.00000004 * x + 0.00002
muh2o = -0.000000000004 * x ^ 2 + 0.00000004 * x + 0.000008
muso2 = -0.000000000008 * x ^ 2 + 0.00000004 * x + 0.00001
mumix = (muh2o * xw * (18) ^ 0.5 + muo2 * 32 * (xo2) ^ 0.5 + muso2 * 64 * (xso2) ^ 0.5 + mun2 * 28 * (xn2) ^ 0.5 + muco2 * 44 * (xco2) ^ 0.5) / (xw * (18) ^ 0.5 + 64 * (xso2) ^ 0.5 + 32 * (xo2) ^ 0.5 + 28 * (xn2) ^ 0.5 + 44 * (xco2) ^ 0.5)

'Knematic Viscosity

'vco2 = 0.00000000004 * x ^ 2 + 0.00000007 * x + 0.000005
'vn2 = 0.00000000008 * x ^ 2 + 0.00000004 * x + 0.00002
'vo2 = 0.00000000006 * x ^ 2 + 0.0000001 * x + 0.00001
'vh2o = 0.0000000001 * x ^ 2 + 0.0000001 * x + 0.000001


'Density

'dco2 = -0.41 * Log(x) + 3.304
'dn2 = -0.26 * Log(x) + 2.103
'do2 = -0.3 * Log(x) + 2.402
'dh2o = -0.16 * Log(x) + 1.352





Tw = 600 'assumption
Em = 0.9 'input

'mf1 = 200 'input
MF1 = val(Text9.Text)
cv1 = val(Text8.Text)
cv = cv1 * 1000 * 4.186
mf = MF1 / 3600



Ar = 22 / 7 * h * d

ht = Ar * (5.67 * 10 ^ (-8) * (T1 ^ 4 - Tw ^ 4)) / (1 / 0.9 + 1 / Eg - 1)

ts = val(Text10.Text)

T2 = (mf * cv - ht) / (mf * gf * smix) + ts



dT = Abs(T2 - T1)

T1 = T1 + 0.1 * (T2 - T1)

GoTo iter
End If



''Print
''Print "Results from Radiation Calculations:"
''Print
''Print Eg
''Print
''Print T2
''Print
''Print ht
''Print

'Convection

mg = 200 * gf / 3600 'correlation


u1 = 35 'Assumption
u2 = 40 'Assumption
u3 = 45 'Assumption

'Assumed :D = 1;1.1;1.8;1.9 m
A1 = 13.823
A2 = 22.619
A3 = 23.876


mfl = 19.44 'input

cpfl = 2000 ' from equation

qgen = mf * cv

'dci
'dco
'dti
'dto

si = 0.0000000567
rad = Ar * si / ((1 / Em) + (1 / Eg) - 1)

'In convection
'hi
'ho
'km



'Assumptions:

Tg1 = 300
Tg2 = 1200
Tg3 = 900
Tg4 = 600
Tf1 = 520
Tf2 = 530
Tf3 = 540

iter1:

f1 = rad * (Tg2 ^ 4 - Tw ^ 4) - qgen + mg * smix * (Tg2 - Tg1)
f2 = u1 * A1 * ((Tg2 - Tf2) - (Tg3 - Tf3)) / Log((Tg2 - Tf2) / (Tg3 - Tf3)) + u2 * A2 * ((Tg2 - Tf2) - (Tg3 - Tf1)) / Log((Tg2 - Tf2) / (Tg3 - Tf1)) - mg * smix * (Tg2 - Tg3)
f3 = u3 * A3 * ((Tg3 - Tf1) - (Tg4 - Tf2)) / Log((Tg3 - Tf1) / (Tg4 - Tf2)) - mg * smix * (Tg3 - Tg4)
f4 = rad * (Tg2 ^ 4 - Tw ^ 4) + u1 * A1 * ((Tg2 - Tf2) - (Tg3 - Tf3)) / Log((Tg2 - Tf2) / (Tg3 - Tf3)) - mfl * cpfl * (Tf3 - Tf2)
f5 = u3 * A3 * ((Tg3 - Tf1) - (Tg4 - Tf2)) / Log((Tg3 - Tf1) / (Tg4 - Tf2)) + u2 * A2 * ((Tg2 - Tf2) - (Tg3 - Tf1)) / Log((Tg2 - Tf2) / (Tg3 - Tf1)) - mfl * cpfl * (Tf2 - Tf1)

'dtg2

tg21 = Tg2 * 1.000001

n2f1 = rad * (tg21 ^ 4 - Tw ^ 4) - qgen + mg * smix * (tg21 - Tg1)
n2f2 = u1 * A1 * ((tg21 - Tf2) - (Tg3 - Tf3)) / Log((tg21 - Tf2) / (Tg3 - Tf3)) + u2 * A2 * ((tg21 - Tf2) - (Tg3 - Tf1)) / Log((tg21 - Tf2) / (Tg3 - Tf1)) - mg * smix * (tg21 - Tg3)
n2f3 = u3 * A3 * ((Tg3 - Tf1) - (Tg4 - Tf2)) / Log((Tg3 - Tf1) / (Tg4 - Tf2)) - mg * smix * (Tg3 - Tg4)
n2f4 = rad * (tg21 ^ 4 - Tw ^ 4) + u1 * A1 * ((tg21 - Tf2) - (Tg3 - Tf3)) / Log((tg21 - Tf2) / (Tg3 - Tf3)) - mfl * cpfl * (Tf3 - Tf2)
n2f5 = u3 * A3 * ((Tg3 - Tf1) - (Tg4 - Tf2)) / Log((Tg3 - Tf1) / (Tg4 - Tf2)) + u2 * A2 * ((tg21 - Tf2) - (Tg3 - Tf1)) / Log((tg21 - Tf2) / (Tg3 - Tf1)) - mfl * cpfl * (Tf2 - Tf1)

m(1, 1) = (n2f1 - f1) / (0.000001 * Tg2)
m(2, 1) = (n2f2 - f2) / (0.000001 * Tg2)
m(3, 1) = (n2f3 - f3) / (0.000001 * Tg2)
m(4, 1) = (n2f4 - f4) / (0.000001 * Tg2)
m(5, 1) = (n2f5 - f5) / (0.000001 * Tg2)




'dtg3

tg31 = Tg3 * 1.000001

n3f1 = rad * (Tg2 ^ 4 - Tw ^ 4) - qgen + mg * smix * (Tg2 - Tg1)
n3f2 = u1 * A1 * ((Tg2 - Tf2) - (tg31 - Tf3)) / Log((Tg2 - Tf2) / (tg31 - Tf3)) + u2 * A2 * ((Tg2 - Tf2) - (tg31 - Tf1)) / Log((Tg2 - Tf2) / (tg31 - Tf1)) - mg * smix * (Tg2 - tg31)
n3f3 = u3 * A3 * ((tg31 - Tf1) - (Tg4 - Tf2)) / Log((tg31 - Tf1) / (Tg4 - Tf2)) - mg * smix * (tg31 - Tg4)
n3f4 = rad * (Tg2 ^ 4 - Tw ^ 4) + u1 * A1 * ((Tg2 - Tf2) - (tg31 - Tf3)) / Log((Tg2 - Tf2) / (tg31 - Tf3)) - mfl * cpfl * (Tf3 - Tf2)
n3f5 = u3 * A3 * ((tg31 - Tf1) - (Tg4 - Tf2)) / Log((tg31 - Tf1) / (Tg4 - Tf2)) + u2 * A2 * ((Tg2 - Tf2) - (tg31 - Tf1)) / Log((Tg2 - Tf2) / (tg31 - Tf1)) - mfl * cpfl * (Tf2 - Tf1)

m(1, 2) = (n3f1 - f1) / (0.000001 * Tg3)
m(2, 2) = (n3f2 - f2) / (0.000001 * Tg3)
m(3, 2) = (n3f3 - f3) / (0.000001 * Tg3)
m(4, 2) = (n3f4 - f4) / (0.000001 * Tg3)
m(5, 2) = (n3f5 - f5) / (0.000001 * Tg3)



'dtg4

tg41 = Tg4 * 1.000001

n4f1 = rad * (Tg2 ^ 4 - Tw ^ 4) - qgen + mg * smix * (Tg2 - Tg1)
n4f2 = u1 * A1 * ((Tg2 - Tf2) - (Tg3 - Tf3)) / Log((Tg2 - Tf2) / (Tg3 - Tf3)) + u2 * A2 * ((Tg2 - Tf2) - (Tg3 - Tf1)) / Log((Tg2 - Tf2) / (Tg3 - Tf1)) - mg * smix * (Tg2 - Tg3)
n4f3 = u3 * A3 * ((Tg3 - Tf1) - (tg41 - Tf2)) / Log((Tg3 - Tf1) / (tg41 - Tf2)) - mg * smix * (Tg3 - tg41)
n4f4 = rad * (Tg2 ^ 4 - Tw ^ 4) + u1 * A1 * ((Tg2 - Tf2) - (Tg3 - Tf3)) / Log((Tg2 - Tf2) / (Tg3 - Tf3)) - mfl * cpfl * (Tf3 - Tf2)
n4f5 = u3 * A3 * ((Tg3 - Tf1) - (tg41 - Tf2)) / Log((Tg3 - Tf1) / (tg41 - Tf2)) + u2 * A2 * ((Tg2 - Tf2) - (Tg3 - Tf1)) / Log((Tg2 - Tf2) / (Tg3 - Tf1)) - mfl * cpfl * (Tf2 - Tf1)

m(1, 3) = (n4f1 - f1) / (0.000001 * Tg4)
m(2, 3) = (n4f2 - f2) / (0.000001 * Tg4)
m(3, 3) = (n4f3 - f3) / (0.000001 * Tg4)
m(4, 3) = (n4f4 - f4) / (0.000001 * Tg4)
m(5, 3) = (n4f5 - f5) / (0.000001 * Tg4)


'dtf2

tf21 = Tf2 * 1.000001

n5f1 = rad * (Tg2 ^ 4 - Tw ^ 4) - qgen + mg * smix * (Tg2 - Tg1)
n5f2 = u1 * A1 * ((Tg2 - tf21) - (Tg3 - Tf3)) / Log((Tg2 - tf21) / (Tg3 - Tf3)) + u2 * A2 * ((Tg2 - tf21) - (Tg3 - Tf1)) / Log((Tg2 - tf21) / (Tg3 - Tf1)) - mg * smix * (Tg2 - Tg3)
n5f3 = u3 * A3 * ((Tg3 - Tf1) - (Tg4 - tf21)) / Log((Tg3 - Tf1) / (Tg4 - tf21)) - mg * smix * (Tg3 - Tg4)
n5f4 = rad * (Tg2 ^ 4 - Tw ^ 4) + u1 * A1 * ((Tg2 - tf21) - (Tg3 - Tf3)) / Log((Tg2 - tf21) / (Tg3 - Tf3)) - mfl * cpfl * (Tf3 - tf21)
n5f5 = u3 * A3 * ((Tg3 - Tf1) - (Tg4 - tf21)) / Log((Tg3 - Tf1) / (Tg4 - tf21)) + u2 * A2 * ((Tg2 - tf21) - (Tg3 - Tf1)) / Log((Tg2 - tf21) / (Tg3 - Tf1)) - mfl * cpfl * (tf21 - Tf1)

m(1, 4) = (n5f1 - f1) / (0.000001 * Tf2)
m(2, 4) = (n5f2 - f2) / (0.000001 * Tf2)
m(3, 4) = (n5f3 - f3) / (0.000001 * Tf2)
m(4, 4) = (n5f4 - f4) / (0.000001 * Tf2)
m(5, 4) = (n5f5 - f5) / (0.000001 * Tf2)


'dtf3

tf31 = Tf3 * 1.000001

n6f1 = rad * (Tg2 ^ 4 - Tw ^ 4) - qgen + mg * smix * (Tg2 - Tg1)
n6f2 = u1 * A1 * ((Tg2 - Tf2) - (Tg3 - tf31)) / Log((Tg2 - Tf2) / (Tg3 - tf31)) + u2 * A2 * ((Tg2 - Tf2) - (Tg3 - Tf1)) / Log((Tg2 - Tf2) / (Tg3 - Tf1)) - mg * smix * (Tg2 - Tg3)
n6f3 = u3 * A3 * ((Tg3 - Tf1) - (Tg4 - Tf2)) / Log((Tg3 - Tf1) / (Tg4 - Tf2)) - mg * smix * (Tg3 - Tg4)
n6f4 = rad * (Tg2 ^ 4 - Tw ^ 4) + u1 * A1 * ((Tg2 - Tf2) - (Tg3 - tf31)) / Log((Tg2 - Tf2) / (Tg3 - tf31)) - mfl * cpfl * (tf31 - Tf2)
n6f5 = u3 * A3 * ((Tg3 - Tf1) - (Tg4 - Tf2)) / Log((Tg3 - Tf1) / (Tg4 - Tf2)) + u2 * A2 * ((Tg2 - Tf2) - (Tg3 - Tf1)) / Log((Tg2 - Tf2) / (Tg3 - Tf1)) - mfl * cpfl * (Tf2 - Tf1)

m(1, 5) = (n6f1 - f1) / (0.000001 * Tf3)
m(2, 5) = (n6f2 - f2) / (0.000001 * Tf3)
m(3, 5) = (n6f3 - f3) / (0.000001 * Tf3)
m(4, 5) = (n6f4 - f4) / (0.000001 * Tf3)
m(5, 5) = (n6f5 - f5) / (0.000001 * Tf3)



va(1, 1) = f1
va(2, 1) = f2
va(3, 1) = f3
va(4, 1) = f4
va(5, 1) = f5

'Print
'Print "          MATRIX [ M ]"
'Print
'Print m(1, 1), m(1, 2), m(1, 3), m(1, 4), m(1, 5)
'Print
'Print m(2, 1), m(2, 2), m(2, 3), m(2, 4), m(2, 5)
'Print
'Print m(3, 1), m(3, 2), m(3, 3), m(3, 4), m(3, 5)
'Print
'Print m(4, 1), m(4, 2), m(4, 3), m(4, 4), m(4, 5)
'Print
'Print m(5, 1), m(5, 2), m(5, 3), m(5, 4), m(5, 5)
'Print

For r = 1 To 5
For c = 1 To 5
cp(r, c) = m(r, c)
Next c
Next r


' Create the identity matrix

For r = 1 To 5
For c = 1 To 5
id(r, c) = 0
Next c
Next r
For i = 1 To 5
id(i, i) = 1
Next i




'ROW I


For c = 1 To 5
m(1, c) = m(1, c) / cp(1, 1)
id(1, c) = id(1, c) / cp(1, 1)

Next c

For r = 2 To 5
For c = 1 To 5
m(r, c) = m(r, c) - cp(r, 1) * m(1, c)
id(r, c) = id(r, c) - cp(r, 1) * id(1, c)
Next c
Next r

For r = 1 To 5
cp(r, 2) = m(r, 2)
Next r


'ROW II


For c = 1 To 5
m(2, c) = m(2, c) / cp(2, 2)
id(2, c) = id(2, c) / cp(2, 2)

Next c

For r = 3 To 5

For c = 1 To 5
m(r, c) = m(r, c) - cp(r, 2) * m(2, c)
id(r, c) = id(r, c) - cp(r, 2) * id(2, c)
Next c
Next r

For c = 1 To 5
m(1, c) = m(1, c) - cp(1, 2) * m(2, c)
id(1, c) = id(1, c) - cp(1, 2) * id(2, c)
Next c

For r = 1 To 5
cp(r, 3) = m(r, 3)
Next r

'ROW III


For c = 1 To 5
m(3, c) = m(3, c) / cp(3, 3)
id(3, c) = id(3, c) / cp(3, 3)

Next c

For r = 1 To 2

For c = 1 To 5
m(r, c) = m(r, c) - cp(r, 3) * m(3, c)
id(r, c) = id(r, c) - cp(r, 3) * id(3, c)
Next c
Next r

For r = 4 To 5
For c = 1 To 5
m(r, c) = m(r, c) - cp(r, 3) * m(3, c)
id(r, c) = id(r, c) - cp(r, 3) * id(3, c)
Next c
Next r

For r = 1 To 5
cp(r, 4) = m(r, 4)
Next r


'ROW IV


For c = 1 To 5
m(4, c) = m(4, c) / cp(4, 4)
id(4, c) = id(4, c) / cp(4, 4)

Next c

For r = 1 To 3

For c = 1 To 5
m(r, c) = m(r, c) - cp(r, 4) * m(4, c)
id(r, c) = id(r, c) - cp(r, 4) * id(4, c)
Next c
Next r


For c = 1 To 5
m(5, c) = m(5, c) - cp(5, 4) * m(4, c)
id(5, c) = id(5, c) - cp(5, 4) * id(4, c)
Next c


For r = 1 To 5
cp(r, 5) = m(r, 5)
Next r

'ROW V


For c = 1 To 5
m(5, c) = m(5, c) / cp(5, 5)
id(5, c) = id(5, c) / cp(5, 5)

Next c

For r = 1 To 4
For c = 1 To 5
m(r, c) = m(r, c) - cp(r, 5) * m(5, c)
id(r, c) = id(r, c) - cp(r, 5) * id(5, c)
Next c
Next r



sol(1, 1) = va(1, 1) * id(1, 1) + va(2, 1) * id(1, 2) + va(3, 1) * id(1, 3) + va(4, 1) * id(1, 4) + va(5, 1) * id(1, 5)
sol(2, 1) = va(1, 1) * id(2, 1) + va(2, 1) * id(2, 2) + va(3, 1) * id(2, 3) + va(4, 1) * id(2, 4) + va(5, 1) * id(2, 5)
sol(3, 1) = va(1, 1) * id(3, 1) + va(2, 1) * id(3, 2) + va(3, 1) * id(3, 3) + va(4, 1) * id(3, 4) + va(5, 1) * id(3, 5)
sol(4, 1) = va(1, 1) * id(4, 1) + va(2, 1) * id(4, 2) + va(3, 1) * id(4, 3) + va(4, 1) * id(4, 4) + va(5, 1) * id(4, 5)
sol(5, 1) = va(1, 1) * id(5, 1) + va(2, 1) * id(5, 2) + va(3, 1) * id(5, 3) + va(4, 1) * id(5, 4) + va(5, 1) * id(5, 5)






'tg1 = 300
Tg2 = Tg2 - 0.1 * sol(1, 1)
Tg3 = Tg3 - 0.1 * sol(2, 1)
Tg4 = Tg4 - 0.1 * sol(3, 1)
Tf2 = Tf2 - 0.1 * sol(4, 1)
Tf3 = Tf3 - 0.1 * sol(5, 1)

If Abs(f1) > 1 Or Abs(f2) > 1 Or Abs(f3) > 1 Or Abs(f4) > 1 Or Abs(f5) > 1 Then

GoTo iter1

End If




Print
'Print "           INVERSE"
'Print
'Print id(1, 1), id(1, 2), id(1, 3), id(1, 4), id(1, 5)
'Print
'Print id(2, 1), id(2, 2), id(2, 3), id(2, 4), id(2, 5)
'Print
'Print id(3, 1), id(3, 2), id(3, 3), id(3, 4), id(3, 5)
'Print
'Print id(4, 1), id(4, 2), id(4, 3), id(4, 4), id(4, 5)
'Print
'Print id(5, 1), id(5, 2), id(5, 3), id(5, 4), id(5, 5)
Print

''Print "  JACOBIAN INVERSE                     F{To}"
''Print
''Print sol(1, 1), f1
''Print
''Print sol(2, 1), f2
''Print
''Print sol(3, 1), f3
''Print
''Print sol(4, 1), f4
''Print
''Print sol(5, 1), f5
''Print
''Print

''Print "        TEMPERATURES "
''Print
''Print "      Tg1    =   "; tg1
''Print
''Print "      Tg2    =    "; tg2
''Print
''Print "      Tg3    =    "; tg3
''Print
''Print "      Tg4    =    "; tg4
''Print
''Print "      Tf1    =    "; tf1
''Print
''Print "      Tf2    =    "; tf2
''Print
''Print "      Tf3    =    "; tf3
''Print

Text24.Text = Tg1
Text25.Text = Tg2
Text26.Text = Tg3
Text27.Text = Tg4
Text28.Text = Tf1
Text29.Text = Tf2
Text30.Text = Tf3


End Sub
