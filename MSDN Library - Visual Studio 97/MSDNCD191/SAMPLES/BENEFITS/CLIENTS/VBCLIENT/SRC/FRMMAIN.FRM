VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Object = "{0AC835B7-756D-11D0-8317-00AA006C3AB9}#2.0#0"; "sbarvb.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Corporate Benefits"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   1.9821e5
   ClientWidth     =   9300
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Tag             =   "2100"
   Begin SideBarVB.SideBar sbCategories 
      Height          =   6120
      Left            =   15
      TabIndex        =   164
      Top             =   1185
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   10795
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SmallImageList  =   "ilSideBar"
      LargeImageList  =   "ilSideBar"
   End
   Begin VB.PictureBox pctBanks 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   6105
      Left            =   2280
      ScaleHeight     =   6105
      ScaleWidth      =   5955
      TabIndex        =   109
      TabStop         =   0   'False
      Tag             =   "Banks"
      Top             =   2505
      Visible         =   0   'False
      Width           =   5955
      Begin VB.Frame gbBanks 
         Caption         =   "Banks"
         Height          =   5565
         Left            =   75
         TabIndex        =   126
         Tag             =   "2128"
         Top             =   90
         Width           =   4185
         Begin VB.ListBox lstAllBanks 
            Height          =   2205
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   61
            Top             =   240
            Width           =   3960
         End
         Begin VB.TextBox txtBankCity 
            Height          =   330
            Left            =   510
            TabIndex        =   66
            Top             =   5100
            Width           =   3495
         End
         Begin VB.TextBox txtBankAddress2 
            Height          =   330
            Left            =   510
            TabIndex        =   65
            Top             =   4695
            Width           =   3495
         End
         Begin VB.TextBox txtBankAddress1 
            Height          =   330
            Left            =   510
            TabIndex        =   64
            Top             =   4275
            Width           =   3495
         End
         Begin VB.TextBox txtBankName 
            DataField       =   "FirstName"
            DataSource      =   "dataEmployees"
            Height          =   330
            Left            =   90
            TabIndex        =   62
            Top             =   2955
            Width           =   3960
         End
         Begin VB.TextBox txtBankCode 
            Height          =   330
            Left            =   90
            TabIndex        =   63
            Top             =   3600
            Width           =   1905
         End
         Begin VB.Label lblBankAddress 
            Caption         =   "&Address :"
            Height          =   255
            Left            =   90
            TabIndex        =   130
            Tag             =   "2131"
            Top             =   4035
            Width           =   1530
         End
         Begin VB.Label lblCityBank 
            Caption         =   "Cit&y :"
            Height          =   210
            Left            =   90
            TabIndex        =   129
            Tag             =   "2132"
            Top             =   5160
            Width           =   375
         End
         Begin VB.Label lblBankName 
            Caption         =   "&Name :"
            Height          =   210
            Left            =   90
            TabIndex        =   128
            Tag             =   "2129"
            Top             =   2715
            Width           =   945
         End
         Begin VB.Label lblBankCode 
            Caption         =   "&Code :"
            Height          =   195
            Left            =   90
            TabIndex        =   127
            Tag             =   "2130"
            Top             =   3360
            Width           =   990
         End
      End
      Begin VB.CommandButton cmdDeleteBank 
         Caption         =   "&Delete Bank"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4360
         TabIndex        =   71
         Tag             =   "2135"
         Top             =   1740
         Width           =   1530
      End
      Begin VB.CommandButton cmdEditBank 
         Caption         =   "&Edit Bank"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4360
         TabIndex        =   70
         Tag             =   "2134"
         Top             =   1350
         Width           =   1530
      End
      Begin VB.CommandButton cmdAddBank 
         Caption         =   "&Add Bank"
         Height          =   345
         Left            =   4360
         TabIndex        =   69
         Tag             =   "2133"
         Top             =   975
         Width           =   1530
      End
      Begin VB.CommandButton cmdBankOK 
         Caption         =   "OK"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4740
         TabIndex        =   67
         Tag             =   "2004"
         Top             =   90
         Width           =   1150
      End
      Begin VB.CommandButton cmdBankCancel 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4740
         TabIndex        =   68
         Tag             =   "2005"
         Top             =   480
         Width           =   1150
      End
   End
   Begin VB.PictureBox pctDependants 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   6120
      Left            =   2160
      ScaleHeight     =   6120
      ScaleWidth      =   5955
      TabIndex        =   100
      TabStop         =   0   'False
      Tag             =   "Dependants"
      Top             =   2295
      Visible         =   0   'False
      Width           =   5955
      Begin VB.CommandButton cmdEditDependant 
         Caption         =   "&Edit Dependant"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4360
         TabIndex        =   59
         Tag             =   "2173"
         Top             =   1350
         Width           =   1530
      End
      Begin VB.CommandButton cmdAddDependant 
         Caption         =   "&Add Dependant"
         Height          =   345
         Left            =   4360
         TabIndex        =   58
         Tag             =   "2172"
         Top             =   975
         Width           =   1530
      End
      Begin VB.Frame gbDependants 
         Caption         =   "Employees De&pendants"
         Height          =   4560
         Left            =   75
         TabIndex        =   103
         Tag             =   "2161"
         Top             =   975
         Width           =   4185
         Begin VB.TextBox txtDepDOB 
            DataField       =   "FirstName"
            DataSource      =   "dataEmployees"
            Height          =   330
            Left            =   1155
            TabIndex        =   50
            Top             =   2625
            Width           =   2175
         End
         Begin VB.ComboBox cbDepRelationship 
            Height          =   315
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   4065
            Width           =   2850
         End
         Begin VB.Frame gbDepStatus 
            Caption         =   "Dependant &Status"
            Height          =   825
            Left            =   1950
            TabIndex        =   108
            Tag             =   "2168"
            Top             =   3045
            Width           =   1755
            Begin VB.OptionButton optDepStatusDependant 
               Caption         =   "Depe&ndant"
               Height          =   285
               Left            =   135
               TabIndex        =   53
               Tag             =   "2169"
               Top             =   210
               Value           =   -1  'True
               Width           =   1170
            End
            Begin VB.OptionButton optDepStatusBeneficiary 
               Caption         =   "&Beneficiary"
               Height          =   285
               Left            =   135
               TabIndex        =   54
               Tag             =   "2170"
               Top             =   495
               Width           =   1095
            End
         End
         Begin VB.TextBox txtDepFirstName 
            DataField       =   "FirstName"
            DataSource      =   "dataEmployees"
            Height          =   330
            Left            =   1155
            TabIndex        =   48
            Top             =   1740
            Width           =   2175
         End
         Begin VB.TextBox txtDepLastName 
            DataField       =   "LastName"
            DataSource      =   "dataEmployees"
            Height          =   330
            Left            =   1155
            TabIndex        =   49
            Top             =   2175
            Width           =   2850
         End
         Begin VB.Frame gbDepGender 
            Caption         =   "&Gender"
            Height          =   825
            Left            =   270
            TabIndex        =   104
            Tag             =   "2149"
            Top             =   3045
            Width           =   1305
            Begin VB.OptionButton optDepGenderFemale 
               Caption         =   "Fem&ale"
               Height          =   285
               Left            =   135
               TabIndex        =   52
               Tag             =   "2151"
               Top             =   495
               Width           =   810
            End
            Begin VB.OptionButton optDepGenderMale 
               Caption         =   "&Male"
               Height          =   285
               Left            =   135
               TabIndex        =   51
               Tag             =   "2150"
               Top             =   210
               Value           =   -1  'True
               Width           =   900
            End
         End
         Begin VB.ListBox lstDependants 
            Height          =   1425
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   47
            Top             =   225
            Width           =   3960
         End
         Begin VB.Label lblDepDOB 
            Caption         =   "Date of Birth :"
            Height          =   210
            Left            =   120
            TabIndex        =   132
            Tag             =   "2164"
            Top             =   2685
            Width           =   975
         End
         Begin VB.Label lblDepFirstName 
            Caption         =   "&First Name :"
            Height          =   210
            Left            =   120
            TabIndex        =   107
            Tag             =   "2162"
            Top             =   1800
            Width           =   840
         End
         Begin VB.Label lblDepLastName 
            Caption         =   "&Last Name :"
            Height          =   210
            Left            =   120
            TabIndex        =   106
            Tag             =   "2163"
            Top             =   2242
            Width           =   870
         End
         Begin VB.Label lblDepRelationship 
            Caption         =   "&Relationship :"
            Height          =   210
            Left            =   135
            TabIndex        =   105
            Tag             =   "2171"
            Top             =   4095
            Width           =   990
         End
      End
      Begin VB.Frame gbEmployee 
         Caption         =   "Employee"
         Height          =   795
         Left            =   75
         TabIndex        =   101
         Tag             =   "2160"
         Top             =   90
         Width           =   4185
         Begin VB.Label lblEmployeeName 
            Height          =   240
            Left            =   120
            TabIndex        =   102
            Top             =   345
            Width           =   3960
         End
      End
      Begin VB.CommandButton cmdDependantOK 
         Caption         =   "OK"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4740
         TabIndex        =   56
         Tag             =   "2004"
         Top             =   90
         Width           =   1150
      End
      Begin VB.CommandButton cmdDependantCancel 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4740
         TabIndex        =   57
         Tag             =   "2005"
         Top             =   480
         Width           =   1150
      End
      Begin VB.CommandButton cmdDeleteDependant 
         Caption         =   "&Delete Dependant"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4360
         TabIndex        =   60
         Tag             =   "2174"
         Top             =   1740
         Width           =   1530
      End
   End
   Begin VB.PictureBox pctEmployees 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   6120
      Left            =   2040
      ScaleHeight     =   6120
      ScaleWidth      =   5955
      TabIndex        =   86
      TabStop         =   0   'False
      Tag             =   "Employees"
      Top             =   2070
      Visible         =   0   'False
      Width           =   5955
      Begin VB.TextBox txtEmplPosition 
         Height          =   330
         Left            =   1050
         TabIndex        =   29
         Top             =   1440
         Width           =   3060
      End
      Begin VB.ComboBox cbAllEmployees 
         Height          =   315
         Left            =   90
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   300
         Width           =   4005
      End
      Begin VB.CommandButton cmdEmployeeOK 
         Caption         =   "OK"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4770
         TabIndex        =   42
         Tag             =   "2004"
         Top             =   90
         Width           =   1150
      End
      Begin VB.TextBox txtEmplSalary 
         Height          =   330
         Left            =   1605
         TabIndex        =   34
         Top             =   3705
         Width           =   1305
      End
      Begin VB.TextBox txtEmplUserName 
         Height          =   330
         Left            =   1605
         TabIndex        =   33
         Top             =   3300
         Width           =   1305
      End
      Begin VB.TextBox txtEmplStartDate 
         Height          =   330
         Left            =   4575
         TabIndex        =   38
         Top             =   3735
         Width           =   1305
      End
      Begin VB.TextBox txtEmplEndDate 
         Height          =   330
         Left            =   4575
         TabIndex        =   39
         Top             =   4110
         Width           =   1305
      End
      Begin VB.CommandButton cmdEmployeeCancel 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4740
         TabIndex        =   43
         Tag             =   "2005"
         Top             =   480
         Width           =   1150
      End
      Begin VB.TextBox txtEmplDOB 
         Height          =   330
         Left            =   1605
         TabIndex        =   35
         Top             =   4110
         Width           =   1305
      End
      Begin VB.CommandButton cmdEditEmployee 
         Caption         =   "&Edit Employee"
         Height          =   345
         Left            =   4360
         TabIndex        =   45
         Tag             =   "2156"
         Top             =   1350
         Width           =   1530
      End
      Begin VB.CommandButton cmdAddEmployee 
         Caption         =   "&Add Employee"
         Height          =   345
         Left            =   4360
         TabIndex        =   44
         Tag             =   "2155"
         Top             =   975
         Width           =   1530
      End
      Begin VB.Frame gbEmplBank 
         Caption         =   "Bank Details"
         Height          =   960
         Left            =   60
         TabIndex        =   92
         Tag             =   "2152"
         Top             =   4440
         Width           =   5835
         Begin VB.TextBox txtEmplBankAccount 
            Height          =   330
            Left            =   1200
            TabIndex        =   41
            Top             =   540
            Width           =   1905
         End
         Begin VB.ComboBox cbEmplBanks 
            Height          =   315
            ItemData        =   "frmMain.frx":0442
            Left            =   1200
            List            =   "frmMain.frx":0444
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   180
            Width           =   4515
         End
         Begin VB.Label lblEmplBankAccount 
            Caption         =   "A&ccount No. :"
            Height          =   195
            Left            =   120
            TabIndex        =   94
            Tag             =   "2154"
            Top             =   600
            Width           =   990
         End
         Begin VB.Label lblEmplBankList 
            Caption         =   "Ban&k :"
            Height          =   195
            Left            =   120
            TabIndex        =   93
            Tag             =   "2153"
            Top             =   255
            Width           =   495
         End
      End
      Begin VB.Frame gbEmplGender 
         Caption         =   "&Gender"
         Height          =   510
         Left            =   3015
         TabIndex        =   91
         Tag             =   "2149"
         Top             =   3180
         Width           =   2865
         Begin VB.OptionButton optEmplGenderFemale 
            Caption         =   "F&emale"
            Height          =   285
            Left            =   1560
            TabIndex        =   37
            Tag             =   "2151"
            Top             =   195
            Width           =   810
         End
         Begin VB.OptionButton optEmplGenderMale 
            Caption         =   "&Male"
            Height          =   285
            Left            =   420
            TabIndex        =   36
            Tag             =   "2150"
            Top             =   195
            Value           =   -1  'True
            Width           =   765
         End
      End
      Begin VB.Frame gbEmplAddress 
         Caption         =   "Addre&ss"
         Height          =   1410
         Left            =   60
         TabIndex        =   89
         Tag             =   "2142"
         Top             =   1770
         Width           =   4185
         Begin VB.TextBox txtEmplCity 
            Height          =   330
            Left            =   510
            TabIndex        =   32
            Top             =   990
            Width           =   3525
         End
         Begin VB.TextBox txtEmplAddress2 
            Height          =   330
            Left            =   510
            TabIndex        =   31
            Top             =   615
            Width           =   3525
         End
         Begin VB.TextBox txtEmplAddress1 
            Height          =   330
            Left            =   510
            TabIndex        =   30
            Top             =   225
            Width           =   3525
         End
         Begin VB.Label lblEmplCity 
            Caption         =   "Cit&y :"
            Height          =   210
            Left            =   75
            TabIndex        =   90
            Tag             =   "2143"
            Top             =   1050
            Width           =   375
         End
      End
      Begin VB.TextBox txtEmplLastName 
         Height          =   330
         Left            =   1050
         TabIndex        =   28
         Top             =   1065
         Width           =   3060
      End
      Begin VB.TextBox txtEmplFirstName 
         Height          =   330
         Left            =   1050
         TabIndex        =   27
         Top             =   675
         Width           =   3060
      End
      Begin VB.CommandButton cmdDeleteEmployee 
         Caption         =   "&Delete Employee"
         Height          =   345
         Left            =   4360
         TabIndex        =   46
         Tag             =   "2157"
         Top             =   1740
         Width           =   1530
      End
      Begin VB.Label lblEmplPosition 
         Caption         =   "&Position :"
         Height          =   210
         Left            =   105
         TabIndex        =   135
         Tag             =   "2141"
         Top             =   1500
         Width           =   900
      End
      Begin VB.Label lblEmployeeCurrency 
         Alignment       =   1  'Right Justify
         Caption         =   "$"
         Height          =   210
         Left            =   1155
         TabIndex        =   134
         Tag             =   "2175"
         Top             =   3765
         Width           =   405
      End
      Begin VB.Label lblEmployeesCombo 
         Caption         =   "Employees"
         Height          =   210
         Left            =   75
         TabIndex        =   131
         Tag             =   "2138"
         Top             =   90
         Width           =   2370
      End
      Begin VB.Label lblEmplEndDate 
         Caption         =   "End of Empl&oyment :"
         Height          =   210
         Left            =   3030
         TabIndex        =   111
         Tag             =   "2145"
         Top             =   4170
         Width           =   1515
      End
      Begin VB.Label lblEmplStartDate 
         Caption         =   "Start of Employment :"
         Height          =   210
         Left            =   3030
         TabIndex        =   110
         Tag             =   "2144"
         Top             =   3795
         Width           =   1515
      End
      Begin VB.Label lblEmplUserName 
         Caption         =   "&User Name :"
         Height          =   195
         Left            =   90
         TabIndex        =   99
         Tag             =   "2146"
         Top             =   3360
         Width           =   1275
      End
      Begin VB.Label lblEmplSalary 
         Caption         =   "Sala&ry :"
         Height          =   210
         Left            =   90
         TabIndex        =   98
         Tag             =   "2147"
         Top             =   3750
         Width           =   1275
      End
      Begin VB.Label lblEmplDOB 
         Caption         =   "Date of &Birth :"
         Height          =   210
         Left            =   90
         TabIndex        =   97
         Tag             =   "2148"
         Top             =   4155
         Width           =   1275
      End
      Begin VB.Label lblEmplLastName 
         Caption         =   "Last Na&me :"
         Height          =   210
         Left            =   105
         TabIndex        =   88
         Tag             =   "2140"
         Top             =   1125
         Width           =   900
      End
      Begin VB.Label lblEmplFirstName 
         Caption         =   "F&irst Name :"
         Height          =   210
         Left            =   105
         TabIndex        =   87
         Tag             =   "2139"
         Top             =   735
         Width           =   900
      End
   End
   Begin VB.PictureBox pctBenefits 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   6120
      Left            =   1800
      ScaleHeight     =   6120
      ScaleWidth      =   5940
      TabIndex        =   112
      TabStop         =   0   'False
      Tag             =   "Benefits"
      Top             =   1845
      Visible         =   0   'False
      Width           =   5940
      Begin VB.Frame gbBenefitDetails 
         Caption         =   "Benefits"
         Height          =   5355
         Left            =   75
         TabIndex        =   113
         Tag             =   "2114"
         Top             =   90
         Width           =   4185
         Begin VB.ComboBox cbBenefitType 
            Height          =   315
            ItemData        =   "frmMain.frx":0446
            Left            =   1815
            List            =   "frmMain.frx":0448
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   2430
            Width           =   1515
         End
         Begin VB.TextBox txtBenefitMultiples 
            Height          =   330
            Left            =   1815
            TabIndex        =   19
            Top             =   4500
            Width           =   1260
         End
         Begin VB.TextBox txtBenefitIncrement 
            Height          =   330
            Left            =   1815
            TabIndex        =   18
            Top             =   4095
            Width           =   1260
         End
         Begin VB.TextBox txtBenefitMaximum 
            Height          =   330
            Left            =   1815
            TabIndex        =   17
            Top             =   3675
            Width           =   1260
         End
         Begin VB.TextBox txtBenefitMinimum 
            Height          =   330
            Left            =   1815
            TabIndex        =   16
            Top             =   3255
            Width           =   1260
         End
         Begin VB.TextBox txtBenefitCost 
            Height          =   330
            Left            =   1815
            TabIndex        =   15
            Top             =   2820
            Width           =   1260
         End
         Begin VB.ListBox lstAllBenefits 
            Height          =   1425
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   12
            Top             =   240
            Width           =   3960
         End
         Begin VB.TextBox txtBenefitDescription 
            DataField       =   "LastName"
            DataSource      =   "dataEmployees"
            Height          =   330
            Left            =   120
            TabIndex        =   13
            Top             =   2025
            Width           =   3960
         End
         Begin VB.ComboBox cbBenefitTaxStatus 
            Height          =   315
            ItemData        =   "frmMain.frx":044A
            Left            =   1815
            List            =   "frmMain.frx":044C
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   4905
            Width           =   1515
         End
         Begin VB.Label lblBenefitCurrency 
            Alignment       =   1  'Right Justify
            Caption         =   "$"
            Height          =   210
            Left            =   1365
            TabIndex        =   133
            Tag             =   "2175"
            Top             =   2880
            Width           =   405
         End
         Begin VB.Label lblBenefitType 
            Caption         =   "&Benefit Categories :"
            Height          =   210
            Left            =   135
            TabIndex        =   121
            Tag             =   "2116"
            Top             =   2475
            Width           =   1635
         End
         Begin VB.Label lblBenefitMultiples 
            Caption         =   "Benefit &Multiples :"
            Height          =   210
            Left            =   135
            TabIndex        =   120
            Tag             =   "2121"
            Top             =   4560
            Width           =   1620
         End
         Begin VB.Label lblBenefitIncrement 
            Caption         =   "Benefit &Increment :"
            Height          =   210
            Left            =   135
            TabIndex        =   119
            Tag             =   "2120"
            Top             =   4155
            Width           =   1620
         End
         Begin VB.Label lblBenefitMaximum 
            Caption         =   "Benefit Ma&ximum :"
            Height          =   210
            Left            =   135
            TabIndex        =   118
            Tag             =   "2119"
            Top             =   3735
            Width           =   1620
         End
         Begin VB.Label lblBenefitMinimum 
            Caption         =   "Benefit Mi&nimum :"
            Height          =   210
            Left            =   135
            TabIndex        =   117
            Tag             =   "2118"
            Top             =   3315
            Width           =   1620
         End
         Begin VB.Label lblBenefitCost 
            Caption         =   "&Cost :"
            Height          =   210
            Left            =   135
            TabIndex        =   116
            Tag             =   "2117"
            Top             =   2880
            Width           =   1005
         End
         Begin VB.Label lblBenefitTaxStatus 
            Caption         =   "&Tax Status :"
            Height          =   210
            Left            =   135
            TabIndex        =   115
            Tag             =   "2122"
            Top             =   4965
            Width           =   1620
         End
         Begin VB.Label lblDescription 
            Caption         =   "Benefit &Description :"
            Height          =   210
            Left            =   120
            TabIndex        =   114
            Tag             =   "2115"
            Top             =   1815
            Width           =   3825
         End
      End
      Begin VB.CommandButton cmdDeleteBenefit 
         Caption         =   "&Delete Benefit"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4360
         TabIndex        =   25
         Tag             =   "2125"
         Top             =   1740
         Width           =   1530
      End
      Begin VB.CommandButton cmdAddBenefit 
         Caption         =   "&Add Benefit"
         Height          =   345
         Left            =   4360
         TabIndex        =   23
         Tag             =   "2123"
         Top             =   975
         Width           =   1530
      End
      Begin VB.CommandButton cmdEditBenefit 
         Caption         =   "&Edit Benefit"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4360
         TabIndex        =   24
         Tag             =   "2124"
         Top             =   1350
         Width           =   1530
      End
      Begin VB.CommandButton cmdBenefitCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4740
         TabIndex        =   22
         Tag             =   "2005"
         Top             =   480
         Width           =   1150
      End
      Begin VB.CommandButton cmdBenefitOK 
         Caption         =   "OK"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4740
         TabIndex        =   21
         Tag             =   "2004"
         Top             =   90
         Width           =   1150
      End
   End
   Begin VB.PictureBox pctResults 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   5880
      Left            =   1710
      ScaleHeight     =   5880
      ScaleWidth      =   5895
      TabIndex        =   84
      TabStop         =   0   'False
      Tag             =   "Results"
      Top             =   1680
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton cmdAnotherQuery 
         Caption         =   "Another &Query"
         Height          =   345
         Left            =   4360
         TabIndex        =   11
         Tag             =   "2127"
         ToolTipText     =   "Click here to fire off query"
         Top             =   90
         Width           =   1530
      End
      Begin VB.Frame gbBenefit 
         Caption         =   "Benefit Query"
         Height          =   990
         Left            =   75
         TabIndex        =   85
         Tag             =   "2126"
         ToolTipText     =   "Benefit queried on"
         Top             =   75
         Width           =   4185
         Begin VB.Label lblQueryResultsText 
            Height          =   570
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   3960
            WordWrap        =   -1  'True
         End
      End
      Begin MSChartLib.MSChart ctlMSChart 
         Height          =   4665
         Left            =   75
         OleObjectBlob   =   "frmMain.frx":044E
         TabIndex        =   155
         Top             =   1170
         Width           =   5760
      End
      Begin VB.OLE oleMSGraph 
         Height          =   4665
         Left            =   75
         SizeMode        =   1  'Stretch
         TabIndex        =   124
         Top             =   1170
         Width           =   5760
      End
   End
   Begin VB.PictureBox pctQuery 
      BorderStyle     =   0  'None
      Height          =   6120
      Left            =   1470
      ScaleHeight     =   6120
      ScaleWidth      =   5940
      TabIndex        =   79
      TabStop         =   0   'False
      Tag             =   "Queries"
      Top             =   1350
      Width           =   5940
      Begin VB.CommandButton cmdBenefitGlossary 
         Caption         =   "Benefit Glossary..."
         Height          =   345
         Left            =   4360
         TabIndex        =   9
         Tag             =   "2113"
         Top             =   975
         Width           =   1530
      End
      Begin VB.Frame gbBenefitQueries 
         Caption         =   "A&ll Benefits"
         Height          =   1665
         Left            =   75
         TabIndex        =   83
         Tag             =   "2104"
         Top             =   90
         Width           =   4185
         Begin VB.ComboBox cbQueryBenefits 
            Height          =   315
            ItemData        =   "frmMain.frx":291C
            Left            =   105
            List            =   "frmMain.frx":291E
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   1215
            Width           =   3960
         End
         Begin VB.ComboBox cbQueryBenefitTypes 
            Height          =   315
            ItemData        =   "frmMain.frx":2920
            Left            =   105
            List            =   "frmMain.frx":2922
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   510
            Width           =   3960
         End
         Begin VB.Label lblBenefits 
            Caption         =   "&Available Benefits :"
            Height          =   270
            Left            =   105
            TabIndex        =   96
            Tag             =   "2106"
            Top             =   960
            Width           =   1710
         End
         Begin VB.Label lblBenCats 
            Caption         =   "&Benefit Categories :"
            Height          =   270
            Left            =   105
            TabIndex        =   95
            Tag             =   "2105"
            Top             =   255
            Width           =   1710
         End
      End
      Begin VB.Frame gbQueryFactors 
         Caption         =   "Benefit &Factors"
         Height          =   2805
         Left            =   75
         TabIndex        =   80
         Tag             =   "2107"
         Top             =   2085
         Width           =   4185
         Begin VB.Frame gbQueryGender 
            Caption         =   "&Gender"
            Height          =   1110
            Left            =   105
            TabIndex        =   125
            Tag             =   "2149"
            Top             =   240
            Width           =   1650
            Begin VB.OptionButton optQueryGender 
               Caption         =   "&Ignore Gender"
               Height          =   285
               Index           =   0
               Left            =   135
               TabIndex        =   2
               Tag             =   "2108"
               Top             =   225
               Value           =   -1  'True
               Width           =   1350
            End
            Begin VB.OptionButton optQueryGender 
               Caption         =   "&Male"
               Height          =   285
               Index           =   1
               Left            =   135
               TabIndex        =   3
               Tag             =   "2150"
               Top             =   495
               Width           =   765
            End
            Begin VB.OptionButton optQueryGender 
               Caption         =   "F&emale"
               Height          =   285
               Index           =   2
               Left            =   135
               TabIndex        =   4
               Tag             =   "2151"
               Top             =   765
               Width           =   810
            End
         End
         Begin VB.ComboBox cbQueryCoverage 
            Height          =   315
            ItemData        =   "frmMain.frx":2924
            Left            =   105
            List            =   "frmMain.frx":2926
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1680
            Width           =   3960
         End
         Begin VB.ComboBox cbQueryTax 
            Height          =   315
            ItemData        =   "frmMain.frx":2928
            Left            =   105
            List            =   "frmMain.frx":292A
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   2325
            Width           =   1965
         End
         Begin VB.Label lblCoverage 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "&Coverage :"
            Height          =   195
            Left            =   105
            TabIndex        =   82
            Tag             =   "2109"
            Top             =   1440
            Width           =   780
         End
         Begin VB.Label lblTax 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "&Tax :"
            Height          =   195
            Left            =   105
            TabIndex        =   81
            Tag             =   "2110"
            Top             =   2070
            Width           =   360
         End
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "&Query"
         Height          =   345
         Left            =   4740
         TabIndex        =   7
         Tag             =   "2111"
         Top             =   90
         Width           =   1150
      End
      Begin VB.CommandButton cmdQueryReset 
         Appearance      =   0  'Flat
         Caption         =   "&Reset"
         Height          =   345
         Left            =   4740
         TabIndex        =   8
         Tag             =   "2112"
         Top             =   480
         Width           =   1150
      End
   End
   Begin ComctlLib.StatusBar statMain 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   150
      Top             =   8655
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327680
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   13098
            MinWidth        =   530
            Text            =   "Ready"
            TextSave        =   "Ready"
            Key             =   "sbStatus"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Current Application Status"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
            Key             =   "sbCaps"
            Object.Tag             =   ""
            Object.ToolTipText     =   "CAPS Lock"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
            Key             =   "sbNum"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Num Lock"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "INS"
            Key             =   "sbIns"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Insert Mode"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmMain.frx":292C
   End
   Begin VB.PictureBox pctDataServices 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   6120
      Left            =   1320
      ScaleHeight     =   408
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   405
      TabIndex        =   122
      TabStop         =   0   'False
      Tag             =   "Data Services"
      Top             =   1245
      Visible         =   0   'False
      Width           =   6075
      Begin VB.PictureBox picData 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Index           =   5
         Left            =   5190
         ScaleHeight     =   960
         ScaleWidth      =   855
         TabIndex        =   163
         Top             =   225
         Width           =   855
      End
      Begin VB.PictureBox picData 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Index           =   4
         Left            =   4155
         ScaleHeight     =   960
         ScaleWidth      =   1035
         TabIndex        =   162
         Top             =   225
         Width           =   1035
      End
      Begin VB.PictureBox picData 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Index           =   3
         Left            =   3120
         ScaleHeight     =   960
         ScaleWidth      =   1035
         TabIndex        =   161
         Top             =   225
         Width           =   1035
      End
      Begin VB.PictureBox picData 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Index           =   2
         Left            =   2085
         ScaleHeight     =   960
         ScaleWidth      =   1035
         TabIndex        =   160
         Top             =   225
         Width           =   1035
      End
      Begin VB.PictureBox picData 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   960
         Index           =   1
         Left            =   1050
         ScaleHeight     =   960
         ScaleWidth      =   1035
         TabIndex        =   159
         Top             =   225
         Width           =   1035
      End
      Begin VB.PictureBox picData 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   960
         Index           =   0
         Left            =   15
         ScaleHeight     =   960
         ScaleWidth      =   1035
         TabIndex        =   158
         Top             =   225
         Width           =   1035
      End
      Begin VB.CommandButton cmdSettingsCancel 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         Height          =   345
         Left            =   3645
         TabIndex        =   157
         Tag             =   "2005"
         Top             =   5595
         Visible         =   0   'False
         Width           =   1150
      End
      Begin VB.CommandButton cmdSettingsOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   345
         Left            =   4815
         TabIndex        =   156
         Tag             =   "2004"
         Top             =   5595
         Visible         =   0   'False
         Width           =   1150
      End
      Begin VB.Frame gbDatabase 
         Caption         =   "Database"
         Height          =   1710
         Left            =   165
         TabIndex        =   140
         Tag             =   "2185"
         Top             =   2055
         Visible         =   0   'False
         Width           =   5775
         Begin VB.OptionButton optDBChoice 
            Caption         =   "&Access Database"
            Height          =   255
            Index           =   0
            Left            =   195
            TabIndex        =   73
            Tag             =   "2202"
            Top             =   285
            Width           =   3255
         End
         Begin VB.OptionButton optDBChoice 
            Caption         =   "&ODBC Database"
            Height          =   255
            Index           =   1
            Left            =   195
            TabIndex        =   76
            Tag             =   "2205"
            Top             =   990
            Width           =   3255
         End
         Begin VB.TextBox txtMDBPath 
            Height          =   300
            Left            =   195
            TabIndex        =   74
            Top             =   555
            Width           =   4200
         End
         Begin VB.TextBox txtDSN 
            Height          =   300
            Left            =   195
            TabIndex        =   77
            Top             =   1260
            Width           =   4200
         End
         Begin VB.CommandButton cmdMDBBrowse 
            Caption         =   "&Browse..."
            Height          =   315
            Left            =   4425
            TabIndex        =   75
            Tag             =   "2204"
            Top             =   555
            Width           =   1185
         End
         Begin VB.CommandButton cmdDSNChoose 
            Caption         =   "&Select DSN..."
            Height          =   315
            Left            =   4425
            TabIndex        =   78
            Tag             =   "2210"
            Top             =   1260
            Width           =   1185
         End
      End
      Begin VB.Frame gbNetComm 
         Caption         =   "Network Communications"
         Height          =   2490
         Left            =   165
         TabIndex        =   139
         Tag             =   "2182"
         Top             =   2070
         Visible         =   0   'False
         Width           =   5775
         Begin VB.CommandButton cmdRunRACMgr 
            Caption         =   "Run &RAC Manager..."
            Height          =   345
            Left            =   3840
            TabIndex        =   72
            Tag             =   "2211"
            Top             =   2025
            Width           =   1815
         End
         Begin VB.Label lblNetComm 
            Caption         =   $"frmMain.frx":2948
            Height          =   1470
            Left            =   75
            TabIndex        =   141
            Tag             =   "2212"
            Top             =   375
            Width           =   5580
         End
      End
      Begin VB.Frame gbRemote 
         Caption         =   "Remote Server"
         Height          =   3450
         Left            =   165
         TabIndex        =   137
         Tag             =   "2183"
         Top             =   2040
         Visible         =   0   'False
         Width           =   5775
         Begin VB.ComboBox cbRemoteServer 
            Height          =   315
            Left            =   2835
            Style           =   2  'Dropdown List
            TabIndex        =   151
            Top             =   615
            Width           =   2760
         End
         Begin VB.Label lblRemFuncNA 
            Enabled         =   0   'False
            Height          =   1140
            Left            =   330
            TabIndex        =   153
            Tag             =   "2215"
            Top             =   1350
            Visible         =   0   'False
            Width           =   5235
            WordWrap        =   -1  'True
         End
         Begin VB.Image imgRemoteServer 
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   1920
            Left            =   1755
            Stretch         =   -1  'True
            Top             =   1065
            Visible         =   0   'False
            Width           =   2130
         End
         Begin VB.Label lblRemServer 
            Caption         =   "&Select the required Corporate Benefits ActiveX Server :"
            Height          =   270
            Left            =   165
            TabIndex        =   147
            Tag             =   "2213"
            Top             =   345
            Width           =   5355
         End
      End
      Begin VB.Frame gbRemoteLink 
         Caption         =   "Remote Link Server"
         Height          =   3390
         Left            =   165
         TabIndex        =   138
         Tag             =   "2184"
         Top             =   2100
         Visible         =   0   'False
         Width           =   5775
         Begin VB.ComboBox cbRemoteLinkServer 
            Height          =   315
            Left            =   2835
            Style           =   2  'Dropdown List
            TabIndex        =   152
            Top             =   615
            Width           =   2760
         End
         Begin VB.Image imgRemoteLinkServer 
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   1920
            Left            =   1755
            Stretch         =   -1  'True
            Top             =   1065
            Visible         =   0   'False
            Width           =   2130
         End
         Begin VB.Label lblRemLinkServer 
            Caption         =   "&Select the required Corporate Benefits ActiveX Server :"
            Height          =   270
            Left            =   165
            TabIndex        =   148
            Tag             =   "2213"
            Top             =   345
            Width           =   5355
         End
         Begin VB.Label lblRemLinkFuncNA 
            Enabled         =   0   'False
            Height          =   1140
            Left            =   330
            TabIndex        =   154
            Tag             =   "2215"
            Top             =   1350
            Visible         =   0   'False
            Width           =   5235
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Label lblDataAdminHelp 
         Alignment       =   2  'Center
         Caption         =   "Click on one of the Configuration objects above to change it's setting."
         Height          =   795
         Left            =   120
         TabIndex        =   149
         Tag             =   "2214"
         Top             =   2160
         Width           =   5745
      End
      Begin VB.Label lblDataServices 
         Caption         =   "Database"
         Height          =   615
         Index           =   5
         Left            =   5265
         TabIndex        =   146
         Tag             =   "2185"
         Top             =   1335
         Width           =   690
      End
      Begin VB.Label lblDataServices 
         Alignment       =   2  'Center
         Height          =   615
         Index           =   4
         Left            =   4245
         TabIndex        =   145
         Tag             =   "2184"
         Top             =   1335
         Width           =   1065
      End
      Begin VB.Label lblDataServices 
         Alignment       =   2  'Center
         Height          =   615
         Index           =   3
         Left            =   3180
         TabIndex        =   144
         Tag             =   "2183"
         Top             =   1335
         Width           =   1065
      End
      Begin VB.Label lblDataServices 
         Alignment       =   2  'Center
         Height          =   615
         Index           =   2
         Left            =   2115
         TabIndex        =   143
         Tag             =   "2182"
         Top             =   1335
         Width           =   1065
      End
      Begin VB.Label lblDataServices 
         Alignment       =   2  'Center
         Height          =   615
         Index           =   1
         Left            =   1050
         TabIndex        =   142
         Tag             =   "2181"
         Top             =   1335
         Width           =   1065
      End
      Begin VB.Line lineBlack 
         X1              =   17
         X2              =   390
         Y1              =   130
         Y2              =   130
      End
      Begin VB.Line lineWhite 
         BorderColor     =   &H80000014&
         X1              =   17
         X2              =   390
         Y1              =   131
         Y2              =   131
      End
      Begin VB.Label lblDataServices 
         Alignment       =   2  'Center
         Caption         =   "Front-end Application"
         Height          =   615
         Index           =   0
         Left            =   -15
         TabIndex        =   136
         Tag             =   "2180"
         Top             =   1335
         Width           =   1065
      End
      Begin ComctlLib.ImageList ilDataServices 
         Left            =   300
         Top             =   2190
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   71
         ImageHeight     =   64
         MaskColor       =   12632256
         _Version        =   327680
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   12
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":2A96
               Key             =   "VB Application"
               Object.Tag             =   "VB Application"
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":3428
               Key             =   "Passthrough"
               Object.Tag             =   "Passthrough"
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":3DBA
               Key             =   "C++ Server DLL"
               Object.Tag             =   "C++ Server DLL"
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":474C
               Key             =   "C++ Server EXE"
               Object.Tag             =   "C++ Server EXE"
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":50DE
               Key             =   "VB Server DLL"
               Object.Tag             =   "VB Server DLL"
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":5A70
               Key             =   "VB Server EXE"
               Object.Tag             =   "VB Server EXE"
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":6402
               Key             =   "ATL Server DLL"
               Object.Tag             =   "ATL Server DLL"
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":6D94
               Key             =   "ATL Server EXE"
               Object.Tag             =   "ATL Server EXE"
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":7726
               Key             =   "MFC Server DLL"
               Object.Tag             =   "MFC Server DLL"
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":80B8
               Key             =   "MFC Server EXE"
               Object.Tag             =   "MFC Server EXE"
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":8A4A
               Key             =   "Network Communications"
               Object.Tag             =   "Network Communications"
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmMain.frx":93DC
               Key             =   "Database"
               Object.Tag             =   "Database"
            EndProperty
         EndProperty
      End
      Begin VB.Label lblDataServicesTitle 
         Alignment       =   2  'Center
         Caption         =   "Current System Configuration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   45
         TabIndex        =   123
         Tag             =   "2179"
         Top             =   -15
         Width           =   5835
      End
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   1125
      Left            =   0
      Picture         =   "frmMain.frx":9D6E
      Top             =   0
      Width           =   12000
   End
   Begin ComctlLib.ImageList ilSideBar 
      Left            =   240
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1061A
            Key             =   "Data Services"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":108AC
            Key             =   "Banks"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":10B3E
            Key             =   "Employees"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":10DD0
            Key             =   "Queries"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":11062
            Key             =   "Dependants"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":112F4
            Key             =   "Benefits"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Visual Studio '97  - Corporate Benefits Sample
'frmMain.frm        - Copyright (C) 1997 Microsoft Corporation
'                   - All rights reserved.
'Date               - 10/01/97
'Authors            - Simon J. Poile
'Purpose            - Main Form for the Corporate Benefits Client Application

Option Explicit

Private oCurrentEmployee        As CEmployee        'Keep a record of the current employee
Private cDataOptions            As Collection       'Data Options frames collection
Private cScreens                As Collection       'Collection of the screens offered by the application
Private bChanged                As Boolean
Private bServers(stTotalServers) As Boolean         'Data Options frames collection
Private giQueryGender           As Integer          'Flag as to which gender is being queried, if any

'BanksFieldsEnable
'Purpose     - Enable bank fields
'Inputs      - (Boolean)    bEnable
Private Sub BankFieldsEnable(ByVal bEnable As Boolean)
txtBankName.Enabled = bEnable
txtBankCode.Enabled = bEnable
txtBankAddress1.Enabled = bEnable
txtBankAddress2.Enabled = bEnable
txtBankCity.Enabled = bEnable
End Sub

'BenefitFieldsEnable
'Purpose     - Enable benefit fields
'Inputs      - (Boolean)    bEnable
Private Sub BenefitFieldsEnable(ByVal bEnable As Boolean)
txtBenefitDescription.Enabled = bEnable
cbBenefitType.Enabled = bEnable
txtBenefitCost.Enabled = bEnable
txtBenefitMinimum.Enabled = bEnable
txtBenefitMaximum.Enabled = bEnable
txtBenefitIncrement.Enabled = bEnable
txtBenefitMultiples.Enabled = bEnable
cbBenefitTaxStatus.Enabled = bEnable
End Sub

'DoConfigure
'Purpose     - Setup screen to accept changes
'Outputs     - (Boolean)    Success
Public Function DoConfigure() As Boolean
Dim ctlTemp As Control
Dim lHeight As Long
Dim lWidth  As Long
Dim sCap    As String

'Reset the look of the main form,
'just for taking the initial settings
For Each ctlTemp In cScreens
    ctlTemp.Enabled = False
    ctlTemp.Visible = False
Next ctlTemp

sCap = Caption
lHeight = Height
lWidth = Width
Caption = LoadResString(es.InitSetup)
statMain.Visible = False
sbCategories.Visible = False
pctDataServices.Enabled = True
pctDataServices.Visible = True
pctDataServices.Top = 1125 '0
pctDataServices.Left = 0
pctDataServices.ZOrder
optDBChoice(1) = True
cmdSettingsOK.Enabled = False
cmdSettingsCancel.Enabled = True
cmdSettingsOK.Visible = True
cmdSettingsCancel.Visible = True
Height = pctDataServices.Height + imgBanner.Height + (Height - ScaleHeight)
Width = pctDataServices.Width + (Width - ScaleWidth)

tSet.lSrc = icDoDataDirect
tSet.bChanged = False
DoPaint

'Display Settings screen modally
Show vbModal

'Reset screens back to normal
Caption = sCap
Height = lHeight
Width = lWidth
statMain.Visible = True
sbCategories.Visible = True
cmdSettingsOK.Enabled = False
cmdSettingsCancel.Enabled = False
cmdSettingsOK.Visible = False
cmdSettingsCancel.Visible = False

For Each ctlTemp In cDataOptions
    ctlTemp.Enabled = False
    ctlTemp.Visible = False
Next ctlTemp

pctDataServices.Visible = False
pctDataServices.Enabled = False
pctDataServices.ZOrder 1
If CBool(Len(tSet.sType)) And CBool(Len(tSet.sDB)) Then
    DoConfigure = True
Else
    DoConfigure = False
End If
End Function

'DoPaint
'Purpose     - Main data services image map paint routine
Private Sub DoPaint()
If tSet.lRem Then
    'Paint the network communications bitmap
    'if a Remote Server is selected
    picData(2).Picture = ilDataServices.ListImages("Network Communications").Picture
    picData(4).Enabled = True
Else
    picData(2).Picture = ilDataServices.ListImages("Passthrough").Picture
    picData(4).Enabled = False
End If
lblDataServices(2).Caption = LoadResString(es.DSNetComm)
'Paint the relevant Remote Server bitmap
picData(3).Picture = ilDataServices.ListImages(tSet.lRem + 2).Picture
lblDataServices(3).Caption = LoadResString(es.DOPASSTHROUGH + tSet.lRem)
'Paint the relevant Remote Link Server bitmap
picData(4).Picture = ilDataServices.ListImages(tSet.lLnk + 2).Picture
lblDataServices(4).Caption = LoadResString(es.DOPASSTHROUGH + tSet.lLnk)
imgRemoteLinkServer.Picture = ilDataServices.ListImages(tSet.lLnk + 2).Picture
imgRemoteServer.Picture = ilDataServices.ListImages(tSet.lRem + 2).Picture
lblDataServices(5).Caption = tSet.sType & vbCrLf & tSet.sDB
End Sub

'EmployeeFieldsEnable
'Purpose     - Enable employee fields
'Inputs      - (Boolean)    bEnable
Private Sub EmployeeFieldsEnable(ByVal bEnable As Boolean)
txtEmplFirstName.Enabled = bEnable
txtEmplLastName.Enabled = bEnable
txtEmplPosition.Enabled = bEnable
txtEmplAddress1.Enabled = bEnable
txtEmplAddress2.Enabled = bEnable
txtEmplCity.Enabled = bEnable
txtEmplUserName.Enabled = bEnable
txtEmplSalary.Enabled = bEnable
txtEmplDOB.Enabled = bEnable
optEmplGenderMale.Enabled = bEnable
optEmplGenderFemale.Enabled = bEnable
txtEmplStartDate.Enabled = bEnable
txtEmplEndDate.Enabled = bEnable
cbEmplBanks.Enabled = bEnable
txtEmplBankAccount.Enabled = bEnable
End Sub

'Init
'Purpose     - Main form load routine
'Outputs     - (Boolean)    Success
Public Function Init() As Boolean
'Objects used to get initial Admin data
Dim oBank           As CBank
Dim oBenefit        As CBenefit
Dim oEmployee       As CEmployee

'Generic CLists used to fill list-boxes & combo-boxes
Dim lstFull         As New CList

Dim bRetVal         As Boolean
Dim hKey            As Long
Dim lRetVal         As Long

'Strings to hold localizable string resources
Dim sAdmin          As String
Dim sDataAdmin      As String
Dim sDataServices   As String
Dim sStatistics     As String

Dim sBanks          As String
Dim sBenefits       As String
Dim sDependants     As String
Dim sEmployees      As String
Dim sQueries        As String
Dim sResults        As String

Screen.MousePointer = vbHourglass

Height = 7995
Width = 7320

pctQuery.Top = 1125
pctQuery.Left = 1140

'Load Screens localized text strings
sAdmin = LoadResString(es.Admin)
sDataAdmin = LoadResString(es.DataAdmin)
sDataServices = LoadResString(es.DataServices)
sStatistics = LoadResString(es.Statistics)

sBanks = LoadResString(es.Banks)
sBenefits = LoadResString(es.Benefits)
sDependants = LoadResString(es.Dependants)
sEmployees = LoadResString(es.Employees)
sQueries = LoadResString(es.Queries)
sResults = LoadResString(es.Results)

'Set the Query Genders field initial value
giQueryGender = 0
    
'Setup sidebar control
With sbCategories
    .AddFolder sStatistics, sStatistics
    .Folders(sStatistics).AddItem sQueries, sQueries, sQueries
    .Folders(sStatistics).DefaultItem = sQueries
    'Don't add Folders & Items if User Rights don't allow it
    If iRights = icValidAdmin Then
        'User is Admin so, add Admin functionaloity to SideBar
        .AddFolder sAdmin, sAdmin
        .Folders(sAdmin).AddItem sBenefits, sBenefits, sBenefits
        .Folders(sAdmin).AddItem sEmployees, sEmployees, sEmployees
        .Folders(sAdmin).AddItem sDependants, sDependants, sDependants
        .Folders(sAdmin).AddItem sBanks, sBanks, sBanks
        .Folders(sAdmin).DefaultItem = sBenefits
        .AddFolder sDataAdmin, sDataAdmin
        .Folders(sDataAdmin).AddItem sDataServices, sDataServices, sDataServices
        .Folders(sDataAdmin).DefaultItem = sDataServices
    End If
End With

'Setup of the Query Screen

On Error GoTo errInit

oAdmin.FillListBenefits

'Fill Availble Benefits combobox
cbQueryBenefits.AddItem LoadResString(es.qryAllBenefits)
cbQueryBenefits.ItemData(cbQueryBenefits.NewIndex) = 0
BenefitsToControl cbQueryBenefits
cbQueryBenefits.ListIndex = 0

'Fill Benefit Types/Categories combobox
cbQueryBenefitTypes.AddItem LoadResString(es.qryAllBenefitTypes)
cbQueryBenefitTypes.ItemData(cbQueryBenefitTypes.NewIndex) = 0

If oAdmin.GetBenefitTypes(lstFull) Then
    DomainListToControl lstFull, cbQueryBenefitTypes
    cbQueryBenefitTypes.ListIndex = 0
Else
    MsgError es.BenefitsErrorTitle, es.DBBenTypesFailed
End If

'Fill Coverage combobox
cbQueryCoverage.AddItem LoadResString(es.qryIgnoreCoverage)
cbQueryCoverage.ItemData(cbQueryCoverage.NewIndex) = 0

If oAdmin.GetEmployeeBenefitCoverageTypes(lstFull) Then
    DomainListToControl lstFull, cbQueryCoverage
    cbQueryCoverage.ListIndex = 0
Else
    MsgError es.BenefitsErrorTitle, es.DBCovTypesFailed
End If

'Fill Tax combobox
cbQueryTax.AddItem LoadResString(es.qryPaymentOfTax)
cbQueryTax.ItemData(cbQueryTax.NewIndex) = 0

If oAdmin.GetEmployeeBenefitTaxTypes(lstFull) Then
    DomainListToControl lstFull, cbQueryTax
    cbQueryTax.ListIndex = 0
Else
    MsgError es.BenefitsErrorTitle, es.DBQueryTaxTypesFailed
End If

'Only setup the following screens if the current user has Admin rights
If iRights = icValidAdmin Then
    'Set up of the Employees Screen
    'Fill All Employees & All Banks List
    oAdmin.FillListEmployees
    oAdmin.FillListBanks
    
    'Setup of the Dependants Screen
    'Fill Relationship combobox
    If oAdmin.GetDependantTypes(lstFull) Then
        DomainListToControl lstFull, cbDepRelationship
    Else
        MsgError es.BenefitsErrorTitle, es.DBDepRelFailed
    End If
        
    'Setup of the Benefits Screen
    'Fill Categories combobox
    If oAdmin.GetBenefitTypes(lstFull) Then
        DomainListToControl lstFull, cbBenefitType
    Else
        MsgError es.BenefitsErrorTitle, es.DBBenTypesFailed
    End If
        
    'Fill Tax combobox
    If oAdmin.GetBenefitTaxTypes(lstFull) Then
        DomainListToControl lstFull, cbBenefitTaxStatus
    Else
        MsgError es.BenefitsErrorTitle, es.DBBenTaxTypesFailed
    End If
    
    'Disable editable fields
    BankFieldsEnable (False)
    BenefitFieldsEnable (False)
    DependantFieldsEnable (False)
    EmployeeFieldsEnable (False)
End If

'Load MS Graph OLE Prog Id for later use
sGraphPID = LoadResString(es.MSGraphProgID)

'Setup the UI Strings (for localization)
LoadScreenResStrings
    
'Check whether app is going to be able to use MS Graph
lRetVal = RegOpenKeyEx(HKEY_CLASSES_ROOT, sGraphPID, 0, KEY_QUERY_VALUE, hKey)
If lRetVal = ERROR_NONE Then
    'MS Graph IS installed
    bUseMSGraph = True
    'Unload Key as we don't need it
    RegCloseKey (hKey)
Else
    'MS Graph IS NOT installed
    bUseMSGraph = False
End If

'Set settings changed value to False for a clean start

Screen.MousePointer = vbDefault
Init = True
Exit Function

errInit:
    If Err > vbObjectError Then
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err - vbObjectError) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    Else
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    End If

End Function

'LoadScreenResStrings
'Purpose     - Load all/any localized string resources
'            - for each control on the form
Sub LoadScreenResStrings()
On Error Resume Next
    
Dim ctl As Control
Dim obj As Object
    
'Set the form's caption
Me.Caption = LoadResString(CInt(Me.Tag))
    
'Set the controls'captions using the caption
'property for menu items and the Tag property
'for all other controls
For Each ctl In Me.Controls
    Err.Clear
    If TypeName(ctl) = "Menu" Then
        If IsNumeric(ctl.Caption) Then
            If Err = 0 Then
                ctl.Caption = LoadResString(CInt(ctl.Caption))
            End If
        End If
    ElseIf TypeName(ctl) = "PictureBox" Then
        'Don't do anything
    Else
        If IsNumeric(ctl.Tag) Then
            If Err = 0 Then
            ctl.Caption = LoadResString(CInt(ctl.Tag))
            End If
        End If
        'Check for a tooltip
        If IsNumeric(ctl.ToolTipText) Then
            If Err = 0 Then
                ctl.ToolTipText = LoadResString(CInt(ctl.ToolTipText))
            End If
        End If
    End If
Next
End Sub

'UpdateBankData
'Purpose     - Load, save or clear the fields for a Bank object
'Inputs      - (Integer)    iUpdateFlag
'            - (CBank)      Optional oBank
Private Sub UpdateBankData(iUpdateFlag As Integer, Optional oBank As CBank)
On Error GoTo errUpdateBankData

Select Case (iUpdateFlag)
Case eUpdate.emClearFields:
    'Clear all field values
    txtBankName = ""
    txtBankCode = ""
    txtBankAddress1 = ""
    txtBankAddress2 = ""
    txtBankCity = ""
Case eUpdate.emFieldsToData
    With oBank
        .Name = txtBankName
        .Code = txtBankCode
        .Address1 = txtBankAddress1
        .Address2 = txtBankAddress2
        .City = txtBankCity
    End With
Case eUpdate.emDataToFields
    With oBank
        txtBankName = .Name
        txtBankCode = .Code
        txtBankAddress1 = .Address1
        txtBankAddress2 = .Address2
        txtBankCity = .City
    End With
End Select
Exit Sub

errUpdateBankData:
    If Err > vbObjectError Then
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err - vbObjectError) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    Else
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    End If
    Resume Next
End Sub

'UpdateBenefitData
'Purpose     - Load, save or clear the fields for a benefit object
'Inputs      - (Integer)    iUpdateFlag
'            - (CBenefit)   Optional oBenefit
Private Sub UpdateBenefitData(iUpdateFlag As Integer, Optional oBenefit As CBenefit)
Dim iCount As Integer
On Error GoTo errUpdateBenefitData
    
Select Case (iUpdateFlag)
Case eUpdate.emClearFields:
    txtBenefitDescription = ""
    cbBenefitType.ListIndex = 0
    txtBenefitCost = ""
    txtBenefitMinimum = ""
    txtBenefitMaximum = ""
    txtBenefitIncrement = ""
    txtBenefitMultiples = ""
    cbBenefitTaxStatus.ListIndex = 0
Case eUpdate.emFieldsToData
    With oBenefit
        .Description = txtBenefitDescription
        .Cost = CLng(IIf(txtBenefitCost = "", "0", txtBenefitCost))
        .Minimum = CLng(IIf(txtBenefitMinimum = "", "0", txtBenefitMinimum))
        .Maximum = CLng(IIf(txtBenefitMaximum = "", "0", txtBenefitMaximum))
        .Increment = CLng(IIf(txtBenefitIncrement = "", "0", txtBenefitIncrement))
        .Multiples = CLng(IIf(txtBenefitMultiples = "", "0", txtBenefitMultiples))
        .BnfTypeId = cbBenefitType.ItemData(cbBenefitType.ListIndex)
        .BnfTaxId = cbBenefitTaxStatus.ItemData(cbBenefitTaxStatus.ListIndex)
    End With
Case eUpdate.emDataToFields:
    With oBenefit
        txtBenefitDescription = .Description
        txtBenefitCost = .Cost
        txtBenefitMinimum = .Minimum
        txtBenefitMaximum = .Maximum
        txtBenefitIncrement = .Increment
        txtBenefitMultiples = .Multiples
        'Find the correct entry in the Type Combobox
        FindItem cbBenefitType, .BnfTypeId
        'Find the correct entry in the Tax Status Combobox
        FindItem cbBenefitTaxStatus, .BnfTaxId
    End With
End Select
Exit Sub
    
errUpdateBenefitData:
    If Err > vbObjectError Then
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err - vbObjectError) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    Else
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    End If
    Resume Next
End Sub

'UpdateDependantData
'Purpose     - Load, save or clear the fields for a dependant object
'Inputs      - (Integer)    iUpdateFlag
'            - (CDependant) Optional oDependant
Private Sub UpdateDependantData(iUpdateFlag As Integer, Optional oDependant As CDependant)
Dim iCount As Integer
On Error GoTo errUpdateDependantData

Select Case (iUpdateFlag)
Case eUpdate.emClearFields:
    txtDepFirstName = ""
    txtDepLastName = ""
    txtDepDOB = ""
    optDepGenderMale.Value = True
    optDepStatusDependant.Value = True
    cbDepRelationship.ListIndex = -1
Case eUpdate.emFieldsToData:
    With oDependant
        .FirstName = txtDepFirstName
        .LastName = txtDepLastName
        If (txtDepDOB <> "") Then
            .DateOfBirth = CDate(txtDepDOB)
        End If
        .Gender = optDepGenderMale
        .Status = optDepStatusDependant
        .DepTypeId = cbDepRelationship.ItemData(cbDepRelationship.ListIndex)
        Set .Employee = oCurrentEmployee
    End With
Case eUpdate.emDataToFields:
        With oDependant
            txtDepFirstName = .FirstName
            txtDepLastName = .LastName
            If (DatePart("yyyy", .DateOfBirth) = "1899") Then
                txtDepDOB = ""
            Else
                txtDepDOB = .DateOfBirth
            End If
            If (.Gender = True) Then
                optDepGenderFemale = True
            Else
                optDepGenderMale = True
            End If
            If (.Status = True) Then
                optDepStatusDependant = True
            Else
                optDepStatusBeneficiary = True
            End If
            'Find the correct entry in the Tax Status Combobox
            FindItem cbDepRelationship, .DepTypeId
        End With
End Select
Exit Sub

errUpdateDependantData:
    If Err > vbObjectError Then
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err - vbObjectError) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    Else
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    End If
    Resume Next
End Sub

'UpdateEmployeeData
'Purpose     - Load, save or clear the fields for a employee object
'Inputs      - (Integer)    Integer
'            - (CEmployee)  Optional oEmployee
Private Sub UpdateEmployeeData(iUpdateFlag As Integer, Optional oEmployee As CEmployee)
Dim iCount As Integer
On Error GoTo errUpdateEmployeeData

Select Case (iUpdateFlag)
Case eUpdate.emClearFields:
    txtEmplFirstName = ""
    txtEmplLastName = ""
    txtEmplPosition = ""
    txtEmplAddress1 = ""
    txtEmplAddress2 = ""
    txtEmplCity = ""
    txtEmplUserName = ""
    txtEmplSalary = ""
    txtEmplDOB = ""
    optEmplGenderMale = True
    txtEmplStartDate = ""
    txtEmplEndDate = ""
    cbEmplBanks.ListIndex = 0
    txtEmplBankAccount = ""
Case eUpdate.emFieldsToData:
    With oEmployee
        .FirstName = txtEmplFirstName
        .LastName = txtEmplLastName
        .Position = txtEmplPosition
        .Address1 = txtEmplAddress1
        .Address2 = txtEmplAddress2
        .City = txtEmplCity
        .UserName = txtEmplUserName
        .Salary = CCur(IIf(txtEmplSalary = "", "0", txtEmplSalary))
        If (txtEmplDOB <> "") Then
            .DateOfBirth = CDate(txtEmplDOB)
        End If
        .Gender = optEmplGenderFemale
        If (txtEmplStartDate <> "") Then
            .StartDate = CDate(txtEmplStartDate)
        End If
        If (txtEmplEndDate <> "") Then
            .EndDate = CDate(txtEmplEndDate)
        End If
        If cbEmplBanks.ListIndex > 0 Then
            Set .Bank = oAdmin.Banks(cbEmplBanks.ItemData(cbEmplBanks.ListIndex))
        Else
            Set .Bank = Nothing
        End If
        .BankAccount = txtEmplBankAccount
    End With
Case eUpdate.emDataToFields:
    With oEmployee
        txtEmplFirstName = .FirstName
        txtEmplLastName = .LastName
        txtEmplPosition = .Position
        txtEmplAddress1 = .Address1
        txtEmplAddress2 = .Address2
        txtEmplCity = .City
        txtEmplUserName = .UserName
        txtEmplSalary = .Salary
        If (DatePart("yyyy", .DateOfBirth) = "1899") Then
            txtEmplDOB = ""
        Else
            txtEmplDOB = .DateOfBirth
        End If
        If (.Gender = True) Then
            optEmplGenderFemale = True
        Else
            optEmplGenderMale = True
        End If
        If (DatePart("yyyy", .StartDate) = "1899") Then
            txtEmplStartDate = ""
        Else
            txtEmplStartDate = .StartDate
        End If
        If (DatePart("yyyy", .EndDate) = "1899") Then
            txtEmplEndDate = ""
        Else
            txtEmplEndDate = .EndDate
        End If
            FindItem cbEmplBanks, .Bank.PKId
        txtEmplBankAccount = .BankAccount
    End With
End Select
Exit Sub

errUpdateEmployeeData:
    If Err > vbObjectError Then
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err - vbObjectError) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    Else
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    End If
    Resume Next
End Sub

Private Sub cbAllEmployees_Click()
'Click Event for the list of All Employees
Dim iCount          As Integer
Dim lEmployeePKId   As Long

'Is an employee selected
If cbAllEmployees.ListIndex = -1 Then
    'Clear all field values
    UpdateEmployeeData eUpdate.emClearFields
    cmdEditEmployee.Enabled = False
    cmdDeleteEmployee.Enabled = False
    Exit Sub
Else
    cmdEditEmployee.Enabled = True
    cmdDeleteEmployee.Enabled = True
End If

'Get the selected Employees ID
lEmployeePKId = cbAllEmployees.ItemData(cbAllEmployees.ListIndex)

'Get the Employees object from it's ID
Set oCurrentEmployee = oAdmin.Employees(lEmployeePKId)

'Update all the fields to reflect the Employees values
UpdateEmployeeData eUpdate.emDataToFields, oCurrentEmployee
End Sub

Private Sub cbQueryBenefitTypes_Click()
'Reset Available Benefits List with
'only those Benefits of this type/category
Dim oBenTemp As New CBenefit

'Don't bother resetting if nothing is selected
    If cbQueryBenefitTypes.ListIndex = -1 Then
        Exit Sub
    End If
    
'Clear the combo
    cbQueryBenefits.Clear
'Add "All Benefits"
    cbQueryBenefits.AddItem LoadResString(es.qryAllBenefits)

    On Error GoTo ErrHandler

'Add each benefit from the database
    For Each oBenTemp In oAdmin.Benefits
        If cbQueryBenefitTypes.ListIndex = 0 Or oBenTemp.BnfTypeId = cbQueryBenefitTypes.ItemData(cbQueryBenefitTypes.ListIndex) Then
            cbQueryBenefits.AddItem oBenTemp.Description
            cbQueryBenefits.ItemData(cbQueryBenefits.NewIndex) = oBenTemp.PKId
        End If
    Next
    
    cbQueryBenefits.ListIndex = 0
    
    Exit Sub

ErrHandler:
    If Err > vbObjectError Then
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err - vbObjectError) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    Else
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    End If
    Exit Sub
End Sub

Private Sub cbRemoteLinkServer_Click()
'Handles selection of a Remote Link Server configuration
Dim rct As RECT

'Only do the following if this screen is visible
If gbRemoteLink.Visible Then
    'Update the Remote Link setting with the new selection
    tSet.lLnk = cbRemoteLinkServer.ItemData(cbRemoteLinkServer.ListIndex)
    If tSet.lLnk = stPassthrough Then
        If tSet.lRem Then
            tSet.lSrc = icDoDataRemote
        Else
            tSet.lSrc = icDoDataDirect
        End If
    Else
        tSet.lSrc = icDoDataLinked
    End If
    'Update the graphical representation
    DoPaint
End If
End Sub

Private Sub cbRemoteServer_Click()
'Handles selection of a Remote Server configuration
Dim rct As RECT

'Only do the following if this screen is visible
If gbRemote.Visible Then
    'Update the Remote setting with the new selection
    tSet.lRem = cbRemoteServer.ItemData(cbRemoteServer.ListIndex)
    If tSet.lRem Then
        tSet.lSrc = icDoDataRemote
    Else
        tSet.lSrc = icDoDataDirect
        tSet.lLnk = 0
    End If
    'Update the graphical representation
    DoPaint
End If
End Sub

Private Sub cmdAddBenefit_Click()
'Add a new Benefit

'De-select and disable any item in the Benefits List
lstAllBenefits.ListIndex = -1
lstAllBenefits.Enabled = False
    
'Disable sidebar & other buttons
sbCategories.Enabled = False
cmdAddBenefit.Enabled = False
cmdEditBenefit.Enabled = False
cmdDeleteBenefit.Enabled = False

'Enable editable fields
BenefitFieldsEnable (True)

'Enable OK & Cancel
cmdBenefitOK.Enabled = True
cmdBenefitOK.Default = True
cmdBenefitCancel.Enabled = True
    
'Set focus to first editable field
txtBenefitDescription.SetFocus
End Sub

Private Sub cmdBenefitGlossary_Click()
'View the glossary of terms used in this benefits application
Dim bValidFile      As Boolean
Dim oIE             As Object
Dim sGlossaryPath   As String

Screen.MousePointer = vbHourglass

'Get the file name
sGlossaryPath = CurDir$ & LoadResString(es.DirDelimiter) & LoadResString(es.BenGlossaryAddress)
    
'Check that the file exists
On Error GoTo NoFile
bValidFile = True
Open sGlossaryPath For Input As #1
Close #1
On Error GoTo 0
    
If (bValidFile = True) Then
    'If the file exists, the app trys to open the HTML document
    'using OLE Automation & Internet Explorer
    Set oIE = CreateObject("InternetExplorer.Application")
    If IsObject(oIE) Then
        'Internet Explorer was successfully loaded - now display the document
        oIE.Visible = True
        oIE.navigate "file://" & sGlossaryPath
    Else
        'Internet Explorer could not be found on this system -
        'therefore try to load whatever application is associated with .HTML
        ShellExecute Me.hwnd, 0, sGlossaryPath, "", "", SW_SHOWNORMAL
    End If
    Screen.MousePointer = vbDefault
Else
    Screen.MousePointer = vbDefault
    'If the file does not exist
    MsgError es.BenefitsErrorTitle, es.BenGlossaryNotFound
End If
    
Exit Sub

NoFile:
    bValidFile = False
    Resume Next
End Sub

Private Sub cmdDSNChoose_Click()
Dim sDBS As String
Dim sDSN As String
Dim sPWD As String
Dim sUID As String

On Error Resume Next
If GetDSN(sDSN, sUID, sPWD, sDBS) Then
    txtDSN = sDSN
    ' Enable the OK button if configuring for the first time
    If Not sbCategories.Visible Then cmdSettingsOK.Enabled = True
End If
End Sub

Private Sub cmdEditBenefit_Click()
'Edit the currently selected Benefit

'Disable the Benefits List
lstAllBenefits.Enabled = False

'Disable sidebar & other buttons
sbCategories.Enabled = False
cmdAddBenefit.Enabled = False
cmdEditBenefit.Enabled = False
cmdDeleteBenefit.Enabled = False

'Enable editable fields
BenefitFieldsEnable (True)

'Enable OK & Cancel
cmdBenefitOK.Enabled = True
cmdBenefitOK.Default = True
cmdBenefitCancel.Enabled = True
   
'Set focus to first editable field
txtBenefitDescription.SetFocus
End Sub

Private Sub cmdDeleteBenefit_Click()
'Delete the currently selected Benefit
Dim oBen            As CBenefit
Dim mbResult        As VbMsgBoxResult
Dim bRetVal         As Boolean
Dim sDeleteBenefit  As String
        
'Is a benefit selected ?
If lstAllBenefits.ListIndex = -1 Then
    MsgError es.BenefitsErrorTitle, es.NoBenefitErrorText
    Exit Sub
End If
        
'Confirm deletion of benefit
sDeleteBenefit = LoadResString(es.DeleteBenefit)
mbResult = MsgBox(sDeleteBenefit & lstAllBenefits.Text, vbQuestion Or vbYesNo Or vbDefaultButton2, es.AppName)
    
If mbResult = vbNo Then
    'Answer was no so end without deleting benefit
    Exit Sub
Else
    Set oBen = oAdmin.Benefits(lstAllBenefits.ItemData(lstAllBenefits.ListIndex))
End If
        
On Error GoTo ErrHandler

bRetVal = oAdmin.Delete(oBenefit:=oBen)
    
If bRetVal = True Then
    'Benefit successfully deleted.
    'Now we need to delete the entry from the listbox
    lstAllBenefits.RemoveItem lstAllBenefits.ListIndex
    lstAllBenefits.ListIndex = -1
    UpdateBenefitData eUpdate.emClearFields
Else
    'Failed to delete benefit
    sDeleteBenefit = LoadResString(es.DbFailedDeleteBenefit)
    MsgBox sDeleteBenefit & lstAllBenefits.Text, vbExclamation, es.BenefitsErrorTitle
End If
Exit Sub
    
ErrHandler:
    If Err > vbObjectError Then
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err - vbObjectError) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    Else
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    End If
    bRetVal = False
    Resume Next
End Sub

Private Sub cmdBenefitCancel_Click()
'User selected Cancel to either an Add or Edit of a Benefit

'Re-enable the Benefits list
lstAllBenefits.Enabled = True
lstAllBenefits_Click
    
'Re-enable sidebar & other buttons
sbCategories.Enabled = True
cmdAddBenefit.Enabled = True

'Disable editable fields
BenefitFieldsEnable (False)

'Disable OK & Cancel
cmdBenefitOK.Enabled = False
cmdBenefitCancel.Enabled = False
End Sub

Private Sub cmdBenefitOK_Click()
'User selected OK to either an Add or Edit of a Benefit
Dim oBen        As CBenefit
Dim oBenTemp    As CBenefit
Dim bRetVal     As Boolean
Dim iCurSel     As Integer
    
On Error GoTo ErrHandler

Set oBen = New CBenefit
If (lstAllBenefits.ListIndex = -1) Then
    'No currently selected item in the Benefits List
    'This means that we must be adding a new benefit
    'Load the new values into the object
    UpdateBenefitData eUpdate.emFieldsToData, oBen
        
    bRetVal = oAdmin.Insert(oBenefit:=oBen)
    
    'Update the Benefits list
    lstAllBenefits.Clear
    BenefitsToControl lstAllBenefits
    
    If bRetVal Then
        FindItemByName lstAllBenefits, txtBenefitDescription
    Else
        MsgError es.BenefitsErrorTitle, es.DbFailAddBenefit
    End If
Else
    'An item is currently selected in the Benefits List
    'This means that we must be editing an existing benefit
    
    'Get the actual Benefit object that we are editing
    Set oBen = oAdmin.Benefits(lstAllBenefits.ItemData(lstAllBenefits.ListIndex))
        
    'Load the edited values into the object
    UpdateBenefitData eUpdate.emFieldsToData, oBen
        
    'Update the Db Record of this benefit itself
    bRetVal = oAdmin.Update(oBenefit:=oBen)
    
    oAdmin.FillListBanks
    
    'Recheck the list
    lstAllBenefits.Clear
    BenefitsToControl lstAllBenefits
    
    If bRetVal Then
        FindItemByName lstAllBenefits, txtBenefitDescription
    Else
        MsgError es.BenefitsErrorTitle, es.DbFailEditBenefit
    End If
End If

'Re-enable the Benefits list
lstAllBenefits.Enabled = True

'Re-enable sidebar & other buttons
sbCategories.Enabled = True
cmdAddBenefit.Enabled = True
cmdEditBenefit.Enabled = True
cmdDeleteBenefit.Enabled = True

'Disable editable fields
BenefitFieldsEnable (False)

'Disable OK & Cancel
cmdBenefitOK.Enabled = False
cmdBenefitCancel.Enabled = False
Exit Sub
    
ErrHandler:
    If Err > vbObjectError Then
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err - vbObjectError) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    Else
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    End If
    bRetVal = False
    UpdateBenefitData eUpdate.emClearFields
    Resume Next
End Sub

Private Sub cmdAddBank_Click()
'Add a new Bank

'De-select and disable any item in the Benefits List
lstAllBanks.ListIndex = -1
lstAllBanks.Enabled = False
    
'Disable sidebar & other buttons
sbCategories.Enabled = False
cmdAddBank.Enabled = False
cmdEditBank.Enabled = False
cmdDeleteBank.Enabled = False

'Enable editable fields
BankFieldsEnable (True)

'Enable OK & Cancel
cmdBankOK.Enabled = True
cmdBankOK.Default = True
cmdBankCancel.Enabled = True
    
'Set focus to first editable field
txtBankName.SetFocus
End Sub

Private Sub cmdEditBank_Click()
'Edit the current Bank

'Disable the currently selected item in the Benefits List
lstAllBanks.Enabled = False

'Disable sidebar & other buttons
sbCategories.Enabled = False
cmdAddBank.Enabled = False
cmdEditBank.Enabled = False
cmdDeleteBank.Enabled = False

'Enable editable fields
BankFieldsEnable (True)

'Enable OK & Cancel
cmdBankOK.Enabled = True
cmdBankOK.Default = True
cmdBankCancel.Enabled = True
    
'Set focus to first editable field
txtBankName.SetFocus
End Sub

Private Sub cmdDeleteBank_Click()
'Delete the current Bank
Dim oBank       As CBank
Dim oEmplTemp   As CEmployee
Dim mbResult    As VbMsgBoxResult
Dim bRetVal     As Boolean
Dim sDeleteBank As String

'Is a bank selected ?
If lstAllBanks.ListIndex = -1 Then
    MsgError es.BenefitsErrorTitle, es.NoBankErrorText
    Exit Sub
Else
    Set oBank = oAdmin.Banks(lstAllBanks.ItemData(lstAllBanks.ListIndex))
End If
        
'Is the Bank currently used by anyone ?
For Each oEmplTemp In oAdmin.Employees
    If (oEmplTemp.Bank.PKId = oBank.PKId) Then
        MsgError es.BenefitsErrorTitle, es.BankInUse
        Exit Sub
    End If
Next oEmplTemp
        
'Confirm deletion of bank
sDeleteBank = LoadResString(es.DeleteBank)
mbResult = MsgBox(sDeleteBank & lstAllBanks.Text, vbQuestion Or vbYesNo Or vbDefaultButton2, es.AppName)
    
If mbResult = vbNo Then
    'Answer was no so end without deleting bank
    Exit Sub
End If
        
On Error GoTo ErrHandler
                
bRetVal = oAdmin.Delete(oBank:=oBank)
If bRetVal = True Then
    'Bank successfully deleted.
    'Now we need to delete the entry from the listbox
    lstAllBanks.RemoveItem lstAllBanks.ListIndex
    lstAllBanks.ListIndex = -1
    UpdateBankData eUpdate.emClearFields
Else
    'Failed to delete bank
    sDeleteBank = LoadResString(es.DbFailedDeleteBank)
    MsgBox sDeleteBank & lstAllBanks.Text, vbExclamation, es.BenefitsErrorTitle
End If
Exit Sub
    
ErrHandler:
    If Err > vbObjectError Then
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err - vbObjectError) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    Else
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    End If
    bRetVal = False
    Resume Next
End Sub

Private Sub cmdBankCancel_Click()
'User selected Cancel to either an Add or Edit of a Bank

'Re-enable sidebar & other buttons
lstAllBanks.Enabled = True
lstAllBanks_Click
    
sbCategories.Enabled = True
cmdAddBank.Enabled = True

'Disable editable fields
BankFieldsEnable (False)

'Disable OK & Cancel
cmdBankOK.Enabled = False
cmdBankCancel.Enabled = False
End Sub

Private Sub cmdBankOK_Click()
'User selected OK to either an Add or Edit of a Bank
Dim oBank       As CBank
Dim oBankTemp   As CBank
Dim bRetVal     As Boolean
Dim iCurSel     As Integer

On Error GoTo ErrHandler

Set oBank = New CBank
If (lstAllBanks.ListIndex = -1) Then
    'No currently selected item in the Banks List
    'This means that we must be adding a new bank
    'Load the new values into the object
    UpdateBankData eUpdate.emFieldsToData, oBank
        
    bRetVal = oAdmin.Insert(oBank:=oBank)
    
    'Add the new Bank Name to the list
    lstAllBanks.Clear
    BanksToControl lstAllBanks
    
    If bRetVal Then
        FindItemByName lstAllBanks, txtBankName
    Else
        MsgError es.BenefitsErrorTitle, es.DbFailAddBank
    End If
Else
    'An item is currently selected in the Banks List
    'This means that we must be editing an existing bank
    
    'Get the actual Bank object that we are editing
    Set oBank = oAdmin.Banks(lstAllBanks.ItemData(lstAllBanks.ListIndex))
 
    'Load the edited values into the object
    UpdateBankData eUpdate.emFieldsToData, oBank
        
    'Update the Db Record of this bank itself
    bRetVal = oAdmin.Update(oBank:=oBank)
    
    'Recheck the list
    lstAllBanks.Clear
    BanksToControl lstAllBanks
    
    If bRetVal Then
        FindItemByName lstAllBanks, txtBankName
    Else
        MsgError es.BenefitsErrorTitle, es.DbFailEditBank
    End If
    
End If

'Re-enable the Benefits list
lstAllBanks.Enabled = True

'Re-enable sidebar & other buttons
sbCategories.Enabled = True
cmdAddBank.Enabled = True
cmdEditBank.Enabled = True
cmdDeleteBank.Enabled = True

'Disable editable fields
BankFieldsEnable (False)

'Disable OK & Cancel
cmdBankOK.Enabled = False
cmdBankCancel.Enabled = False

Exit Sub

ErrHandler:
    If Err > vbObjectError Then
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err - vbObjectError) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    Else
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    End If
    bRetVal = False
    UpdateBankData eUpdate.emClearFields
    Resume Next
End Sub

Private Sub cmdAddEmployee_Click()
'Add a new Employee

'De-select and disable any item in the Employees Combobox
cbAllEmployees.ListIndex = -1
cbAllEmployees.Enabled = False
    
'Disable sidebar & other buttons
sbCategories.Enabled = False
cmdAddEmployee.Enabled = False
cmdEditEmployee.Enabled = False
cmdDeleteEmployee.Enabled = False

'Enable editable fields
EmployeeFieldsEnable (True)

'Enable OK & Cancel
cmdEmployeeOK.Enabled = True
cmdEmployeeOK.Default = True
cmdEmployeeCancel.Enabled = True

'Set focus to first editable field
txtEmplFirstName.SetFocus
End Sub

Private Sub cmdEditEmployee_Click()
'Edit the current Employee

'Disable the Employees Combobox
cbAllEmployees.Enabled = False

'Disable sidebar & other buttons
cbAllEmployees.Enabled = False
    
sbCategories.Enabled = False
cmdAddEmployee.Enabled = False
cmdEditEmployee.Enabled = False
cmdDeleteEmployee.Enabled = False

'Enable editable fields
EmployeeFieldsEnable (True)

'Enable OK & Cancel
cmdEmployeeOK.Enabled = True
cmdEmployeeOK.Default = True
cmdEmployeeCancel.Enabled = True

'Set focus to first editable field
txtEmplFirstName.SetFocus
End Sub

Private Sub cmdDeleteEmployee_Click()
'Delete the currently selected Employee
Dim oEmpl           As CEmployee
Dim mbResult        As VbMsgBoxResult
Dim bRetVal         As Boolean
Dim sDeleteEmployee As String

'Is an employee selected ?
If cbAllEmployees.ListIndex = -1 Then
    MsgError es.BenefitsErrorTitle, es.NoEmplErrorText
    Exit Sub
End If

'Confirm deletion of employee
sDeleteEmployee = LoadResString(es.DeleteEmployee)
mbResult = MsgBox(sDeleteEmployee & txtEmplFirstName & " " & txtEmplLastName, _
                  vbQuestion Or vbYesNo Or vbDefaultButton2)
    
If mbResult = vbNo Then
    'Answer was no so end without deleting employee
     Exit Sub
Else
     Set oEmpl = oAdmin.Employees(cbAllEmployees.ItemData(cbAllEmployees.ListIndex))
End If
       
On Error GoTo ErrHandler
        
bRetVal = oAdmin.Delete(oEmployee:=oEmpl)
If bRetVal = True Then
    'Benefit successfully deleted.
    'Now we need to delete the entry from the listbox
    cbAllEmployees.RemoveItem cbAllEmployees.ListIndex
    cbAllEmployees.ListIndex = -1
    UpdateEmployeeData eUpdate.emClearFields
Else
    'Failed to delete benefit
    sDeleteEmployee = LoadResString(es.DbFailedDeleteEmployee)
    MsgBox sDeleteEmployee & cbAllEmployees.Text, vbExclamation, es.BenefitsErrorTitle
End If
Exit Sub
    
ErrHandler:
    If Err > vbObjectError Then
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err - vbObjectError) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    Else
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    End If
    bRetVal = False
    Resume Next
End Sub

Private Sub cmdEmployeeCancel_Click()
'User selected Cancel to either an Add or Edit of a Employee

'Re-enable sidebar & other buttons
cbAllEmployees.Enabled = True
cbAllEmployees_Click
    
sbCategories.Enabled = True
cmdAddEmployee.Enabled = True

'Disable editable fields
EmployeeFieldsEnable (False)

'Disable OK & Cancel
cmdEmployeeOK.Enabled = False
cmdEmployeeCancel.Enabled = False
End Sub

Private Sub cmdEmployeeOK_Click()
'User selected OK to either an Add or Edit of a Bank
Dim oEmpl       As CEmployee
Dim oEmplTemp   As CEmployee
Dim bRetVal     As Boolean
Dim iCurSel     As Integer

On Error GoTo ErrHandler
        
Set oEmpl = New CEmployee
If (cbAllEmployees.ListIndex = -1) Then
    'No currently selected item in the Employee Combobox
    'This means that we must be adding a new employee
    'Load the new values into the object
    UpdateEmployeeData eUpdate.emFieldsToData, oEmpl
    
    bRetVal = oAdmin.Insert(oEmployee:=oEmpl)
    
    'Update the Employee list
    cbAllEmployees.Clear
    EmplsToControl cbAllEmployees
    
    If bRetVal Then
        FindItemByName cbAllEmployees, txtEmplLastName & ", " & txtEmplFirstName
    Else
        MsgError es.BenefitsErrorTitle, es.DbFailAddEmployee
    End If
Else
    'An item is currently selected in the Employees Combobox
    'This means that we must be editing an existing employee
    
    'Get the actual employee object that we are editing
    Set oEmpl = oAdmin.Employees(cbAllEmployees.ItemData(cbAllEmployees.ListIndex))
        
    'Load the edited values into the object
    UpdateEmployeeData eUpdate.emFieldsToData, oEmpl
        
    'Update the Db Record of this employee itself
    bRetVal = oAdmin.Update(oEmployee:=oEmpl)
    
    'Recheck the list
    cbAllEmployees.Clear
    EmplsToControl cbAllEmployees
    
    If bRetVal Then
        FindItemByName cbAllEmployees, txtEmplLastName & ", " & txtEmplFirstName
    Else
        MsgError es.BenefitsErrorTitle, es.DbFailEditEmployee
    End If
    
End If

'Re-enable sidebar & other buttons
cbAllEmployees.Enabled = True

sbCategories.Enabled = True
cmdAddEmployee.Enabled = True
cmdEditEmployee.Enabled = True
cmdDeleteEmployee.Enabled = True

'Disable editable fields
EmployeeFieldsEnable (False)

'Disable OK & Cancel
cmdEmployeeOK.Enabled = False
cmdEmployeeCancel.Enabled = False
Exit Sub

ErrHandler:
    If Err > vbObjectError Then
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err - vbObjectError) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    Else
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    End If
    bRetVal = False
    UpdateEmployeeData eUpdate.emClearFields
    Resume Next
End Sub

Private Sub cmdAddDependant_Click()
'Add a new dependant for the current Employee

'De-select and disable any item in the Dependents List
lstDependants.ListIndex = -1
lstDependants.Enabled = False

'Clear all field values
UpdateDependantData eUpdate.emClearFields

'Disable sidebar & other buttons
sbCategories.Enabled = False
cmdAddDependant.Enabled = False
cmdEditDependant.Enabled = False
cmdDeleteDependant.Enabled = False
    
'Enable editable fields
DependantFieldsEnable (True)

'Enable OK & Cancel
cmdDependantOK.Enabled = True
cmdDependantOK.Default = True
cmdDependantCancel.Enabled = True

'Set focus to first editable field
txtDepFirstName.SetFocus
End Sub

Private Sub cmdEditDependant_Click()
'Edit the currently selected dependant for the current Employee

'Disable the Dependants List
lstDependants.Enabled = False
    
'Disable sidebar & other buttons
sbCategories.Enabled = False
cmdAddDependant.Enabled = False
cmdEditDependant.Enabled = False
cmdDeleteDependant.Enabled = False
    
'Enable editable fields
DependantFieldsEnable (True)

'Enable OK & Cancel
cmdDependantOK.Enabled = True
cmdDependantOK.Default = True
cmdDependantCancel.Enabled = True

'Set focus to first editable field
txtDepFirstName.SetFocus
End Sub

Private Sub cmdDeleteDependant_Click()
'Delete the currently selected dependant
Dim oDep                As CDependant
Dim mbResult            As VbMsgBoxResult
Dim bRetVal             As Boolean
Dim sDeleteDependant    As String
Dim sDepPKId            As String
    
'Is a dependant selected ?
If lstDependants.ListIndex = -1 Then
    MsgError es.BenefitsErrorTitle, es.NoDepErrorText
    Exit Sub
End If
    
'Confirm deletion of dependant
sDeleteDependant = LoadResString(es.DeleteDependant)
mbResult = MsgBox(sDeleteDependant & txtDepFirstName & " " & txtDepLastName, _
                  vbQuestion Or vbYesNo Or vbDefaultButton2)
    
If mbResult = vbNo Then
    'Answer was no, so end without deleting dependant
     Exit Sub
End If
    
On Error GoTo ErrHandler
    
'Get the current dependants object
Set oDep = oCurrentEmployee.Dependants(lstDependants.ItemData(lstDependants.ListIndex))
    
'Store the PKId for later use
sDepPKId = oDep.PKId
        
bRetVal = oAdmin.Delete(oDependant:=oDep)
If bRetVal = True Then
    'Dependant successfully deleted.
    'Refill the employees dependants list
    oAdmin.FillListEmployees

    'Now we need to delete the entry from the listbox
    lstDependants.RemoveItem lstDependants.ListIndex
    lstDependants.ListIndex = -1
    UpdateDependantData eUpdate.emClearFields
Else
    'Failed to delete dependant
    sDeleteDependant = LoadResString(es.DbFailedDeleteDependant)
    MsgBox sDeleteDependant & txtDepFirstName & " " & txtDepLastName, vbExclamation, es.BenefitsErrorTitle
End If
Exit Sub
    
ErrHandler:
    If Err > vbObjectError Then
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err - vbObjectError) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    Else
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    End If
    bRetVal = False
    Resume Next
End Sub

Private Sub cmdDependantCancel_Click()
'User selected Cancel to either an Add or Edit of a Dependant

'Re-enable sidebar & other buttons
lstDependants.Enabled = True
lstDependants_Click

sbCategories.Enabled = True
cmdAddDependant.Enabled = True

'Disable editable fields
DependantFieldsEnable (False)

'Disable OK & Cancel
cmdDependantOK.Enabled = False
cmdDependantCancel.Enabled = False
End Sub

Private Sub cmdDependantOK_Click()
'User selected OK to either an Add or Edit of a Dependant
Dim oDep        As CDependant
Dim oDepTemp    As CDependant
Dim bRetVal     As Boolean
Dim iCurSel     As Integer
Dim lTempPKId   As Long

On Error GoTo ErrHandler
    
Set oDep = New CDependant
If (lstDependants.ListIndex = -1) Then
    'No currently selected item in the Dependants List
    'This means that we must be adding a new dependant
    'Load the new values into the object
    UpdateDependantData eUpdate.emFieldsToData, oDep
    'Drop the reference to Current Employee so
    'that you can reset it with the new Dependant
    lTempPKId = oCurrentEmployee.PKId
    oCurrentEmployee.Term
    Set oCurrentEmployee = Nothing
    
    bRetVal = oAdmin.Insert(oDependant:=oDep)
        
    'Re-get the CurrentEmployee with the
    'updated dependants list
    Set oCurrentEmployee = oAdmin.Employees(lTempPKId)
        
    'Update the Depedants list
    lstDependants.Clear
    DependantsToControl oCurrentEmployee.Dependants, lstDependants
    
    If bRetVal Then
        FindItemByName lstDependants, txtDepLastName & ", " & txtDepFirstName
    Else
        MsgError es.BenefitsErrorTitle, es.DbFailAddDependant
    End If
Else
    'An item is currently selected in the Dependants List
    'This means that we must be editing an existing dependant
    
    'Get the actual dependant object that we are editing
    Set oDep = oCurrentEmployee.Dependants(lstDependants.ItemData(lstDependants.ListIndex))
        
    'Load the edited values into the object
    UpdateDependantData eUpdate.emFieldsToData, oDep
        
    'Update the Db Record of this dependant itself
    bRetVal = oAdmin.Update(oDependant:=oDep)
    
    'Recheck the list
    lstDependants.Clear
    DependantsToControl oCurrentEmployee.Dependants, lstDependants
    
    If bRetVal Then
        FindItemByName lstDependants, txtDepLastName & ", " & txtDepFirstName
    Else
        MsgError es.BenefitsErrorTitle, es.DbFailEditDependant
    End If
End If

'Re-enable sidebar & other buttons
lstDependants.Enabled = True
    
sbCategories.Enabled = True
cmdAddDependant.Enabled = True
cmdEditDependant.Enabled = True
cmdDeleteDependant.Enabled = True

'Disable editable fields
DependantFieldsEnable (False)

'Disable OK & Cancel
cmdDependantOK.Enabled = False
cmdDependantCancel.Enabled = False
Exit Sub

ErrHandler:
    If Err > vbObjectError Then
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err - vbObjectError) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    Else
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    End If
    bRetVal = False
    UpdateDependantData eUpdate.emClearFields
    Resume Next
End Sub

Private Sub cmdMDBBrowse_Click()
'This function uses the Open File dialog to help the user select a valid Access Database
Dim OpenFile    As OPENFILENAME
Dim lReturn     As Long
Dim sFilter     As String
    
OpenFile.lStructSize = Len(OpenFile)
OpenFile.hwndOwner = frmMain.hwnd
OpenFile.hInstance = App.hInstance
sFilter = LoadResString(es.MDBFileDesc) & Chr(0) & LoadResString(es.MDBFileExt) & Chr(0)
OpenFile.lpstrFilter = sFilter
OpenFile.nFilterIndex = 1
OpenFile.lpstrFile = String(257, 0)
OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
OpenFile.lpstrFileTitle = OpenFile.lpstrFile
OpenFile.nMaxFileTitle = OpenFile.nMaxFile
OpenFile.lpstrInitialDir = CurDir$
OpenFile.lpstrTitle = LoadResString(es.MDBSelectText)
OpenFile.flags = 0
lReturn = GetOpenFileName(OpenFile)
    
If lReturn Then
    'A valid Access Database was selected
    optDBChoice(0).Value = True
    txtMDBPath = OpenFile.lpstrFile
    txtMDBPath.SetFocus
    ' Enable the OK button if configuring for the first time
    If Not sbCategories.Visible Then cmdSettingsOK.Enabled = True
End If
End Sub

Private Sub cmdQueryReset_Click()
'Reset query parameters

cbQueryBenefitTypes.ListIndex = 0
cbQueryBenefits.ListIndex = 0
    
optQueryGender(0) = True
cbQueryCoverage.ListIndex = 0
cbQueryTax.ListIndex = 0
    
End Sub

Private Sub cmdSettingsCancel_Click()
'This button is only visible when
'the application is prompting for
'initial Data Services settings
Dim vbrCancel As VbMsgBoxResult

'User chose to cancel therefore
'the application will close
'Confirm closing of application
vbrCancel = MsgBox(LoadResString(es.CancelSetupText), vbQuestion Or vbYesNo, LoadResString(es.CancelSetupTitle))
If vbrCancel = vbYes Then Unload Me
End Sub

Private Sub cmdSettingsOK_Click()
'This button is only visible when
'the application is prompting for
'initial Data Services settings

If Len(txtMDBPath) < 1 And Len(txtDSN) < 1 Then
    ' No database has been set yet. Exit now
    Beep
    Exit Sub
End If

'Update db field values before continuing
If optDBChoice(0).Value Then
    tSet.sType = "Access"
    tSet.sDB = Trim$(txtMDBPath)
End If
If optDBChoice(1).Value Then
    tSet.sType = "ODBC"
    tSet.sDB = Trim$(txtDSN)
End If

SaveSettings

If Not sbCategories.Visible Then
    Unload Me
End If
End Sub

Private Sub cmdAnotherQuery_Click()
'Reset results
'to allow for another query

Dim pctRes      As PictureBox
Dim pctStats    As PictureBox

Set pctStats = cScreens("Queries")
pctStats.Enabled = True
pctStats.Visible = True
    
Set pctRes = cScreens("Results")
pctRes.Enabled = False
pctRes.Visible = False
End Sub

Private Sub cmdRunRACMgr_Click()
'Run Remote Automation Manager
'This option is supplied to allow for
'the user to select the different methods of
'remote communication.
'1) Local
'2) Network using Remote Automation
'3) Network using Distributed COM

Shell LoadResString(es.RACMgr), vbNormalFocus
End Sub

Private Sub cmdQuery_Click()
'Query function for pulling together the statistice of
'Corporate Benefit selections by employees.

'Used by MS Graph OLE Object
Dim sData       As String

'Used by MS Chart OLE Control
Dim arrData()   As Integer
Dim iRows       As Integer
Dim iCols       As Integer
    
Dim bRetVal     As Boolean
Dim iCount      As Integer
Dim iCoverage   As Integer
Dim iTax        As Integer
Dim lQueryTotal As Long

'Used for reporting errors with localizable text
Dim sErrorTitle As String
Dim sErrorText  As String

'Query Collections
Dim colFields   As Collection
Dim colFullList As Collection

'Generic CLists used to fill list-boxes & combo-boxes
Dim lstFields   As New CList
Dim lstFull     As New CList

'Temporary Benefit Object
Dim oBenTemp    As CBenefit

On Error GoTo QueryErrHandler

Screen.MousePointer = vbHourglass

'Collect any necessary parameters to the query
iCoverage = cbQueryCoverage.ListIndex
iTax = cbQueryTax.ListIndex
    
If cbQueryBenefitTypes.ListIndex = 0 Then
    'All Benefit Types to be listed
    If cbQueryBenefits.ListIndex = 0 Then
        'All Benefits Types to be queried on
        lblQueryResultsText = LoadResString(es.qryEmplsPerCat) & vbCrLf
        If (giQueryGender = 0) Then
            bRetVal = oAdmin.GetNumEmplsPerBenefitType(lstFull, iCoverage, iTax)
        Else
            bRetVal = oAdmin.GetNumEmplsPerBenefitType(lstFull, iCoverage, iTax, CBool(giQueryGender = 2))
        End If
    Else
        'All Benefits are to be queried for
        lblQueryResultsText = LoadResString(es.qryEmplsPerBen) & vbCrLf
        If (giQueryGender = 0) Then
            bRetVal = oAdmin.GetNumEmplsPerBenefit(lstFull, iCoverage, iTax)
        Else
            bRetVal = oAdmin.GetNumEmplsPerBenefit(lstFull, iCoverage, iTax, CBool(giQueryGender = 2))
        End If
    End If
Else
    'A specific Benefit Type is to be listed
    If cbQueryBenefits.ListIndex = 0 Then
        'All Benefits of a specific type are to be queried on
        lblQueryResultsText = LoadResString(es.qryEmplsOneCat) & cbQueryBenefitTypes.Text & "." & vbCrLf
        If (giQueryGender = 0) Then
            bRetVal = oAdmin.GetNumEmplsPerBenefit(lstFull, iCoverage, iTax)
        Else
            bRetVal = oAdmin.GetNumEmplsPerBenefit(lstFull, iCoverage, iTax, CBool(giQueryGender = 2))
        End If
    Else
        'A specific Benefit is to be queried on
        lblQueryResultsText = LoadResString(es.qryEmplsOneBen) & cbQueryBenefitTypes.Text & "." & vbCrLf
        If (giQueryGender = 0) Then
            bRetVal = oAdmin.GetNumEmplsPerBenefit(lstFull, iCoverage, iTax)
        Else
            bRetVal = oAdmin.GetNumEmplsPerBenefit(lstFull, iCoverage, iTax, CBool(giQueryGender = 2))
        End If
    End If
End If
        
If bRetVal = False Then
    'Something went wrong during the query, report the error
    If (Err.Number = 0) Then
        'It was a valid error because nobody used the benefit queried on
        MsgError es.DbNoEmplsFoundTitle, es.DbNoEmplsFoundText
    Else
        'Report an operational error
        MsgError es.BenefitsErrorTitle, es.DBBenTypesFailed
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
End If
    
On Error GoTo ResultsErrHandler
    
'Now we've got the query results, let's plot them...
If bUseMSGraph = True Then
    'MS Graph is installed, so set it up for use.

    oleMSGraph.Visible = True
    oleMSGraph.ZOrder
    ctlMSChart.Visible = False
        
    'Set the file format to text.
    oleMSGraph.Format = LoadResString(es.CFTEXT)
    'Create an the MS Graph OLE Object.
    oleMSGraph.CreateEmbed "", sGraphPID
    
    'Put data into a string that the MSGraph OLE Object understands
    iCount = 0
    If cbQueryBenefitTypes.ListIndex = 0 Then
        'All Benefit Types to be listed
        If cbQueryBenefits.ListIndex = 0 Then
            'All Benefits Types to be queried on
            For Each lstFields In lstFull
                sData = sData + vbNewLine + lstFields.Item("Benefit Type Name") & vbTab & lstFields.Item("Employee Count")
                iCount = iCount + 1
            Next
        Else
            'All Benefits are to be queried for
            For Each oBenTemp In oAdmin.Benefits
                bRetVal = False
                'Cycle throught the returned list, checking for each benefit in turn
                For Each lstFields In lstFull
                    If ((lstFields.Item("Benefit Name") = oBenTemp.Description)) Then
                        sData = sData + vbNewLine + lstFields.Item("Benefit Name") & vbTab & lstFields.Item("Employee Count")
                        bRetVal = True
                        Exit For
                    End If
                Next
                If (bRetVal = False) Then
                    'Nobody has this benefit
                    sData = sData + vbNewLine + oBenTemp.Description & vbTab & 0
                End If
                iCount = iCount + 1
            Next
        End If
    Else
        'A specific Benefit Type is to be listed
        If cbQueryBenefits.ListIndex = 0 Then
            'All Benefits of a specific type are to be queried on
            For Each lstFields In lstFull
                For Each oBenTemp In oAdmin.Benefits
                    If ((lstFields.Item("Benefit Name") = oBenTemp.Description) And _
                        (oBenTemp.BnfTypeId = cbQueryBenefitTypes.ItemData(cbQueryBenefitTypes.ListIndex))) Then
                           sData = sData + vbNewLine + lstFields.Item("Benefit Name") & vbTab & lstFields.Item("Employee Count")
                           iCount = iCount + lstFields.Item("Employee Count")
                    End If
                Next
            Next
        Else
            'A specific Benefit is to be queried on
            For Each lstFields In lstFull
                If (lstFields.Item("Benefit Name") = cbQueryBenefits.Text) Then
                    sData = sData + vbNewLine + lstFields.Item("Benefit Name") & vbTab & lstFields.Item("Employee Count")
                    iCount = iCount + lstFields.Item("Employee Count")
                End If
            Next
        End If
    End If
        
    'Send the data using the DataText property.
    'Activate MSGRAPH as hidden.
    oleMSGraph.DoVerb vbOLEHide
        
    If oleMSGraph.AppIsRunning Then
        oleMSGraph.DataText = sData
            
        'Set the Query Results Text
        lblQueryResultsText = lblQueryResultsText & LoadResString(es.qryResults) & iCount
            
        'Update the object.
        oleMSGraph.Update
    Else
        'MS Graph is registered but failed to run,
        'report the error then do this again, this time using
        'the Chart control (MSChart.ocx)
        sErrorTitle = LoadResString(es.ChartErrorTitle)
        sErrorText = LoadResString(es.ChartErrorText)
        MsgBox sErrorText, vbExclamation, sErrorTitle
        bUseMSGraph = False
        cmdQuery_Click
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
        
Else
    'MS Graph IS NOT avaiable, so use the MS Chart control instead.

    ctlMSChart.Visible = True
    ctlMSChart.ZOrder
    oleMSGraph.Visible = False
        
    'Setup how many rows & columns of data there are
    iRows = 0
    iCols = lstFull.Count - 1
            
    ctlMSChart.ColumnLabelCount = 2

    ctlMSChart.ColumnCount = 0
    ctlMSChart.RowCount = 1
    iCount = 1
        
    If (cbQueryBenefitTypes.ListIndex = 0) Then
        For Each lstFields In lstFull
            ctlMSChart.ColumnCount = ctlMSChart.ColumnCount + 1
            ctlMSChart.Column = iCount
            'Set the Column Label
            ctlMSChart.ColumnLabelIndex = 1
            ctlMSChart.ColumnLabel = CInt(lstFields.Item("Employee Count"))
            ctlMSChart.ColumnLabelIndex = 2
            If (cbQueryBenefits.ListIndex = 0) Then
               ctlMSChart.ColumnLabel = lstFields.Item("Benefit Type Name")
            Else
               ctlMSChart.ColumnLabel = lstFields.Item("Benefit Name")
            End If
            'Set the Column Data
            ctlMSChart.Data = CInt(lstFields.Item("Employee Count"))
            iCount = iCount + 1
        Next
    Else
        If (cbQueryBenefits.ListIndex = 0) Then
            For Each oBenTemp In oAdmin.Benefits
                If (oBenTemp.BnfTypeId = cbQueryBenefitTypes.ItemData(cbQueryBenefitTypes.ListIndex)) Then
                    For Each lstFields In lstFull
                        If (oBenTemp.Description = lstFields.Item("Benefit Name")) Then
                            ctlMSChart.ColumnCount = ctlMSChart.ColumnCount + 1
                            ctlMSChart.Column = iCount
                            'Set the Column Label
                            ctlMSChart.ColumnLabelIndex = 1
                            ctlMSChart.ColumnLabel = CInt(lstFields.Item("Employee Count"))
                            ctlMSChart.ColumnLabelIndex = 2
                            ctlMSChart.ColumnLabel = lstFields.Item("Benefit Name")
                            'Set the Column Data
                            ctlMSChart.Data = CInt(lstFields.Item("Employee Count"))
                            iCount = iCount + 1
                        End If
                    Next
                End If
            Next
        Else
            For Each lstFields In lstFull
                If (lstFields.Item("Benefit Name") = cbQueryBenefits.Text) Then
                    ctlMSChart.ColumnCount = ctlMSChart.ColumnCount + 1
                    ctlMSChart.Column = iCount
                    'Set the Column Label
                    ctlMSChart.ColumnLabelIndex = 1
                    ctlMSChart.ColumnLabel = CInt(lstFields.Item("Employee Count"))
                    ctlMSChart.ColumnLabelIndex = 2
                    ctlMSChart.ColumnLabel = lstFields.Item("Benefit Name")
                    'Set the Column Data
                    ctlMSChart.Data = CInt(lstFields.Item("Employee Count"))
                    iCount = iCount + 1
                End If
            Next
        End If
    End If
        
    'Set the Query Results Text
    lblQueryResultsText = lblQueryResultsText & _
                         LoadResString(es.qryResults) & ctlMSChart.ColumnCount
    
End If
    
'Having setup the Query Result window, lets display it...
pctResults.Enabled = True
pctResults.Visible = True
cmdAnotherQuery.Default = True
pctResults.Top = 1125
pctResults.Left = 1140
pctResults.ZOrder
    
'Hide the Query Screen
pctQuery.Enabled = False
pctQuery.Visible = False
    
Screen.MousePointer = vbDefault
    
Exit Sub

QueryErrHandler:
    If Err > vbObjectError Then
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err - vbObjectError) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    Else
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    End If
    bRetVal = False
    Resume Next
    
ResultsErrHandler:
    If Err > vbObjectError Then
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err - vbObjectError) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    Else
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
End Sub

Private Sub lstAllBanks_Click()
'Update all the Benefit detail fields
'with the selected items values

Dim oBank       As CBank
Dim iCount      As Integer
Dim lBankPKId   As Long

If (lstAllBanks.ListIndex = -1) Then
    'Clear all field values
    UpdateBankData eUpdate.emClearFields
    cmdEditBank.Enabled = False
    cmdDeleteBank.Enabled = False
    Exit Sub
Else
    cmdEditBank.Enabled = True
    cmdDeleteBank.Enabled = True
End If

'Get the selected Benefits ID
lBankPKId = lstAllBanks.ItemData(lstAllBanks.ListIndex)
    
'Get the Benefits object from it's ID
Set oBank = oAdmin.Banks(lBankPKId)
       
'Update all the fields to reflect the Banks values
UpdateBankData eUpdate.emDataToFields, oBank
End Sub

Private Sub lstAllBenefits_Click()
'Update all the Benefit detail fields
'with the selected items values

Dim oBenefit        As CBenefit
Dim iCount          As Integer
Dim lBenefitPKId    As Long

If (lstAllBenefits.ListIndex = -1) Then
    'Clear all field values
    UpdateBenefitData eUpdate.emClearFields
    cmdEditBenefit.Enabled = False
    cmdDeleteBenefit.Enabled = False
    Exit Sub
Else
    cmdEditBenefit.Enabled = True
    cmdDeleteBenefit.Enabled = True
End If

'Get the selected Benefits ID
lBenefitPKId = lstAllBenefits.ItemData(lstAllBenefits.ListIndex)
    
'Get the Benefits object from it's ID
Set oBenefit = oAdmin.Benefits(lBenefitPKId)
       
'Update all the fields to reflect the Benefits values
UpdateBenefitData eUpdate.emDataToFields, oBenefit
End Sub

Private Sub lstDependants_Click()
'Update all the Dependant detail fields
'with the selected items values

Dim oDepTemp    As CDependant
Dim iCount      As Integer
Dim lDepPKId    As Long

UpdateDependantData eUpdate.emClearFields
If lstDependants.ListIndex = -1 Then
    cmdEditDependant.Enabled = False
    cmdDeleteDependant.Enabled = False
    Exit Sub
Else
    cmdEditDependant.Enabled = True
    cmdDeleteDependant.Enabled = True
End If

'Get the selected Benefits ID
lDepPKId = lstDependants.ItemData(lstDependants.ListIndex)
   
'Get the Benefits object from it's ID
Set oDepTemp = oCurrentEmployee.Dependants(lDepPKId)
    
UpdateDependantData eUpdate.emDataToFields, oDepTemp
       
End Sub

Private Sub optDBChoice_Click(Index As Integer)
'Called when a user changes database types (Access vs. ODBC)
Select Case (Index)
Case 0:
    lblDataServices(5).Caption = "Access" & vbCrLf & Trim$(txtMDBPath)
    txtMDBPath.Enabled = True
    cmdMDBBrowse.Enabled = True
    txtDSN.Enabled = False
    cmdDSNChoose.Enabled = False
Case 1:
    lblDataServices(5).Caption = "ODBC" & vbCrLf & Trim$(txtDSN)
    txtMDBPath.Enabled = False
    cmdMDBBrowse.Enabled = False
    txtDSN.Enabled = True
    cmdDSNChoose.Enabled = True
End Select
DoPaint
End Sub

Private Sub optQueryGender_Click(Index As Integer)
'This is for the gender options involved in a query
giQueryGender = Index
End Sub

Private Sub picData_Click(Index As Integer)
Dim iLoop As Integer

Select Case Index
Case 2: 'Network Communication
    If tSet.lRem Then
        'Explain the use of RA & DCOM
        lblNetComm.Caption = LoadResString(NetComm)
        cmdRunRACMgr.Enabled = True
    Else
        'No Remote Server has been selected - tell the user
        lblNetComm.Caption = LoadResString(RemFuncNA)
        cmdRunRACMgr.Enabled = False
    End If
    
Case 3: 'Remote Server
    'Set up the list of ActiveX Benefits Servers available for Remote
    cbRemoteServer.Clear
    cbRemoteServer.AddItem LoadResString(DOPASSTHROUGH)
    cbRemoteServer.ItemData(cbRemoteServer.NewIndex) = stPassthrough
    'Can't have a remote server of the same type as the local
    If stLocal <> stVBEXE Then
        If bServers(stVBEXE) Then
            'The VB Server is registered on the this computer so add it to the list
            cbRemoteServer.AddItem LoadResString(DOVBEXE)
            cbRemoteServer.ItemData(cbRemoteServer.NewIndex) = stVBEXE
        End If
    End If
    If stLocal <> stCPPEXE Then
        If bServers(stCPPEXE) Then
            'The C++ Server is registered on the this computer so add it to the list
            cbRemoteServer.AddItem LoadResString(DOCPPEXE)
            cbRemoteServer.ItemData(cbRemoteServer.NewIndex) = stCPPEXE
        End If
    End If
    If stLocal <> stATLEXE Then
        If bServers(stATLEXE) Then
            'The ATL Server is registered on the this computer so add it to the list
            cbRemoteServer.AddItem LoadResString(DOATLEXE)
            cbRemoteServer.ItemData(cbRemoteServer.NewIndex) = stATLEXE
        End If
    End If
    If stLocal <> stMFCEXE Then
        If bServers(stMFCEXE) Then
            'The MFC Server is registered on the this computer so add it to the list
            cbRemoteServer.AddItem LoadResString(DOMFCEXE)
            cbRemoteServer.ItemData(cbRemoteServer.NewIndex) = stMFCEXE
        End If
    End If
    If cbRemoteServer.ListCount = 1 Then
        cbRemoteServer.ListIndex = 0
        cbRemoteServer.Enabled = False
        lblRemFuncNA.Visible = True
        lblRemFuncNA.Enabled = True
        imgRemoteServer.Visible = False
        If tSet.lRem Then
            'Something went wrong with the Server settings so set Remote and Remote Link to Passthrough
            tSet.lSrc = icDoDataDirect
            tSet.lRem = 0
            tSet.lLnk = 0
        End If
    Else
        'Select the current Remote Server from the combo-box
        For iLoop = 0 To cbRemoteServer.ListCount - 1
            If cbRemoteServer.ItemData(iLoop) = tSet.lRem Then
                cbRemoteServer.ListIndex = iLoop
                Exit For
            End If
        Next iLoop
        cbRemoteServer.Enabled = True
        lblRemFuncNA.Visible = False
        lblRemFuncNA.Enabled = False
        imgRemoteServer.Visible = True
    End If
Case 4: 'Remote Link Server
    'Set up the list of ActiveX Benefits Servers available for Remote Link
    cbRemoteLinkServer.Clear
    cbRemoteLinkServer.AddItem LoadResString(DOPASSTHROUGH)
    cbRemoteLinkServer.ItemData(cbRemoteLinkServer.NewIndex) = stPassthrough
    If ((tSet.lRem <> stVBEXE) And bServers(stVBEXE)) Then
        'The Remote Link Server can not be of the same type as the Remote Server
        cbRemoteLinkServer.AddItem LoadResString(DOVBEXE)
        cbRemoteLinkServer.ItemData(cbRemoteLinkServer.NewIndex) = stVBEXE
    End If
    If ((tSet.lRem <> stCPPEXE) And bServers(stCPPEXE)) Then
        'The Remote Link Server can not be of the same type as the Remote Server
        cbRemoteLinkServer.AddItem LoadResString(DOCPPEXE)
        cbRemoteLinkServer.ItemData(cbRemoteLinkServer.NewIndex) = stCPPEXE
    End If
    If ((tSet.lRem <> stATLEXE) And bServers(stATLEXE)) Then
        'The Remote Link Server can not be of the same type as the Remote Server
        cbRemoteLinkServer.AddItem LoadResString(DOATLEXE)
        cbRemoteLinkServer.ItemData(cbRemoteLinkServer.NewIndex) = stATLEXE
    End If
    If ((tSet.lRem <> stMFCEXE) And bServers(stMFCEXE)) Then
        'The Remote Link Server can not be of the same type as the Remote Server
        cbRemoteLinkServer.AddItem LoadResString(DOMFCEXE)
        cbRemoteLinkServer.ItemData(cbRemoteLinkServer.NewIndex) = stMFCEXE
    End If
    If (tSet.lLnk = tSet.lRem) Or (tSet.lRem = stPassthrough) Or (cbRemoteLinkServer.ListCount = 1) Then
        cbRemoteLinkServer.ListIndex = 0
        cbRemoteLinkServer.Enabled = False
        lblRemLinkFuncNA.Visible = True
        lblRemLinkFuncNA.Enabled = True
        imgRemoteLinkServer.Visible = False
        If tSet.lLnk Then
            'Something went wrong with the Server settings so set Remote and Remote Link to Passthrough
            tSet.lSrc = icDoDataDirect
            tSet.lRem = 0
            tSet.lLnk = 0
        End If
        If tSet.lRem = stPassthrough Then
            'Tell the user that they must select a Remote Server before they can set a Remote Link Server
            lblRemLinkFuncNA.Caption = LoadResString(es.RemFuncNA)
        Else
            'Tell the user that there are no more servers available for use
            lblRemLinkFuncNA.Caption = LoadResString(es.RemFuncNoSrvAvail)
        End If
    Else
        For iLoop = 0 To cbRemoteLinkServer.ListCount - 1
            If (cbRemoteLinkServer.ItemData(iLoop) = tSet.lLnk) Then
                cbRemoteLinkServer.ListIndex = iLoop
                 Exit For
            End If
        Next iLoop
        cbRemoteLinkServer.Enabled = True
        lblRemLinkFuncNA.Visible = False
        lblRemLinkFuncNA.Enabled = False
        imgRemoteLinkServer.Visible = True
    End If
Case 5: 'Database
    If tSet.sType = "Access" Then
        'The current settings are for an Access Database, display the path
        txtMDBPath = tSet.sDB
    Else
        'The current settings are for an ODBC Database, display the relevant settings
        txtDSN = tSet.sDB
    End If
End Select
                            
DoPaint
For iLoop = 1 To cDataOptions.Count
    If iLoop = (Index - 1) Then
        cDataOptions(Index - 1).Top = 142
        cDataOptions(Index - 1).Left = 8
        cDataOptions(Index - 1).Enabled = True
        cDataOptions(Index - 1).Visible = True
        cDataOptions(Index - 1).ZOrder
    Else
        cDataOptions.Item(iLoop).Enabled = False
        cDataOptions.Item(iLoop).Visible = False
        cDataOptions.Item(iLoop).ZOrder 1
    End If
Next iLoop
End Sub

Private Sub picData_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim iLoop As Integer
For iLoop = 2 To 5
    If iLoop = Index Then
        picData(iLoop).BorderStyle = 1
    Else
        picData(iLoop).BorderStyle = 0
    End If
Next iLoop
End Sub

Private Sub sbCategories_FolderClick(Folder As SideBarVB.Folder)
'Handler for SideBar control to select
'the correct screen from the tab selected
Dim gbScreen    As Frame
Dim pctScreen   As PictureBox
   
'Trap to see if we are entering the DataOptions Screen
If (Folder.Key = LoadResString(es.DataAdmin)) Then
    If (tSet.sType = "Access") Then
        optDBChoice(0) = True
        txtMDBPath = tSet.sDB
    Else
        optDBChoice(1) = True
        txtDSN = tSet.sDB
    End If
End If

'Trap to see if we are leaving the DataOptions Screen
If pctDataServices.Enabled Then
   'Disable all the Data Option Screens
    For Each gbScreen In cDataOptions
        gbScreen.Enabled = False
        gbScreen.Visible = False
    Next gbScreen
End If
   
For Each pctScreen In cScreens
    If pctScreen.Tag = Folder.Key Then
        pctScreen.Enabled = True
        pctScreen.Visible = True
        pctScreen.Top = 1125
        pctScreen.Left = 1140
        pctScreen.ZOrder
    Else
        pctScreen.Enabled = False
        pctScreen.Visible = False
    End If
Next pctScreen
End Sub

Private Sub sbCategories_FolderItemClick(FolderItem As SideBarVB.FolderItem)
'Handler for SideBar control to select
'the correct screen from the selected item

Dim lEmployeePKId   As Integer
Dim pctScreen       As PictureBox
    
On Error GoTo ErrHandler
    
For Each pctScreen In cScreens
    If pctScreen.Tag = FolderItem.Key Then
    
        Select Case FolderItem.Key
        Case "Benefits"
            'The Benefits screen has been selected.
            'Make sure that the Benefits List is upto date.
            lstAllBenefits.Clear
            BenefitsToControl lstAllBenefits
            
        Case "Employees"
            'The Employees screen has been selected.
            'Make sure that the Employees Lists is upto date.
            cbAllEmployees.Clear
            EmplsToControl cbAllEmployees
            
            cbEmplBanks.Clear
            cbEmplBanks.AddItem LoadResString(es.emplBankUnknown)
            cbEmplBanks.ItemData(cbEmplBanks.NewIndex) = 0
            BanksToControl cbEmplBanks
            
            If Not (oCurrentEmployee Is Nothing) Then
                FindItem cbAllEmployees, oCurrentEmployee.PKId
            Else
                'No current Employee, so select the firtst in the list
                If cbAllEmployees.ListCount > 0 Then
                    cbAllEmployees.ListIndex = 0
                    cbAllEmployees_Click
                End If
            End If
            
        Case "Dependants"
            'The dependants screen has been selected.
            'We need to fill the screen details with
            'the current employees dependants
            
            lstDependants.Clear
            
            'Set Employee label with Employee Name
            If Not (oCurrentEmployee Is Nothing) Then
                lblEmployeeName.Caption = oCurrentEmployee.FirstName & " " & oCurrentEmployee.LastName
                'Fill list Dependants
                DependantsToControl oCurrentEmployee.Dependants, lstDependants
            End If
            
            'Select first dependant or clear fields
            If lstDependants.ListCount > 0 Then
                lstDependants.ListIndex = 0
            End If
            lstDependants_Click
            
        Case "Banks"
            'The dependants screen has been selected.
            'Set the current bank to the current
            'employees bank, if any
            lstAllBanks.Clear
            BanksToControl lstAllBanks
            
            If Not (oCurrentEmployee Is Nothing) Then
                If Not IsEmpty(oCurrentEmployee.Bank) Then
                    FindItem lstAllBanks, oCurrentEmployee.Bank.PKId
                End If
            End If
        End Select

        pctScreen.Enabled = True
        pctScreen.Visible = True
        pctScreen.Top = 1125
        pctScreen.Left = 1140
        pctScreen.ZOrder
    Else
        pctScreen.Enabled = False
    End If
Next
Exit Sub

ErrHandler:
    If Err > vbObjectError Then
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err - vbObjectError) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    Else
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    End If
    Resume Next
End Sub

Private Sub DependantFieldsEnable(ByVal bEnable As Boolean)
txtDepFirstName.Enabled = bEnable
txtDepLastName.Enabled = bEnable
txtDepDOB.Enabled = bEnable
optDepGenderMale.Enabled = bEnable
optDepGenderFemale.Enabled = bEnable
optDepStatusDependant.Enabled = bEnable
optDepStatusBeneficiary.Enabled = bEnable
cbDepRelationship.Enabled = bEnable
End Sub

Private Sub Form_Load()
'Initialize values
Dim iLoop           As Integer
Dim hKey            As Long
Dim lRetVal         As Long
Dim sServerProgID   As String

'Setup collection of available screens
Set cScreens = New Collection
With cScreens
    .Add pctQuery, LoadResString(es.Queries)
    .Add pctResults, LoadResString(es.Results)
    .Add pctEmployees, LoadResString(es.Employees)
    .Add pctBenefits, LoadResString(es.Benefits)
    .Add pctDependants, LoadResString(es.Dependants)
    .Add pctBanks, LoadResString(es.Banks)
    .Add pctDataServices, LoadResString(es.DataServices)
End With

'Check the registry for what ActiveX Servers are available for use
For iLoop = 1 To stTotalServers
    sServerProgID = LoadResString(es.srvPassthrough + iLoop) & ".CAdmin"
    lRetVal = RegOpenKeyEx(HKEY_CLASSES_ROOT, sServerProgID, 0, KEY_QUERY_VALUE, hKey)
    If lRetVal = ERROR_NONE Then
        'ActiveXServer IS installed
        bServers(iLoop) = True
        'Unload Key as we don't need it
        RegCloseKey (hKey)
    Else
        'ActiveXServer IS NOT installed
        bServers(iLoop) = False
    End If
Next iLoop

'Setup of the Database Options Screen
Set cDataOptions = New Collection
With cDataOptions
    .Add gbNetComm, LoadResString(es.DSNetComm)
    .Add gbRemote, LoadResString(es.DSRemote)
    .Add gbRemoteLink, LoadResString(es.DSRemoteLink)
    .Add gbDatabase, LoadResString(es.DSDatabase)
End With
'Paint the application bitmap
picData(0).Picture = ilDataServices.ListImages("VB Application").Picture
'Paint the local server bitmap
picData(1).Picture = ilDataServices.ListImages(stLocal + 2).Picture
lblDataServices(1).Caption = LoadResString(DOPASSTHROUGH + stLocal)
'Paint the Database bitmap
picData(5).Picture = ilDataServices.ListImages("Database").Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
SaveSettings
End Sub

'Generic Fill Listbox or Combo Box from supplied list Procedure
Private Sub DomainListToControl(lst As CList, ctl As Control)
Dim lstFields As CList
    
    For Each lstFields In lst
        ctl.AddItem lstFields.Item("Description")
        ctl.ItemData(ctl.NewIndex) = lstFields.Item("PKId")
    Next
End Sub

'Fill Listbox or Combo Box with all Benefits Procedure
Private Sub BenefitsToControl(ctl As Control)
Dim oBenefit As CBenefit
    For Each oBenefit In oAdmin.Benefits
        ctl.AddItem oBenefit.Description
        ctl.ItemData(ctl.NewIndex) = oBenefit.PKId
    Next
End Sub

'Fill Listbox or Combo Box with all Employees Procedure
Private Sub EmplsToControl(ctl As Control)
Dim oEmployee As CEmployee
    For Each oEmployee In oAdmin.Employees
        ctl.AddItem oEmployee.LastName & ", " & oEmployee.FirstName
        ctl.ItemData(ctl.NewIndex) = oEmployee.PKId
    Next
End Sub
        
'Fill Listbox or Combo Box with all Dependants Procedure
Private Sub DependantsToControl(lst As CListDependant, ctl As Control)
Dim oDep As CDependant

    For Each oDep In lst
        ctl.AddItem oDep.FirstName & " " & oDep.LastName
        ctl.ItemData(ctl.NewIndex) = oDep.PKId
    Next
End Sub
        
'Fill Listbox or Combo Box with all Banks Procedure
Private Sub BanksToControl(ctl As Control)
Dim oBank As CBank
    For Each oBank In oAdmin.Banks
        ctl.AddItem oBank.Name
        ctl.ItemData(ctl.NewIndex) = oBank.PKId
    Next
End Sub

'Find an item in a listbox or combobox by its itemdata integer value
Private Function FindItem(ctl As Control, iD As Integer) As Integer
Dim iCount As Integer

    For iCount = 0 To ctl.ListCount - 1
        If (ctl.ItemData(iCount) = iD) Then
            ctl.ListIndex = iCount
            FindItem = iCount
            Exit Function
        End If
    Next iCount
    FindItem = -1
End Function

'Find an item in a listbox or combobox by its text
Private Function FindItemByName(ctl As Control, Name As String) As Integer
Dim iCount As Integer
    
    For iCount = 0 To ctl.ListCount - 1
        If (ctl.List(iCount) = Name) Then
            ctl.ListIndex = iCount
            FindItemByName = iCount
            Exit Function
        End If
    Next iCount
    FindItemByName = -1
End Function

Private Sub txtDSN_Change()
' Enable the OK button if configuring for the first time
If Not sbCategories.Visible And Len(txtDSN) > 0 Then
    cmdSettingsOK.Enabled = True
Else
    cmdSettingsOK.Enabled = False
End If
End Sub

Private Sub txtMDBPath_Change()
' Enable the OK button if configuring for the first time
If Not sbCategories.Visible And Len(txtMDBPath) > 0 Then
    cmdSettingsOK.Enabled = True
Else
    cmdSettingsOK.Enabled = False
End If
End Sub
