VERSION 5.00
Begin VB.Form frmSoda 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF0000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5805
   ClientLeft      =   6075
   ClientTop       =   3510
   ClientWidth     =   3825
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   Icon            =   "frmSoda.frx":0000
   LinkTopic       =   "frmSoda"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   3825
   Begin VB.CommandButton cmdRestock1 
      Caption         =   "1"
      Height          =   550
      Left            =   3650
      TabIndex        =   43
      Top             =   1080
      Width           =   135
   End
   Begin VB.CommandButton cmdRestock2 
      Caption         =   "2"
      Height          =   550
      Left            =   3650
      TabIndex        =   42
      Top             =   1800
      Width           =   135
   End
   Begin VB.CommandButton cmdRestock3 
      Caption         =   "3"
      Height          =   550
      Left            =   3650
      TabIndex        =   41
      Top             =   2520
      Width           =   135
   End
   Begin VB.CommandButton cmdRestock4 
      Caption         =   "4"
      Height          =   550
      Left            =   3650
      TabIndex        =   40
      Top             =   3240
      Width           =   135
   End
   Begin VB.CommandButton cmdRestock5 
      Caption         =   "5"
      Height          =   550
      Left            =   3650
      TabIndex        =   39
      Top             =   3960
      Width           =   135
   End
   Begin VB.CommandButton cmdMaint 
      BackColor       =   &H00FF0000&
      Height          =   315
      Left            =   -240
      TabIndex        =   38
      Top             =   5640
      Width           =   4095
   End
   Begin VB.CommandButton cmdAddSoda 
      Caption         =   "Add Soda"
      Height          =   550
      Left            =   120
      TabIndex        =   36
      Top             =   6000
      Width           =   3375
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   6840
   End
   Begin VB.TextBox txtSomeSong 
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Text            =   "Some Snack"
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Snack!"
      Height          =   550
      Index           =   2
      Left            =   2640
      TabIndex        =   12
      Top             =   6720
      Width           =   975
   End
   Begin VB.Frame fraMachine 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000E&
      Height          =   3855
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   3600
      Begin VB.TextBox txtQtyMountainWew 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000040&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3300
         TabIndex        =   48
         Text            =   "0"
         Top             =   3300
         Width           =   300
      End
      Begin VB.TextBox txtQtyDietPepso 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000040&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3300
         TabIndex        =   47
         Text            =   "0"
         Top             =   2600
         Width           =   300
      End
      Begin VB.TextBox txtQtyPepso 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000040&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3300
         TabIndex        =   46
         Text            =   "0"
         Top             =   1870
         Width           =   300
      End
      Begin VB.TextBox txtQtyDietCola 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000040&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3300
         TabIndex        =   45
         Text            =   "0"
         Top             =   1140
         Width           =   300
      End
      Begin VB.TextBox txtQtyCola 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000040&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3300
         TabIndex        =   44
         Text            =   "0"
         Top             =   400
         Width           =   300
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   1680
         ScaleHeight     =   1358.416
         ScaleMode       =   0  'User
         ScaleWidth      =   225
         TabIndex        =   35
         Top             =   1310
         Width           =   255
         Begin VB.Image Image1 
            Height          =   1530
            Left            =   0
            Picture         =   "frmSoda.frx":000C
            Top             =   0
            Width           =   240
         End
      End
      Begin VB.TextBox txtCoinSafeTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   720
         TabIndex        =   34
         Text            =   "0.00"
         Top             =   2520
         Width           =   975
      End
      Begin VB.Frame fraSoda 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   650
         Index           =   6
         Left            =   75
         TabIndex        =   27
         Top             =   3000
         Width           =   1075
         Begin VB.CommandButton cmdSoda 
            BackColor       =   &H00404040&
            Height          =   550
            Index           =   6
            Left            =   50
            Picture         =   "frmSoda.frx":136E
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   50
            Width           =   975
         End
      End
      Begin VB.Frame fraSoda 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   650
         Index           =   0
         Left            =   75
         TabIndex        =   25
         Top             =   600
         Width           =   1075
         Begin VB.CommandButton cmdSoda 
            Height          =   550
            Index           =   0
            Left            =   50
            MaskColor       =   &H8000000F&
            Picture         =   "frmSoda.frx":3D18
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   50
            Width           =   975
         End
      End
      Begin VB.Frame fraSoda 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   650
         Index           =   5
         Left            =   2250
         TabIndex        =   18
         Top             =   3000
         Width           =   1075
         Begin VB.CommandButton cmdSoda 
            BackColor       =   &H00008000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   550
            Index           =   5
            Left            =   50
            Picture         =   "frmSoda.frx":65EE
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   50
            Width           =   975
         End
      End
      Begin VB.Frame fraSoda 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   650
         Index           =   4
         Left            =   2250
         TabIndex        =   17
         Top             =   2280
         Width           =   1075
         Begin VB.CommandButton cmdSoda 
            BackColor       =   &H00FFC0C0&
            Height          =   550
            Index           =   4
            Left            =   50
            Picture         =   "frmSoda.frx":8DF0
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   50
            Width           =   975
         End
      End
      Begin VB.Frame fraSoda 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   650
         Index           =   3
         Left            =   2250
         TabIndex        =   16
         Top             =   1560
         Width           =   1075
         Begin VB.CommandButton cmdSoda 
            BackColor       =   &H00FF8080&
            Height          =   550
            Index           =   3
            Left            =   50
            MaskColor       =   &H8000000F&
            Picture         =   "frmSoda.frx":B122
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   50
            Width           =   975
         End
      End
      Begin VB.Frame fraSoda 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   650
         Index           =   2
         Left            =   2250
         TabIndex        =   15
         Top             =   840
         Width           =   1075
         Begin VB.CommandButton cmdSoda 
            BackColor       =   &H00C0C0FF&
            Height          =   550
            Index           =   2
            Left            =   50
            Picture         =   "frmSoda.frx":DA5C
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   50
            Width           =   975
         End
      End
      Begin VB.Frame fraSoda 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   650
         Index           =   1
         Left            =   2250
         TabIndex        =   14
         Top             =   120
         Width           =   1075
         Begin VB.CommandButton Command2 
            Caption         =   "Command2"
            Height          =   615
            Left            =   1080
            TabIndex        =   37
            Top             =   0
            Width           =   75
         End
         Begin VB.CommandButton cmdSoda 
            BackColor       =   &H008080FF&
            Height          =   550
            Index           =   1
            Left            =   50
            Picture         =   "frmSoda.frx":104F2
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   50
            Width           =   975
         End
      End
      Begin VB.OptionButton OptPenny 
         BackColor       =   &H00FF0000&
         Caption         =   "Penny"
         ForeColor       =   &H0000FFFF&
         Height          =   200
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OptDollar 
         BackColor       =   &H00FF0000&
         Caption         =   "Dollar"
         ForeColor       =   &H0000FFFF&
         Height          =   200
         Left            =   240
         TabIndex        =   9
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000040&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   195
         Width           =   975
      End
      Begin VB.OptionButton optQuarter 
         BackColor       =   &H00FF0000&
         Caption         =   "Quarter"
         ForeColor       =   &H0000FFFF&
         Height          =   200
         Left            =   240
         TabIndex        =   6
         Top             =   2040
         Width           =   855
      End
      Begin VB.OptionButton optDime 
         BackColor       =   &H00FF0000&
         Caption         =   "Dime"
         ForeColor       =   &H0000FFFF&
         Height          =   200
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   855
      End
      Begin VB.OptionButton optNickel 
         BackColor       =   &H00FF0000&
         Caption         =   "Nickel"
         ForeColor       =   &H0000FFFF&
         Height          =   200
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblDollarsCount 
         Height          =   255
         Left            =   1200
         TabIndex        =   33
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label lblQuarters 
         Height          =   255
         Left            =   1200
         TabIndex        =   32
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblDimesCount 
         Height          =   255
         Left            =   1200
         TabIndex        =   31
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblNickelsCount 
         Height          =   255
         Left            =   1200
         TabIndex        =   30
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblPenniesCount 
         Height          =   255
         Left            =   1200
         TabIndex        =   29
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF0000&
         Caption         =   "Balance $ "
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   -30
         Width           =   975
      End
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   795
      Left            =   120
      ScaleHeight     =   735
      ScaleWidth      =   3315
      TabIndex        =   2
      Top             =   4800
      Width           =   3375
      Begin VB.TextBox txtRefund 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   0
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Image imgMain 
         Height          =   495
         Left            =   1080
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame fraAd 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.Line Line1 
         X1              =   3600
         X2              =   3600
         Y1              =   0
         Y2              =   840
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         Caption         =   "50¢"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   2880
         TabIndex        =   10
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblHeader 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Beverage "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   3375
      End
   End
   Begin VB.Line Line2 
      X1              =   3600
      X2              =   3600
      Y1              =   0
      Y2              =   7200
   End
End
Attribute VB_Name = "frmSoda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'//**************************************************************************
' VendingMachine.vbg - A simple flexible enterprise vending application
' by John C. Kirwin
'//**************************************************************************
' VendingMachine.vbp - A soda vending machine UI
'//**************************************************************************
' Displays the $ balance calculated from ICoinSlot interface in CSodaVend
' o   Displays adjustments to balance from deposits
'     o   Allows depositing of coins one-at-a-time
'     o   Enumerates eTender
'         ex. Pennies, Nickles, Dimes, etc...
' o   Displays transactions from purchase selections
' o   Displays adjustments to balance from refunds
'     o   Handles change in any combinations of eTender.
'     o   Indicates refund amount
'     o   Specifies coins returned
'         ex. 1 Dollar, 2 Quarters, 3 Dimes, 3 Nickels, and 4 Pennies.

Dim cDeposit As Currency
Dim intSoda As Integer
Dim iLogFile As Integer
Dim iMaint As Boolean

Dim iQtyCola As Integer
Dim iQtyDietCola As Integer
Dim iQtyPepso As Integer
Dim iQtyDietPepso As Integer
Dim iQtyMountainWew As Integer

Private Const kSodaCost As Currency = (EPrice.Soda * 0.01)
Private m_oSodaVend As CSodaVend
Private m_oCoinSlot As ICoinSlot
Private m_oCoinSafe As ICoinSafe

Private Sub Form_Load()
LogFile ("Private Sub Form_Load()")
'//**************************************************************************
' Set object references and add stock to local vending machine
'//**************************************************************************
On Error GoTo EH
    
    
    ' Initialize
    intSoda = 0
    iMaint = False
    frmSoda.Width = 3720
    frmSoda.Height = 6100
        
        
        
    Set m_oSodaVend = New CSodaVend                                       ' Instantiate
    Set m_oCoinSlot = m_oSodaVend                                         ' Query Interface 1
    Set m_oCoinSafe = m_oSodaVend                                         ' Query Interface 2

 
 'SQL Server using OLE DB Provider
 m_oSodaVend.ConnectionString = "Provider        = sqloledb;" & _
                                "Data Source     = (local);" & _
                                "Initial Catalog = pubs;" & _
                                "User Id         = sa;" & _
                                "Password        = ; "
                                
                                
                                
  
  'Add Inventory and stock the Vending Machine
  AddInventory
  AddStock
  CheckStock
Exit Sub
EH:
      MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub Deposit()
LogFile ("Private Sub Deposit()")
'//**************************************************************************

'//**************************************************************************
On Error GoTo EH

    If OptPenny = True Then                                               ' Penny
        cDeposit = eTender.Pennies * 0.01
        m_oSodaVend.ICoinSafe_PenniesCount = 1
      
      ElseIf optNickel = True Then                                        ' Nickel
        cDeposit = eTender.Nickles * 0.01
        m_oSodaVend.ICoinSafe_NickelsCount = 1
      
      ElseIf optDime = True Then                                          ' Dime
        cDeposit = eTender.Dimes * 0.01
        m_oSodaVend.ICoinSafe_DimesCount = 1
      
      ElseIf optQuarter = True Then                                       ' Quarter
        cDeposit = eTender.Quarters * 0.01
        m_oSodaVend.ICoinSafe_QuartersCount = 1
      
      ElseIf OptDollar = True Then                                        ' Dollar
        cDeposit = eTender.Dollar * 0.01
        m_oSodaVend.ICoinSafe_DollarsCount = 1
    
    End If

    m_oSodaVend.ICoinSlot_Deposit = cDeposit

Exit Sub
EH:
      MsgBox Err.Number & ": " & Err.Description

End Sub


Private Sub cmdSoda_Click(Index As Integer)
LogFile ("Private Sub cmdSoda_Click(Index As Integer)")
'//**************************************************************************
' Handle Soda choice click
'//**************************************************************************
On Error GoTo EH

RemoveSoda

    Select Case Index
      Case 0                                                              ' Deposit
        fraSoda(0).BackColor = &HFFFF&
        EmptyCoinReturn
        Deposit
        Balance
        CoinCount
      Case 1                                                              ' Cola
        fraSoda(1).BackColor = &HFFFF&
        MakeSelection (Index)
      Case 2                                                              ' DietCola
        fraSoda(2).BackColor = &HFFFF&
        MakeSelection (Index)

      Case 3                                                              ' Pepso
        fraSoda(3).BackColor = &HFFFF&
        MakeSelection (Index)

      Case 4                                                              ' DietPepso
        fraSoda(4).BackColor = &HFFFF&
        MakeSelection (Index)

      Case 5                                                              ' MountainWew
        fraSoda(5).BackColor = &HFFFF&
        MakeSelection (Index)

      Case 6                                                              ' Refund
        fraSoda(6).BackColor = &HFFFF&
        Refund

      Case Else
        Debug.Assert False
    End Select

CheckStock

Exit Sub
EH:
      MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub Balance()
LogFile ("Private Sub Balance()")
'//**************************************************************************
' Display Balance
'//**************************************************************************
On Error GoTo EH

    txtBalance.Text = Format$(m_oSodaVend.ICoinSlot_Balance, "$#,##0.00") ' Display Formatted Balance


Exit Sub
EH:
      MsgBox Err.Number & ": " & Err.Description

End Sub
Private Sub CoinCount()
LogFile ("Private Sub CoinCount()")
'//**************************************************************************

'//**************************************************************************
On Error GoTo EH

    ' Display Coin Counts
    lblPenniesCount.Caption = m_oCoinSafe.PenniesCount
    lblNickelsCount.Caption = m_oCoinSafe.NickelsCount
    lblDimesCount.Caption = m_oCoinSafe.DimesCount
    lblQuarters.Caption = m_oCoinSafe.QuartersCount
    lblDollarsCount.Caption = m_oCoinSafe.DollarsCount
    
    txtCoinSafeTotal.Text = Format$(m_oSodaVend.CoinSafeTotal, "$#,##0.00") ' Display Formatted CoinSafe Total
    

Exit Sub
EH:
      MsgBox Err.Number & ": " & Err.Description
    
End Sub
Private Sub MakeSelection(intSoda)
LogFile ("Private Sub MakeSelection(intSoda)")
'//**************************************************************************

'//**************************************************************************
On Error GoTo EH
  
  ' Clear the Coin return
    EmptyCoinReturn

    If m_oSodaVend.ICoinSlot_Balance >= kSodaCost Then
        m_oCoinSlot.BuyIt intSoda
        DropSoda (intSoda)
        
        Balance
        CoinCount
    Else
        Exit Sub
    End If

Exit Sub
EH:
      MsgBox Err.Number & ": " & Err.Description
      
End Sub

Private Sub DropSoda(intSoda)
LogFile ("Private Sub DropSoda(intSoda)")
'//**************************************************************************
' Display image to indicate which soda Dropped
'//**************************************************************************
On Error GoTo EH

    Select Case intSoda
      Case 0                                                              ' Doposit drops nothing

      Case 1                                                              ' Cola
        imgMain.Picture = LoadPicture("Cola.bmp")

      Case 2                                                              ' DietCola
        imgMain.Picture = LoadPicture("DietCola.bmp")

      Case 3                                                              ' Pepso
        imgMain.Picture = LoadPicture("Pepso.bmp")

      Case 4                                                              ' DietPepso
        imgMain.Picture = LoadPicture("DietPepso.bmp")

      Case 5                                                              ' MountainWew
        imgMain.Picture = LoadPicture("MountainWew.bmp")

      Case 6                                                              ' Refund Coins

      Case Else
        Debug.Assert False

    End Select


Exit Sub
EH:
      MsgBox Err.Number & ": " & Err.Description

End Sub
Private Sub Refund()
LogFile ("Private Sub Refund()")
'//**************************************************************************

'//**************************************************************************
On Error GoTo EH

    txtRefund.Visible = True
    
    intSoda = 0
    
    txtRefund.Text = Format$(m_oCoinSlot.Refund, "$#,##0.00")
    txtBalance.Text = Format$(m_oSodaVend.ICoinSlot_Balance, "$#,##0.00")
   
    CoinCount

Exit Sub
EH:
      MsgBox Err.Number & ": " & Err.Description
      
End Sub
Private Sub EmptyCoinReturn()
LogFile ("Private Sub EmptyCoinReturn()")
'//**************************************************************************
' Empty Coin Return
'//**************************************************************************
On Error GoTo EH

  ' Clear money from Coin Return
    txtRefund.Text = 0#
    txtRefund.Visible = False

Exit Sub
EH:
      MsgBox Err.Number & ": " & Err.Description
      
End Sub
Private Sub Change()
LogFile ("Private Sub Change()")
'//**************************************************************************
' Handle Change
'//**************************************************************************
On Error GoTo EH
    
    txtRefund.Visible = True
    intSoda = 0

Exit Sub
EH:
      MsgBox Err.Number & ": " & Err.Description

End Sub
Private Sub RemoveSoda()
LogFile ("Private Sub RemoveSoda()")
'//**************************************************************************
' Indicate Soda removal
'//**************************************************************************
On Error GoTo EH
    
    imgMain.Picture = Nothing

Exit Sub
EH:
      MsgBox Err.Number & ": " & Err.Description

End Sub
Private Sub Command1_Click(Index As Integer)
LogFile ("Private Sub Command1_Click(Index As Integer)")
'//**************************************************************************

'//**************************************************************************
On Error GoTo EH
    
    Select Case Index
      Case 0                                                              ' Deposit
        Deposit
      Case 1                                                              ' Refund
        Refund
      Case 2                                                              ' Commit
        MakeSelection (intSoda)
      Case Else
        Debug.Assert False
    End Select

    Balance

Exit Sub
EH:
      MsgBox Err.Number & ": " & Err.Description

End Sub
Private Sub Timer1_Timer()
'LogFile ("Private Sub Timer1_Timer()")
'//**************************************************************************
' Timer
'//**************************************************************************

   
   Dim i As Integer
   'Reset Selection Indicator Frames
   For i = 0 To 6
    fraSoda(i).BackColor = &HFF0000
   Next i

Exit Sub
EH:
      MsgBox Err.Number & ": " & Err.Description

End Sub
Private Sub cmdAddSoda_Click()
Dim SQL As String
Dim sSoda As String
On Error GoTo EH

sSoda = InputBox("What soda would like like to add to inventory?")

SQL = "INSERT Beverages (vcDrink, mPrice) VALUES ('" & sSoda & "', .50)"

 m_oSodaVend.Insert (SQL)
 
Exit Sub
EH:
      MsgBox Err.Number & ": " & Err.Description

End Sub

Public Sub LogFile(Message As String)
Debug.Print Message
'//**************************************************************************

'//**************************************************************************

 Dim LogFile As Integer
 LogFile = FreeFile
 Open "C:\LogFile.log" For Append As #LogFile
 Print #LogFile, Message
 Close #LogFile

Exit Sub
EH:
      MsgBox Err.Number & ": " & Err.Description

End Sub
Private Sub fraMachine_Click()
LogFile ("Private Sub fraMachine_Click()")
'//**************************************************************************

'//**************************************************************************
On Error GoTo EH

    RemoveSoda
    EmptyCoinReturn

Exit Sub
EH:
      MsgBox Err.Number & ": " & Err.Description

End Sub
Private Sub imgMain_Click()
LogFile ("Private Sub imgMain_Click()")
'//**************************************************************************

'//**************************************************************************
On Error GoTo EH

    RemoveSoda
    EmptyCoinReturn

Exit Sub
EH:
      MsgBox Err.Number & ": " & Err.Description

End Sub
Private Sub picMain_Click()
LogFile ("Private Sub picMain_Click()")
'//**************************************************************************

'//**************************************************************************
On Error GoTo EH

    RemoveSoda
    EmptyCoinReturn

Exit Sub
EH:
      MsgBox Err.Number & ": " & Err.Description
      
End Sub


Private Sub cmdMaint_Click()
LogFile ("Private Sub cmdMaint_Click()")
'//**************************************************************************
' Soda Machine Maintenance
'//**************************************************************************
Dim sKey As String
On Error GoTo EH

Select Case iMaint

Case 0
    Debug.Print "False"
    sKey = InputBox("What is the pin/password to open this Vending Machine for maintenance?")
    ' Check password
    If sKey = "" Then
        iMaint = True
        frmSoda.Width = 3900
        frmSoda.Height = 6200
    ElseIf sKey <> " " Then
        Exit Sub
    End If

Case 1
    Debug.Print "True"
    iMaint = False
    frmSoda.Width = 3720
    frmSoda.Height = 6100

End Select

Exit Sub
EH:
      MsgBox Err.Number & ": " & Err.Description

End Sub


Private Sub cmdRestock1_Click()
LogFile ("Private Sub cmdRestock1_Click()")
'//**************************************************************************
' Restock
'//**************************************************************************
Dim SQL As String
On Error GoTo EH

SQL = "INSERT TranLog (iItemID, iQuantity, mAmount, vcTranComment) VALUES (1, 10 , 0, 'Stock 10 Cola')"
      
m_oSodaVend.Insert (SQL)

CheckStock

Exit Sub
EH:
      MsgBox Err.Number & ": " & Err.Description

End Sub


Private Sub cmdRestock2_Click()
LogFile ("Private Sub cmdRestock2_Click()")
'//**************************************************************************
' Restock
'//**************************************************************************
Dim SQL As String
On Error GoTo EH

SQL = "INSERT TranLog (iItemID, iQuantity, mAmount, vcTranComment) VALUES (2, 10 , 0, 'Stock 10 DietCola')"
      
m_oSodaVend.Insert (SQL)

CheckStock

Exit Sub
EH:
      MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub cmdRestock3_Click()
LogFile ("Private Sub cmdRestock3_Click()")
'//**************************************************************************
' Restock
'//**************************************************************************
Dim SQL As String
On Error GoTo EH

SQL = "INSERT TranLog (iItemID, iQuantity, mAmount, vcTranComment) VALUES (3, 10 , 0, 'Stock 10 Pepso')"
      
m_oSodaVend.Insert (SQL)

CheckStock

Exit Sub
EH:
      MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub cmdRestock4_Click()
LogFile ("Private Sub cmdRestock4_Click()")
'//**************************************************************************
' Restock
'//**************************************************************************
Dim SQL As String
On Error GoTo EH

SQL = "INSERT TranLog (iItemID, iQuantity, mAmount, vcTranComment) VALUES (4, 10 , 0, 'Stock 10 DietPepso')"
      
m_oSodaVend.Insert (SQL)

CheckStock

Exit Sub
EH:
      MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub cmdRestock5_Click()
LogFile ("Private Sub cmdRestock5_Click()")
'//**************************************************************************
' Restock
'//**************************************************************************
Dim SQL As String
On Error GoTo EH

SQL = "INSERT TranLog (iItemID, iQuantity, mAmount, vcTranComment) VALUES (5, 10 , 0, 'Stock 10 MountainWew')"
      
m_oSodaVend.Insert (SQL)

CheckStock

Exit Sub
EH:
      MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub AddStock()
LogFile ("Private Sub AddStock()")
'//**************************************************************************
' Initialize Stock to 10 each at startup
'//**************************************************************************
 Dim SQL As String
 On Error GoTo EH
  
  SQL = "INSERT TranLog (iItemID, iQuantity, mAmount, vcTranComment) VALUES (1, 10 , 0, 'Stock 10 Cola')"
  m_oSodaVend.Insert (SQL)
  SQL = "INSERT TranLog (iItemID, iQuantity, mAmount, vcTranComment) VALUES (2, 10 , 0, 'Stock 10 DietCola')"
  m_oSodaVend.Insert (SQL)
  SQL = "INSERT TranLog (iItemID, iQuantity, mAmount, vcTranComment) VALUES (3, 10 , 0, 'Stock 10 Pepso')"
  m_oSodaVend.Insert (SQL)
  SQL = "INSERT TranLog (iItemID, iQuantity, mAmount, vcTranComment) VALUES (4, 10 , 0, 'Stock 10 DietPepso')"
  m_oSodaVend.Insert (SQL)
  SQL = "INSERT TranLog (iItemID, iQuantity, mAmount, vcTranComment) VALUES (5, 10 , 0, 'Stock 10 MountainWew')"
  m_oSodaVend.Insert (SQL)

Exit Sub
EH:
      MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub AddInventory()
LogFile ("Private Sub AddInventory()")
'//**************************************************************************
' Add inventory
'//**************************************************************************
 Dim SQL As String
 On Error GoTo EH
  SQL = "INSERT Beverages (vcDrink, mPrice) VALUES ('Cola', .50)"
  m_oSodaVend.Insert (SQL)
  SQL = "INSERT Beverages (vcDrink, mPrice) VALUES ('DietCola', .50)"
  m_oSodaVend.Insert (SQL)
  SQL = "INSERT Beverages (vcDrink, mPrice) VALUES ('Pepso', .50)"
  m_oSodaVend.Insert (SQL)
  SQL = "INSERT Beverages (vcDrink, mPrice) VALUES ('DietPepso', .50)"
  m_oSodaVend.Insert (SQL)
  SQL = "INSERT Beverages (vcDrink, mPrice) VALUES ('MountainWew', .50)"
  m_oSodaVend.Insert (SQL)

Exit Sub
EH:
      MsgBox Err.Number & ": " & Err.Description

End Sub

Private Sub CheckStock()
'//**************************************************************************
' Check stock
'//**************************************************************************
Dim sSQL As String


sSQL = "SELECT SUM(iQuantity) 'iReturn' FROM TranLog WHERE iItemID = 1"
iQtyCola = m_oSodaVend.CheckStock(sSQL)

sSQL = "SELECT SUM(iQuantity) 'iReturn' FROM TranLog WHERE iItemID = 2"
iQtyDietCola = m_oSodaVend.CheckStock(sSQL)

sSQL = "SELECT SUM(iQuantity) 'iReturn' FROM TranLog WHERE iItemID = 3"
iQtyPepso = m_oSodaVend.CheckStock(sSQL)

sSQL = "SELECT SUM(iQuantity) 'iReturn' FROM TranLog WHERE iItemID = 4"
iQtyDietPepso = m_oSodaVend.CheckStock(sSQL)

sSQL = "SELECT SUM(iQuantity) 'iReturn' FROM TranLog WHERE iItemID = 5"
iQtyMountainWew = m_oSodaVend.CheckStock(sSQL)

txtQtyCola.Text = iQtyCola
txtQtyDietCola.Text = iQtyDietCola
txtQtyPepso.Text = iQtyPepso
txtQtyDietPepso.Text = iQtyDietPepso
txtQtyMountainWew.Text = iQtyMountainWew
End Sub
