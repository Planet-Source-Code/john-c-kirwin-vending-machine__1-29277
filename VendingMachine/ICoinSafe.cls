VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICoinSafe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'==========================================================================
' ICoinSafe - Coinbox like count of coins
'==========================================================================


'==========================================================================
' Private Variables
'==========================================================================
Private iPennies As Integer
Private iNickels As Integer
Private iDimes As Integer
Private iQuarters As Integer
Private iDollars As Integer
Private Property Let PenniesCount(iCoinCount As Integer)
'==========================================================================
' Pennies
'==========================================================================
    
    iPennies = iPennies + iCoinCount

End Property
Public Property Get PenniesCount() As Integer
'==========================================================================
' Return Count of Pennies
'==========================================================================
    
    PenniesCount = iPennies

End Property
Private Property Let NickelsCount(iCoinCount As Integer)
'==========================================================================
' Nickels
'==========================================================================
    
    iNickels = iNickels + iCoinCount

End Property
Public Property Get NickelsCount() As Integer
'==========================================================================
' Return Count of Nickels
'==========================================================================
    
    NickelsCount = iNickels

End Property
Private Property Let DimesCount(iCoinCount As Integer)
'==========================================================================
' Dimes
'==========================================================================
    
    iDimes = iDimes + iCoinCount

End Property
Public Property Get DimesCount() As Integer
'==========================================================================
' Return Count of Dimes
'==========================================================================
    
    DimesCount = iDimes

End Property


Private Property Let QuartersCount(iCoinCount As Integer)
'==========================================================================
' Quarters
'==========================================================================
    
    iQuarters = iQuarters + iCoinCount

End Property
Public Property Get QuartersCount() As Integer
'==========================================================================
' Return Count of Quarters
'==========================================================================
    
    QuartersCount = iQuarters

End Property
Private Property Let DollarsCount(iCoinCount As Integer)
'==========================================================================
' Dollars
'==========================================================================
    
    iDollars = iDollars + iCoinCount

End Property
Public Property Get DollarsCount() As Integer
'==========================================================================
' Return Count of Dollars
'==========================================================================
    
    DollarsCount = iDollars

End Property
