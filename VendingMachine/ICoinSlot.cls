VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICoinSlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'==========================================================================
' IcoinSlot - Coin input and output handling of transactions
'==========================================================================

Private m_cBalance As Currency
Private m_iSoda As String
Public Function BuyIt(ByVal iSoda As Integer) As Boolean
'==========================================================================
' Indicate and identify purchase
'==========================================================================
Debug.Print "Public Function BuyIt(ByVal iSoda As Integer) As Boolean"
    m_iSoda = iSoda

End Function
Friend Function Commit(cAmount As Currency) As Boolean
Debug.Print "Friend Function Commit(cAmount As Currency) As Boolean"
'==========================================================================
' Commit the transaction amount
'==========================================================================

    Commit = False
    If (m_cBalance >= cAmount) Then
        m_cBalance = m_cBalance - cAmount
        Commit = True
      Else
        Debug.Assert False                                                ' Not enough
    End If

End Function
Public Property Get Balance() As Currency
Debug.Print "Public Property Get Balance() As Currency"
'==========================================================================
' Balance
'==========================================================================

    Balance = m_cBalance

End Property
Public Property Let Deposit(cAmount As Currency)
Debug.Print "Public Property Let Deposit(cAmount As Currency)"
'==========================================================================
' Deposit
'==========================================================================

    m_cBalance = m_cBalance + cAmount

End Property
Public Function Refund() As Currency
Debug.Print "Public Function Refund() As Currency"
'==========================================================================
' Refund
'==========================================================================
    
    Refund = m_cBalance
    m_cBalance = 0

End Function
Private Sub Class_Initialize()
Debug.Print "Private Sub Class_Initialize()"
'==========================================================================
' Initialize ICoinSlot
'==========================================================================

    m_cBalance = 0

End Sub







