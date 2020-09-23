VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSodaVend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'==========================================================================
' SodaVend.vbp - Middle-tier business logic for vending
'==========================================================================
' CSodaVend
'     o   Implements IcoinSlot & ICoinSafe
' ICoinSlot Interface
'     o   Handles Deposit, Refunds and tracks Balance
'     o   cAmount As Currency
'     o   Commit as Boolean
' ICoinSafe Interface
'     o   Maintains count of eTenders

'==========================================================================
' Implements ICoinSlot & ICoinSafe Interface Classes
'==========================================================================
Implements ICoinSlot
Implements ICoinSafe

Private m_oCoinSlot As ICoinSlot
Private m_oCoinSafe As ICoinSafe

Private m_PenniesCount As Integer
Private m_NickelsCount As Integer
Private m_DimesCount As Integer
Private m_QuartersCount As Integer
Private m_DollarsCount As Integer


'==========================================================================
' Enumerations
'==========================================================================
Public Enum eTender                                                       ' Enumeration eTender
    Pennies = 1                                                           ' = 1     Penny
    Nickles = 5                                                           ' = 5     Pennies
    Dimes = 10                                                            ' = 10    Pennies
    Quarters = 25                                                         ' = 25    Pennies
    Dollar = 100                                                          ' = 100   Pennies
End Enum

Public Enum EPrice                                                        ' Enumeration EPrice
    Soda = 50                                                             ' Price   cents
End Enum

'==========================================================================
' Private Variables
'==========================================================================
Private mcAmount As Currency
Private m_iSoda As Integer
Private iCoinCount As Integer

'==========================================================================
' Constants
'==========================================================================
Private Const kPriceSoda As Currency = EPrice.Soda * 0.01                 ' Price of Soda from Enumeration EPrice
Private Const kPennies As Currency = eTender.Pennies * 0.01               ' 1       cent  from Enumeration eTender
Private Const kNickels As Currency = eTender.Nickles * 0.01               ' 5       cents from Enumeration eTender
Private Const kDimes  As Currency = eTender.Dimes * 0.01                  ' 10      cents from Enumeration eTender
Private Const kQuarters  As Currency = eTender.Quarters * 0.01            ' 25      cents from Enumeration eTender
Private Const kDollars  As Currency = eTender.Dollar * 0.01               ' 100     cents from Enumeration eTender


'==========================================================================
' ADO Utilities Class - Wrapper for ActiveX Data Objects 2.5 Library
'                       database access/manipulation.
' "Microsoft ActiveX Data Objects 2.5 Library" in project and declare the class.
'==========================================================================

Private m_sConnectionString As String
Private m_sLastError As String
Private m_bIsSQL As Boolean
Public Property Get ConnectionString() As String
'==========================================================================
' Get ConnectionString
'==========================================================================
    ConnectionString = m_sConnectionString
End Property

Public Property Let ConnectionString(ByVal ConnString As String)
'==========================================================================
' Let ConnectionString
'
' SQL Server using an ODBC Data Source
'm_sConnectionString = "DSN=pubs SQLServer;" & _
'                      "UID=sa;PWD=;" & _
'                      "DATABASE=Pubs"
'
'SQL Server using OLE DB Provider
'm_sConnectionString = "Provider        = sqloledb;" & _
'                      "Data Source     = (local);" & _
'                      "Initial Catalog = pubs;" & _
'                      "User Id         = sa;" & _
'                      "Password        = ; "
'
' Jet MDB
'm_sConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51" & _
'                      ";Data Source=..\Soda.mdb"
'==========================================================================


    
    m_sConnectionString = ConnString
    
End Property
Public Function Insert(SQL As String) As Long
'==========================================================================
' Insert using ADO into the server, database, and table indicated
' by the ConnectionString
'
' Returns the value of IDENTITY COLUMN of a newly inserted record
' SQL is a valid Insert statement.
' ConnetionString property has been set to a valid Connection String
'==========================================================================

On Error GoTo LocalError
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim lIdentity As Long
' Assume failure
lIdentity = 0


With rs

'Prepare the RecordSet
.CursorLocation = adUseServer
 .CursorType = adOpenForwardOnly
 .LockType = adLockReadOnly
 .Source = "SELECT @@Identity"
End With

With cn
 .ConnectionString = ConnectionString
 .CursorLocation = adUseServer
 .Open
 .BeginTrans
 .Execute SQL, , adCmdText + adExecuteNoRecords

        With rs

        .ActiveConnection = cn
        .Open , , , , adCmdText
        lIdentity = rs(0).Value
        .Close
        End With
 .CommitTrans
   Debug.Print lIdentity
   
 .Close
End With
Set rs = Nothing
Set cn = Nothing


Exit Function


LocalError:
 Debug.Print Err.Number & " - " & Err.Description

 lIdentity = 0

End Function
Public Function CheckStock(SQL As String) As Long
'==========================================================================
' Insert using ADO into the server, database, and table indicated
' by the ConnectionString
'
' Returns the value of IDENTITY COLUMN of a newly inserted record
' SQL is a valid Insert statement.
' ConnetionString property has been set to a valid Connection String
'==========================================================================

On Error GoTo LocalError
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim iCheckStock As Integer

    
    ' Set the connection
    cn.ConnectionString = m_sConnectionString
    ' Set the mode of the connection
    cn.Mode = adModeReadWrite
    ' Set Cursor location to client
    cn.CursorLocation = adUseClient
    cn.Open

Set rs = New Recordset
    rs.Open SQL, cn, adOpenForwardOnly, adLockReadOnly

    iCheckStock = rs.Fields("iReturn").Value
    CheckStock = iCheckStock
 rs.Close

Set rs = Nothing
Set cn = Nothing

Exit Function

LocalError:
 Debug.Print Err.Number & " - " & Err.Description

 CheckStock = 0

End Function

Public Function CoinSafeTotal()
'==========================================================================
' Calculate CoinSafe total
'==========================================================================
Dim cCoinSafeCash As Currency

cCoinSafeCash = (m_DollarsCount * kDollars) + (m_QuartersCount * kQuarters) + (m_DimesCount * kDimes) + (m_NickelsCount * kNickels) + (m_PenniesCount * kPennies)

CoinSafeTotal = cCoinSafeCash
End Function
Public Function ICoinSlot_Refund() As Currency
Debug.Print "Public Function ICoinSlot_Refund() As Currency"
'==========================================================================
' Handle Transaction Refunds indicating total refund and coins returned
'==========================================================================
Dim iDollarsReturn As Integer
Dim iQuartersReturn As Integer
Dim iDimesReturn As Integer
Dim iNickelsReturn As Integer
Dim iPenniesReturn As Integer
Dim cRefAmt As Currency
Dim cCoinReturn As Currency
    
    cRefAmt = m_oCoinSlot.Refund
    cCoinReturn = cRefAmt
    
    
    If cCoinReturn > 0.99 Then                                            ' Dollars to Return
    iDollarsReturn = Int(cCoinReturn)
    cCoinReturn = cCoinReturn - (iDollarsReturn * kDollars)
    End If

    If cCoinReturn > 0.24 Then                                            ' Quarters to Return
    iQuartersReturn = Int(cCoinReturn / kQuarters)
    cCoinReturn = cCoinReturn - (iQuartersReturn * kQuarters)
    End If

    If cCoinReturn > 0.09 Then                                            ' Dimes to Return
    iDimesReturn = Int(cCoinReturn / kDimes)
    cCoinReturn = cCoinReturn - (iDimesReturn * kDimes)
    End If

    If cCoinReturn > 0.04 Then                                            ' Nickels to Return
    iNickelsReturn = Int(cCoinReturn / kNickels)
    cCoinReturn = cCoinReturn - (iNickelsReturn * kNickels)
    End If
    
    iPenniesReturn = Int(cCoinReturn * 100)                               ' Pennies to Return


    ' Drop Coins
    m_DollarsCount = m_DollarsCount - iDollarsReturn
    m_QuartersCount = m_QuartersCount - iQuartersReturn
    m_DimesCount = m_DimesCount - iDimesReturn
    m_NickelsCount = m_NickelsCount - iNickelsReturn
    m_PenniesCount = m_PenniesCount - iPenniesReturn
    
    ' Indicate Coins Dropped
    MsgBox "Refund " & cRefAmt & " as " & iDollarsReturn & " Dollars " & iQuartersReturn & _
    " Quarters " & iDimesReturn & " Dimes " & iNickelsReturn & " Nickels " & iPenniesReturn & " Pennies "
    
    ICoinSlot_Refund = cRefAmt
    
End Function
Public Property Let ICoinSafe_PenniesCount(iCoinCount As Integer)
Debug.Print "Private Property Let ICoinSafe_PenniesCount(iCoinCount As Integer)"
'==========================================================================

'==========================================================================

    m_PenniesCount = iCoinCount + m_PenniesCount
End Property
Public Property Get ICoinSafe_PenniesCount() As Integer
Debug.Print "Private Property Get ICoinSafe_PenniesCount() As Integer"
'==========================================================================

'==========================================================================

    ICoinSafe_PenniesCount = m_PenniesCount
End Property
Public Property Let ICoinSafe_NickelsCount(iCoinCount As Integer)
Debug.Print "Private Property Let ICoinSafe_NickelsCount(iCoinCount As Integer)"
'==========================================================================

'==========================================================================

    m_NickelsCount = iCoinCount + m_NickelsCount
End Property
Public Property Get ICoinSafe_NickelsCount() As Integer
Debug.Print "Private Property Get ICoinSafe_NickelsCount() As Integer"
'==========================================================================

'==========================================================================

    ICoinSafe_NickelsCount = m_NickelsCount
End Property
Public Property Let ICoinSafe_DimesCount(iCoinCount As Integer)
Debug.Print "Private Property Let ICoinSafe_DimesCount(iCoinCount As Integer)"
'==========================================================================

'==========================================================================

    m_DimesCount = iCoinCount + m_DimesCount
End Property
Public Property Get ICoinSafe_DimesCount() As Integer
Debug.Print "Private Property Get ICoinSafe_DimesCount() As Integer"
'==========================================================================

'==========================================================================

    ICoinSafe_DimesCount = m_DimesCount
End Property
Public Property Let ICoinSafe_QuartersCount(iCoinCount As Integer)
Debug.Print "Private Property Let ICoinSafe_QuartersCount(iCoinCount As Integer)"
'==========================================================================

'==========================================================================

    m_QuartersCount = iCoinCount + m_QuartersCount
End Property
Public Property Get ICoinSafe_QuartersCount() As Integer
Debug.Print "Private Property Get ICoinSafe_QuartersCount() As Integer"
'==========================================================================

'==========================================================================

    ICoinSafe_QuartersCount = m_QuartersCount
End Property
Public Property Let ICoinSafe_DollarsCount(iCoinCount As Integer)
Debug.Print "Private Property Let ICoinSafe_DollarsCount(iCoinCount As Integer)"
'==========================================================================

'==========================================================================

    m_DollarsCount = iCoinCount + m_DollarsCount
End Property
Public Property Get ICoinSafe_DollarsCount() As Integer
Debug.Print "Private Property Get ICoinSafe_DollarsCount() As Integer"
'==========================================================================

'==========================================================================

    ICoinSafe_DollarsCount = m_DollarsCount
End Property
Public Property Get ICoinSlot_Balance() As Currency
Debug.Print "Public Property Get ICoinSlot_Balance() As Currency"
'==========================================================================
'
'==========================================================================
    ICoinSlot_Balance = m_oCoinSlot.Balance
End Property
Public Property Let ICoinSlot_Deposit(cAmount As Currency)
Debug.Print "Public Property Let ICoinSlot_Deposit(cAmount As Currency)"
'==========================================================================

'==========================================================================

    m_oCoinSlot.Deposit = cAmount

End Property
Public Function cBalance(m_cBalance As Currency, cDeposit As Currency) As Currency
Debug.Print "Public Function cBalance(m_cBalance As Currency, cDeposit As Currency) As Currency"
'==========================================================================
' Calculates Balance after a deposit
'==========================================================================
  
  m_cBalance = m_cBalance + cDeposit

    cBalance = m_cBalance - cDeposit

End Function
Public Function ICoinSlot_BuyIt(ByVal iSoda As Integer) As Boolean
Debug.Print "Public Function ICoinSlot_BuyIt(ByVal iSoda As Integer) As Boolean"
'==========================================================================
' Returns false if not enough money to make purchase
'==========================================================================
Dim SQL As String
    
    ICoinSlot_BuyIt = False                                               ' Assume Failure

    If m_oCoinSlot.Commit(kPriceSoda) Then
        m_iSoda = iSoda
        ICoinSlot_BuyIt = True                                            ' Approved

        
      SQL = "INSERT TranLog (iItemID, iQuantity, mAmount, vcTranComment) VALUES (" & iSoda & ", -1 , .50, 'Purchase')"

      Insert (SQL)
      
      Else
        Debug.Assert False                                                ' Not enough money!
    End If

End Function
Private Sub Class_Initialize()
Debug.Print "Private Sub Class_Initialize()"
'==========================================================================
' Initialize
'==========================================================================

    Set m_oCoinSlot = New ICoinSlot
    Set m_oCoinSafe = New ICoinSafe

End Sub
Private Sub Class_Terminate()
Debug.Print "Private Sub Class_Terminate()"
'==========================================================================
' Cleanup
'==========================================================================

    Set m_oCoinSlot = Nothing
    Set m_oCoinSafe = Nothing

End Sub
