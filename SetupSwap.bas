Sub SetUpSwapAutoITv0()

Dim CRow As Integer

Dim NearDate As Date
Dim FarDate As Date
Dim clientName As String
Dim MMRef As String
Dim BuySell As String
Dim BaseAmt As Double
Dim BaseCcy As String
Dim CounterCcy As String
Dim CIF As String
Dim CcyPair As String
Dim Rate As String

CRow = ActiveCell.Row
If rows(CRow).Hidden Then
    MsgBox "Row is Hidden"
    Exit Sub
End If

NearDate = Cells(CRow, 1).Value
clientName = Cells(CRow, 2).Value
MMRef = Cells(CRow, 3).Value
BuySell = Cells(CRow, 6).Value
BaseAmt = Abs(Cells(CRow, 7).Value)
BaseCcy = Cells(CRow, 8).Value
CounterCcy = Cells(CRow, 10).Value
Rate = Cells(CRow, 11).Value
Dim VLDeets As String
Dim SpreadPip As Double
Dim SpreadBP As Double
CcyPair = BaseCcy & CounterCcy

CIF = WorksheetFunction.Index(Sheets("Setup").Range("C2:C200"), _
        WorksheetFunction.Match(clientName, Sheets("Setup").Range("B2:B200"), 0))

FarDate = WorksheetFunction.Index(Sheets("Setup").Range("N2:N200"), _
            WorksheetFunction.Match(clientName & BaseCcy & CounterCcy, Sheets("Setup").Range("R2:R200"), 0))
            
If FarDate < Date Then
MsgBox "Far Date looks wrong"
Exit Sub
End If
            
            
VLDeets = WorksheetFunction.Index(Sheets("Setup").Range("F2:F200"), _
        WorksheetFunction.Match(clientName, Sheets("Setup").Range("B2:B200"), 0))

SpreadPip = WorksheetFunction.Index(Sheets("Setup").Range("G2:G200"), _
        WorksheetFunction.Match(clientName, Sheets("Setup").Range("B2:B200"), 0))
        
SpreadBP = WorksheetFunction.Index(Sheets("Setup").Range("G2:G200"), _
        WorksheetFunction.Match(clientName, Sheets("Setup").Range("B2:B200"), 0))

'FarDate = WorksheetFunction.Index(Sheets("Setup").Range(OfficeYCol & "2:G50"), _
        WorksheetFunction.Match(1, (ClientName = Sheets("Setup").Range("E2:E50")) * _
                                    (BaseCcy & CounterCcy = Sheets("Setup").Range(OfficeXCol & "2:F50")), 0))
'FarDate = WorksheetFunction.Index(Sheets("Setup").Range(OfficeYCol & "2:G50"), WorksheetFunction.Match(ClientName & BaseCcy & CounterCcy, Sheets("Setup").Range("I2:I50"), 0))
If IsError(FarDate) Then
    MsgBox "Problem with FarDate"
    Exit Sub
End If

Dim RootX As Integer, RootY As Integer
Dim OfficeColOffset As Integer

If Sheets("Setup").Range("AA2").Value2 = "Office" Then
    OfficeColOffset = 0
Else
    OfficeColOffset = 2
End If
Dim OfficeXCol As String
Dim OfficeYCol As String
OfficeXCol = "AB"
OfficeYCol = "AC"

Dim SwapX As Integer, SwapY As Integer
SwapX = Sheets("Setup").Range(OfficeXCol & "5").Offset(0, OfficeColOffset).Value2
SwapY = Sheets("Setup").Range(OfficeYCol & "5").Offset(0, OfficeColOffset).Value2

Dim CIFX As Integer, CIFY As Integer
CIFX = Sheets("Setup").Range(OfficeXCol & "6").Offset(0, OfficeColOffset).Value2
CIFY = Sheets("Setup").Range(OfficeYCol & "6").Offset(0, OfficeColOffset).Value2

Dim CcyPairX As Integer, CcyPairY As Integer
CcyPairX = Sheets("Setup").Range(OfficeXCol & "7").Offset(0, OfficeColOffset).Value2
CcyPairY = Sheets("Setup").Range(OfficeYCol & "7").Offset(0, OfficeColOffset).Value2

Dim CPDropDownX As Integer, CPDropDownY As Integer
CPDropDownX = Sheets("Setup").Range(OfficeXCol & "8").Offset(0, OfficeColOffset).Value2
CPDropDownY = Sheets("Setup").Range(OfficeYCol & "8").Offset(0, OfficeColOffset).Value2

Dim DecisionMakerClickX As Integer, DecisionMakerClickY As Integer
DecisionMakerClickX = Sheets("Setup").Range(OfficeXCol & "38").Offset(0, OfficeColOffset).Value2
DecisionMakerClickY = Sheets("Setup").Range(OfficeYCol & "38").Offset(0, OfficeColOffset).Value2

Dim DecisionMakerDD As Integer
DecisionMakerDD = WorksheetFunction.Index(Sheets("Setup").Range("Q2:Q200"), _
    WorksheetFunction.Match(clientName & BaseCcy & CounterCcy, Sheets("Setup").Range("R2:R200"), 0))
    
Dim DecisionMakerDDX As Integer, DecisionMakerDDY As Integer

        DecisionMakerDDX = Sheets("Setup").Range(OfficeXCol & "38").Offset(DecisionMakerDD, OfficeColOffset).Value2
        DecisionMakerDDY = Sheets("Setup").Range(OfficeYCol & "38").Offset(DecisionMakerDD, OfficeColOffset).Value2
'Select Case DecisionMakerDD
    'Case 1
      '  DecisionMakerDDX = Sheets("Setup").Range(OfficeXCol & "39").Offset(0, OfficeColOffset).Value2
      '  DecisionMakerDDY = Sheets("Setup").Range(OfficeYCol & "39").Offset(0, OfficeColOffset).Value2
  '  Case 2
     '   DecisionMakerDDX = Sheets("Setup").Range(OfficeXCol & "40").Offset(0, OfficeColOffset).Value2
     '   DecisionMakerDDY = Sheets("Setup").Range(OfficeYCol & "40").Offset(0, OfficeColOffset).Value2
   ' Case 3
      '  DecisionMakerDDX = Sheets("Setup").Range(OfficeXCol & "41").Offset(0, OfficeColOffset).Value2
      '  DecisionMakerDDY = Sheets("Setup").Range(OfficeYCol & "41").Offset(0, OfficeColOffset).Value2
   ' Case 4
      '  DecisionMakerDDX = Sheets("Setup").Range(OfficeXCol & "42").Offset(0, OfficeColOffset).Value2
      '  DecisionMakerDDY = Sheets("Setup").Range(OfficeYCol & "42").Offset(0, OfficeColOffset).Value2
        
    'Case Else
       ' MsgBox "Problem with DecisionMaker DropDown"
    'Exit Sub
'End Select

Dim spotDate As Date
spotDate = WorksheetFunction.Index(Sheets("Setup").Range("S2:S200"), _
    WorksheetFunction.Match(clientName & BaseCcy & CounterCcy, Sheets("Setup").Range("R2:R200"), 0))
If IsError(spotDate) Then
    MsgBox "Problem with Spot Date Setup"
    Exit Sub
End If
Dim tomDate As Date
tomDate = WorksheetFunction.Index(Sheets("Setup").Range("V2:V200"), _
    WorksheetFunction.Match(clientName & BaseCcy & CounterCcy, Sheets("Setup").Range("R2:R200"), 0))
If IsError(tomDate) Then
    MsgBox "Problem with Tom Date Setup"
    Exit Sub
End If

Dim NearDateClickX As Integer, NearDateClickY As Integer
NearDateClickX = Sheets("Setup").Range(OfficeXCol & "9").Offset(0, OfficeColOffset).Value2
NearDateClickY = Sheets("Setup").Range(OfficeYCol & "9").Offset(0, OfficeColOffset).Value2

Dim NearDateDropDownX As Integer, NearDateDropDownY As Integer
Select Case NearDate
Case spotDate
    NearDateDropDownX = Sheets("Setup").Range(OfficeXCol & "12").Offset(0, OfficeColOffset).Value2
    NearDateDropDownY = Sheets("Setup").Range(OfficeYCol & "12").Offset(0, OfficeColOffset).Value2
Case tomDate
    NearDateDropDownX = Sheets("Setup").Range(OfficeXCol & "11").Offset(0, OfficeColOffset).Value2
    NearDateDropDownY = Sheets("Setup").Range(OfficeYCol & "11").Offset(0, OfficeColOffset).Value2
Case Date
    NearDateDropDownX = Sheets("Setup").Range(OfficeXCol & "10").Offset(0, OfficeColOffset).Value2
    NearDateDropDownY = Sheets("Setup").Range(OfficeYCol & "10").Offset(0, OfficeColOffset).Value2
Case Else
    MsgBox "Problem with Near Date"
    Exit Sub
End Select

Dim NextMonthClick As Integer
If Month(FarDate) = Month(Date) Then
    NextMonthClick = 0
Else
    NextMonthClick = 1
End If

Dim FarDateClickX As Integer, FarDateClickY As Integer
FarDateClickX = Sheets("Setup").Range(OfficeXCol & "13").Offset(0, OfficeColOffset).Value2
FarDateClickY = Sheets("Setup").Range(OfficeYCol & "13").Offset(0, OfficeColOffset).Value2

Dim NextMonthClickX As Integer, NextMonthClickY As Integer
NextMonthClickX = Sheets("Setup").Range(OfficeXCol & "14").Offset(0, OfficeColOffset).Value2
NextMonthClickY = Sheets("Setup").Range(OfficeYCol & "14").Offset(0, OfficeColOffset).Value2

Dim FarDateRow As Integer, FarDateColumn As Integer
FarDateColumn = ((FarDate - 1) Mod 7) + 1
FarDateRow = FarDateRowCalc(FarDate)

'Debug.Print FarDateColumn, FarDateRow

Dim FarDateDropDownX As Integer, FarDateDropDownY As Integer
FarDateDropDownX = Sheets("Setup").Range(OfficeXCol & "15").Offset(FarDateColumn, OfficeColOffset).Value2
FarDateDropDownY = Sheets("Setup").Range(OfficeYCol & "15").Offset(FarDateRow, OfficeColOffset).Value2

'if underlying trade is a buy, then select sell leg on near date
Dim BuySellX As Integer, BuySellY As Integer
If LCase(BuySell) = "buy" Then
    BuySellX = Sheets("Setup").Range(OfficeXCol & "23").Offset(0, OfficeColOffset).Value2
    BuySellY = Sheets("Setup").Range(OfficeYCol & "23").Offset(0, OfficeColOffset).Value2
Else
    BuySellX = Sheets("Setup").Range(OfficeXCol & "24").Offset(0, OfficeColOffset).Value2
    BuySellY = Sheets("Setup").Range(OfficeYCol & "24").Offset(0, OfficeColOffset).Value2
End If

Dim PortfolioClickX As Integer, PortfolioClickY As Integer
PortfolioClickX = Sheets("Setup").Range(OfficeXCol & "25").Offset(0, OfficeColOffset).Value2
PortfolioClickY = Sheets("Setup").Range(OfficeYCol & "25").Offset(0, OfficeColOffset).Value2

Dim PortfolioDropdown As Integer
'if underlying trade is a buy, then select sell leg on near date
If LCase(BuySell) = "buy" Then
    PortfolioDropdown = WorksheetFunction.Index(Sheets("Setup").Range("O2:O200"), _
            WorksheetFunction.Match(clientName & BaseCcy & CounterCcy, Sheets("Setup").Range("R2:R200"), 0))
Else
    PortfolioDropdown = WorksheetFunction.Index(Sheets("Setup").Range("P2:P200"), _
            WorksheetFunction.Match(clientName & BaseCcy & CounterCcy, Sheets("Setup").Range("R2:R200"), 0))
End If
If PortfolioDropdown > 3 Then
    MsgBox "Problem with Portfolio dropdown value"
    Exit Sub
End If

Dim PortfolioDropDownX As Integer, PortfolioDropDownY As Integer
PortfolioDropDownX = Sheets("Setup").Range(OfficeXCol & "26").Offset(PortfolioDropdown - 1, OfficeColOffset).Value2
PortfolioDropDownY = Sheets("Setup").Range(OfficeYCol & "26").Offset(PortfolioDropdown - 1, OfficeColOffset).Value2


Dim TradeActionClickX As Integer, TradeActionClickY As Integer
TradeActionClickX = Sheets("Setup").Range(OfficeXCol & "29").Offset(0, OfficeColOffset).Value2
TradeActionClickY = Sheets("Setup").Range(OfficeYCol & "29").Offset(0, OfficeColOffset).Value2

Dim TradeActionDropDownX As Integer, TradeActionDropDownY As Integer
TradeActionDropDownX = Sheets("Setup").Range(OfficeXCol & "30").Offset(0, OfficeColOffset).Value2
TradeActionDropDownY = Sheets("Setup").Range(OfficeYCol & "30").Offset(0, OfficeColOffset).Value2

Dim MMRefBoxX As Integer, MMRefBoxY As Integer
MMRefBoxX = Sheets("Setup").Range(OfficeXCol & "31").Offset(0, OfficeColOffset).Value2
MMRefBoxY = Sheets("Setup").Range(OfficeYCol & "31").Offset(0, OfficeColOffset).Value2

Dim VLBoxX As Integer, VLBoxY As Integer
VLBoxX = Sheets("Setup").Range(OfficeXCol & "32").Offset(0, OfficeColOffset).Value2
VLBoxY = Sheets("Setup").Range(OfficeYCol & "32").Offset(0, OfficeColOffset).Value2

Dim SpreadBoxX As Integer, SpreadBoxY As Integer
SpreadBoxX = Sheets("Setup").Range(OfficeXCol & "33").Offset(0, OfficeColOffset).Value2
SpreadBoxY = Sheets("Setup").Range(OfficeYCol & "33").Offset(0, OfficeColOffset).Value2

Dim AmountBoxX As Integer, AmountBoxY As Integer
If LCase(BuySell) = "buy" Then
    AmountBoxX = Sheets("Setup").Range(OfficeXCol & "34").Offset(0, OfficeColOffset).Value2
    AmountBoxY = Sheets("Setup").Range(OfficeYCol & "34").Offset(0, OfficeColOffset).Value2
Else
    AmountBoxX = Sheets("Setup").Range(OfficeXCol & "35").Offset(0, OfficeColOffset).Value2
    AmountBoxY = Sheets("Setup").Range(OfficeYCol & "35").Offset(0, OfficeColOffset).Value2
End If

Dim QuoteButtonX As Integer, QuoteButtonY As Integer
QuoteButtonX = Sheets("Setup").Range(OfficeXCol & "36").Offset(0, OfficeColOffset).Value2
QuoteButtonY = Sheets("Setup").Range(OfficeYCol & "36").Offset(0, OfficeColOffset).Value2

Dim NewOrderButtonX As Integer, NewOrderButtonY As Integer
NewOrderButtonX = Sheets("Setup").Range(OfficeXCol & "37").Offset(0, OfficeColOffset).Value2
NewOrderButtonY = Sheets("Setup").Range(OfficeYCol & "37").Offset(0, OfficeColOffset).Value2


Dim AutoITRunfile As String
'AutoITRunfile = """C:\Users\abc\OneDrive\Desktop\Systems\AutoITScripts\WIP\SWAPExcel.exe"""
'AutoITRunfile = """C:\Users\abc\OneDrive\Documents\Tools\AutoITScripts\WIP\SWAPExcelNoExtend.exe"""
AutoITRunfile = """C:\Users\abc\scbapps\Scripts\SWAPExcelNoExtend.exe"""

Dim AutoITCommand As String
'AutoITCommand = WorksheetFunction.TextJoin(" ", False, AutoITRunfile, _
                                        SwapX, SwapY, _
                                        CIF, CIFX, CIFY, _
                                        CcyPair, CcyPairX, CcyPairY, CPDropDownX, CPDropDownY, _
                                        NearDateClickX, NearDateClickY, NearDateDropDownX, NearDateDropDownY, _
                                        FarDateClickX, FarDateClickY, _
                                        NextMonthClick, NextMonthClickX, NextMonthClickY, _
                                        FarDateDropDownX, FarDateDropDownY, _
                                        BuySellX, BuySellY, _
                                        PortfolioClickX, PortfolioClickY, PortfolioDropDownX, PortfolioDropDownY, _
                                        TradeActionClickX, TradeActionClickY, TradeActionDropDownX, TradeActionDropDownY, _
                                        MMRef, MMRefBoxX, MMRefBoxY, _
                                       VLBoxX, VLBoxY, _
                                       SpreadBoxX, SpreadBoxY, _
                                       BaseAmt, AmountBoxX, AmountBoxY)

AutoITCommand = WorksheetFunction.TextJoin(" ", False, AutoITRunfile, _
                                        SwapX, SwapY, _
                                        CIF, CIFX, CIFY, _
                                        CcyPair, CcyPairX, CcyPairY, _
                                        CPDropDownX, CPDropDownY, _
                                        NearDateClickX, NearDateClickY, NearDateDropDownX, NearDateDropDownY, _
                                        FarDateClickX, FarDateClickY, _
                                        NextMonthClick, NextMonthClickX, NextMonthClickY, _
                                        FarDateDropDownX, FarDateDropDownY, _
                                        BuySellX, BuySellY)
                                        
AutoITCommand = WorksheetFunction.TextJoin(" ", False, AutoITCommand, _
                                        PortfolioClickX, PortfolioClickY, PortfolioDropDownX, PortfolioDropDownY, _
                                        TradeActionClickX, TradeActionClickY, TradeActionDropDownX, TradeActionDropDownY, _
                                        MMRef, MMRefBoxX, MMRefBoxY, _
                                       VLBoxX, VLBoxY, _
                                       SpreadBoxX, SpreadBoxY, _
                                       BaseAmt, AmountBoxX, AmountBoxY, _
                                       QuoteButtonX, QuoteButtonY, _
                                       Rate, NewOrderButtonX, NewOrderButtonY)


AutoITCommand = WorksheetFunction.TextJoin(" ", False, AutoITCommand, _
                                        DecisionMakerClickX, DecisionMakerClickY, _
                                        DecisionMakerDDX, DecisionMakerDDY)
'Debug.Print AutoITCommand
If FarDate > LastDateofNextMonth(Date) Then
    MsgBox "Far Date beyond 1-month." _
            & vbNewLine & "1. Key " & Format(FarDate, "dd-mmm-yy") & " manually" _
            & vbNewLine & "2. Set spread to 0"
End If
Dim RetVal As Variant
RetVal = Shell(AutoITCommand)
'Dim wsh As Object
'Set wsh = CreateObject("WScript.Shell")
'AutoITCommand = "notepad.exe"
'wsh.Run AutoITCommand
'Debug.Print RetVal

End Sub


