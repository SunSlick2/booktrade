Sub SetUpDailyRolloverSwapAutoITv1()

    Dim CRow As Long
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
    Dim VLDeets As String
    Dim SpreadPip As Double
    Dim SpreadBP As Double
    Dim spotDate As Date
    Dim tomDate As Date
    Dim PortfolioDropdown As Long
    Dim DecisionMakerDD As Long
    Dim wsSetup As Worksheet
    Dim clientMatchRow As Variant
    Dim ccyPairMatchRow As Variant
    Dim clientRange As Range
    Dim ccyPairRange As Range
    
    CRow = ActiveCell.Row
    If rows(CRow).Hidden Then
        MsgBox "Row is Hidden"
        Exit Sub
    End If

    ' Turn off screen updating and calculations for performance
    Application.ScreenUpdating = False

    ' Assign worksheet to a variable for faster access
    Set wsSetup = ThisWorkbook.Sheets("Setup")

    ' --- Step 1: Read values from the active sheet ---
    NearDate = Cells(CRow, 1).Value
    clientName = Cells(CRow, 2).Value
    MMRef = Cells(CRow, 3).Value
    BuySell = Cells(CRow, 6).Value
    BaseAmt = Abs(Cells(CRow, 7).Value)
    BaseCcy = Cells(CRow, 8).Value
    CounterCcy = Cells(CRow, 10).Value
    Rate = Cells(CRow, 11).Value
    CcyPair = BaseCcy & CounterCcy

    ' --- Step 2: Perform a single set of lookups on the "Setup" sheet ---
    ' Define the lookup ranges once
    Set clientRange = wsSetup.Range("B2:B200")
    Set ccyPairRange = wsSetup.Range("R2:R200")
    
    ' Find the rows for the client and the currency pair once
    clientMatchRow = Application.Match(clientName, clientRange, 0)
    CIF = wsSetup.Cells(clientMatchRow + 1, "C").Value ' +1 for the 1-based index from Match
    ccyPairMatchRow = Application.Match(CIF & BaseCcy & CounterCcy, ccyPairRange, 0)

    ' Check for lookup errors immediately
    If IsError(clientMatchRow) Then
        MsgBox "Client '" & clientName & "' not found in Setup sheet."
        GoTo CleanUp
    End If

    If IsError(ccyPairMatchRow) Then
        MsgBox "Currency Pair for '" & clientName & "' not found in Setup sheet."
        GoTo CleanUp
    End If

    ' Use the stored row numbers for all subsequent lookups

    VLDeets = wsSetup.Cells(clientMatchRow + 1, "F").Value
    SpreadPip = wsSetup.Cells(clientMatchRow + 1, "G").Value
    SpreadBP = wsSetup.Cells(clientMatchRow + 1, "G").Value ' As per your clarification, these are the same for now

    FarDate = wsSetup.Cells(ccyPairMatchRow + 1, "N").Value
    DecisionMakerDD = wsSetup.Cells(ccyPairMatchRow + 1, "Q").Value
    spotDate = wsSetup.Cells(ccyPairMatchRow + 1, "S").Value
    tomDate = wsSetup.Cells(ccyPairMatchRow + 1, "V").Value
    
    If LCase(BuySell) = "buy" Then
        PortfolioDropdown = wsSetup.Cells(ccyPairMatchRow + 1, "O").Value
    Else
        PortfolioDropdown = wsSetup.Cells(ccyPairMatchRow + 1, "P").Value
    End If
    
    ' --- Step 3: Error checks and other calculations ---
    If FarDate < Date Then
        MsgBox "Far Date looks wrong"
        GoTo CleanUp
    End If
    
    If PortfolioDropdown > 3 Then
        MsgBox "Problem with Portfolio dropdown value"
        GoTo CleanUp
    End If
    
    If IsError(spotDate) Then
        MsgBox "Problem with Spot Date Setup"
        GoTo CleanUp
    End If
    
    If IsError(tomDate) Then
        MsgBox "Problem with Tom Date Setup"
        GoTo CleanUp
    End If
    
    ' --- Step 4: Continue with the rest of your code using the variables ---
    ' ... (rest of the code is unchanged) ...
    Dim RootX As Long, RootY As Long
    Dim OfficeColOffset As Long

    If wsSetup.Range("AA2").Value2 = "Office" Then
        OfficeColOffset = 0
    Else
        OfficeColOffset = 2
    End If
    Dim OfficeXCol As String
    Dim OfficeYCol As String
    OfficeXCol = "AB"
    OfficeYCol = "AC"
    
    Dim SwapX As Long, SwapY As Long
    SwapX = wsSetup.Range(OfficeXCol & "5").Offset(0, OfficeColOffset).Value2
    SwapY = wsSetup.Range(OfficeYCol & "5").Offset(0, OfficeColOffset).Value2
    
    Dim CIFX As Long, CIFY As Long
    CIFX = wsSetup.Range(OfficeXCol & "6").Offset(0, OfficeColOffset).Value2
    CIFY = wsSetup.Range(OfficeYCol & "6").Offset(0, OfficeColOffset).Value2
    
    Dim CcyPairX As Long, CcyPairY As Long
    CcyPairX = wsSetup.Range(OfficeXCol & "7").Offset(0, OfficeColOffset).Value2
    CcyPairY = wsSetup.Range(OfficeYCol & "7").Offset(0, OfficeColOffset).Value2
    
    Dim CPDropDownX As Long, CPDropDownY As Long
    CPDropDownX = wsSetup.Range(OfficeXCol & "8").Offset(0, OfficeColOffset).Value2
    CPDropDownY = wsSetup.Range(OfficeYCol & "8").Offset(0, OfficeColOffset).Value2
    
    Dim DecisionMakerClickX As Long, DecisionMakerClickY As Long
    DecisionMakerClickX = wsSetup.Range(OfficeXCol & "38").Offset(0, OfficeColOffset).Value2
    DecisionMakerClickY = wsSetup.Range(OfficeYCol & "38").Offset(0, OfficeColOffset).Value2
    
    Dim DecisionMakerDDX As Long, DecisionMakerDDY As Long
    DecisionMakerDDX = wsSetup.Range(OfficeXCol & "38").Offset(DecisionMakerDD, OfficeColOffset).Value2
    DecisionMakerDDY = wsSetup.Range(OfficeYCol & "38").Offset(DecisionMakerDD, OfficeColOffset).Value2
    
    Dim NearDateClickX As Long, NearDateClickY As Long
    NearDateClickX = wsSetup.Range(OfficeXCol & "9").Offset(0, OfficeColOffset).Value2
    NearDateClickY = wsSetup.Range(OfficeYCol & "9").Offset(0, OfficeColOffset).Value2
    
    Dim NearDateDropDownX As Long, NearDateDropDownY As Long
    Select Case NearDate
    Case spotDate
        NearDateDropDownX = wsSetup.Range(OfficeXCol & "12").Offset(0, OfficeColOffset).Value2
        NearDateDropDownY = wsSetup.Range(OfficeYCol & "12").Offset(0, OfficeColOffset).Value2
    Case tomDate
        NearDateDropDownX = wsSetup.Range(OfficeXCol & "11").Offset(0, OfficeColOffset).Value2
        NearDateDropDownY = wsSetup.Range(OfficeYCol & "11").Offset(0, OfficeColOffset).Value2
    Case Date
        NearDateDropDownX = wsSetup.Range(OfficeXCol & "10").Offset(0, OfficeColOffset).Value2
        NearDateDropDownY = wsSetup.Range(OfficeYCol & "10").Offset(0, OfficeColOffset).Value2
    Case Else
        MsgBox "Problem with Near Date"
        GoTo CleanUp
    End Select
    
    Dim NextMonthClick As Long
    If Month(FarDate) = Month(Date) Then
        NextMonthClick = 0
    Else
        NextMonthClick = 1
    End If
    
    Dim FarDateClickX As Long, FarDateClickY As Long
    FarDateClickX = wsSetup.Range(OfficeXCol & "13").Offset(0, OfficeColOffset).Value2
    FarDateClickY = wsSetup.Range(OfficeYCol & "13").Offset(0, OfficeColOffset).Value2
    
    Dim NextMonthClickX As Long, NextMonthClickY As Long
    NextMonthClickX = wsSetup.Range(OfficeXCol & "14").Offset(0, OfficeColOffset).Value2
    NextMonthClickY = wsSetup.Range(OfficeYCol & "14").Offset(0, OfficeColOffset).Value2
    
    Dim FarDateRow As Long, FarDateColumn As Long
    FarDateColumn = ((FarDate - 1) Mod 7) + 1
    FarDateRow = FarDateRowCalc(FarDate)
    
    Dim FarDateDropDownX As Long, FarDateDropDownY As Long
    FarDateDropDownX = wsSetup.Range(OfficeXCol & "15").Offset(FarDateColumn, OfficeColOffset).Value2
    FarDateDropDownY = wsSetup.Range(OfficeYCol & "15").Offset(FarDateRow, OfficeColOffset).Value2
    
    Dim BuySellX As Long, BuySellY As Long
    If LCase(BuySell) = "buy" Then
        BuySellX = wsSetup.Range(OfficeXCol & "23").Offset(0, OfficeColOffset).Value2
        BuySellY = wsSetup.Range(OfficeYCol & "23").Offset(0, OfficeColOffset).Value2
    Else
        BuySellX = wsSetup.Range(OfficeXCol & "24").Offset(0, OfficeColOffset).Value2
        BuySellY = wsSetup.Range(OfficeYCol & "24").Offset(0, OfficeColOffset).Value2
    End If
    
    Dim PortfolioClickX As Long, PortfolioClickY As Long
    PortfolioClickX = wsSetup.Range(OfficeXCol & "25").Offset(0, OfficeColOffset).Value2
    PortfolioClickY = wsSetup.Range(OfficeYCol & "25").Offset(0, OfficeColOffset).Value2
    
    Dim PortfolioDropDownX As Long, PortfolioDropDownY As Long
    PortfolioDropDownX = wsSetup.Range(OfficeXCol & "26").Offset(PortfolioDropdown - 1, OfficeColOffset).Value2
    PortfolioDropDownY = wsSetup.Range(OfficeYCol & "26").Offset(PortfolioDropdown - 1, OfficeColOffset).Value2
    
    Dim TradeActionClickX As Long, TradeActionClickY As Long
    TradeActionClickX = wsSetup.Range(OfficeXCol & "29").Offset(0, OfficeColOffset).Value2
    TradeActionClickY = wsSetup.Range(OfficeYCol & "29").Offset(0, OfficeColOffset).Value2
    
    Dim TradeActionDropDownX As Long, TradeActionDropDownY As Long
    TradeActionDropDownX = wsSetup.Range(OfficeXCol & "30").Offset(0, OfficeColOffset).Value2
    TradeActionDropDownY = wsSetup.Range(OfficeYCol & "30").Offset(0, OfficeColOffset).Value2
    
    Dim MMRefBoxX As Long, MMRefBoxY As Long
    MMRefBoxX = wsSetup.Range(OfficeXCol & "31").Offset(0, OfficeColOffset).Value2
    MMRefBoxY = wsSetup.Range(OfficeYCol & "31").Offset(0, OfficeColOffset).Value2
    
    Dim VLBoxX As Long, VLBoxY As Long
    VLBoxX = wsSetup.Range(OfficeXCol & "32").Offset(0, OfficeColOffset).Value2
    VLBoxY = wsSetup.Range(OfficeYCol & "32").Offset(0, OfficeColOffset).Value2
    
    Dim SpreadBoxX As Long, SpreadBoxY As Long
    SpreadBoxX = wsSetup.Range(OfficeXCol & "33").Offset(0, OfficeColOffset).Value2
    SpreadBoxY = wsSetup.Range(OfficeYCol & "33").Offset(0, OfficeColOffset).Value2
    
    Dim AmountBoxX As Long, AmountBoxY As Long
    If LCase(BuySell) = "buy" Then
        AmountBoxX = wsSetup.Range(OfficeXCol & "34").Offset(0, OfficeColOffset).Value2
        AmountBoxY = wsSetup.Range(OfficeYCol & "34").Offset(0, OfficeColOffset).Value2
    Else
        AmountBoxX = wsSetup.Range(OfficeXCol & "35").Offset(0, OfficeColOffset).Value2
        AmountBoxY = wsSetup.Range(OfficeYCol & "35").Offset(0, OfficeColOffset).Value2
    End If
    
    Dim QuoteButtonX As Long, QuoteButtonY As Long
    QuoteButtonX = wsSetup.Range(OfficeXCol & "36").Offset(0, OfficeColOffset).Value2
    QuoteButtonY = wsSetup.Range(OfficeYCol & "36").Offset(0, OfficeColOffset).Value2
    
    Dim NewOrderButtonX As Long, NewOrderButtonY As Long
    NewOrderButtonX = wsSetup.Range(OfficeXCol & "37").Offset(0, OfficeColOffset).Value2
    NewOrderButtonY = wsSetup.Range(OfficeYCol & "37").Offset(0, OfficeColOffset).Value2

    Dim AutoITRunfile As String
    AutoITRunfile = """C:\Users\abc\mnoapps\Scripts\SWAPExcelNoExtend.exe"""

    Dim AutoITCommand As String
    
    AutoITCommand = JoinArgs(AutoITRunfile, SwapX, SwapY, _
                            CIF, CIFX, CIFY, _
                            CcyPair, CcyPairX, CcyPairY, _
                            CPDropDownX, CPDropDownY, _
                            NearDateClickX, NearDateClickY, NearDateDropDownX, NearDateDropDownY, _
                            FarDateClickX, FarDateClickY, _
                            NextMonthClick, NextMonthClickX, NextMonthClickY, _
                            FarDateDropDownX, FarDateDropDownY, _
                            BuySellX, BuySellY, _
                            PortfolioClickX, PortfolioClickY, PortfolioDropDownX, PortfolioDropDownY, _
                            TradeActionClickX, TradeActionClickY, TradeActionDropDownX, TradeActionDropDownY, _
                            MMRef, MMRefBoxX, MMRefBoxY, _
                            VLBoxX, VLBoxY, SpreadBoxX, SpreadBoxY, _
                            BaseAmt, AmountBoxX, AmountBoxY, _
                            QuoteButtonX, QuoteButtonY, _
                            Rate, NewOrderButtonX, NewOrderButtonY, _
                            DecisionMakerClickX, DecisionMakerClickY, DecisionMakerDDX, DecisionMakerDDY, _
                            SwapX, SwapY, CIF, CIFX, CIFY, _
                            CcyPair, CcyPairX, CcyPairY, _
                            CPDropDownX, CPDropDownY, _
                            NearDateClickX, NearDateClickY, NearDateDropDownX, NearDateDropDownY, _
                            FarDateClickX, FarDateClickY, _
                            NextMonthClick, NextMonthClickX, NextMonthClickY, _
                            FarDateDropDownX, FarDateDropDownY, _
                            BuySellX, BuySellY)

    If FarDate > LastDateofNextMonth(Date) Then
        MsgBox "Far Date beyond 1-month." _
             & vbNewLine & "1. Key " & Format(FarDate, "dd-mmm-yy") & " manually" _
             & vbNewLine & "2. Set spread to 0"
    End If

    Dim RetVal As Variant
    RetVal = Shell(AutoITCommand)

CleanUp:
    Application.ScreenUpdating = True
    
End Sub
