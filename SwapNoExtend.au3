    ;Home Runs with 90% zoom
opt("MustDeclareVars",1)
Local $SwapX = $CmdLine[1]
Local $SwapY = $CmdLine[2]
Local $CIF = $CmdLine[3]
Local $CIFX = $CmdLine[4]
Local $CIFY = $CmdLine[5]
Local $CcyPair = $CmdLine[6]
Local $CcyPairX = $CmdLine[7]
Local $CcyPairY = $CmdLine[8]
Local $CPDropDownX = $CmdLine[9]
Local $CPDropDownY = $CmdLine[10]
Local $NearDateClickX = $CmdLine[11]
Local $NearDateClickY = $CmdLine[12]
Local $NearDateDropDownX = $CmdLine[13]
Local $NearDateDropDownY = $CmdLine[14]
Local $FarDateClickX = $CmdLine[15]
Local $FarDateClickY = $CmdLine[16]
Local $NextMonth = $CmdLine[17]
Local $NextMonthClickX = $CmdLine[18]
Local $NextMonthClickY = $CmdLine[19]
Local $FarDateDropDownX = $CmdLine[20]
Local $FarDateDropDownY = $CmdLine[21]
Local $BuySellX = $CmdLine[22]
Local $BuySellY = $CmdLine[23]
Local $PortfolioClickX =$CmdLine[24]
Local $PortfolioClickY =$CmdLine[25]
Local $PortfolioDropDownX =$CmdLine[26]
Local $PortfolioDropDownY =$CmdLine[27]
Local $TradeActionClickX =$CmdLine[28]
Local $TradeActionClickY =$CmdLine[29]
Local $TradeActionDropDownX =$CmdLine[30]
Local $TradeActionDropDownY =$CmdLine[31]
Local $MMRef = $CmdLine[32]
Local $MMRefBoxX = $CmdLine[33]
Local $MMRefBoxY = $CmdLine[34]
Local $VLBoxX = $CmdLine[35]
Local $VLBoxY = $CmdLine[36]
Local $SpreadBoxX = $CmdLine[37]
Local $SpreadBoxY = $CmdLine[38]
Local $Amount = $CmdLine[39]
Local $AmountBoxX = $CmdLine[40]
Local $AmountBoxY = $CmdLine[41]
Local $QuoteButtonX = $CmdLine[42]
Local $QuoteButtonY = $CmdLine[43]
Local $Rate = $CmdLine[44]
Local $NewOrderButtonX = $CmdLine[45]
Local $NewOrderButtonY = $CmdLine[46]
Local $DecisionMakerClickX = $CmdLine[47]
Local $DecisionMakerClickY = $CmdLine[48]
Local $DecisionMakerDDX = $CmdLine[49]
Local $DecisionMakerDDY = $CmdLine[50]
;Click NewOrderButton
WinActivate("OFX")
MouseClick("left",$NewOrderButtonX,$NewOrderButtonY)
;Input client CIF
;ClipPut("")
;ClipPut($CIF)
WinActivate("OFX")
MouseClick("left",$CIFX,$CIFY)
;Send("^a")
;Send("^v")
Send($CIF)
Send("{TAB}")
Sleep(1000)
;Key in currencypair, then click dropdown
;copy paste didn't work for dropdown
;ClipPut($CcyPair)
WinActivate("OFX")
MouseClick("left",$CcyPairX,$CcyPairY)
Send($CcyPair)
;Send("^a")
;Send("^v")
;Sleep(1000)
MouseClick("left",$CPDropDownX,$CPDropDownY)
;Activate Swaps
Sleep(500)
WinActivate("OFX")
MouseClick("left",$SwapX,$SwapY)
;Sleep(500)
;activate near date dropdown then select tod/tom/spot
;WinActivate("OFX")
MouseClick("left",$NearDateClickX,$NearDateClickY)
;MouseClick("left",$NearDateClickX,$NearDateClickY)
Sleep(1000)
MouseClick("left",$NearDateDropDownX,$NearDateDropDownY)
Sleep(500)
;activate fardate dropdown then select next month if required, then date 
WinActivate("OFX")
MouseClick("left",$FarDateClickX,$FarDateClickY)
If Not(stringcompare($NextMonth,"1")) Then 
MouseClick("left",$NextMonthClickX,$NextMonthClickY)
EndIf 
MouseClick("left",$FarDateDropDownX,$FarDateDropDownY)
;click BuySell
;msgbox(4096,"BuySell", $BuySellX & " " & $BuySellY)
WinActivate("OFX")
Sleep(2000)
MouseClick("left",$BuySellX,$BuySellY)
;activate portfoliobox then select portfolio
WinActivate("OFX")
Sleep(2000)
MouseClick("left",$PortfolioClickX,$PortfolioClickY)
Sleep(500)
MouseClick("left",$PortfolioDropDownX,$PortfolioDropDownY)
;activate tradeaction then select Extend
;WinActivate("OFX")
;MouseClick("left",$TradeActionClickX,$TradeActionClickY)
;MouseClick("left",$TradeActionDropDownX,$TradeActionDropDownY)
;Input MMRef
;ClipPut("")
;ClipPut($MMRef)
;WinActivate("OFX")
;MouseClick("left",$MMRefBoxX,$MMRefBoxY)
;Sleep(500)
;Send($MMRef)
;Send("^a")
;Send("^v")
;Input VL deets
;ClipPut("")
;ClipPut("SI OC " & $Rate)
WinActivate("OFX")
MouseClick("left",$VLBoxX,$VLBoxY)
Sleep(500)
Send("^a")
Send("{BACKSPACE}")
;Send("SI OC " & $Rate)
Send("SI ROLL")
;Send("^a")
;Send("^v")
;Override spread to 0
;ClipPut("")
;ClipPut("0")
WinActivate("OFX")
MouseClick("left",$DecisionMakerClickX,$DecisionMakerClickY)
MouseClick("left",$DecisionMakerDDX,$DecisionMakerDDY)
WinActivate("OFX")
MouseClick("left",$SpreadBoxX,$SpreadBoxY)
Sleep(500)
Send("^a")
Send("{BACKSPACE}")
;Send("^v")
Send("0")
;Input Amount
;ClipPut("")
;ClipPut($Amount)
WinActivate("OFX")
MouseClick("left",$AmountBoxX,$AmountBoxY)
Sleep(500)
;Send("^a")
;Send("^v")
Send($Amount)
WinActivate("OFX")
Mousemove($QuoteButtonX,$QuoteButtonY)

