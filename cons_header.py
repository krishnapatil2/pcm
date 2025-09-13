
# Required columns
DATE = 'Date'
CP_CODE = 'CP Code'
SI = 'Segment Indicator'
FINANCIAL_LEDGER_BALANCE = 'Financial Ledger balance (clear)-B in the books of TM for clients and in the books of CM for TM (Pro) and in the books of CM for CP'
columns_to_keep = [DATE, CP_CODE, SI, FINANCIAL_LEDGER_BALANCE]
FNO = "FO"
CD = "CD"
NSE_MEMBER_CODE = "90123"
C = "C"

SEGMENTS = {
    "FNO": FNO,
    "CD": CD,
}


NSE_AND_MCX = "NSE_AND_MCX"
LEDGER = "LEDGER"

TM_CP_CODE = 'TM/CP code'
CASH_COLLECTED = 'Cash allocated (b)'
CLIENT_CODE = 'ClientCode'
FO_MARGIN = 'TotalCollateral'

col_1 = [TM_CP_CODE, CASH_COLLECTED]
col_2 = [CLIENT_CODE, FO_MARGIN]

########### obligation file header
SGMT = "Sgmt"
SRC = "Src"
CLRMMBID = "ClrMmbId"
BRKRORCTDNPTCPTID = "BrkrOrCtdnPtcptId"
FININSTRMID = "FinInstrmId"
ISIN = "ISIN"
TCKRSYMB = "TckrSymb"
SCTYSRS = "SctySrs"
STTLMTP = "SttlmTp"
SCTIESSTTLMTXID = "SctiesSttlmTxId"
TRADREGNORGN = "TradRegnOrgn"
CLNTID = "ClntId"
CTDNID = "CtdnId"
CTDNPTCPTID = "CtdnPtcptId"
RPTGDT = "RptgDt"
FNDSPAYINDT = "FndsPayInDt"
FNDSPAYOUTDT = "FndsPayOutDt"
CMMDTYORSCRTYPAYINDT = "CmmdtyOrScrtyPayInDt"
CMMDTYORSCRTYPAYOUTDT = "CmmdtyOrScrtyPayOutDt"
DALYBUYTRADGVOL = "DalyBuyTradgVol"
DALYSELLTRADGVOL = "DalySellTradgVol"
DALYBUYTRADGVAL = "DalyBuyTradgVal"
DALYSELLTRADGVAL = "DalySellTradgVal"
CMLTVBUYVOL = "CmltvBuyVol"
CMLTVSELLVOL = "CmltvSellVol"
CMLTVBUYAMT = "CmltvBuyAmt"
CMLTVSELLAMT = "CmltvSellAmt"
FNLOBLGTNFLG = "FnlOblgtnFlg"
RMKS = "Rmks"
RSVD1 = "Rsvd1"
RSVD2 = "Rsvd2"
RSVD3 = "Rsvd3"
RSVD4 = "Rsvd4"


########################### stamp duty header
RPTHDR = "RptHdr"
SGMT = "Sgmt"
SRC = "Src"
STTLMTP = "SttlmTp"
SCTIESSTTLMTXID = "SctiesSttlmTxId"
CLRMMBID = "ClrMmbId"
BRKRORCTDNPTCPTID = "BrkrOrCtdnPtcptId"
CLCTNDT = "ClctnDt"
DUEDT = "DueDt"
CLNTID = "ClntId"
CTRYSUBDVSN = "CtrySubDvsn"
TCKRSYMB = "TckrSymb"
SCTYSRS = "SctySrs"
FININSTRMID = "FinInstrmId"
FININSTRMTP = "FinInstrmTp"
ISIN = "ISIN"
XPRYDT = "XpryDt"
STRKPRIC = "StrkPric"
OPTNTP = "OptnTp"
TTLBUTRADGVOL = "TtlBuyTradgVol"
TTLBUTRFVAL = "TtlBuyTrfVal"
TTLSLLTRADGVOL = "TtlSellTradgVol"
TTLSLLTRFVAL = "TtlSellTrfVal"
BUYDLVRYQTY = "BuyDlvryQty"
BUYDLVRYVAL = "BuyDlvryVal"
BUYOTHRTHANDLVRYQTY = "BuyOthrThanDlvryQty"
BUYOTHRTHANDLVRYVAL = "BuyOthrThanDlvryVal"
BUYSTMPDTY = "BuyStmpDty"
SELLSTMPDTY = "SellStmpDty"
STTLMPRIC = "SttlmPric"
BUYDLVRYSTMPDTY = "BuyDlvryStmpDty"
BUYOTHRTHANDLVRYSTMPDTY = "BuyOthrThanDlvryStmpDty"
STMPDTYAMT = "StmpDtyAmt"
RMKS = "Rmks"
RSVD1 = "Rsvd1"
RSVD2 = "Rsvd2"
RSVD3 = "Rsvd3"
RSVD4 = "Rsvd4"

########################### stt header
RPTHDR = "RptHdr"
SGMT = "Sgmt"
SRC = "Src"
TRADDT = "TradDt"
CLCTNDT = "ClctnDt"
DUEDT = "DueDt"
STTLMTP = "SttlmTp"
SCTIESSTTLMTXID = "SctiesSttlmTxId"
CLRMMBID = "ClrMmbId"
BRKRORCTDNPTCPTID = "BrkrOrCtdnPtcptId"
CLNTID = "ClntId"
TCKRSYMB = "TckrSymb"
SCTYSRS = "SctySrs"
FININSTRMID = "FinInstrmId"
FININSTRMTP = "FinInstrmTp"
ISIN = "ISIN"
XPRYDT = "XpryDt"
OPTNTP = "OptnTp"
STRKPRIC = "StrkPric"
STTLMPRIC = "SttlmPric"
TTLBUYTRADGVOL = "TtlBuyTradgVol"
TTLBUYTRFVAL = "TtlBuyTrfVal"
TTLSELLTRADGVOL = "TtlSellTradgVol"
TTLSELLTRFVAL = "TtlSellTrfVal"
AVRGPRIC = "AvrgPric"
BUYDLVRBLQTY = "BuyDlvrblQty"
SELLDLVRBLQTY = "SellDlvrblQty"
SELLOTHRTHANDLVRQTY = "SellOthrThanDlvryQty"
BUYDLVRBLVAL = "BuyDlvrblVal"
SELLDLVRBLVAL = "SellDlvrblVal"
SELLOTHRTHANDLVRVAL = "SellOthrThanDlvryVal"
BUYDELVRYTTLTXS = "BuyDelvryTtlTaxs"
SELLDELVRYTTLTXS = "SellDelvryTtlTaxs"
SELLOTHRTHANDELVRYTTLTXS = "SellOthrThanDelvryTtlTaxs"
TAXBLSELLFUTRSVAL = "TaxblSellFutrsVal"
TAXBLSELLOPTNVAL = "TaxblSellOptnVal"
OPTNEXRCQTY = "OptnExrcQty"
OPTNEXRCVAL = "OptnExrcVal"
TAXBLEXRCVAL = "TaxblExrcVal"
FUTRSTTLMTTXS = "FutrsTtlTaxs"
OPTNTTLTXS = "OptnTtlTaxs"
TAXBLBUYVALCALLAUCTN = "TaxblBuyValCallAuctn"
TAXBLSELLVALCALLAUCTN = "TaxblSellValCallAuctn"
TTLTXS = "TtlTaxs"
RMKS = "Rmks"
RSVD1 = "Rsvd1"
RSVD2 = "Rsvd2"
RSVD3 = "Rsvd3"
RSVD4 = "Rsvd4"

EXTRA_COLUMNS = [
    "Buy STT",
    "Sell STT",
    "Sell Stamp Duty",
    "Buy Stamp Duty",
    "Buy Payable Amount",
    "Sell Receivable Amount",
    "Net Receivable \\ Payable"
]

OBLIGATION_HEADER = [
    "Sgmt", "Src", "ClrMmbId", "BrkrOrCtdnPtcptId", "FinInstrmId", "ISIN", "TckrSymb", "SctySrs", "SttlmTp",
    "SctiesSttlmTxId", "TradRegnOrgn", "ClntId", "CtdnId", "CtdnPtcptId", "RptgDt", "FndsPayInDt", "FndsPayOutDt",
    "CmmdtyOrScrtyPayInDt", "CmmdtyOrScrtyPayOutDt", "DalyBuyTradgVol", "DalySellTradgVol", "DalyBuyTradgVal",
    "DalySellTradgVal", "CmltvBuyVol", "CmltvSellVol", "CmltvBuyAmt", "CmltvSellAmt", "FnlOblgtnFlg",
]

STAMP_DUTY_HEADER = [
    "RptHdr", "Sgmt", "Src", "SttlmTp", "SctiesSttlmTxId", "ClrMmbId", "BrkrOrCtdnPtcptId", "ClctnDt", "DueDt",
    "ClntId", "CtrySubDvsn", "TckrSymb", "SctySrs", "FinInstrmId", "FinInstrmTp", "ISIN", "XpryDt", "StrkPric",
    "OptnTp", "TtlBuyTradgVol", "TtlBuyTrfVal", "TtlSellTradgVol", "TtlSellTrfVal", "BuyDlvryQty", "BuyDlvryVal",
    "BuyOthrThanDlvryQty", "BuyOthrThanDlvryVal", "BuyStmpDty", "SellStmpDty", "SttlmPric", "BuyDlvryStmpDty",
    "BuyOthrThanDlvryStmpDty", "StmpDtyAmt", "Rmks", "Rsvd1", "Rsvd2", "Rsvd3", "Rsvd4"
]

STT_HEADER = [
    "RptHdr", "Sgmt", "Src", "TradDt", "ClctnDt", "DueDt", "SttlmTp", "SctiesSttlmTxId", "ClrMmbId",
    "BrkrOrCtdnPtcptId", "ClntId", "TckrSymb", "SctySrs", "FinInstrmId", "FinInstrmTp", "ISIN", "XpryDt",
    "OptnTp", "StrkPric", "SttlmPric", "TtlBuyTradgVol", "TtlBuyTrfVal", "TtlSellTradgVol", "TtlSellTrfVal",
    "AvrgPric", "BuyDlvrblQty", "SellDlvrblQty", "SellOthrThanDlvryQty", "BuyDlvrblVal", "SellDlvrblVal",
    "SellOthrThanDlvryVal", "BuyDelvryTtlTaxs", "SellDelvryTtlTaxs", "SellOthrThanDelvryTtlTaxs",
    "TaxblSellFutrsVal", "TaxblSellOptnVal", "OptnExrcQty", "OptnExrcVal", "TaxblExrcVal", "FutrsTtlTaxs",
    "OptnTtlTaxs", "TaxblBuyValCallAuctn", "TaxblSellValCallAuctn", "TtlTaxs", "Rmks", "Rsvd1", "Rsvd2", "Rsvd3", "Rsvd4"
]
