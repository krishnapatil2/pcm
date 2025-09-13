A = "Date"
B = "Clearing Member PAN"
C = "Trading member PAN"
D = "CP Code"
E = "CP PAN"
F = "Client PAN"
G = "Account Type"
H = "Segment Indicator"
I = "UCC Code"
J = "Financial Ledger balance-A in the books of TM for clients and in the books of CM for TM (Pro) and in the books of CM for CP"
K = "Financial Ledger balance (clear)-B in the books of TM for clients and in the books of CM for TM (Pro) and in the books of CM for CP"
L = "Peak Financial Ledger Balance (Clear)-C in the books of TM for clients and in the books of CM for TM (Pro) and in the books of CM for CP"
M = "Bank Guarantee (BG) received by TM from clients and by CM from TM (Pro) and from CPs"
N = "Fixed Deposit Receipt (FDR) received by TM from clients and by CM from TM(Pro) and from CPs"
O = "Approved Securities Cash Component received by TM from clients and by CM from TM(Pro) and from CPs"
P = "Approved Securities Non-cash component received by TM from clients and by CM from TM(Pro) and from CPs"
Q = "Non-Approved Securities received by TM from clients and by CM from TM(Pro) and from CPs"
R = "Value of CC approved Commodities received by TM from clients and by CM from TM(Pro) and from CPs"
S = "Other collaterals received by TM from clients and by CM from TM(Pro) and from CPs"
T = "Credit entry in ledger in lieu of EPI for clients / TM Pro"
U = "Pool Account for clients / TM Pro"
V = "Cash Retained by TM"
W = "Bank Guarantee (BG) Retained by TM"
X = "Fixed Deposit Receipt (FDR) Retained by TM"
Y = "Approved Securities Cash Component Retained by TM"
Z = "Approved Securities Non-cash component Retained by TM"
AA = "Non-Approved Securities Retained by TM"
AB = "Value of CC approved Commodities Retained by TM"
AC = "Other Collaterals Retained by TM"
AD = "Cash placed with CM"
AE = "Bank Guarantee (BG) placed with CM"
AF = "Fixed deposit receipt (FDR) placed with CM"
AG = "Approved Securities Cash Component placed with CM"
AH = "Approved Securities Non-cash component placed with CM"
AI = "Non-Approved Securities placed with CM"
AJ = "Value of CC approved Commodities placed with CM"
AK = "Other Collaterals placed with CM"
AL = "Cash Retained with CM"
AM = "Bank Guarantee (BG) retained with CM"
AN = "Fixed deposit receipt (FDR) retained with CM"
AO = "Approved Securities Cash Component retained with CM"
AP = "Approved Securities Non-cash component retained with CM"
AQ = "Non-Approved Securities retained with CM"
AR = "Value of CC approved Commodities retained with CM"
AS = "Other Collaterals Retained with CM"
AT = "Cash placed with NCL"
AU = "Bank Guarantee (BG) placed with NCL"
AV = "Fixed deposit receipt (FDR) placed with NCL"
AW = "Approved Securities Cash Component placed with NCL"
AX = "Approved Securities Non-cash component placed with NCL"
AY = "Value of CC approved Commodities placed with NCL"
AZ = "MTF /Non MTF indicator/Reason Code"
BA = "Uncleared Receipts"
BB = "Govt Securities / T-bills received by TM from clients and by CM from TM(Pro) and from CPs"
BC = "Govt Securities /T-bills Retained by TM"
BD = "Govt Securities/T-bills placed with CM"
BE = "Govt Securities/T bills retained with CM"
BF = "Govt Securities/T bills placed with NCL"
BG = "Bank Guarantee (BG) Funded portion retained with CM"
BH = "Bank Guarantee (BG) Non funded portion retained with CM"
BI = "Bank Guarantee (BG) Funded portion placed with NCL"
BJ = "Bank Guarantee (BG) Non funded portion placed with NCL"
BK = "Settlement Amount"
BL = "Unclaimed/Unsettled Client Funds"
BM = "Cash Collateral for MTF positions"

segregation_headers = [A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,AA,AB,AC,AD,AE,AF,AG,AH,AI,AJ,AK,AL,AM,AN,AO,AP,AQ,AR,AS,AT,AU,AV,AW,AX,AY,AZ,BA,BB,BC,BD,BE,BF,BG,BH,BI,BJ,BK,BL,BM]

CO_FIXED_DATA = [[
    "08-09-2025","AACCO4820B","AACCO4820B","DBSBK0000189","AAGCD0792B","", "C","CO","", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0","0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "NA","0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "NA", "0"],
[
    "08-09-2025","AACCO4820B","AACCO4820B","ICICI0006090","AAJCN6787F","", "C","CO","", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0","0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "NA","0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "NA", "0"]
]

# F_CPMaster_data : FO
# X_CPMaster_data : CD
A # ask in frontend - Date
D # 'CP Code'
E # ask in frontend - CP PAN
G # ?
H # Segment Indicator based on file name FO/CD
B # Clearing Member PAN
C # Trading member PAN


# Files - CashCollateral_CDS , CashCollateral_FNO | col - TotalCollateral
J 

# Files - Daily Margin Report NSECR, Daily Margin Report NSEFNO | col - Funds
K
L

# File - Collateral Valuation Report and skip this 90072
O # CashEquivalent
P # NonCash


# same columns data copy in 
K # -> AD AND AV
O # -> AG AND AW
P # -> AH AND AX

# File - F_90123_SEC_PLEDGE_09092025_02 | row - 219 col - GSEC below header -> this col for calculation = Client/CP code, ISIN, 
# this below formula for extra client code which is extra percentage
# col - val = GROSS VALUE * HAIRCUT =H221*I221% 
# FINAL EFFECTIVE VALUE = GROSS VALUE - val =H221-J219
# BB,BD,BF in this column paste this final effective value