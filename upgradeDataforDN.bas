Attribute VB_Name = "upgradeDataforDN"

Sub upgradeDataforDN()


Dim pathc2file As String, nmBrand$
Dim ar_code_Brand(6, 1)
Dim LastRow_CC, LastColumns_CC As Integer
Dim num_month%, CnqMonthNum%
Dim patchTR$, actTR$, LastRow, in_data, status_head   As String, MregName$, nm_REG$
Dim ALLTIME, EDU_PY, EDU_TY, place, AVG_HD As Variant
Dim f_brnd, iii, f_i, x, y, frqOrder   As Integer
Dim ar_nmAVG_Order(), ar_type_clients(1 To 4, 1 To 12)
Dim ar_Data(), ar_brand(), ar_nmHead(150), ar_PYPer_PRTN_VAL, ar_TYPer_PRTN_VAL, ar_PYPer_LOR_VAL(), ar_TYPer_LOR_VAL(), ar_nm_month(), ar_nmMregEN(), ar_nmMregLT()
Dim ThisYear%, VarYear%, f_YearTR As Integer, CnqYearDate As Integer

Dim mreg As String, reg As String


NF = ActiveWorkbook.Name
ThisMonth = CInt(InputBox("Month"))
VarYear = CInt(InputBox("year"))
ThisYear = 2016

'colums CA PRTNN VAL for LTM
'---------------------------------------------------------------------------------------------------------

'colums CA LOREAL VAL
str_PYper_LOR_VAL = 106
str_TYper_LOR_VAL = 93
str_PYper_PRTN_VAL = 79
str_TYper_PRTN_VAL = 66


'---------------------------------------------------------------------------------------------------------

myLib.VBA_Start
status_head = 0

in_data = "in_TR"
Workbooks(NF).Activate
Sheets(in_data).Select
NF_LastRow = myLib.GetLastRow

StartRowThisYear = Empty
For f_nf = 1 To NF_LastRow
    If Cells(f_nf, 1) = VarYear Then
        StartRowThisYear = f_nf
        Exit For
    End If
Next f_nf
If IsEmpty(StartRowThisYear) Then StartRowThisYear = 2


ReDim Preserve ar_Data(1 To 999999, 1 To 150)

ar_brand = Array("LP", "MX", "KR", "RD", "ES", "DE", "CR")
num_ar_brand = UBound(ar_brand)
array_row = 0
iii = 0

For f_YearTR = VarYear To ThisYear
    actual_TY = f_YearTR
    actual_PY = f_YearTR - 1

    act_month_Y = IIf(f_YearTR <> ThisYear, 12, ThisMonth)
       

    For f_brnd = 0 To num_ar_brand
        nmBrand = ar_brand(f_brnd)
        patchTR = GetPatchHistTR(nmBrand, ThisYear, f_YearTR, 12, 12)
        actTR = myLib.OpenFile(patchTR, nmBrand, False)
        If Len(actTR) > 0 Then
            LastRow = myLib.GetLastRow
    
            For f_i = 4 To LastRow
                
                MregName = myLib.GetMregWhitoutBrand(Cells(f_i, 4))
                
                If Not IsEmpty(MregName) Then
                    iii = iii + 1
                    nm_REG = Cells(f_i, 5)
                    MregNameExt = myLib.GetMregExt(MregName, nm_REG)
                    CnqMonthNum = myLib.GetMonthNumeric(Cells(f_i, 64))
                    CnqMonthName = myLib.GetNameMonthEN(CnqMonthNum)
                    CnqYearDate = myLib.GetYearType(f_YearTR, myLib.GetNum2num0(Cells(f_i, 65)), 1)
                    CnqYearGA = myLib.GetYearType(f_YearTR, CnqYearDate, 3)

                    cd_brand_row = nmBrand & Cells(f_i, 1)
                    cd_Univers = Cells(f_i, 2)
                    cd_Univers = IIf(Len(cd_Univers) <> 9, cd_brand_row, cd_Univers)
                    
                    ClientName = myLib.GetSalonName(Cells(f_i, 9), Cells(f_i, 13), Cells(f_i, 10))

                    n = 1: ar_Data(iii, n) = f_YearTR: ar_nmHead(n) = "TR_year"
                    n = n + 1: ar_Data(iii, n) = nmBrand: ar_nmHead(n) = "brand"
                    n = n + 1: ar_Data(iii, n) = myLib.GetTypeBusiness(nmBrand): ar_nmHead(n) = "bussines"
                    n = n + 1: ar_Data(iii, n) = Cells(f_i, 1): ar_nmHead(n) = "rowTR"
                    n = n + 1: ar_Data(iii, n) = cd_brand_row: ar_nmHead(n) = "BRAND_rowTR"
                    n = n + 1: ar_Data(iii, n) = cd_Univers: ar_nmHead(n) = "unvCD"
                    n = n + 1: ar_Data(iii, n) = nmBrand & Cells(f_i, 2): ar_nmHead(n) = "BRAND_unvCD"
                    n = n + 1: ar_Data(iii, n) = MregName: ar_nmHead(n) = "mreg"
                    n = n + 1: ar_Data(iii, n) = MregNameExt: ar_nmHead(n) = "mreg_EXT"
                    n = n + 1: ar_Data(iii, n) = Cells(f_i, 5): ar_nmHead(n) = "REG"
                    n = n + 1: ar_Data(iii, n) = Cells(f_i, 165): ar_nmHead(n) = "FLSM"
                    n = n + 1: ar_Data(iii, n) = Cells(f_i, 6): ar_nmHead(n) = "SEC"
                    n = n + 1: ar_Data(iii, n) = Cells(f_i, 7): ar_nmHead(n) = "SREP"
                    n = n + 1: ar_Data(iii, n) = ClientName: ar_nmHead(n) = "salon"
                    n = n + 1: ar_Data(iii, n) = Cells(f_i, 19): ar_nmHead(n) = "Chain_name"
                    n = n + 1: ar_Data(iii, n) = Cells(f_i, 11): ar_nmHead(n) = "city"
                    n = n + 1: ar_Data(iii, n) = myLib.GetClntType(Cells(f_i, 18), 1): ar_nmHead(n) = "type_SLN"
                    n = n + 1: ar_Data(iii, n) = myLib.GetClntType(Cells(f_i, 18), 2): ar_nmHead(n) = "salon_type_eng"
                    n = n + 1: ar_Data(iii, n) = myLib.GetClntType(Cells(f_i, 18), 3): ar_nmHead(n) = "salon_type_short_eng"
                    n = n + 1: ar_Data(iii, n) = myLib.GetClntType(Cells(f_i, 18), 4): ar_nmHead(n) = "salon_type_chain_eng"
                    n = n + 1: ar_Data(iii, n) = DateSerial(CnqYearDate, CnqMonthNum, 1): ar_nmHead(n) = "date_CNQ_Y"
                    n = n + 1: ar_Data(iii, n) = CnqMonthNum: ar_nmHead(n) = "date_month_num"
                    n = n + 1: ar_Data(iii, n) = CnqMonthName: ar_nmHead(n) = "date_month_name"
                    n = n + 1: ar_Data(iii, n) = CnqYearDate: ar_nmHead(n) = "date_year"
                    n = n + 1: ar_Data(iii, n) = CnqYearGA: ar_nmHead(n) = "GA_YEAR"
                    
                    '---------------------------------------------------------------------------------------------------------
                    'creat ca val loreal monthly
                    '---------------------------------------------------------------------------------------------------------
                    For f_ga = 1 To 2
                    cum_val = Empty
                    val_cell = Empty

                        Select Case f_ga
                                Case 1: str_clm = str_PYper_LOR_VAL: head_name = "CA_PY"
                                Case 2: str_clm = str_TYper_LOR_VAL: head_name = "CA_TY"
                        End Select

                        For f_m = 0 To 11
                            clm_m = str_clm + f_m
                            val_cell = Cells(f_i, clm_m)
                            cum_val = val_cell + cum_val

                            If val_cell = 0 Or Len(val_cell) = 0 Then
                                val_cell = Empty
                                ElseIf f_ga = 2 and f_YearTR = ThisYear And f_m > ThisMonth - 1 Then
                                    val_cell = Empty
                                    Else: val_cell = val_cell
                            End If

                            n = n + 1: ar_Data(iii, n) = myLib.num2numNull(myLib.getNumInThrousend(val_cell)): ar_nmHead(n) = head_name & "_M" & f_m + 1

                            If cum_val = 0 Or Len(cum_val) = 0 Then
                                cum_val = Empty
                                ElseIf f_ga =2 and f_YearTR = ThisYear And f_m > ThisMonth - 1 Then
                                    cum_val = Empty
                                    Else: cum_val = cum_val
                            End If
                            

                            nn = n + 24: ar_Data(iii, nn) = myLib.num2numNull(myLib.getNumInThrousend(cum_val)): ar_nmHead(nn) = head_name & "_YTD" & f_m + 1
                        Next f_m
                    Next f_ga
            
                    '---------------------------------------------------------------------------------------------------------
                    ' first conq order
                    '---------------------------------------------------------------------------------------------------------
                            val_cnq_TY = Empty
                            val_cnq_PY = Empty
                    Select Case CnqYearDate
                        Case actual_TY
                            clmStart = 93
                            clmCnq = clmStart + CnqMonthNum - 1
                            val_cnq_TY = IIf(Not IsEmpty(Cells(f_i, clmCnq)), Cells(f_i, clmCnq), Empty)

                        Case actual_PY
                            clmStart = 106
                            clmCnq = clmStart + CnqMonthNum - 1
                            val_cnq_PY = IIf(Not IsEmpty(Cells(f_i, clmCnq)), Cells(f_i, clmCnq), Empty)
                    End Select


                    n = nn + 1: ar_Data(iii, n) = myLib.num2numNull(myLib.getNumInThrousend(val_cnq_PY)): ar_nmHead(n) = "PY_CNQ_Order"
                    n = n + 1: ar_Data(iii, n) = myLib.num2numNull(myLib.getNumInThrousend(val_cnq_TY)): ar_nmHead(n) = "TY_CNQ_Order"
                End If
            Next f_i
            myLib.IsOpenTRtoClsd
        End If
    Next f_brnd
Next f_YearTR
    
Workbooks(NF).Activate
Sheets(in_data).Activate


Range(Cells(StartRowThisYear, 1), Cells(GetLastRow, GetLastColumn)).Cells.ClearContents
Range(Cells(1, 1), Cells(1, GetLastColumn)).Cells.ClearContents

For t = 1 To n
    Cells(1, t) = ar_nmHead(t)
    Cells(1, t).Select
    ActiveWorkbook.Names.Add Name:=ar_nmHead(t), RefersTo:="=" & ActiveSheet.Name & "!" & ActiveCell.Address()
Next t


ActiveSheet.Cells(StartRowThisYear, 1).Resize(iii + StartRowThisYear, n) = ar_Data()

'ActiveWorkbook.Names.Add Name:="SOURCE", RefersToR1C1:="=OFFSET(in_TR!R1C1,0,0,COUNTA(in_TR!R1C1:R999999C1),COUNTA(in_TR!R1C1:R1C1000))"
'ActiveWorkbook.Names("SOURCE").Comment = ""
ActiveWorkbook.RefreshAll

'---------------------------------------------------------------------------------------------------------

myLib.VBA_End

End Sub