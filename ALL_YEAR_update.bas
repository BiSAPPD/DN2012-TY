Attribute VB_Name = "ALL_YEAR_update"
Sub data_CC_in_GC()


Dim pathc2file As String
Dim ar_code_Brand(6, 1)
Dim LastRow_CC, LastColumns_CC As Integer
Dim num_month

Dim patchTR, actTR, ar_LastRow, in_data, status_head   As String
Dim ALLTIME, EDU_PY, EDU_TY, place, AVG_HD As Variant
Dim b, iii, i, x, y, frqOrder   As Integer
Dim ar_nmAVG_Order(), ar_type_clients(1 To 4, 1 To 12)

'Dim   As Worksheet
Dim ar_Data(), ar_brand(), ar_nmHead(150), ar_PYPer_PRTN_VAL, ar_TYPer_PRTN_VAL, ar_PYPer_LOR_VAL(), ar_TYPer_LOR_VAL(), ar_nm_month(), ar_nmMregEN(), ar_nmMregLT()

NF = ActiveWorkbook.Name
act_month = InputBox("Month")
act_month = CInt(act_month)

'colums CA PRTNN VAL for LTM
'---------------------------------------------------------------------------------------------------------
ar_PYPer_PRTN_VAL = Array(0, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89)
ar_TYPer_PRTN_VAL = Array(0, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77)


'colums CA LOREAL VAL
str_PYper_LOR_VAL = 106
str_TYper_LOR_VAL = 93
str_PYper_PRTN_VAL = 79
str_TYper_PRTN_VAL = 66

ar_nm_month = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

ar_nmMregEN = Array("MOSCOW", "GR", "NORTHWEST", "CENTER", "VOLGA", "SOUTH", "URAL", "SIBERIA", "FAR EAST")
ar_nmMregLT = Array("Moscou", "GR", "Nord-Ouest", "Centre", "Volga-Centre", "Sud", "Oural", "Siberie", "EO")
ar_nmCompetitors = Array("Estel", "Schwarzkopf", "Wella", "Londa", "Keune", "Revlon", "Goldwell", "Cutrin", "Kadus", "Indola", "Paul Mitchell", "Label", "Syoss", "Chi", "Seah", "Kydra", "Sebastian", "American Crew", "Alterna", "Other")


ar_type_clients(1, 1) = "салон"
ar_type_clients(2, 1) = "salon"
ar_type_clients(3, 1) = "salon"
ar_type_clients(4, 1) = "single"
ar_type_clients(1, 2) = "сеть салонов"
ar_type_clients(2, 2) = "chain_salons"
ar_type_clients(3, 2) = "salon"
ar_type_clients(4, 2) = "chain"
ar_type_clients(1, 3) = "ч/м"
ar_type_clients(2, 3) = "hdres"
ar_type_clients(3, 3) = "salon"
ar_type_clients(4, 3) = "single"
ar_type_clients(1, 4) = "сеть магазинов"
ar_type_clients(2, 4) = "chain_shops"
ar_type_clients(3, 4) = "shop"
ar_type_clients(4, 4) = "chain"
ar_type_clients(1, 5) = "магазин"
ar_type_clients(2, 5) = "shop"
ar_type_clients(3, 5) = "shop"
ar_type_clients(4, 5) = "single"
ar_type_clients(1, 6) = "салон-маг."
ar_type_clients(2, 6) = "salon"
ar_type_clients(3, 6) = "salon"
ar_type_clients(4, 6) = "single"
ar_type_clients(1, 7) = "(пусто)"
ar_type_clients(2, 7) = "other"
ar_type_clients(3, 7) = "other"
ar_type_clients(4, 7) = "single"
ar_type_clients(1, 8) = "школа"
ar_type_clients(2, 8) = "school"
ar_type_clients(3, 8) = "school"
ar_type_clients(4, 8) = "single"
ar_type_clients(1, 9) = "другое"
ar_type_clients(2, 9) = "other"
ar_type_clients(3, 9) = "other"
ar_type_clients(4, 9) = "single"
ar_type_clients(1, 10) = "нейл-бар"
ar_type_clients(2, 10) = "nails_bar"
ar_type_clients(3, 10) = "nails"
ar_type_clients(4, 10) = "single"
ar_type_clients(1, 11) = "сеть нейл-баров"
ar_type_clients(2, 11) = "chain_nails"
ar_type_clients(3, 11) = "nails"
ar_type_clients(4, 11) = "chain"
ar_type_clients(1, 12) = "e-commerce"
ar_type_clients(2, 12) = "e-commerce"
ar_type_clients(3, 12) = "e-commerce"
ar_type_clients(4, 12) = "single"






ar_nmAVG_Order = Array(0, 2.5, 5, 10, 15, 20, 25, 30, 50, 60, 70, 100000)
'---------------------------------------------------------------------------------------------------------

With Application
.ScreenUpdating = False
.EnableEvents = False
.Calculation = xlCalculationManual
'.DisplayPageBreaks = False
.DisplayAlerts = False
End With

status_head = 0

ReDim Preserve ar_Data(999999, 150)

ar_brand = Array("LP", "KR", "RD", "MX", "ES", "DE", "CR")
num_ar_brand = UBound(ar_brand)
array_row = 0

iii = 0
For f_year = 2011 To 2016
actual_TY = f_year
actual_PY = f_year - 1

If f_year <> 2016 Then
act_month_Y = 12
Else
act_month_Y = act_month
End If


For b = 0 To UBound(ar_brand)
        
patchTR = "p:\DPP\Business development\Book commercial\" & ar_brand(b) & "\Top Russia Total " & f_year & " " & ar_brand(b) & ".xlsm"
in_data = "in_TR"

If Dir(patchTR) = "" Then
Exit For
Else
Workbooks.Open Filename:=patchTR, Notify:=False
End If

actTR = ActiveWorkbook.Name
Sheets(ar_brand(b)).Select
ActiveSheet.AutoFilterMode = False
ar_LastRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1  ' ????????? ??????

array_row = array_row + ar_LastRow - 4




    
    For i = 4 To ar_LastRow
    
    num_colums = 0
    ar_Data(iii, num_colums) = f_year
    ar_nmHead(num_colums) = "TR_year"
    
    
    num_colums = num_colums + 1
    nm_brand = ar_brand(b)
    ar_Data(iii, num_colums) = ar_brand(b)
    ar_nmHead(num_colums) = "brand"
    
    
    num_colums = num_colums + 1
     If ar_brand(b) = "LP" Or ar_brand(b) = "MX" Or ar_brand(b) = "KR" Or ar_brand(b) = "RD" Then type_brand = "Hair"
          
     If ar_brand(b) = "ES" Then type_brand = "Nails"
          
     If ar_brand(b) = "DE" Or ar_brand(b) = "CR" Then type_brand = "Skin"
    
    ar_Data(iii, num_colums) = type_brand
    ar_nmHead(num_colums) = "bussines"
            
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = Cells(i, 1)
    ar_nmHead(num_colums) = "rowTR"
    
    num_colums = num_colums + 1
    cd_brand_row = ar_brand(b) & Cells(i, 1)
    ar_Data(iii, num_colums) = cd_brand_row
    ar_nmHead(num_colums) = "BRAND_rowTR"
    
    num_colums = num_colums + 1
    cd_Univers = Cells(i, 2)
    If Len(cd_Univers) <> 9 Then
    cd_Univers = cd_brand_row
    Else: cd_Univers = cd_Univers
    End If
    ar_Data(iii, num_colums) = cd_Univers
    ar_nmHead(num_colums) = "unvCD"

    
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = ar_brand(b) & Cells(i, 2)
    ar_nmHead(num_colums) = "BRAND_unvCD"
    
    num_colums = num_colums + 1
    nm_Mreg = Right(Cells(i, 4), Len(Cells(i, 4).Value) - 3)
    ar_Data(iii, num_colums) = nm_Mreg
    ar_nmHead(num_colums) = "mreg"
    
'Mreg LT-> EN + split Moscou GR
'---------------------------------------------------------------------------------------------------------
            
    num_colums = num_colums + 1
    textPos = 0
    
    If nm_Mreg = "Moscou GR" Then
    nm_REG = Cells(i, 5)
    textPos = InStr(nm_REG, "MSK")
    textPos = InStr(nm_REG, "Moscou") + textPos
        If textPos > 0 Then
        nm_Mreg = "Moscou"
        Else
        nm_Mreg = "GR"

        End If
    End If
    
    For f_mr = 0 To UBound(ar_nmMregLT)
    If ar_nmMregLT(f_mr) = nm_Mreg Then
    nm_Mreg = ar_nmMregEN(f_mr)
    End If
    Next f_mr
    
    ar_Data(iii, num_colums) = nm_Mreg
    ar_nmHead(num_colums) = "mreg_EXT"
    
'---------------------------------------------------------------------------------------------------------
        
    
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = Cells(i, 5)
    ar_nmHead(num_colums) = "REG"
    
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = Cells(i, 165)
    ar_nmHead(num_colums) = "FLSM"
    
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = Cells(i, 6)
    ar_nmHead(num_colums) = "SEC"
    
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = Cells(i, 7)
    ar_nmHead(num_colums) = "SREP"
        
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = Left(Cells(i, 9), 40) & ". " & Left(Cells(i, 12), 50) & ", " & Left(Cells(i, 13), 20) & ", " & Left(Cells(i, 11), 20)
    ar_nmHead(num_colums) = "salon"
    
 
     num_colums = num_colums + 1
    ar_Data(iii, num_colums) = Cells(i, 19)
    ar_nmHead(num_colums) = "Chain_name"
    
     num_colums = num_colums + 1
    ar_Data(iii, num_colums) = Cells(i, 11)
    ar_nmHead(num_colums) = "city"
    
    
    num_colums = num_colums + 1
    type_sln_rus = Trim(Cells(i, 18))
    If Len(type_sln_rus) = 0 Then type_sln_rus = "салон"
    ar_Data(iii, num_colums) = type_sln_rus
    ar_nmHead(num_colums) = "type_SLN"
    
    For f_sl = 1 To 12
     
 
    If StrComp(ar_type_clients(1, f_sl), type_sln_rus, vbTextCompare) = 0 Then
        
        nm_salon_type_eng = ar_type_clients(2, f_sl)
        nm_salon_type_short_eng = ar_type_clients(3, f_sl)
        nm_salon_type_chain_eng = ar_type_clients(4, f_sl)
        Exit For
        Else
        nm_salon_type_eng = ""
        nm_salon_type_short_eng = ""
        nm_salon_type_chain_eng = ""
    End If
    Next f_sl
    
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = nm_salon_type_eng
    ar_nmHead(num_colums) = "salon_type_eng"
    
    num_colums = num_colums + 1


ar_Data(iii, num_colums) = nm_salon_type_short_eng
    ar_nmHead(num_colums) = "salon_type_short_eng"
    
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = nm_salon_type_chain_eng
    ar_nmHead(num_colums) = "salon_type_chain_eng"
    
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = Cells(i, 42)
    ar_nmHead(num_colums) = "type_CLUB"
    
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = Cells(i, 40)
    ar_nmHead(num_colums) = "type_confirmed_CLUB"
   
   
    
    num_colums = num_colums + 1
        If Cells(i, 161) <> "" Then cdMonth = Cells(i, 161) Else cdMonth = 1
        If Len(Cells(i, 65)) = 4 Then cdYear = Cells(i, 65) Else cdYear = 2008
    
    ar_Data(iii, num_colums) = cdMonth & "-" & cdYear
    ar_nmHead(num_colums) = "date_CNQ_Y"
    
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = cdMonth
    ar_nmHead(num_colums) = "date_month_num"
    
    
    num_colums = num_colums + 1

    For f_m = 0 To 11
    
    If cdMonth - 1 = f_m Then
    nmMonth = ar_nm_month(f_m)
    Exit For
    End If
    Next f_m
    
    ar_Data(iii, num_colums) = nmMonth
    ar_nmHead(num_colums) = "date_month_name"
    
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = cdYear
    ar_nmHead(num_colums) = "date_year"
    
'---------------------------------------------------------------------------------------------------------

Select Case CInt(cdYear)

    Case actual_TY
    GA_Y = "CNQ_TY"

    Case actual_PY
    GA_Y = "CNQ_PY"

    Case Else
    GA_Y = "PPY"


End Select
  
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = GA_Y
    ar_nmHead(num_colums) = "GA_YEAR"
    
        
    num_colums = num_colums + 1
    vl_mag = Cells(i, 160)
    If Len(vl_mag) <> 2 Then vl_mag = Null
    ar_Data(iii, num_colums) = vl_mag
    ar_nmHead(num_colums) = "type_MAG"
    
    num_colums = num_colums + 1
      ar_Data(iii, num_colums) = Cells(i, 158)
    ar_nmHead(num_colums) = "type_MAG_PRICE"
    
    num_colums = num_colums + 1
      ar_Data(iii, num_colums) = Cells(i, 159)
    ar_nmHead(num_colums) = "type_MAG_type_place"
    
    num_colums = num_colums + 1
    st_dn_cln = Cells(i, 8)
    ar_Data(iii, num_colums) = st_dn_cln
    ar_nmHead(num_colums) = "status_DN_num"
    
    num_colums = num_colums + 1
    If Cells(i, 8) = 1 Then
    st_cln_base = "Active"

    Else
    st_cln_base = "Closed"

    End If
    ar_Data(iii, num_colums) = st_cln_base
    ar_nmHead(num_colums) = "status_DN_name"
       
'---------------------------------------------------------------------------------------------------------
'   calculate LTM AVG CA & FrqRate
'---------------------------------------------------------------------------------------------------------
    sumCA12M = 0
    frqOrder = 0
    
    
    For iq = act_month_Y To 11
    
    
        If isNumeric(Cells(i, ar_PYPer_PRTN_VAL(iq))) Then
        CA = Cells(i, ar_PYPer_PRTN_VAL(iq))
        Else
        CA = 0
        End If
        
        sumCA12M = sumCA12M + CA
        If Cells(i, ar_PYPer_PRTN_VAL(iq)) <> "" And Cells(i, ar_PYPer_PRTN_VAL(iq)) > 0 Then
        frqOrder = frqOrder + 1
        End If
    
    Next iq
    
    For iw = 1 To act_month_Y
    
    If isNumeric(Cells(i, ar_TYPer_PRTN_VAL(iw))) Then
        CA = Cells(i, ar_TYPer_PRTN_VAL(iw))
        Else
        CA = 0
        End If
    
    sumCA12M = sumCA12M + CA
        If Cells(i, ar_TYPer_PRTN_VAL(iw)) <> "" And Cells(i, ar_TYPer_PRTN_VAL(iw)) > 0 Then
        frqOrder = frqOrder + 1
        End If
    
    Next iw
            
        If sumCA12M <> 0 Then
        AVG_CA_LTM = Round(sumCA12M / 12 / 1000, 1)
        Else
        AVG_CA_LTM = ""
        End If
'---------------------------------------------------------------------------------------------------------
    
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = AVG_CA_LTM
    ar_nmHead(num_colums) = "CA_AVG_LTM"
   
   
   
    num_colums = num_colums + 1
    For f_avg = 1 To UBound(ar_nmAVG_Order())
        If AVG_CA_LTM <= ar_nmAVG_Order(f_avg) And AVG_CA_LTM > ar_nmAVG_Order(f_avg - 1) Then
        
        nm_avg_CA = "'" & ar_nmAVG_Order(f_avg - 1) & "-" & ar_nmAVG_Order(f_avg)
        Exit For
        Else
        nm_avg_CA = Null
        End If
    Next f_avg
    
        If nm_avg_CA = 100000 Then nm_avg_CA = ">70"
       
    
    ar_Data(iii, num_colums) = nm_avg_CA
    ar_nmHead(num_colums) = "CA_AVG_LTM_name"
    
    
    

    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = frqOrder & "\12" '
    ar_nmHead(num_colums) = "frq_order_LTM"
    
    
    num_colums = num_colums + 1
        ev_ca = Cells(i, 92)

        If isNumeric(ev_ca) Then
        ev_ca = Round(ev_ca, 2)



        Else
        ev_ca = Null
        End If
    ar_Data(iii, num_colums) = ev_ca
    ar_nmHead(num_colums) = "CA_ev"
   
 ' ev CA vector
 '---------------------------------------------------------------------------------------------------------
        num_colums = num_colums + 1
        ev_ca = Cells(i, 92)



        If isNumeric(ev_ca) Then
        Select Case ev_ca
        Case Is > 0
        nm_ev_ca = "+"

        Case Is < 0
        nm_ev_ca = "-"

        Case Else
        nm_ev_ca = Null


        End Select
        Else
        nm_ev_ca = Null


End If
                 
    ar_Data(iii, num_colums) = nm_ev_ca
    ar_nmHead(num_colums) = "CA_ev_name"
 '---------------------------------------------------------------------------------------------------------
    
    
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = Cells(i, 29)
    ar_nmHead(num_colums) = "EDU_id_ECAD"
    
    num_colums = num_colums + 1
    EDU_ALLTIME_MSTR = Cells(i, 30)
        If isNumeric(EDU_ALLTIME_MSTR) And EDU_ALLTIME_MSTR <> 0 Then
        EDU_ALLTIME_MSTR = Round(EDU_ALLTIME_MSTR, 0)
        Else
        EDU_ALLTIME_MSTR = ""
        End If
    ar_Data(iii, num_colums) = EDU_ALLTIME_MSTR
    ar_nmHead(num_colums) = "EDU_ALLTIME_MSTR"
    
    num_colums = num_colums + 1
    EDU_PY_MSTR = Cells(i, 31)
        If isNumeric(EDU_PY_MSTR) And EDU_PY_MSTR <> 0 Then
        EDU_PY = Round(EDU_PY_MSTR, 0)
        Else
        EDU_PY = ""
        End If
    ar_Data(iii, num_colums) = EDU_PY_MSTR
    ar_nmHead(num_colums) = "EDU_PY_MSTR"
    
    num_colums = num_colums + 1
    EDU_TY_MSTR = Cells(i, 32)
        If isNumeric(EDU_TY_MSTR) And EDU_TY_MSTR <> 0 Then
        EDU_TY_MSTR = Round(EDU_TY_MSTR, 0)
        Else
        EDU_TY = ""
        End If
    ar_Data(iii, num_colums) = EDU_TY_MSTR
    ar_nmHead(num_colums) = "EDU_TY_MSTR"
    
    
    num_colums = num_colums + 1
    EDU_ALLTIME_CNTCT = Cells(i, 33)
        If isNumeric(EDU_ALLTIME_CNTCT) And EDU_ALLTIME_CNTCT <> 0 Then
        EDU_ALLTIME = Round(EDU_ALLTIME_CNTCT, 0)
        Else
        EDU_ALLTIME = ""
        End If
    ar_Data(iii, num_colums) = EDU_ALLTIME_CNTCT
    ar_nmHead(num_colums) = "EDU_ALLTIME_CNTCT"
    
    num_colums = num_colums + 1
    EDU_PY_CNTCT = Cells(i, 34)
        If isNumeric(EDU_PY_CNTCT) And EDU_PY_CNTCT <> 0 Then
        EDU_PY = Round(EDU_PY_CNTCT, 0)
        Else
        EDU_PY = ""
        End If
    ar_Data(iii, num_colums) = EDU_PY_CNTCT
    ar_nmHead(num_colums) = "EDU_PY_CNTCT"
    
    num_colums = num_colums + 1
    EDU_TY_CNTCT = Cells(i, 35)
        If isNumeric(EDU_TY_CNTCT) And EDU_TY_CNTCT <> 0 Then
        EDU_TY = Round(EDU_TY_CNTCT, 0)
        Else
        EDU_TY = ""
        End If
    ar_Data(iii, num_colums) = EDU_TY_CNTCT
    ar_nmHead(num_colums) = "EDU_TY_CNTCT"
    
      
    
    
    num_colums = num_colums + 1
    place = Cells(i, 27)
        If isNumeric(place) Then
        place = Round(place, 0)
        Else
        place = ""
        End If
    ar_Data(iii, num_colums) = place
    ar_nmHead(num_colums) = "type_place_HD"
    
    num_colums = num_colums + 1
    AVG_HD = Cells(i, 28)
        If isNumeric(AVG_HD) Then
        AVG_HD = Round(AVG_HD, 0)
        Else
        place = ""
        End If
    ar_Data(iii, num_colums) = AVG_HD
    ar_nmHead(num_colums) = "type_AVG_HD"
    
    
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = Cells(i, 209)
    ar_nmHead(num_colums) = "com_KPI"
    
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = Cells(i, 167)
    ar_nmHead(num_colums) = "nm_PRTNner"
          
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = Cells(i, 173)
    ar_nmHead(num_colums) = "cd_PRTNner"
'---------------------------------------------------------------------------------------------------------
'creat ca val loreal monthly
'---------------------------------------------------------------------------------------------------------

    For f_m = 0 To 11
    num_colums = num_colums + 1
    clm_m = str_PYper_LOR_VAL + f_m
    If Cells(i, clm_m) = 0 Then
    m_val = Null

    Else
    m_val = Cells(i, clm_m) / 1000

    End If
    ar_Data(iii, num_colums) = m_val
    ar_nmHead(num_colums) = "CA_PY_M" & f_m + 1

    Next f_m
    

    For f_m = 0 To 11
    num_colums = num_colums + 1
    clm_m = str_TYper_LOR_VAL + f_m
    If Cells(i, clm_m) = 0 Then
    m_val = Null

    Else
    m_val = Cells(i, clm_m) / 1000

    End If
    ar_Data(iii, num_colums) = m_val
    ar_nmHead(num_colums) = "CA_TY_M" & f_m + 1


Next f_m
 '---------------------------------------------------------------------------------------------------------
  'creat ca val loreal cumul
'---------------------------------------------------------------------------------------------------------
    
    m_valP = 0

    For f_m = 0 To 11
    num_colums = num_colums + 1
    clm_m = str_PYper_LOR_VAL + f_m
    m_val = (Cells(i, clm_m) / 1000) + m_valP
    
    
    If m_val = 0 Then  ' del 0 value out
    ar_Data(iii, num_colums) = Null

    Else
    ar_Data(iii, num_colums) = m_val

    End If
    
    ar_nmHead(num_colums) = "CA_PY_YTD" & f_m + 1
    m_valP = m_val

    Next f_m
    
    
    m_valP = 0
    For f_m = 0 To 11  ' limit tange by actuale period
    num_colums = num_colums + 1
    If f_m < CInt(act_month_Y) Then
    clm_m = str_TYper_LOR_VAL + f_m
    m_val = (Cells(i, clm_m) / 1000) + m_valP

    Else
    m_val = Null

    End If
    
    If m_val = 0 Then  ' del 0 value out
    ar_Data(iii, num_colums) = Null

    Else
    ar_Data(iii, num_colums) = m_val

    End If
    
    ar_nmHead(num_colums) = "CA_TY_YTD" & f_m + 1
    m_valP = m_val

    Next f_m
    
'---------------------------------------------------------------------------------------------------------
'creat  ca val loreal Quarter
'---------------------------------------------------------------------------------------------------------
 
    q_m_c = 0
    For f_q = 0 To 3
    num_colums = num_colums + 1
    m_val_q = 0
    m_val = 0
    
        For f_m = 0 To 2
        clm_m = str_PYper_LOR_VAL + q_m_c
        m_val = Cells(i, clm_m)
        m_val_q = m_val_q + m_val
        
        q_m_c = q_m_c + 1
        
        Next f_m
        
        If m_val_q = 0 Then
        m_val_q = Null

        Else
        m_val_q = m_val_q / 1000
        End If
           
    ar_Data(iii, num_colums) = m_val_q
    ar_nmHead(num_colums) = "CA_PY_Q_" & f_q + 1
    Next f_q
    
    
   q_m_c = 0
    For f_q = 0 To 3
    num_colums = num_colums + 1
    m_val_q = 0
    m_val = 0
    
        For f_m = 0 To 2
        clm_m = str_TYper_LOR_VAL + q_m_c
        m_val = Cells(i, clm_m)
        m_val_q = m_val_q + m_val
        
        q_m_c = q_m_c + 1
        
        Next f_m
        
        If m_val_q = 0 Then
        m_val_q = Null

        Else
        m_val_q = m_val_q / 1000
        End If
           
    ar_Data(iii, num_colums) = m_val_q
    ar_nmHead(num_colums) = "CA_TY_Q" & f_q + 1
    Next f_q
    
 '---------------------------------------------------------------------------------------------------------
 ' first conq order
 '---------------------------------------------------------------------------------------------------------
    

    
    num_cnq_year = CInt(cdYear)
    num_cnq_month = CInt(act_month_Y)
    
    
        On Error Resume Next
          Select Case num_cnq_year
          Case actual_PY
          fst_order_LOR_PY = Cells(i, str_PYper_LOR_VAL + cdMonth - 1) / 1000
          fst_order_PRTN_PY = Cells(i, str_PYper_PRTN_VAL + cdMonth - 1) / 1000
          fst_order_LOR_TY = Null
          fst_order_PRTN_TY = Null
          Case actual_TY
          fst_order_LOR_TY = Cells(i, str_TYper_LOR_VAL + cdMonth - 1) / 1000
          fst_order_PRTN_TY = Cells(i, str_TYper_PRTN_VAL + cdMonth - 1) / 1000
          fst_order_LOR_PY = Null
          fst_order_PRTN_PY = Null
          Case Else
          fst_order_LOR_PY = Null
          fst_order_PRTN_PY = Null
          fst_order_LOR_TY = Null
          fst_order_PRTN_TY = Null
          End Select


num_colums = num_colums + 1
    ar_Data(iii, num_colums) = fst_order_LOR_PY
    ar_nmHead(num_colums) = "PY_CNQ_Order"
    
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = fst_order_LOR_TY
    ar_nmHead(num_colums) = "TY_CNQ_Order"
    
    
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = fst_order_PRTN_PY
    ar_nmHead(num_colums) = "PY_CNQ_Order_PRTN_CA"
    
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = fst_order_PRTN_TY
    ar_nmHead(num_colums) = "TY_CNQ_Order_PRTN_CA"
    
  
'---------------------------------------------------------------------------------------------------------
  'creat ca val loreal PYvsTY YTD
'---------------------------------------------------------------------------------------------------------
    
    
    num_colums = num_colums + 1
    m_valP = 0
    For f_m = 0 To 11  ' limit tange by actuale period
    If f_m < CInt(act_month_Y) Then
    clm_m = str_PYper_LOR_VAL + f_m
    m_val = (Cells(i, clm_m) / 1000) + m_valP
    Else
    Exit For
    End If
      
    m_valP = m_val

    Next f_m
    
    If m_val = 0 Then  ' del 0 value out
    ar_Data(iii, num_colums) = Null

    Else
    ar_Data(iii, num_colums) = m_val

    End If
    ar_nmHead(num_colums) = "CA_PY_YTD"
    ca_ytd_PY = m_val
    
    
    num_colums = num_colums + 1
    m_valP = 0
    For f_m = 0 To 11  ' limit tange by actuale period
    If f_m < CInt(act_month_Y) Then
    clm_m = str_TYper_LOR_VAL + f_m
    m_val = (Cells(i, clm_m) / 1000) + m_valP



    Else
    Exit For
    End If
    
    m_valP = m_val

    Next f_m
    
    If m_val = 0 Then  ' del 0 value out
    ar_Data(iii, num_colums) = Null

    Else
    ar_Data(iii, num_colums) = m_val

    End If
    ar_nmHead(num_colums) = "CA_TY_YTD"
    ca_ytd_TY = m_val
    
    num_colums = num_colums + 1
    
    
    If ca_ytd_PY <> 0 And ca_ytd_TY = 0 Then
    type_cln_react = "lost"
    ca_ytd_PY_lost = ca_ytd_PY * -1
    End If

    
    If ca_ytd_PY = 0 And ca_ytd_TY = 0 Then
    sts_clnt_act = "null"
    Else
    type_cln_react = "act"
    ca_ytd_PY_lost = Empty
    End If
        
        
    If sts_clnt_act <> 0 Then
        type_cln_react = "lost"
        ca_ytd_PY_lost = ca_ytd_PY
        Else
        type_cln_react = "act"
        ca_ytd_PY_lost = Empty
    End If
    ar_Data(iii, num_colums) = type_cln_react
    ar_nmHead(num_colums) = "type_LOST"
    
         
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = ca_ytd_PY_lost
    ar_nmHead(num_colums) = "CA_LOST_PY"
    
    
'---------------------------------------------------------------------------------------------------------
'dt_constante
'---------------------------------------------------------------------------------------------------------
    num_colums = num_colums + 1
    dt_st = 0
    
    If ca_ytd_PY > 0 Then dt_st = dt_st + 1
    If ca_ytd_TY > 0 Then dt_st = dt_st + 1
    If st_dn_cln = 1 Then dt_st = dt_st + 1
    If GA_Y = "PPY" Then dt_st = dt_st + 1

    If dt_st = 4 Then
    dt_st_nm = 1

    Else
    dt_st_nm = 0

    End If
    ar_Data(iii, num_colums) = dt_st_nm
    ar_nmHead(num_colums) = "LfL"
 
    '---------------------------------------------------------------------------------------------------------
    If dt_st_nm = 1 Then
    ca_ytd_PY_dt = ca_ytd_PY
    ca_ytd_TY_dt = ca_ytd_TY
    Else
    ca_ytd_PY_dt = Null
    ca_ytd_TY_dt = Null

    End If
    
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = ca_ytd_PY_dt
    ar_nmHead(num_colums) = "CA_PY_LfL"
    
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = ca_ytd_TY_dt
    ar_nmHead(num_colums) = "CA_TY_LfL"

'---------------------------------------------------------------------------------------------------------
'CA YTD split by GA


    For f_qe = 1 To 3
    


        Select Case f_qe
        Case 1
        find_GA_Y = "PPY"

        Case 2
        find_GA_Y = "CNQ_PY"

        Case 3
        find_GA_Y = "CNQ_TY"

        End Select

    If GA_Y = find_GA_Y Then
    ca_ytd_PY_GA = ca_ytd_PY
    ca_ytd_TY_GA = ca_ytd_TY


Else
    ca_ytd_PY_GA = Null
    ca_ytd_TY_GA = Null

    End If
          
    If ca_ytd_PY_GA = 0 Then ca_ytd_PY_GA = Null
    If ca_ytd_TY_GA = 0 Then ca_ytd_TY_GA = Null
    
         
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = ca_ytd_PY_GA
    ar_nmHead(num_colums) = "CA_PY_" & find_GA_Y
 
    num_colums = num_colums + 1
    ar_Data(iii, num_colums) = ca_ytd_TY_GA
    ar_nmHead(num_colums) = "CA_TY_" & find_GA_Y
    

    Next f_qe

'---------------------------------------------------------------------------------------------------------
'creat closed data
'---------------------------------------------------------------------------------------------------------
sts_cls_f = False
num_clsd_month = Empty
num_clsd_year = Empty
clm_m = 0

If st_dn_cln = 0 Then


    For f_yy = 1 To 2
        
        Select Case f_yy
            Case 1
            strt_month_clm = str_TYper_LOR_VAL
            Case 2
            strt_month_clm = str_PYper_LOR_VAL
        End Select

       
        
        For f_m = 11 To 0 Step -1


        clm_m = strt_month_clm + f_m
        
If f_yy = 2 Then

End If
            
        If Cells(i, clm_m) <> 0 Then
        num_clsd_month = 1 + f_m
        
        Select Case f_yy
            Case 1
            num_clsd_year = actual_TY
            Case 2
            num_clsd_year = actual_PY
        End Select
        
        sts_cls_f = True
        Exit For
        End If
    
        Next f_m
            If sts_cls_f = True Then
            Exit For
            End If
    Next f_yy
Else

End If
    
num_colums = num_colums + 1
ar_Data(iii, num_colums) = num_clsd_month
ar_nmHead(num_colums) = "Closed_M"

num_colums = num_colums + 1
ar_Data(iii, num_colums) = num_clsd_year
ar_nmHead(num_colums) = "Closed_Y"

'---------------------------------------------------------------------------------------------------------
iii = iii + 1
Next i
array_row = iii


Workbooks(actTR).Close
Application.DisplayAlerts = False

Next b
Next f_year

  
    
Workbooks(NF).Activate
If Sheets(in_data).Visible = False Then
Sheets(in_data).Visible = True
End If
Sheets(in_data).Activate

'clear sheet & create head & create name OR fiil data
'---------------------------------------------------------------------------------------------------------
If status_head = 0 Then
ActiveSheet.UsedRange.Cells.ClearContents
end_POS = iii + 1
start_POS = 2

Dim n As Name
For Each n In ThisWorkbook.Names
    On Error Resume Next
    n.Delete
    Next n

For t = 0 To num_colums
Cells(1, t + 1) = ar_nmHead(t)
Cells(1, t + 1).Select
ActiveWorkbook.Names.Add Name:=ar_nmHead(t), RefersTo:="=" & ActiveSheet.Name & "!" & ActiveCell.Address()
Next t

Else
start_POS = end_POS + 1
end_POS = start_POS + iii - 1
End If

ActiveSheet.Cells(start_POS, 1).Resize(end_POS - start_POS + 1, num_colums + 1) = ar_Data()
status_head = 1


ActiveWorkbook.Names.Add Name:="SOURCE", RefersToR1C1:="=OFFSET(in_TR!R1C1,0,0,COUNTA(in_TR!R1C1:R65535C1),COUNTA(in_TR!R1C1:R1C255))"
'ActiveWorkbook.Names("SOURCE").Comment = ""
Sheets(in_data).Visible = False
ActiveWorkbook.RefreshAll

'---------------------------------------------------------------------------------------------------------

With Application
.ScreenUpdating = True
.Calculation = xlCalculationAutomatic
.EnableEvents = True
.DisplayStatusBar = True
.DisplayAlerts = True
End With

End Sub

