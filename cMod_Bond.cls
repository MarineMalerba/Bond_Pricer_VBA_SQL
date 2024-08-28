VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMod_Bond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

'Attributes %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Private str_Company As String
Private str_Coupon_Rate_Type As String '("Variable" or "Fixed")
Private dbl_Coupon_Rate_Or_Margin As Double
Private dbl_Coupon_Frequency As Double 'number of coupons in a year
Private dbl_Maturity As Double 'in years

'Properties %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Property Get pCompany() As String
    pCompany = str_Company
End Property

Property Let pCompany(Company As String)
    str_Company = Company
End Property

Property Get pCoupon_Rate_Type() As String
    pCoupon_Rate_Type = str_Coupon_Rate_Type
End Property

Property Let pCoupon_Rate_Type(Coupon_Rate_Type As String)
    str_Coupon_Rate_Type = Coupon_Rate_Type
End Property

Property Get pCoupon_Rate_Or_Margin() As Double
    pCoupon_Rate_Or_Margin = dbl_Coupon_Rate_Or_Margin
End Property

Property Let pCoupon_Rate_Or_Margin(Coupon_Rate_Or_Margin As Double)
    dbl_Coupon_Rate_Or_Margin = Coupon_Rate_Or_Margin
End Property

Property Get pCoupon_Frequency() As Double
    pCoupon_Frequency = dbl_Coupon_Frequency
End Property

Property Let pCoupon_Frequency(Coupon_Frequency As Double)
    dbl_Coupon_Frequency = Coupon_Frequency
End Property

Property Get pMaturity() As Double
    pMaturity = dbl_Maturity
End Property

Property Let pMaturity(Maturity As Double)
    dbl_Maturity = Maturity
End Property

'Methods %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Function fn_Price() As Double

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'Returns the price of the bond

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'Variables declaration
Dim dbl_price As Double, i As Double, int_premier_coupon As Double

'Calculate the first coupon date
    int_premier_coupon = dbl_Maturity - Int(dbl_Maturity) + 1 / dbl_Coupon_Frequency
    
'Coupons
    For i = int_premier_coupon To dbl_Maturity Step (1 / dbl_Coupon_Frequency)
        If str_Coupon_Rate_Type = "Fixed" Then
            dbl_price = dbl_price + dbl_Coupon_Rate_Or_Margin / dbl_Coupon_Frequency / (1 + obj_RfCurve.fn_Interpolate(i) + obj_Spread.fn_Interpolate(i)) ^ i
        ElseIf str_Coupon_Rate_Type = "Variable" Then
            dbl_price = dbl_price + (dbl_Coupon_Rate_Or_Margin + obj_LIBOR.fn_Interpolate(i)) * 100 / dbl_Coupon_Frequency / (1 + obj_RfCurve.fn_Interpolate(i) + obj_Spread.fn_Interpolate(i)) ^ i
        End If
    Next i

'Repayment
    fn_Price = dbl_price + 100 / (1 + obj_RfCurve.fn_Interpolate(dbl_Maturity) + obj_Spread.fn_Interpolate(dbl_Maturity)) ^ dbl_Maturity

End Function

Function fn_Duration() As Double

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

' Calculates duration, adjusted for spreads, coupon frequency, and maturity.

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'Variables declaration
Dim dbl_duration As Double, i As Double, int_premier_coupon As Double

'Calculate the first coupon date
    int_premier_coupon = dbl_Maturity - Int(dbl_Maturity) + 1 / dbl_Coupon_Frequency

'Weighted coupons
    For i = int_premier_coupon To dbl_Maturity Step (1 / dbl_Coupon_Frequency)
        If str_Coupon_Rate_Type = "Fixed" Then
            dbl_duration = dbl_duration + (i * dbl_Coupon_Rate_Or_Margin) / dbl_Coupon_Frequency / (1 + obj_RfCurve.fn_Interpolate(i) + obj_Spread.fn_Interpolate(i)) ^ i
        ElseIf str_Coupon_Rate_Type = "Variable" Then
            dbl_duration = dbl_duration + (i * (dbl_Coupon_Rate_Or_Margin + obj_LIBOR.fn_Interpolate(i)) * 100) / dbl_Coupon_Frequency / (1 + obj_RfCurve.fn_Interpolate(i) + obj_Spread.fn_Interpolate(i)) ^ i
        End If
    Next i

'Repayment
    fn_Duration = (dbl_duration + dbl_Maturity * 100 / (1 + obj_RfCurve.fn_Interpolate(dbl_Maturity) + obj_Spread.fn_Interpolate(dbl_Maturity)) ^ dbl_Maturity) / fn_Price

End Function


Sub sub_Schedule()

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'Prints the cashflow schedule on a dedicated sheet

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'Variables declaration
Dim int_nb_coupons As Double, i As Double, int_premier_coupon As Double, dbl_CashFlow As Double, _
    dbl_Risk_Free_Rate As Double, dbl_Spread As Double, dbl_Discount_Factor As Double, dbl_Discounted_CashFlow As Double

'Clear data on sheet
    sht_Schedule.Cells.Clear

'Print titles
    sht_Schedule.Cells(1, 1).Resize(1, 6).Value = Array("Maturity", "CashFlow", "Risk Free Rate", "Spread", "Discount Factor", "Discounted CashFlow")

'Calculate the first coupon date
    int_premier_coupon = dbl_Maturity - Int(dbl_Maturity) + 1 / dbl_Coupon_Frequency
    
'Loop on time periods to fill the table
    For i = int_premier_coupon To dbl_Maturity Step (1 / dbl_Coupon_Frequency)
        
        'CashFlow
        If str_Coupon_Rate_Type = "Fixed" Then
            dbl_CashFlow = dbl_Coupon_Rate_Or_Margin / dbl_Coupon_Frequency
        ElseIf str_Coupon_Rate_Type = "Variable" Then
            dbl_CashFlow = (dbl_Coupon_Rate_Or_Margin + obj_LIBOR.fn_Interpolate(i)) * 100 / dbl_Coupon_Frequency
        End If
        If i = dbl_Maturity Then dbl_CashFlow = dbl_CashFlow + 100 'repayment
        
        'Risk Free Rate
        dbl_Risk_Free_Rate = obj_RfCurve.fn_Interpolate(i)
        
        'Spread
        dbl_Spread = obj_Spread.fn_Interpolate(i)
        
        'Discount Factor
        dbl_Discount_Factor = 1 / (1 + dbl_Risk_Free_Rate + dbl_Spread) ^ i
         
        'Discounted CashFlow
        dbl_Discounted_CashFlow = dbl_Discount_Factor * dbl_CashFlow
        
        ' Printing values to the sheet
        With sht_Schedule
            .Cells(.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = i
            .Cells(.Rows.Count, 2).End(xlUp).Offset(1, 0).Value = dbl_CashFlow / 100
            .Cells(.Rows.Count, 3).End(xlUp).Offset(1, 0).Value = dbl_Risk_Free_Rate
            .Cells(.Rows.Count, 4).End(xlUp).Offset(1, 0).Value = dbl_Spread
            .Cells(.Rows.Count, 5).End(xlUp).Offset(1, 0).Value = dbl_Discount_Factor
            .Cells(.Rows.Count, 6).End(xlUp).Offset(1, 0).Value = dbl_Discounted_CashFlow / 100
        End With
        
    Next i

'Page layout
    With sht_Schedule.UsedRange
        .Columns.AutoFit
        .Borders.Weight = xlThin
        .Rows(1).Font.Bold = True
        .Rows(1).Borders.Weight = xlMedium
    End With
        
'Pourcentage format
    sht_Schedule.Cells(2, 2).Resize(sht_Schedule.UsedRange.Rows.Count - 1, sht_Schedule.UsedRange.Columns.Count - 1).NumberFormat = "0.0000%"
    
End Sub


