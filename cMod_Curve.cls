VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMod_Curve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Attributes %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Private str_name As String
Private str_type As String
Private arr_maturity() As Double
Private arr_rate() As Double

'Properties %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Property Get pName() As String
        pName = str_name
End Property

Property Let pName(Name As String)
        str_name = Name
End Property

Property Get pType() As String
        pType = str_type
End Property

Property Let pType(Types As String)
        str_type = Types
End Property

Property Get pMaturity() As Double()
        pMaturity = arr_maturity
End Property

Property Let pMaturity(Maturity() As Double)
        arr_maturity = Maturity
End Property

Property Get pRate() As Double()
        pRate = arr_rate
End Property

Property Let pRate(Rate() As Double)
        arr_rate = Rate
End Property

'Methods %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Public Function fn_Interpolate(dbl_Maturity As Double) As Double

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'Returns the rate, given a certain maturity, by linear interpolation

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'Variables declaration
Dim dbl_lowerMaturity As Double, dbl_upperMaturity As Double, dbl_lowerRate As Double, dbl_upperRate As Double
Dim i As Integer

'If the specified maturity is below the minimum, the closest rate is used
    If dbl_Maturity < arr_maturity(1) Then
        fn_Interpolate = arr_rate(1)
        Exit Function
    End If

'If the specified maturity is above the maximum, the closest rate is used
    If dbl_Maturity > arr_maturity(UBound(arr_rate)) Then
        fn_Interpolate = arr_rate(UBound(arr_rate))
        Exit Function
    End If
    
'Finding the nearest maturity values with a do-while loop
    i = 1
    Do
        i = i + 1
    Loop While arr_maturity(i) < dbl_Maturity
    dbl_upperMaturity = arr_maturity(i)
    dbl_lowerMaturity = arr_maturity(i - 1)
    dbl_upperRate = arr_rate(i)
    dbl_lowerRate = arr_rate(i - 1)
    
    fn_Interpolate = dbl_lowerRate + (dbl_Maturity - dbl_lowerMaturity) * (dbl_upperRate - dbl_lowerRate) / (dbl_upperMaturity - dbl_lowerMaturity)

End Function
