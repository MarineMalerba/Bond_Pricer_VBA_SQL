Attribute VB_Name = "Mod_Interface_and_Instantiation"
Option Explicit
Option Base 1

'Declaration and static instantiation
    'Risk-free curve
    Public obj_RfCurve As New cMod_Curve
    
    'Spread curve
    Public obj_Spread As New cMod_Curve
    
    'LIBOR curve
    Public obj_LIBOR As New cMod_Curve

Sub sub_Reset_Interface()

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'Creates an interface for a bond pricer, preparing input and output tables
'The private sub Workbook_Open() launches this sub automatically when the file is opened


'       Table of Contents:

'          1. Link between Excel & Access database
'                   • Clear data
'                   • Initializing and opening the connection

'          2. Creating an interface for a bond pricer
'                   • Print the titles
'                   • Page layout
'                   • Name the cells
'                   • Dropdown menus


'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'Variables declaration
Dim obj_Cnn As ADODB.Connection
Dim obj_Rst As ADODB.Recordset
Dim str_SQLRequest As String
Dim i As Integer, j As Integer
Dim arr_Companies() As String
Dim name_cell As Name

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'   1. Link between Excel & Access database
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'Initializing the connection
    Set obj_Cnn = New ADODB.Connection

'Connection parameter
    obj_Cnn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "Data Source= C:\Users\zouli\OneDrive\Documents\Data_Projet.accdb"

'Opening the connection
    obj_Cnn.Open

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'   2. Creating an interface for a bond pricer
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'Clear data on sheet
    sht_interface.Cells.ClearHyperlinks
    sht_interface.Cells.Clear
    
    'Clear previous range names
    For Each name_cell In sht_interface.Names
        name_cell.Delete
    Next name_cell
    
'Print the titles
    With sht_interface
        'Title
        .Cells(2, 2).Value = "Corporate Bond Pricer"
        .Cells(2, 2).Font.Bold = True
        
        'Inputs
        .Cells(5, 2).Value = "Inputs Parameters"
        .Cells(5, 2).Font.Bold = True
        .Cells(6, 2).Value = "Company"
        .Cells(7, 2).Value = "Coupon rate type"
        .Cells(8, 2).Value = "Coupon rate / Margin"
        .Cells(9, 2).Value = "Coupon frequency"
        .Cells(10, 2).Value = "Maturity"
    
        'Outputs
        .Cells(5, 6).Value = "Outputs"
        .Cells(5, 6).Font.Bold = True
        .Cells(6, 6).Value = "Spread to maturity"
        .Cells(7, 6).Value = "Price"
        .Cells(8, 6).Value = "Duration"
        .Cells(10, 6).Value = "Payment schedule"
    End With

'Page layout
    sht_interface.Cells(2, 2).Resize(9, 6).Columns.AutoFit

    'Title
    With sht_interface.Cells(2, 2).Resize(2, 6)
        .Merge
        .Borders.Weight = xlThin
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
    End With
    
    'Inputs
    With sht_interface.Cells(5, 2).Resize(1, 3)
        .Merge
        .Borders.Weight = xlThin
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
    End With
    
    With sht_interface
        .Cells(6, 2).Resize(5, 3).Borders.Weight = xlThin
        .Cells(6, 3).Resize(1, 2).Merge
        .Cells(7, 3).Resize(1, 2).Merge
        .Cells(8, 3).Resize(1, 2).Merge
        .Cells(9, 3).Resize(1, 2).Merge
        .Cells(10, 3).Resize(1, 2).Merge
    End With
    
    'Outputs
    With sht_interface.Cells(5, 6).Resize(1, 2)
        .Merge
        .Borders.Weight = xlThin
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
    End With
   
    With sht_interface
        .Cells(6, 6).Resize(3, 2).Borders.Weight = xlThin
        .Cells(10, 6).Resize(1, 2).Borders.Weight = xlThin
    End With

'Name the cells
    With sht_interface
        'Inputs
        .Range("C6").Name = "rng_interface_Company"
        .Range("C7").Name = "rng_interface_Coupon_Rate_Type"
        .Range("C8").Name = "rng_interface_Rate_Or_Margin"
        .Range("C9").Name = "rng_interface_Coupon_Frequency"
        .Range("C10").Name = "rng_interface_Maturity"
        .Range("C6:C10").Name = "rng_inputs"
        
        'Outputs
        .Range("G6").Name = "rng_spread"
        .Range("G7").Name = "rng_price"
        .Range("G8").Name = "rng_duration"
        .Range("G10").Name = "rng_hypertext_link"
    End With
    
'Dropdown menus
    'Query getting the names of the bond issuers whose records contain all the data necessary to interpolate their spread
        'ADODB objects initialization
        Set obj_Rst = New ADODB.Recordset
    
        'Query counting the number of companies with necessary data
        str_SQLRequest = "SELECT Count(*) FROM CDX_IG_Prices " & _
            "WHERE CDX_IG_Prices.[0,5] IS NOT NULL " & _
            "OR CDX_IG_Prices.[1] IS NOT NULL " & _
            "OR CDX_IG_Prices.[2] IS NOT NULL " & _
            "OR CDX_IG_Prices.[3] IS NOT NULL " & _
            "OR CDX_IG_Prices.[4] IS NOT NULL " & _
            "OR CDX_IG_Prices.[5] IS NOT NULL " & _
            "OR CDX_IG_Prices.[7] IS NOT NULL " & _
            "OR CDX_IG_Prices.[10] IS NOT NULL "

        'Send Query
        Call obj_Rst.Open(str_SQLRequest, obj_Cnn)
        
        'Store the total in an integer
        i = obj_Rst.Fields(0).Value
        
        'ADODB objects initialization
        Set obj_Rst = New ADODB.Recordset
        
        'Query getting the names of the companies whose records contain necessary data
        str_SQLRequest = "SELECT CDX_IG_Prices.Name FROM CDX_IG_Prices " & _
            "WHERE CDX_IG_Prices.[0,5] IS NOT NULL " & _
            "OR CDX_IG_Prices.[1] IS NOT NULL " & _
            "OR CDX_IG_Prices.[2] IS NOT NULL " & _
            "OR CDX_IG_Prices.[3] IS NOT NULL " & _
            "OR CDX_IG_Prices.[4] IS NOT NULL " & _
            "OR CDX_IG_Prices.[5] IS NOT NULL " & _
            "OR CDX_IG_Prices.[7] IS NOT NULL " & _
            "OR CDX_IG_Prices.[10] IS NOT NULL "
    
        'Send Query
        Call obj_Rst.Open(str_SQLRequest, obj_Cnn)
    
        'Store data in an array, resized based on the total previoulsy stored in the variable i
        ReDim arr_Companies(1 To i)
        j = 1
        'Loop on records
        Do While obj_Rst.EOF = False
            arr_Companies(j) = obj_Rst.Fields(0).Value
            'Move to next record
            obj_Rst.MoveNext
            j = j + 1
        Loop
           
    'Create the dropdown menus
        sht_interface.Range("rng_interface_Company").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=Join(arr_Companies, ",")
        sht_interface.Range("rng_interface_Coupon_Rate_Type").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Fixed, Variable"
        sht_interface.Range("rng_interface_Coupon_Frequency").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Yearly, Bi-annually, Quarterly"

End Sub
Sub sub_Bond_Pricer()

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'Bond pricer:
    'Instantiates the curve class module and the bond class module
    'Calculates the spread to maturity, the price and the duration for the bond selected in Interface
    'Displays the results in the output section


'       Table of Contents:

'          1. Link between Excel & Access database
'                   • Initializing and opening the connection

'          2. Interpolation of the spread to maturity
'               A. Creation of a Spread object
'                   • Query getting the spreads, for each company name, and storing the data in arrays
'                   • Fill a dictionary with the company names (=key) and their spread (=item)
'                   • Retrieve from dictionary the spread array of company selected
'                   • Use of let properties
'          3. Creation of a risk-free curve object
'                   • Query getting the risk free maturities and rates
'                   • Use of let properties
'          4. Creation of a LIBOR curve object
'                   • Query getting the LIBOR 3M maturities and rates
'                   • Use of let properties
'          5. Creation of a bond object
'                   • Use of let properties
'          6. Computation of the price, duration and schedule
'                   • Use of the price, duration and schedule methods
'          7. Page Layout


'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'Variables declaration
Dim obj_Cnn As ADODB.Connection
Dim obj_Rst As ADODB.Recordset
Dim dic_Spreads As Dictionary
Dim str_SQLRequest As String
Dim i As Integer, j As Integer, k As Integer
Dim arr_Spread_Maturities_Fields() As Double, arr_Spread_Rates_Records() As Double, arr_spread() As Double, _
    arr_US_Yield_Maturities() As Double, arr_US_Yield_Rates() As Double, arr_LIBOR_Maturities() As Double, _
    arr_LIBOR_Rates() As Double
Dim arr_Companies() As String
Dim var_Field As Variant, var_spreads As Variant

'Declaration and static instantiation
Dim obj_SpreadCurve As New cMod_Curve
Dim obj_Bond As New cMod_Bond

'Create the dictionary
Set dic_Spreads = New Dictionary 'Needs Microsoft Scripting Runtime reference to work

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'Link between Excel & Access database
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'Initializing the connection
    Set obj_Cnn = New ADODB.Connection

'Connection parameter
    obj_Cnn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "Data Source=C:\Users\zouli\OneDrive\Documents\Data_Projet.accdb;"
    
'Opening the connection
    obj_Cnn.Open

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'Interpolation of the spread to maturity
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    'Creation of a Spread object
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

    'Query getting the spreads, for each company name
        'ADODB objects initialization
            Set obj_Rst = New ADODB.Recordset
        
        'Query counting the number of companies with necessary data
            str_SQLRequest = "SELECT Count(*) FROM CDX_IG_Prices " & _
                "WHERE CDX_IG_Prices.[0,5] IS NOT NULL " & _
                "OR CDX_IG_Prices.[1] IS NOT NULL " & _
                "OR CDX_IG_Prices.[2] IS NOT NULL " & _
                "OR CDX_IG_Prices.[3] IS NOT NULL " & _
                "OR CDX_IG_Prices.[4] IS NOT NULL " & _
                "OR CDX_IG_Prices.[5] IS NOT NULL " & _
                "OR CDX_IG_Prices.[7] IS NOT NULL " & _
                "OR CDX_IG_Prices.[10] IS NOT NULL "
    
        'Send Query
            Call obj_Rst.Open(str_SQLRequest, obj_Cnn)
            
        'Store the total in an integer
            i = obj_Rst.Fields(0).Value
            
        'ADODB objects initialization
            Set obj_Rst = New ADODB.Recordset
            
        'Query getting the names of these companies and their spread
            str_SQLRequest = "SELECT CDX_IG_Prices.Name, " & _
                "CDX_IG_Prices.[0,5], " & _
                "CDX_IG_Prices.[1], " & _
                "CDX_IG_Prices.[2], " & _
                "CDX_IG_Prices.[3], " & _
                "CDX_IG_Prices.[4], " & _
                "CDX_IG_Prices.[5], " & _
                "CDX_IG_Prices.[7], " & _
                "CDX_IG_Prices.[10] " & _
                "FROM CDX_IG_Prices "
        
        'Send Query
            Call obj_Rst.Open(str_SQLRequest, obj_Cnn)
        
        'Store data in an array, resized based on the total previoulsy stored in the variable i
            ReDim arr_Companies(1 To i)
            ReDim arr_Spread_Maturities_Fields(1 To 8)
            ReDim arr_Spread_Rates_Records(1 To i, 1 To 8)
            j = 1
            'Loop on records
            Do While obj_Rst.EOF = False
                arr_Companies(j) = obj_Rst.Fields(0).Value
                k = 0
                'Loop on fields
                For Each var_Field In obj_Rst.Fields
                    If k <> 0 Then
                        arr_Spread_Maturities_Fields(k) = var_Field.Name
                        arr_Spread_Rates_Records(j, k) = var_Field.Value
                    End If
                    k = k + 1
                Next
                'Move to next record
                obj_Rst.MoveNext
                j = j + 1
            Loop
    
    'Fill the dictionary with the company names (=key) and their spread (=item)
        ReDim arr_spread(1 To UBound(arr_Spread_Maturities_Fields))
        'Loop on the recorded data, previously stored in arrays and extracting one record per loop, _
        this record becoming one item added to the dictionary
        For i = 1 To UBound(arr_Companies)
            For j = 1 To UBound(arr_Spread_Maturities_Fields)
                arr_spread(j) = arr_Spread_Rates_Records(i, j)
            Next j
            'Filling a spread dictionary
            dic_Spreads.Add Key:=arr_Companies(i), Item:=arr_spread
        Next i
    
    'Retrieve from dictionary the spread array of company selected
        var_spreads = dic_Spreads(sht_interface.Range("rng_interface_Company").Value)
        'Copy the variant values into an array
        For i = LBound(var_spreads) To UBound(var_spreads)
            arr_spread(i) = var_spreads(i) * 0.0001
        Next i
    
    'Use of let properties
        With obj_Spread
            .pName = sht_interface.Range("rng_interface_Company").Value & " spread"
            .pMaturity = arr_Spread_Maturities_Fields
            .pRate = arr_spread
        End With

'Display the interpolated spread to maturity in the output section
    sht_interface.Range("rng_spread").Value = obj_Spread.fn_Interpolate(sht_interface.Range("rng_interface_Maturity").Value)

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'Creation of a risk-free curve object
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'Query getting the risk free maturities and rates
    'ADODB objects initialization
        Set obj_Rst = New ADODB.Recordset
    
    'Query counting the number of maturities and rates in the US Yield Curve sheet
        str_SQLRequest = "SELECT Count(*) FROM [US Yield Curve]"

        'Send Query
        Call obj_Rst.Open(str_SQLRequest, obj_Cnn)
        
        'Store the total in an integer
        i = obj_Rst.Fields(0).Value
        
        'ADODB objects initialization
        Set obj_Rst = New ADODB.Recordset
        
        'Query getting the maturities and rates
        str_SQLRequest = "SELECT [US Yield Curve].Maturity, [US Yield Curve].Rates FROM [US Yield Curve]"

        'Send Query
        Call obj_Rst.Open(str_SQLRequest, obj_Cnn)
    
        'Store data in an array, resized based on the total previoulsy stored in the variable i
        ReDim arr_US_Yield_Maturities(1 To i)
        ReDim arr_US_Yield_Rates(1 To i)
        j = 1
        'Loop on records
        Do While obj_Rst.EOF = False
            arr_US_Yield_Maturities(j) = obj_Rst.Fields(0).Value
            arr_US_Yield_Rates(j) = obj_Rst.Fields(1).Value
            'Move to next record
            obj_Rst.MoveNext
            j = j + 1
        Loop

'Use of let properties
    With obj_RfCurve
        .pMaturity = arr_US_Yield_Maturities
        .pRate = arr_US_Yield_Rates
    End With

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'Creation of a LIBOR curve object
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    
'Query getting the LIBOR 3M maturities and rates
    'ADODB objects initialization
        Set obj_Rst = New ADODB.Recordset
    
    'Query counting the number of maturities and rates in the LIBOR 3M Curve sheet
        str_SQLRequest = "SELECT Count(*) FROM [LIBOR 3M Curve]"

        'Send Query
        Call obj_Rst.Open(str_SQLRequest, obj_Cnn)
        
        'Store the total in an integer
        i = obj_Rst.Fields(0).Value
        
        'ADODB objects initialization
        Set obj_Rst = New ADODB.Recordset
        
        'Query getting the maturities and rates
        str_SQLRequest = "SELECT [LIBOR 3M Curve].Maturity, [LIBOR 3M Curve].Rate FROM [LIBOR 3M Curve]"

        'Send Query
        Call obj_Rst.Open(str_SQLRequest, obj_Cnn)
    
        'Store data in an array, resized based on the total previoulsy stored in the variable i
        ReDim arr_LIBOR_Maturities(1 To i)
        ReDim arr_LIBOR_Rates(1 To i)
        j = 1
        'Loop on records
        Do While obj_Rst.EOF = False
            arr_LIBOR_Maturities(j) = obj_Rst.Fields(0).Value
            arr_LIBOR_Rates(j) = obj_Rst.Fields(1).Value
            'Move to next record
            obj_Rst.MoveNext
            j = j + 1
        Loop

'Use of let properties
    With obj_LIBOR
        .pMaturity = arr_LIBOR_Maturities
        .pRate = arr_LIBOR_Rates
    End With
    
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'Creation of a bond object
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'Use of let properties
    With obj_Bond
        .pCompany = sht_interface.Range("rng_interface_Company").Value
        .pCoupon_Rate_Type = sht_interface.Range("rng_interface_Coupon_Rate_Type").Value
        
        'Rate or Margin
        If sht_interface.Range("rng_interface_Coupon_Rate_Type").Value = "Fixed" Then
            .pCoupon_Rate_Or_Margin = sht_interface.Range("rng_interface_Rate_Or_Margin").Value
        ElseIf sht_interface.Range("rng_interface_Coupon_Rate_Type").Value = "Variable" Then
            .pCoupon_Rate_Or_Margin = sht_interface.Range("rng_interface_Rate_Or_Margin").Value * 0.0001
        End If
        
        'Coupon frequency
        If sht_interface.Range("rng_interface_Coupon_Frequency").Value = "Yearly" Then
            .pCoupon_Frequency = 1
        ElseIf sht_interface.Range("rng_interface_Coupon_Frequency").Value = "Bi-annually" Then
            .pCoupon_Frequency = 2
        ElseIf sht_interface.Range("rng_interface_Coupon_Frequency").Value = "Quarterly" Then
            .pCoupon_Frequency = 4
        End If
        
        .pMaturity = sht_interface.Range("rng_interface_Maturity").Value
    End With
    
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'Computation of the price, duration and schedule
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'Use of Price Method: display the price in the output section
    sht_interface.Range("rng_price").Value = obj_Bond.fn_Price

'Use of Duration Method: display the duration in the output section
    sht_interface.Range("rng_duration").Value = obj_Bond.fn_Duration()
    
'Use of Schedule Method
    obj_Bond.sub_Schedule


'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'Page layout
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    With sht_interface.Range("rng_inputs")
        .Columns.AutoFit
    End With

'hypertext link to the schedule sheet
    sht_interface.Hyperlinks.Add Anchor:=sht_interface.Range("rng_hypertext_link"), _
                                    Address:="", _
                                    SubAddress:="'" & sht_Schedule.Name & "'!A1", _
                                    TextToDisplay:="Go To"

End Sub

