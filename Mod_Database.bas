Attribute VB_Name = "Mod_Database"
Option Explicit
Sub sub_Database_Preparation()
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'Database modification (6M -> 0,5
'                       1Y -> 1
'                       ...)
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'Variables declaration
Dim obj_Cnn As ADODB.Connection
Dim strSQL As String

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
' Rename fields using SQL queries
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'6M
    strSQL = "ALTER TABLE CDX_IG_Prices ADD COLUMN [0,5] DOUBLE;"
    obj_Cnn.Execute strSQL

    strSQL = "UPDATE CDX_IG_Prices SET [0,5] = [6M];"
    obj_Cnn.Execute strSQL

    strSQL = "ALTER TABLE CDX_IG_Prices DROP COLUMN [6M];"
    obj_Cnn.Execute strSQL
    
'1Y
    strSQL = "ALTER TABLE CDX_IG_Prices ADD COLUMN [1] DOUBLE;"
    obj_Cnn.Execute strSQL

    strSQL = "UPDATE CDX_IG_Prices SET [1] = [1Y];"
    obj_Cnn.Execute strSQL

    strSQL = "ALTER TABLE CDX_IG_Prices DROP COLUMN [1Y];"
    obj_Cnn.Execute strSQL
    
'2Y
    strSQL = "ALTER TABLE CDX_IG_Prices ADD COLUMN [2] DOUBLE;"
    obj_Cnn.Execute strSQL
    
    strSQL = "UPDATE CDX_IG_Prices SET [2] = [2Y];"
    obj_Cnn.Execute strSQL
    
    strSQL = "ALTER TABLE CDX_IG_Prices DROP COLUMN [2Y];"
    obj_Cnn.Execute strSQL
    
'3Y
    strSQL = "ALTER TABLE CDX_IG_Prices ADD COLUMN [3] DOUBLE;"
    obj_Cnn.Execute strSQL
    
    strSQL = "UPDATE CDX_IG_Prices SET [3] = [3Y];"
    obj_Cnn.Execute strSQL
    
    strSQL = "ALTER TABLE CDX_IG_Prices DROP COLUMN [3Y];"
    obj_Cnn.Execute strSQL
    
'4Y
    strSQL = "ALTER TABLE CDX_IG_Prices ADD COLUMN [4] DOUBLE;"
    obj_Cnn.Execute strSQL
    
    strSQL = "UPDATE CDX_IG_Prices SET [4] = [4Y];"
    obj_Cnn.Execute strSQL
    
    strSQL = "ALTER TABLE CDX_IG_Prices DROP COLUMN [4Y];"
    obj_Cnn.Execute strSQL
    
'5Y
    strSQL = "ALTER TABLE CDX_IG_Prices ADD COLUMN [5] DOUBLE;"
    obj_Cnn.Execute strSQL
    
    strSQL = "UPDATE CDX_IG_Prices SET [5] = [5Y];"
    obj_Cnn.Execute strSQL
    
    strSQL = "ALTER TABLE CDX_IG_Prices DROP COLUMN [5Y];"
    obj_Cnn.Execute strSQL
    
'7Y
    strSQL = "ALTER TABLE CDX_IG_Prices ADD COLUMN [7] DOUBLE;"
    obj_Cnn.Execute strSQL
    
    strSQL = "UPDATE CDX_IG_Prices SET [7] = [7Y];"
    obj_Cnn.Execute strSQL
    
    strSQL = "ALTER TABLE CDX_IG_Prices DROP COLUMN [7Y];"
    obj_Cnn.Execute strSQL
    
'10Y
    strSQL = "ALTER TABLE CDX_IG_Prices ADD COLUMN [10] DOUBLE;"
    obj_Cnn.Execute strSQL
    
    strSQL = "UPDATE CDX_IG_Prices SET [10] = [10Y];"
    obj_Cnn.Execute strSQL
    
    strSQL = "ALTER TABLE CDX_IG_Prices DROP COLUMN [10Y];"
    obj_Cnn.Execute strSQL

End Sub

Sub sub_Insert_Results()

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'Inserts the price, the inputs and the date in a table, created for this purpose.

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'variables declaration
Dim obj_Cnn As ADODB.Connection
Dim str_SQLCreateB As String, str_SQLInsertB As String

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'Link between Excel & Access database
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
   
'Initializing the connection
    Set obj_Cnn = New ADODB.Connection

'Connection parameter
    obj_Cnn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "Data Source= C:\Users\zouli\OneDrive\Documents\Data_Projet.accdb"
    
'Opening the connection
    obj_Cnn.Open

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'Create a new table in the database
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'Creation request
    str_SQLCreateB = "CREATE TABLE Results (" & _
        "[Name] TEXT(255), " & _
        "[Coupon_Rate_Type] TEXT(255), " & _
        "[Coupon_Rate_Or_Margin] DOUBLE, " & _
        "[Coupon_Frequency] TEXT(255), " & _
        "[Maturity] DOUBLE, " & _
        "[Price] DOUBLE, " & _
        "[Pricing_Date] DATE, " & _
        "PRIMARY KEY([Name],[Coupon_Rate_Type],[Coupon_Rate_Or_Margin],[Coupon_Frequency],[Maturity],[Price],[Pricing_Date]))"
    
'Ignore the error if the table already exists
    'Disable the error management system
    On Error Resume Next
    
    'Execute the request
    obj_Cnn.Execute str_SQLCreateB
    
    'Check if there is an error
    If Err.Number <> 0 Then
        ' Check if the erreur is the result of an already existing table
        If Err.Number = -2147217900 Then
            'If the table already exists, ignore the error
            Err.Clear
        Else
            'Otherwise, inform user
            MsgBox "Erreur"
        End If
    End If
    
    'Enable the error management system
    On Error GoTo 0


'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'Insert date in the new table
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'Insert request and execution
    str_SQLInsertB = "INSERT INTO Results (Name, Coupon_Rate_Type, Coupon_Rate_Or_Margin, Coupon_Frequency, Maturity, Price, Pricing_Date) VALUES (" & _
        "'" & Replace(sht_interface.Range("rng_interface_Company").Value, "'", "''") & "', " & _
        "'" & Replace(sht_interface.Range("rng_interface_Coupon_Rate_Type").Value, "'", "''") & "', " & _
        Replace(sht_interface.Range("rng_interface_Rate_Or_Margin").Value, ",", ".") & ", " & _
        "'" & Replace(sht_interface.Range("rng_interface_Coupon_Frequency").Value, "'", "''") & "', " & _
        Replace(sht_interface.Range("rng_interface_Maturity").Value, ",", ".") & ", " & _
        Replace(sht_interface.Range("rng_price").Value, ",", ".") & ", " & _
        "#" & Format(Date, "yyyy-mm-dd") & "#)"
    
    obj_Cnn.Execute str_SQLInsertB

   
End Sub


