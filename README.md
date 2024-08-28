# Bond_Pricer_VBA_SQL

## cMod_Bond ##
### Overview ###
The **cMod_Bond** module provides a comprehensive set of tools for calculating and analyzing bond prices, durations, and cash flow schedules within an Excel environment using VBA. The module is designed to handle both fixed and variable coupon bonds, leveraging external rate curves (risk-free and spread curves) and LIBOR rates for accurate financial computations.

### Features
+ Bond Properties:

  + ``pCompany``: Represents the company or issuer of the bond.
  + ``pCoupon_Rate_Type``: Indicates whether the bond has a "Fixed" or "Variable" coupon rate.
  + ``pCoupon_Rate_Or_Margin``: The coupon rate for fixed-rate bonds or the margin over the LIBOR rate for variable-rate bonds.
  + ``pCoupon_Frequency``: The number of coupon payments made annually.
  + ``pMaturity``: The maturity of the bond in years.
+ Pricing Functionality:

  + ``fn_Price``: Calculates the bond price based on its properties and interpolated risk-free and spread curves.
+ Duration Calculation:

  + ``fn_Duration``: Computes the Macaulay duration of the bond, considering its coupon payments, spreads, and maturity.
+ Cash Flow Schedule Generation:

  + ``sub_Schedule``: Generates and prints a detailed cash flow schedule on a dedicated worksheet, showing the bond's maturity, cash flows, risk-free rates, spreads, discount factors, and discounted cash flows.

## cMod_Curve
### Overview
The **cMod_Curve** module is a utility designed to manage and perform operations on financial curves, such as yield curves, within an Excel VBA environment. It primarily handles the storage of curve data (maturities and rates) and provides functionality to interpolate rates for any given maturity using linear interpolation.

### Features
+ Curve Properties:

  + ``pName``: Represents the name of the curve (e.g., "Risk-Free Curve", "Spread Curve").
  + ``pType``: Describes the type of the curve (e.g., "Yield", "Discount").
  + ``pMaturity``: Holds an array of maturities corresponding to the curve.
  + ``pRate``: Stores an array of rates associated with each maturity.
+ Interpolation Functionality:

  + ``fn_Interpolate``: This method performs linear interpolation to estimate the rate corresponding to any maturity based on the given curve data.

## Mod_Interface_and_Instantiation
### Overview ###
This module contains procedures for initializing objects and resetting the interface for the bond pricer.
### Features
+ Global Declarations:
  + ``Public obj_RfCurve As New cMod_Curve``: Instantiates an object for the risk-free curve.
  + ``Public obj_Spread As New cMod_Curve``: Instantiates an object for the spread curve.
  + ``Public obj_LIBOR As New cMod_Curve``: Instantiates an object for the LIBOR curve.
+ Sub:
  + ``sub_Reset_Interface``: This subroutine resets the Excel interface for the bond pricer, including:
    + Clearing Data and Names: Clears previous data and range names from the sheet.
    + Printing Titles: Sets up titles for the input and output sections.
    + Page Layout: Adjusts the page layout, merges cells, and sets borders.
    + Naming Cells: Names important cells to be referenced later.
    + Dropdown Menus: Creates dropdown menus for selecting bond-related parameters by querying an Access database for available company data.
  + ``sub_Bond_Pricer``: This subroutine performs the bond pricing operations, including:
    + Database Connection: Opens a connection to the Access database.
    + Spread Interpolation:
      + Queries the database to get the spread data for different maturities.
      + Fills a dictionary with company names as keys and spread data as items.
      + Uses the retrieved spread data to calculate the spread to maturity.
    + Creating Curve Objects:
      + Creates and initializes objects for the spread, risk-free, and LIBOR curves using data from the database.
    + Creating Bond Object:
      + Initializes the bond object with data from the interface.
    + Computing Price, Duration, and Schedule:
      + Uses the bond object to calculate and display the bond price and duration.
      + Creates a hyperlink to the payment schedule sheet.
### Key Concepts in the Code:
+ **Object-Oriented VBA**: The code uses custom classes (cMod_Curve, cMod_Bond) to handle financial data and operations, showcasing object-oriented programming in VBA.
+ **ADODB for Database Connection**: Uses ADODB.Connection and ADODB.Recordset to interact with the Access database, allowing dynamic data retrieval and processing.
+ **Dynamic Interface Managemen**t: The code dynamically creates and manages an Excel-based interface, making it user-friendly and adaptable.
Important Considerations:
+ **Dependencies**: The code relies on the Microsoft ActiveX Data Objects Library (ADODB) for database operations and the Microsoft Scripting Runtime for the dictionary object. This code effectively integrates data from an external database into an Excel interface to perform bond pricing, demonstrating a powerful application of VBA for financial analysis.

## Mod_Database

### Overview
This VBA code is designed to manage and manipulate data within an Access database.
### Features
1. sub_Database_Preparation
This subroutine is used to modify the schema of the CDX_IG_Prices table in the Access database. It performs the following actions:

Establishes a Connection: Connects to the Access database using the ADODB library.
Modifies Columns:
Adds new columns with updated names (e.g., changing 6M to 0,5).
Copies data from the old columns to the new columns.
Drops the old columns.

2. sub_Insert_Results
This subroutine inserts new data into a table named Results. It includes the following steps:

Establishes a Connection: Connects to the Access database using the ADODB library.
Creates Table: Checks if the Results table exists, and if not, creates it with the following columns:
Name (TEXT)
Coupon_Rate_Type (TEXT)
Coupon_Rate_Or_Margin (DOUBLE)
Coupon_Frequency (TEXT)
Maturity (DOUBLE)
Price (DOUBLE)
Pricing_Date (DATE)
Inserts Data: Adds a new row of data into the Results table based on the values provided in specific Excel ranges.
Connection String
Both subroutines use the following connection string to link Excel with the Access database:

vba
Copier le code
Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\zouli\OneDrive\Documents\Data_Projet.accdb
Error Handling
sub_Database_Preparation: Includes basic error handling for SQL execution errors but does not check for specific conditions.
sub_Insert_Results: Includes error handling for the table creation step to ignore errors if the table already exists.
Usage
