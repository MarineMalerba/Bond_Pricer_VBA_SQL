# Bond_Pricer_VBA_SQL

## cMod_Bond ##
### Overview ###
The **cMod_Bond** module provides a comprehensive set of tools for calculating and analyzing bond prices, durations, and cash flow schedules within an Excel environment using VBA. The module is designed to handle both fixed and variable coupon bonds, leveraging external rate curves (risk-free and spread curves) and LIBOR rates for accurate financial computations.

### Features
+ Bond Properties:

  + 'pCompany': Represents the company or issuer of the bond.
  + pCoupon_Rate_Type: Indicates whether the bond has a "Fixed" or "Variable" coupon rate.
  + pCoupon_Rate_Or_Margin: The coupon rate for fixed-rate bonds or the margin over the LIBOR rate for variable-rate bonds.
  + pCoupon_Frequency: The number of coupon payments made annually.
  + pMaturity: The maturity of the bond in years.
+ Pricing Functionality:

  + fn_Price: Calculates the bond price based on its properties and interpolated risk-free and spread curves.
+ Duration Calculation:

  + fn_Duration: Computes the Macaulay duration of the bond, considering its coupon payments, spreads, and maturity.
+ Cash Flow Schedule Generation:

  + sub_Schedule: Generates and prints a detailed cash flow schedule on a dedicated worksheet, showing the bond's maturity, cash flows, risk-free rates, spreads, discount factors, and discounted cash flows.
### Detailed Functionality
Bond Pricing (fn_Price):

This function calculates the bond price by summing the present value of all coupon payments and the repayment of principal at maturity.
Supports both fixed and variable coupon rates, with the latter depending on LIBOR rates.
Uses external risk-free (obj_RfCurve) and spread (obj_Spread) curves for discounting future cash flows.
Duration Calculation (fn_Duration):

The duration is calculated as a weighted average of the times until each payment is received, with the weights proportional to the present value of the cash flows.
Adjusts for the bond's spread and the discounting effects of coupon frequency.
Cash Flow Schedule (sub_Schedule):

The subroutine generates a comprehensive table showing each payment period's cash flow, the corresponding risk-free rate, the spread, the discount factor, and the discounted cash flow.
This output is formatted neatly on an Excel worksheet for easy review and further analysis.
