# VBA-challenge

The VBA code that is saved as a separate file in this repo and is stored in each spreadsheet has the main purpose to create a summary table.

This macro will loop through all the stocks for one year and output the following information on separate columns of each sheet. The script
will output the:

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

The macro will also highlight positive percentage gains in green and negative percentage gains in red.

The macro will also create a second summary table on the first sheet that outputs:

  * Greatest % Increase and associated ticker from all sheets
  
  * Greatest % Decrease and associated ticker from all sheets
  
  * Greatest Total Volume and associated ticker from all sheets