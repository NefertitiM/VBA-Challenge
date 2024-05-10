Multiple Year Stock Data

The For Each loop function was used to extend the main code to all the sheets in the excel file.
  This involved defining the worksheet name so that all reference cells and ranges could be labeled and copied into the other sheets.

Column titles were printed using the Range function.


The first loop was created to fill out the summary table which consists of a ticker, quarterly change, total volume, and percent change.
  This involved defining the lastRow of the dataset using a row counting function (.End(xlUp).Row). 
  The i loop was used to determine when the ticker changed to a different ticker. At that point, the total volume for that ticker, along with the quarterly change, and percent change was calculated and printed in the summary table.

The next loop placed conditional formatting on the summary table.
  This was done by defining the last row of that table and creating a p loop. And then an If Else function was used to shade in positive values green and negative values red.

  The next loop was to identify the tickers with the maximum and minimum values for percent change and total volume. 
    A "Worksheet.Funtion" was used with either the "min" or "max" and applied to the proper range of values.
    Once those variables were defined and the function identified the maximum and minimum values, a J loop was used to identify the values in the summary table and print them out along with the cooresponding ticker it belonged to.
    Lastly, conditional formatting was used to make the values percentages.

  The last loop is similar to the one before it but for the maximum volume. A different loop was created for it since the range, or type of data is different from the previous loop.
  
