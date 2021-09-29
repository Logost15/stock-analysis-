# stock-analysis-
Overview of the Project
  In the beggining was to know if the stock DQ, was a good option of investment, after doing a meticulous search and reasearch we got te conclution, thta we need to review the other stocks.
  Because DQ wasn´t the best option of investment.

Results:
  We use tickers and loops, to determ the percentage of retun the stocks have, using this codes  
  
    If Cells(j, 1).Value = ticker Then
        
            totalVolume = totalVolume + Cells(j, 8).Value
            
            and  
              
             Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuos
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit  
            
  We get te conclution that the year 2017 it was much better that the nest year 2018, only to stocks got good retun in that year.            
  
Summary:
  1 what are the advantages or disadvantages of refactoring code?
    One of the main advantages I saw of refactoring the code is the data I got after going this, also it was more complete the information in the code.
  
  2 How do these pros and cons apply to refactoring the original VBA script?
    In the amount of loops we did, that gave us a more accurate data to make a decision, and know in wich cstock
  
  
