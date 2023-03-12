# Auto-summary


## Goal

This project is to auto-summary the key information from the excel file. The key information is the 6th and 8th sheet of the .xlsx. 

## Steps

1. Extract key information

    1.1. 6th table: the status gap and the sum
   
    1.2. 8th table: change the content to Agreed; if there is no Agreed, change it to None. Then, extract the link gap and the sum

2. Do mathematical calculation

    Both not gap = sum - gap

3. Draw the bar chart

    If the input has comparable folder: indicate the percent of the gap



## Some guidelines/references

Key reference: [Pivottable extraction and modification](https://towardsdatascience.com/automate-excel-with-python-pivot-table-899eab993966)


pywin32: modify the Pivot Table Fields

Format of pivot table: https://miro.medium.com/max/640/1*YSW2nlePQAxSZiv-Q2hQTw.webp

## Comments


I have wrote a lot of comments in the code. You can understand it fully by reading the code.
