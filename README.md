# VBA-challenge
> Overview
Analysis of generated stock market data using Visual Basic in Microsoft Excel

> Note: (Handwritten): For part of this data analysis, I made use of GPT-3.5 and GPT-4. Rather than copy-paste code from there, I opted for a more enriching experience through asking for vague and challenging hints as to what is wrong with my code when I get stuck, to get me to think. This way allowed me to fluently understand the structures in Visual Vasic for Applications. Along the way, I fed it parts of the updated Microsoft documentation (The knowledge cutoff for these Large-Language-Models is 2021) to utilize the tool as a "Second Pair of Eyes"

# File Guide

# Deconstructing the "Ask"

# Part 1.A
> Instructions (From University of Oregon's Data Analytics Program)
Script should loop through all stocks for one year and output the following:

- Ticker symbol
- Difference between opening and closing annual price
- Percentage difference between the annual opening and closing price
- The total stock volume of the stock

# Part 1.B
> Pseudocode (Handwritten)

- FOR-LOOP: 
    - **Will only increment counter after inner loop fully traverses a whole section of a single symbol. This would require unsorted datasets to be sorted.**
    - FOR-LOOP: 
        - **This loop will reset it's counter each time it fully    traverses a single trading symbol, such as AAB, AAF, and so on.**
        - Counter = Counter + 1
    - END INNER LOOP
- END OUTER LOOP