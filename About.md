# What do I actually want for this project?

- To use the current excel template as the base.
- From FX sheet, read in exchange rates and currency to quote. Read in other project information as well.
- From the system sheets, read in internal escalation and mark up as well as the required data
- From the engineering sheet, only read in the required data
- Additional fields are appended at the end

## Excel columns

Be libral with the excel columns. Use as many as required.

1. NO : Number
2. SN : Serial number
3. Qty : Quantity
4. Description
5. Unit : Unit (For quantity)
6. Unit Price : Unit Selling Price
7. Subtotal Price : Subtotal Selling Price
8. Scope : Included or Optional
9. Cur : Currency
10. UC : Uni Cost in original currency
11. SC : Subtotoal Cost in original currency
12. Discount : Discount in percentage
13. Remark : Derived cost should be explained here
14. UCD : Unite Cost after Discount in original currency
15. SCD : Subtotal Cost after Discount in original currency
16. FUP : Fixed Unit Price. If this price exists, the calculation will use this price as the unit selling price. The price must be in quoted currency.
17. Rate : Exchange rates
18. UCDQ : Unit Cost after Discount in Quoted Currency
19. SCDQ : Subtotal Cost after discount in Quoted Currency
20. BUCQ : Based Unit Cost in Quoted Currency. This cost will be inclusive of all escalation. Thinkin of doing this in formula.
21. BSCQ : Based Subtotal Cost in Quoted Currency. This cost will be inclusive of all escalation.
22. Default : Default operation cost normally considered to be 3%. Will be calculated based on subtotal cost.
23. Warranty : Default warranty cost normally considered to be 3%. Will be calculated based on subtotal cost.
24. Freight : Inbound freight cost ranges from 3% to 5%. Will be calculated based on subtotal cost.
25. Special Terms : To put the escalation in % for identified risk such as unfavourable terms. Will be calculated based on subtotal cost.
26. Risk : Inherent risk in project 5%. Will be calculated based on subtotal cost.
27. RUPQ : Recommended Unit Price in Quoted Currency
28. RSPQ : Recommended Subtotal Price in Quoted Currency
29. UP : Unit Price
30. SP : Subtotal Price
31. UPLS : Unit Price Lump Sum
32. SPLS : Subtotal Price Lump Sum
33. Profit
34. Margin
35. Supplier
36. Maker
37. Model
38. Leadtime
39. Format : Format for conditional formatting
40. System : System Name
41. Category : Category Name such as Product or Service
