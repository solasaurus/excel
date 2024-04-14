Attribute VB_Name = "mdl_2_FunctionExamples"
Option Explicit


''''''''''''''''''''''''''''''''''''''''''
'''     META SHEET: Dynamic Defined Names, Lookup Lists, etc.


''' Dynamic Defined Name
'       Placed in the "Refers To" field for the Defined Name / List
'   =OFFSET(META!$B$12,0,0,(ROWS(META!$B$12:$B$29)-COUNTBLANK(META!$B$12:$B$29)),1)
'   Where:
'       B12 is the first record in the list (not the header of the list)
'       B29 is the last record in the list


''' Lookup List
'       Retrieves a unique list of records from a column.
'       Combined with a Dynamic Defined Name, this allows a new inputs into a column to be automatically added to data validation lists.
'   =IFERROR(INDEX(tbl_INVENTORY[MCF],MATCH(0,INDEX(COUNTIF($C$20:C20,tbl_INVENTORY[MCF]),0,0),0)),"")
'   Where:
'       tbl[col] is the list to pull from
'       C20 is header cell for the list
'   Notes:
'       If using to populate a data validation list, be sure to uncheck "error alert" option so that new values can be added to list manually


'   Lookup List based on criteria
'       This method is used to create cascading dropdowns. See AssetMgmt workbook for example
'   {=IFERROR(INDEX(Investments[Investment Name],MATCH(0,COUNTIF(H$138:H145,Investments[Investment Name])+(Investments[State]<>H$138),0),COLUMN($A$1)),"")}
'   Where:
'       tbl[col] is the list to pull from
'       H138 is header cell for the list / criteria
'       A1 is _____. I forgot what this does, I thought it was a counter but then why would it be an absolute reference?

'   Lookup List, removes blanks
'=IFERROR(INDEX(ACQPIPELINE[JV Partner],MATCH(0,IF(ISBLANK(ACQPIPELINE[JV Partner]),1,COUNTIF($I$81:I81,ACQPIPELINE[JV Partner])),0)),"")

''''''''''''''''''''''''''''''''''''''''''
'''     LOOKUPS


''' Lookup based on multiple criteria

'   =SUMPRODUCT(--(B7:B30="Midwest"),--(C7:C30="Masks"),D7:D30)
'   Where:
'       B7:B30/C7:C30 is the criteria range
'       "Midwest"/"Masks" is criteria
'       D7:D30 is the range to bring back

''''''''''''''''''''''''''''''''''''''''''
'''     COUNT


''' COUNT UNIQUE TEXT VALUES WITH CRITERIA
'   =SUM(--(FREQUENCY(IF(C5:C11=G5,MATCH(B5:B11,B5:B11,0)),ROW(B5:B11)-ROW(B5)+1)>0))
'   Where:
'       B5:B11 is the Range being counted
'       C5:C11 is the Criteria Range
'       B5 is the first cell in the Range being counted
'       G5 is the Criteria
'       * Must be entered as array

''' COUNT UNIQUE TEXT VALUES WITH CRITERIA (can handle blanks)
'   {=SUM(--(FREQUENCY(IF(B5:B11<>"",IF(C5:C11=G5,MATCH(B5:B11,B5:B11,0))),ROW(B5:B11)-ROW(B5)+1)>0))}

''' COUNT UNIQUE TEXT VALUES WITH MULTIPLE CRITERIA
'   =SUM(--(FREQUENCY(IF(c1,IF(c2,MATCH(vals,vals,0))),ROW(vals)-ROW(vals.1st)+1)>0))
'   =SUM(--(FREQUENCY(IF(C5:C11=G6,IF(C5:C11=G5,MATCH(B5:B11,B5:B11,0))),ROW(B5:B11)-ROW(B5)+1)>0))
'   Where:
'       G6 is the 2nd Criteria

''' CHECK IF DUPLICATE EXISTS IN COLUMN
'   =IF(COUNTIF([UNIQUE STRING],[@[UNIQUE STRING]])>1,1,"")
'   Where:
'       "UNIQUE STRING" is the column with the unique id


''''''''''''''''''''''''''''''''''''''''''
'''     FIND


'   FIND the first instance of a 4 digit number in a string
'       Used to extract a year from a string
'   =1*MID(A1,MIN(FIND({0,1,2,3,4,5,6,7,8,9},A1&"0123456789")),4)
'   Where:
'       A1 is string


''''''''''''''''''''''''''''''''''''''''''
'''     DATE


''' Find first of month given date
'   =MIN(Loans[Origination Date])-DAY(MIN(Loans[Origination Date]))+1

''' Find end of month given date
'   =EOMONTH(MIN(Loans[Origination Date]),0)

''' Find end of quarter given date
'   =EOMONTH(E$280,MOD(12-MONTH(E$280),3))

''' Find end of year given date
'   =DATE(YEAR(E$280),12,31)

