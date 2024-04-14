Attribute VB_Name = "mdl_z_ToDo"
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' IDEAS / TO DOs


''' ISSUE:
'   worksheets are sometimes organized with a table dimension going horizontally (often a date)
'   to easily analyze this data, it is best to have this data organized in a database format
'   Example:

''  Bad Format:

'           Date ->
' Id        2010    2011    2012    2013    2014
' Record1
' Record2
' Record3

''  Good Format:

' Id      Date
' Record1 2010
' Record1 2011
' Record1 2012
' Record1 2013
' Record1 2014
' Record2 2010
' Record2 2011
' Record2 2012
' Record2 2013
' Record2 2014
' Record3 2010
' Record3 2011
' Record3 2012
' Record3 2013
' Record3 2014

''' SOLUTION:

' I have done this process manually several times, by numbering rows and columns and using lookup functions
' Need macro
'   somehow transpose lateral data
'   input would just be the range?
