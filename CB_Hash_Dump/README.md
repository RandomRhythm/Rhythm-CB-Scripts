#CB Hash Dump - Dumps hashes from CB (Carbon Black) Response
###This script will export a CSV of binary files matching the specified query in CB Response (Carbon Black).

You must edit the code of this script to adjust the query. The following section of code defines the query:

'---Config Section
BoolDebugTrace = False 'Leave this to false unless asked to collect debug logs.
IntDayStartQuery = "*" 'time to go back for start date of query. Set to "*" to query all binaries. Set to "-7" for the last week.
strTimeMeasurement = "d" '"h" for hours "d" for days
IntDayEndQuery = "-1" 'days to go back for end date of query. Set to "*" for no end date. Set to "-1" to stop at yesterday.
strBoolIs_Executable = "True" 'set to "true" to query executables. Set to "false" to query resources (DLLs).
BoolExcludeSRSTRust = True 'Exclude trusted applications from the query
strHostFilter = "" 'computer name to filter to. Use uppercase, is case sensitive 
boolOutputHosts = True ' Set to True to output hostnames for each binary
'---End Config section