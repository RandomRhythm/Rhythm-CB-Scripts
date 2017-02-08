#CB Feeds Dump - Pulls data from the CB Response feeds and dumps to CSV. 

###This script will export two CSV files for each feed/query. One CSV containing all data and a limited CSV containing unique entries.

You must edit the code of this script to adjust the query timeframe and host filter. The following section of code defines the query:


boolEchoInfo = False
IntDayStartQuery = "*" 'days to go back for start date of query. Set to "*" to query all binaries or set to -24 to query last 24 time measurement
IntDayEndQuery = "*" 'days to go back for end date of query. Set to * for no end date
strTimeMeasurement = "d" '"h" for hours "d" for days
strHostFilter = "" 'computer name to filter to. Typically uppercase and is case sensitive.


Script runs addtional queries to identify vulnerable and patched components. Currently supports the following:
* Flash Player
* MS15-065 KB3065822
* MS15-078 KB3079904 not applied
* MS08-067
* Silverlight MS16-006 CVE-2016-0034
* MS16-051 KB3155533
* Internet Explorer Major Version

additional queries can be run via aq.txt in the current directory.
name|query
Example:
knowndll|/api/v1/binary?q=observed_filename:known.dll&digsig_result:Unsigned