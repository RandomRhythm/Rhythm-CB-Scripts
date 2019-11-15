# CB Feeds Dump

### This script will export two CSV files for each feed/watchlist/query. One CSV containing all data and a limited CSV containing unique entries.

Use the Cb_Feeds.ini to adjust the query timeframe and host filter. The following snippet identifies key values:

	'---INI snippet
	[IntegerValues]
	StartTime=* 'Number of time to go back for start date of query. Set to "*" to query all or set to -24 to query last 24 time measurement
	EndTime=* 'days to go back for end date of query. Set to * for no end date
	[StringValues]
	TimeMeasurement=d '"h" for hours "d" for days
	'---End INI snippet

Script runs addtional queries to identify vulnerable and patched components. Currently supports the following checks:
* Flash Player
* MS15-065 KB3065822
* MS15-078 KB3079904 not applied
* MS08-067
* Silverlight MS16-006 CVE-2016-0034
* MS16-051 KB3155533
* Internet Explorer Major Version
* MS17-010
* BlueKeep
* DejaBlue

Additional queries can be run via aq.txt in the current directory. Input format is name|query where the name will be used as the file name for CSV output and the query will be used to pull down the results.

Example:
  
	knowndll|observed_filename:known.dll&digsig_result:Unsigned
	evasion_installutil|process_name:installutil.exe AND parent_name:cmd.exe

To force a query to binary or process include "/api/v1/%type%?q=" before the query:

	knowndll|/api/v1/binary?q=observed_filename:known.dll&digsig_result:Unsigned
	evasion_installutil|/api/v1/process?q=process_name:installutil.exe AND parent_name:cmd.exe
