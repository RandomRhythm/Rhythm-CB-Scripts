# Rhythm-CB-Scripts
Collection of scripts for working with Carbon Black Cb Response API

*Note: If you are getting connection errors it is likely happening because the HTTP Windows API the scripts use by default doesn't support the TLS version configured on the Cb Response console. To work around this problem utilize the compiled executables with SocketTools:
1. Install SocketTools from this repo (InstallSocketTools.exe)
   
   a. Must be run with administrator rights. Will launch regsvr32 to register the ActiveX controls. 
   
2. Browse to the EXE folder for each script and edit the INI file
   
   a. Each executable has its own INI file
   
   b. By default the INI file in the EXE directories are configured with UseSocketTools=True
   
   c. Configuration changes can be made in the INI such as start and end time filters
   
3. Run the EXE

### Configuring the INI file
INI files are provided in each script or EXE directory. The settings in the INI files will overide the default settings in the script/executable. The INI is broken down into three sections:
##### [IntegerValues]
These values should be numeric. Only the StartTime and EndTime can be negative numbers. The StartTime and EndTime are asterix (*) by default which will pull all events. Time is evaluated as the current time so negative numbers are required to filter to events in the past. 
* SleepDelay - miliseconds to sleep between queries
* ReceiveTimeout - Time-out value in seconds
* PagesToPull - Number of pages to pull for each API call (large numbers for certain calls can cause Cb Response console to not return data and could indicate a performance issue)
* SizeLimit - Don't pull more than this number of events 
##### [StringValues]
These are string/text values.
* TimeMeasurement - StartTime and EndTime use this measurement. The following values can be used for the time interval:
    * yyyy	Year
    * q	Quarter
    * m	Month
    * y	Day of year
    * d	Day
    * w	Weekday
    * ww	Week
    * h	Hour
    * n	Minute
    * s	Second
* SensorID - The ID number of a sensor you wish to limit the query to
##### [BooleanValues]
These are boolean values (True or False) to turn on or off features of the script.
* UseSocketTools - Set to True to use SocketTools or False to not use SocketTools

Other Values exist and may be unique to the individual script. The above examples are provided as they are generally available for each INI file.

Problems or questions? email randomrhythm@rhythmengineering.com
