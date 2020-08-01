# Rhythm-CB-Scripts
#### Collection of scripts for working with Carbon Black Cb Response API

This repository contains a folder for each script's purpose.

#### Alerts
The Cb_Alerts script will export alerts from the console in CSV files for each of the feeds and watchlists. The Alerts folder also contains the Cb_Resolve script to resolve alerts within the console.

#### Pull_Events
Process activity generates events, which can be child processes, registry, file, network, or cross-process activity. I call this the API trace, but some may call it the process interactions. The script takes a query and runs it against the API to then output CSV files for each event category. 

#### Feeds_Dump
The Feed_Dump script will output CSV files for each feed or watchlist configured in the console. This script is useful for reviewing feeds and watchlists that are not generating alerts. 

#### Sensor_Dump
This script outputs a CSV file containing each sensor and its associated data. 

#### Hash_Dump
The Hash_Dump script will dump hash values and associated data. Dump all executables, DLL files, or provide a list of hash values to get the associated binary's information. I use this feature to run hashes against hash lookup services, such as VirusTotal, using VTTL.

#### File_Download
Cb Response will provide available files to download the files within zip files. The File_Download script will download the zip files for the provided hash values. 

#### extract_CB_zips
The extract_CB_zips script will utilize 7z to extract File_Download zip files. Files are extracted and renamed to the value of the MD5.

#### Event_Sampler
The Event_Sampler is a branch of the Pull_Events script. Instead of outputting CSV files for the various event categories, the script will output a sampling from each event category into one CSV file.

#### SocketTools
SocketTools requires that it only be used in compiled code. However, compiling VBScript causes many antimalware vendors to detect the resulting executable file. The antimalware detections were causing problems with downloading this repo and thus were removed. If you would like compiled versions, please let me know as currently, that doesn't appear to be a problem these days.  


### Configuring the INI file
INI files are provided in each script directory. The settings in the INI files will override the default settings in the script/executable. The INI is broken down into three sections:
##### [IntegerValues]
These values should be numeric. Only the StartTime and EndTime can be negative numbers. The StartTime and EndTime are asterisks (*) by default, which will pull all events. Time is evaluated at the current time, so negative numbers are required to filter to events in the past. 
* SleepDelay - milliseconds to sleep between queries
* ReceiveTimeout - Time-out value in seconds
* PagesToPull - Number of pages to pull for each API call (large numbers for certain calls can cause Cb Response console not to return data and could indicate a performance issue)
* SizeLimit - Don't pull more than this number of events 
##### [StringValues]
These are string/text values.
* TimeMeasurement - StartTime and EndTime use this measurement. The following values can be used for the time interval:
    * yyyy    Year
    * q    Quarter
    * m    Month
    * y    Day of the year
    * d    Day
    * w    Weekday
    * ww    Week
    * h    Hour
    * n    Minute
    * s    Second
* SensorID - The ID number of a sensor you wish to limit the query to
##### [BooleanValues]
These are boolean values (True or False) to turn on or off features of the script.
* UseSocketTools - Set to True to use SocketTools or False to not use SocketTools

Other Values exist and may be unique to the individual script. The above examples are provided as they are generally available for each INI file.


### Troubleshooting

If you get the message "error on line 1" it is likely due to the file being saved in Unicode. Open up the script in notepad.exe and click File > Save As. In the save as dialog, change the encoding at the bottom of the screen to ANSI.

If you are getting connection errors, it is likely happening because the HTTP Windows API the scripts uses by default doesn't support the TLS version configured on the Cb Response console. To work around the issue, use a modern version of Windows such as Windows 10 or Server 2016. To fix this problem in Windows perform the fix Microsoft describes here:

https://support.microsoft.com/en-us/help/3140245/update-to-enable-tls-1-1-and-tls-1-2-as-default-secure-protocols-in-wi

Another option to work around this problem is to utilize SocketTools. Executables were provided for each script to use SocketTools instead of the Windows API. However, the executables were detected as malware by several vendors, which caused problems with downloading this repo. If you require this workaround, please make a request to have the executables published.

