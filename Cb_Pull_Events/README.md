### Cb Pull Events - Pulls event data from the CB Response API and dumps to CSV. 

Will take the provided query and attempt to pull the following associated event data:
* Network
* Registry
* Module Load
* Child Process
* File Modification
* Cross Process

Example:
Cb_Pull_Events.vbs query

If providing multiple statements within the query you must quote the whole query. Remove drive letters from file paths.