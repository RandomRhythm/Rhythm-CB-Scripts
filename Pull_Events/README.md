### Cb Pull Events - Pulls event data from the CB Response API and dumps to CSV. 

Will take the provided query and attempt to pull the following associated event data:
* Network
* Registry
* Module Load
* Child Process
* File Modification
* Cross Process

Example:
`Cb_Pull_Events.vbs query`

If providing multiple statements within the query you must quote the whole query. Query time frame can be restricted using last_update. 

Example:
`Cb_Pull_Events.vbs "sensor_id:123 AND last_update:-10080m"`

Optional arguments:
* `/a` argument to auto accept pulling down all results.
* `'/b` to baseline. Add letters after the "b" to tell it what to baseline: 
	* `/bmnc`  `"m"` - modules. `"n"` - network. `"c"` - cross process