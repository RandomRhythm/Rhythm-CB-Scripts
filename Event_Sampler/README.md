### Cb Event Sampler - Queries IOCs in Cb Response event data and provides a sampling CSV output.

The Cb Event Sampler script takes a list of IOCs and queries them via the Cb Response API. A CSV row is output for each IOC. The input list can either be a CSV where you specify which column contains the IOCs to lookup or a list of IOCs.

The script will try to pull sample event data from the following categories:
* Network
* Registry
* Module Load
* Child Process
* File Modification
* Cross Process

Domains, IPs, and MD5 hashes are the currently supported lookups. CSV processing can only query values from one column. API CSV output will combine the original CSV input rows. 

This script is useful for querying any items where more context is needed. For example, take low prevalent files from CB_Hash_Dump output and look those up to get context sampling of process activity. 