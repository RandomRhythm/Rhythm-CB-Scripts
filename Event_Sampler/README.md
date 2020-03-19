### Cb Event Sampler - Queries IOCs in Cb Response event data and provides a sampling CSV output.

The script will try to pull sample event data from the following categories:
* Network
* Registry
* Module Load
* Child Process
* File Modification
* Cross Process

Currently domains, IPs, and MD5 hashes are supported lookups. These values are provided in the 
For each IOC in the list/CSV input