This script uses the "SplitPipeline" module for multithreading.

This particular script gets all servers in an environment from Active Directory and filters out Citrix servers based on naming convention.

After we get the list of all servers we remotely check the registry for a specific application; in this case we check if a server has CrowdStrike Installed.
