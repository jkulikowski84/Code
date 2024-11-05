If you use the "VMware.PowerCLI" module in powershell, you will know that it takes a REALLY long time for powershell to load.
Usually in order to perform most of the tasks needed all you really need is the "VMware.VimAutomation.Core" module. 
This module of course also has some prerequisites that it needs, but it's a lot less that the whole powerCLI suite of modules.

This script allows you to type in the module we are interested in (for example) "VMware.VimAutomation.Core", and then recursively checks that modules required modules and goes down each module level until it has everything it needs. The end result is instead of having to load over 80 modules within PowerCLI, instead you only load 4. This will have a dramatic performance improvement with no impact to your script.
