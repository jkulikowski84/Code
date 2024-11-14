To Download MS Store apps manually, you can visit: https://store.rg-adguard.net/

To get the productID info, navigate to the MS Store online here: https://www.microsoft.com/en-us/store/apps/windows
1. Search the app you want and once you find it in the results, click on it to go to that app link.
2. If the app takes you to the link: "https://apps.microsoft.com/detail/9wzdncrfhvn5?hl=en-us&gl=US" - The productID would be "9wzdncrfhvn5"

Now navigate to https://store.rg-adguard.net/
1. Change the drowdown to ProductID and in the textbar search "9wzdncrfhvn5" (you can leave RP as the default)
2. Find the app package you want, but make sure the extension is ".appxbundle" if it's ".eappxbundle" this process won't work. Also if it's a msixbundle, this won't work.

After you download the file, you need to allow developer mode on the machine.

Do a search for Settings, then navigate to "Update & Security" --> "For Developers"
Enable Developer Mode

Now you can install the application.
1. Open up powershell as admin
2. run Add-AppxPackage -Path (path to appx package)

Once it finishes, you should have the application.

