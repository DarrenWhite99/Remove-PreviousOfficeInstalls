# **Remove previous versions of Office**

This PowerShell Script will remove the local MSI installations of Office 2013 and older. The Offscrub vbs scripts are used to remove the MSI installations of Office products.

### **Why This Project Exists**

This fork was created to develop a version that can be called and run dynamically, handling any necessary downloads on the fly.

### **Using the Offscrub scripts**

The Offscrub vbs scripts can be used to automate the removal of Office products. The scripts will uninstall the existing Office products regardless of the current health of the installation. The Remove-PreviousOfficeInstalls.ps1 script will determine which version of Office is currently installed and will call the appropriate Offscrub vbs script to remove the Office products installations.

The Offscrub vbs files included are:

* **OffScrub03.vbs** - Used to remove Office 2003 products.
* **OffScrub07.vbs** - Used to remove Office 2007 products.
* **OffScrub10.vbs** - Used to remove Office 2010 products.
* **OffScrub_O15msi.vbs** - Used to remove Office 2013 MSI products.
* **OffScrub_O16msi.vbs** - Used to remove Office 2016 MSI products.
* **OffScrubc2r.vbs** - Used to remove Click-To-Run products.

More information can be found at: https://blogs.technet.microsoft.com/odsupport/2011/04/08/how-to-obtain-and-use-offscrub-to-automate-the-uninstallation-of-office-products/

### Example

1. Open a PowerShell console.

	From the Run dialog type PowerShell 
		
2. Run this command to download and perform the uninstallation.

	Example: (new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/DarrenWhite99/Remove-PreviousOfficeInstalls/EnableRunFromGitHub-v1/Remove-PreviousOfficeInstalls.ps1') | iex

2. Run this command to perform the uninstallation. Steps 2 and 3 can be combined into a single command line.

	Remove-PreviousOfficeInstalls		

	The version of Office will be detected automatically and the appropriate Offscrub file will be used to remove any Office products. If Office is not detected on the client the script will notify the admin and stop running.
			

	

	

