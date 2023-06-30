# Purpose 
This is a collection of IronPython scripts to be used within the proprietary geospatial software "Trimble Business Center" aka TBC.

See https://geospatial.trimble.com/products-and-solutions/trimble-business-center for details.

You'll need at least the "Survey Advanced" license for TBC in order to execute those scripts.

Since executing and development requires a valid license of TBC this repository is provided as is. At the time that you read this I might not have access to such a license anymore and might not be able to react to improvement requests or bug reports.

# Compatibility
The country I'm living in, and TBC internally, use the metric system. I've added support for imperial units where necessary, but do a sanity check on computed values.

With TBC V5.90 the Ironpython version was changed from 2.7 to 3.4. That syntax change made it necessary to have two separate repositories.

# Installation
The collection of files found in this repository must be stored in the following folder.

until version V5.81
"C:\ProgramData\Trimble\MacroCommands\SCR Macros"
![image](https://github.com/RonnySchneider/SCR_Macros_Public/assets/112836384/d39fb9a5-4379-47c2-a4d0-5e8f18443f07)

from version V5.90 onwards
"C:\ProgramData\Trimble\MacroCommands3\SCR Macros"
![image](https://github.com/RonnySchneider/SCR_Macros_Public/assets/112836384/57d158e5-5a77-40fd-8bbe-3733218dafa7)


Each script needs to import several assemblies at startup.
In order to simplify the maintenance I have outsourced them into one single file "SCR_Imports.py" that must be located in the above mentioned folder.

Each script executes this hard coded imports file at startup.
```python:
import os
execfile("C:\ProgramData\Trimble\MacroCommands\SCR Macros\SCR_Imports.py")
```


"C:\ProgramData\" is usually a hidden folder, just copy and paste it into the address bar of the file explorer.
![image](https://user-images.githubusercontent.com/112836384/233819444-0538e4cb-5e86-4597-8c1f-4b777be79245.png)

Trimble occasionally changes the namespace of some assemblies in a new release of TBC. This is unfortunately not documented anywhere. As developer you'll find out when the import fails in a new version. In that case the imports file needs to be updated, and that's the reason for all those try/except calls in it.

# Developing Macros for TBC
You'll need a free TrimbleID to access the following pages.

See the getting started guide here

https://community.trimble.com/viewdocument/01-welcome-to-trimble-business-cent?CommunityKey=8a262af4-a35e-4e9a-9dd3-191cc785899a&tab=librarydocuments&LibraryFolderKey=639c1ad6-7791-419b-96b3-2330a7eadb72&DefaultView=folder

And the forum for people who may be able to help you out

https://community.trimble.com/communities/community-homepage?communitykey=8a262af4-a35e-4e9a-9dd3-191cc785899a
