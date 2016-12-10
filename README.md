# DocumentDB-Inventory
Inventory Solution to get inventory to DocumentDB (hosted on Azure on on DocumentDB Emulator)

PowerShell to generate an Inventory-File (JSON):

`iex (iwr -URI "https://raw.githubusercontent.com/rzander/DocumentDB-Inventory/master/Inventory%20Agents/Windows-Devices/Create-Inventory.ps1").Content`

Upload the File to DocumentDB by using https://github.com/mingaliu/DocumentDBStudio
