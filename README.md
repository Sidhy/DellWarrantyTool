# DellWarrantyTool
 Automate Dell servicetag uploads


![image](https://user-images.githubusercontent.com/6968113/189205538-f7be6f03-55cf-403a-a6f4-4e9faccd28a7.png)

## Option 1: Set Session Cookie
How to use:
- Open your favorite browser and login to dell.com 
- Press F12 to open the development tools and open the network tab
- In the search bar enter: cookie-domain:www.dell.com
- From the request header section right click the cookie and copy value


## Option 2: Load assets from Dell
This will load all the assets from the Dell website into local memory

## Option 3: Upload assets to Dell
Upload new assets based on service tag to Dell.
When the assets are first loaded locally (option 2) this will check for any duplicates before uploading to Dell.

## Option 4: Save asset info to file
Save asset information (Service Tag, Device Model, Shipment Date, Warranty Status, Warranty End Date) to a file in csv format.

## Option 8: Clear all assets from Dell
This will remove all the assets from the dell portal, important this requires you first load all assets using option 2.


