## Creating Excel Functions that will call Google Maps API

There are two seperate functions that are being created:
1)  Get_Distance: Get travelling distance in km between 2 addresses using Google Maps
2)  Get_Duration: Get travelling duration in minutes between 2 addresses using Google Maps

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. 

Prerequisites:
Requires a reference to Microsoft XML, v6.0, Scripting Library, and Scripting Runtime
Function: https://gist.github.com/cwg999/51bafc6cc5f28308ca219e0b43b1aff2#file-encodeuricomponent-vb
Parse Json: https://github.com/VBA-tools/VBA-JSON
Install both Parse Json and the EncodeUri function in VBA

Function Documentation:
Function 1: Get_Distance(Origin As String, Destination As String)
Function 2: Get_Duration(Origin As String, Destination As String) As Double

Origin and Destination are respective cells that contain the addresses you want to use.

Examples:
Mar-a-Lago: 1100 S Ocean Blvd, Palm Beach, FL 33480, United States
Whitehouse: 1600 Pennsylvania Avenue NW Washington, DC 20500, United States

Get_Distance("1100 S Ocean Blvd, Palm Beach, FL 33480, United States","1600 Pennsylvania Avenue NW Washington, DC 20500, United States")

Get_Duration(""1100 S Ocean Blvd, Palm Beach, FL 33480, United States","1600 Pennsylvania Avenue NW Washington, DC 20500, United States")

## License

This project is licensed under the MIT License - see the LICENSE.md file for details

## Acknowledgments
  
Google API: https://developers.google.com/maps/documentation/distance-matrix/intro#traffic-model
Json Parsing Google Maps: https://stackoverflow.com/questions/36020363/google-api-distancematrix-returning-wrong-json-result-in-vba
EncodeUri Function https://gist.github.com/cwg999/51bafc6cc5f28308ca219e0b43b1aff2#file-encodeuricomponent-vb
Parsing Json: https://github.com/VBA-tools/VBA-JSON
