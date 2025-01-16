GoogleSheetWrapper, as the name says, is a wrapper for the GoogleSheet library for .NET. The intention behind this project is to help the integration of the GoogleSheet library with your project.\
\
It has the basic methods to add values, create sheets, query data without the need of having to write it all yourself.\
\
To use it:\
1- Add the JSON file for your credentials that Google provides you as an Embedded Resource;\
2- Create a Model that inherits from the ISheet interface, put the name you want for the sheet in the property SheetName that the interface requires you to implement;\
3- Then to enable the usage of the Google services, use the static class of GoogleSheetService and use the following methods to properly:\
   3.1- GoogleSheetService.RegisterCredentialJsonFile("The full path and name of the embedded resource, not the path in your folder");\
   3.2- GoogleSheetService.RegisterSheet("The name of the application set within google settings", "The sheet ID you're going to use, that huge string part of the URL after the /d/");\
4- Call GoogleSheetService.Sheet to use the wrapper, the first time around it will do the authentication and let you use the methods available.
