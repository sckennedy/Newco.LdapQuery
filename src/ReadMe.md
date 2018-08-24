## LDAP Query Extractor

To run this, you need to have all of the CEC users ids in text file and then pass the full path of the text file as a parameter to the executable.
The second optional parameter is a full path to the folder where you want the output to be sent.

For example.
The usernames are in a file called usernames.txt at c:\Temp
Output should be sent to c:\Temp

To run the extract run the following from the command line whilst in the folder where this .exe file is
```
LdapQuery.exe C:\Temp\usernames.txt C:\Temp
```