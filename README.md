# Outlook2010CodeSamples
MAPI code samples for Microsoft Outlook 2010.

Documentation for these samples may be found on the [MSDN](http://msdn.microsoft.com/en-us/library/cc839588(office.14).aspx).

**Quick Details**
Date Published: 7/27/2009
Source Language: {"C++"}
Supported Development Platform: Microsoft Visual Studio 2008
Download Size: 186 KB (compressed) / 731 KB (expanded)
Dependencies: These projects are dependent on the [Outlook 2010 MAPI Header Files](http://www.microsoft.com/downloads/details.aspx?FamilyID=f8d01fc8-f7b5-4228-baa3-817488a66db1).

This download includes the following code samples:

**Sample Address Book Provider**
This code sample implements a basic address book provider that you can use as a starting point for further customization. It implements the required features of an address book provider, as well as more advanced features such as name resolution. 

The Outlook 2007 Auxiliary Reference documents a similar sample address book provider for Visual Studio 2005. For more information about that sample address book provider, see [About the Sample Address Book Provider](http://msdn.microsoft.com/en-us/library/bb821134.aspx).

**Sample Message Store Provider**
This code sample uses the PST provider as the back end for storing data. This wrapped PST store provider is intended to be used in conjuction with the Replication API. Most of the functions in this sample pass their arguments directly to the underlying PST provider. 

The Outlook 2007 Auxiliary Reference documents a similar sample message store provider for Visual Studio 2005. For more information about that sample wrapped PST store provider, see [About the Sample Wrapped PST Store Provider](http://msdn.microsoft.com/en-us/library/bb821132.aspx). For more information about the Replication API, see [Replication API](http://msdn.microsoft.com/en-us/library/cc160702.aspx).

**Sample Transport Provider**
This code sample implements a shared network file system transport provider. A transport provider is a DLL that acts as an intermediary between the MAPI subsystem and one or more underlying messaging systems. You can use the code in this sample as a starting point for building a more robust transport provider.

The Outlook 2007 Auxiliary Reference documents a similar sample transport provider for Visual Studio 2005. For more information about that sample transport provider, see [About the Sample Transport Provider](http://msdn.microsoft.com/en-us/library/bb820970.aspx).

**Previous Versions**
[Outlook 2007 Messaging API (MAPI) Code Samples](https://github.com/stephenegriffin/Outlook2007CodeSamples)
