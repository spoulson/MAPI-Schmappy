MAPI-Schmappy
=============

Release 2008-10-21  
Github repo: https://github.com/spoulson/MAPI-Schmappy  
Shawn Poulson, http://explodingcoder.com

MAPI command line tools to set Outlook address list search path and default address list.

As explained in my blog post: http://explodingcoder.com/blog/content/programmatically-updating-outlooks-address-book-options-with-a-command-line-toolAll

INTRODUCTION
------------
### What is MAPI Schmappy?
 - It's a name for a couple command line tools for Outlook written to fulfill a need and I couldn't think of a name.
 - Included are two command line tools:
    - **MAPIAddrListSearch**: Sets the Outlook address list search list for name resolution when sending emails.
    - **MAPIDefaultAddressList**: Sets the default address list when opening the address book.

SYSTEM REQUIREMENTS
-------------------
MAPI Schmappy tools were tested under the following environment:

### Server
 - Exchange Server 2003

### Client
 - Outlook 2000
 - Windows XP SP3

BUILD
-----
 - Using Visual Studio 2010, 2008, or 2005, open the corresponding solution file in the repository root. (e.g. MAPI Schmappy-vs2010.sln)
 - Build as appropriate the platform and configuration.
 - Executables will be found in Debug/Release configuration directories for Win32.  There will be additional x64 subdirectory containing x64 executables.

KNOWN ISSUES
------------
- These tools are known to not work properly under some environments running newer versions of Exchange and/or Outlook.

LICENSING
---------
Unless otherwise attributed, these works are licensed under the Creative Commons Attribution license:  
http://creativecommons.org/licenses/by/3.0/legalcode.
