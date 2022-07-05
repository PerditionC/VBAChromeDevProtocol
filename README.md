# VBAChromeDevProtocol
**VBA** (Excel) based wrapper for **Chrome Developer Protocol (CDP)** - sorta a VBA version of Puppeteer/Selenium

Note: if you can use Puppeteer, Playright, Selenium, or some other tool - then use it!
But if you can only use VBA, then this is meant to provide a means to automate Chrome or Edge based browsers.  (Possibly Firefox via its limited CDP support, though currently untested/nonfunctional). 

See https://chromedevtools.github.io/devtools-protocol/ for overview of Chrome Devloper Protocol

Initial work based on information and clsEdge from https://www.codeproject.com/Tips/5307593/Automate-Chrome-Edge-using-VBA

Primarily connects directly to browser using Chrome/Edge's ability to use the CDP via pipes when started, however, now also has basic support for connecting to browser through standard websocket interface so can reuse already open browser if started with CDP port 9222 listening.

Currently primarily tested with and assumes working with Edge; however will detect and support spawning Chrome and possibly FireFox (Chrome should work at least in websocket mode, Firefox is untested) 

## TODO
- improve/add usage documentation
- generator needs some more work (still has some class names too long/clash, still has some clashes with reserved words, incorrectly assumes class for unspecified _object_)

## Note
when downloading the source files from git, be sure to convert to DOS/Windows CRLF style endings for the text files or Office may import as wrong module type (regular modules instead of class modules) - to be updated to ensure git always does this

## Usage
see Example.xlsm - documentation to be added

## Special thanks / uses source based on 

- clsCDP derived from clsEdge : https://www.codeproject.com/Tips/5307593/Automate-Chrome-Edge-using-VBA - CPOL license copyright ChrisK23
- JsonConverter : https://github.com/VBA-tools/VBA-JSON - BSD license copyright Ryo Yokoyama
- clsProcess : https://stackoverflow.com/questions/62172551/error-with-createpipe-in-vba-office-64bit
- basUtf8FromString : https://www.di-mgt.com.au/basUtf8FromString64.bas.html - MIT copyright David Ireland DI Management Services Pty
- WinHttpCommon / clsWebSocket : https://github.com/EagleAglow/vba-websocket-class - MIT copyright EagleAglow
- ClipboardUtils : https://msdn.microsoft.com/en-us/library/office/ff192913.aspx?f=255&MSPPError=-2147217396
- WinWindowStyle : https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-showwindow

- Denis Lessard for work detecting correct path and which browsers are installed along with improving logic for which one to spawn

Testing and enhancements by Jason Pullen and Kenneth J. Davis
