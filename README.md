# VBAChromeDevProtocol
VBA (Excel) based wrapper for Chrome Developer Protocol (CDP) - sorta a VBA version of Puppeteer/Selenium

Note: if you can use Puppeteer, Playright, Selenium, or some other tool - then use it!
But if you can only use VBA, then this is meant to provide a means to automate Chrome or Edge based browsers.  (Possibly Firefox via its limited CDP support, though currently untested/nonfunctional). 

TODO: this currently relies on Chrome/Edge's ability to do the CDP via pipes when started, a future version will connect to the websocket to allow connecting to existing browsers.

See https://chromedevtools.github.io/devtools-protocol/ for overview of Chrome Devloper Protocol

Special thanks / uses source based on 

JsonConverter : https://github.com/VBA-tools/VBA-JSON - BSD license copyright Ryo Yokoyama
basUtf8FromString : https://www.di-mgt.com.au/basUtf8FromString64.bas.html - MIT copyright David Ireland DI Management Services Pty
WinHttpCommon / wsocket : https://github.com/EagleAglow/vba-websocket-class - MIT copyright ?

