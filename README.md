# office-js-dialog-handler-bug-mac-os

## Description
This small project is example for demonstrate bug of <a href="https://github.com/OfficeDev/office-js">office-js</a> (office.js) library.
<p>
Link to <a href="https://github.com/OfficeDev/office-js/issues/297">bug</a>
<br>
Demonstration <a href="https://www.youtube.com/watch?v=O-Xskhwaenw&feature=youtu.be">video</a> of bug</p>
 
<p>It is plugin for outlook. There is 'open dialog' button which open new dialog window and adds event handler. The event handler gets message from dialog window and write message Id into RESULT panel.</p>

<p>Dialog message has 2 buttons. 'Close dialog' which sends message to parent window for close dialog. 'Send parent message' which generates randow number and sends message with this number to parent window.</p>

The project is derived from <a href="https://github.com/officedev/generator-office">generator-office</a> project.

## Install
Video <a href="https://youtu.be/xBdINuLV9Xg">how to install</a> this plugin for demonstrate <a href="https://github.com/OfficeDev/office-js/issues/297">bug</a>

1. Clone git repository
2. Go to the project directory
3. Run 'npm install'
4. Run 'npm start'

The project runs server on the <b>https://localhost:3000/</b> address

Load 'manifest.xml' file to Add-ins for outlook - plugin will be installed in your outlook client. (<a href='https://github.com/officedev/outlook-add-in-command-demo'>link how to install add-in</a>).

For correct work of https connection of plugin use certificates from 'certs' directory. Add 'ca.crt' file. 
Open Keychain Access on your Mac and go to the Certificates category in your System keychain. 
Once there, import the ca.crt using File > Import Items. 
Double click the imported certificate and change the “When using this certificate:” 
dropdown to Always Trust in the Trust section.