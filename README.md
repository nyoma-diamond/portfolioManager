# portfolioManager
This is a portfolio manager designed for Google Sheets using Google Apps Script.

This program's express purpose is for the streamlining of stock portfolio tracking in Google Sheets through the use of either an add-on or the through the Google Apps Script editor.

This program is designed to be pushed to Google using Clasp. Click [here](https://github.com/google/clasp) to learn more about Clasp.

This program is written in TypeScript, meaning some additional work may be necessary in order to push to Google. 

## Instructions
### Initial setup
Before working on this program its imperative that Clasp is installed. 
First, install Node.js since it is neccessary to install Clasp. You can download Node.js [here](https://nodejs.org/en/download/).

To get Clasp, open a terminal and run `npm i @google/clasp -g`

After installing clasp you need to log in and give Clasp permissions to edit script files on your Google account. To do this run `clasp login` and go through the login screen.

### Creating the project
For testing and use of this program you will need to create a new project in Google Scripts. To do this, open a terminal and run  `clasp create "portfolioManager"`. The project does not need to be named "portfolioManager," but it is recommended for consistency's sake.

### Pushing to Google
Once you have created a project in Google Scripts and are ready to push your work to Google, you need to compile the TypeScript files into JavaScript.
To do this, run `tsc` in a terminal.

Now that the TypeScript has been compiled into JavaScript, you can push the files to Google using `clasp push`. 

### Google Sheets setup
