# portfolioManager
This is a portfolio manager designed for Google Sheets using Google Apps Script.

This program's express purpose is for the streamlining of stock market portfolio tracking in Google Sheets through the use of either a future add-on or through the Google Apps Script editor.

This program is designed to be pushed to Google using `Clasp`. Click [here](https://github.com/google/clasp) to learn more about `Clasp`.

This program is written in `TypeScript`, meaning some additional work will be necessary in order to push to Google. 

It is highly recommended to use `Yarn` instead of `npm` while working on this program. You can install `Yarn` [here](https://yarnpkg.com/lang/en/docs/install/).

## Instructions
### Initial setup
Before working on this program its imperative that `Clasp` is installed. 
To do this you must first install `Node.js`, as it is neccessary to install `Clasp`. You can download `Node.js` [here](https://nodejs.org/en/download/).

To get `Clasp`, open a terminal and run `yarn global add @google/clasp` or `npm i @google/clasp -g` if not using `Yarn`.

After installing `Clasp` you need to log in and give `Clasp` permissions to edit script files on your Google account. To do this run `clasp login` and go through the login screen.

Once logged in you need to enable the `Google Apps Script API` setting in [Google Scripts](https://script.google.com). To do this, go to your [Google Apps Script User Settings](https://script.google.com/home/usersettings), click `Google Apps Script API` and set it to `On` if it isn't already.

### Creating and setting up the project
Unfortunately there is no way to create a project in Google Scripts and link a spreadsheet to it after the fact, so we must first create a spreadsheet in Google Sheets. To do this, go to [Google Drive](https://drive.google.com) and navigate to where you want to keep your spreadsheet, then click `NEW`, followed by `Google Sheets` in the dropdown menu. This will create a blank spreadsheet for you to work with.

Now that you have a spreadsheet to link the script to, open the Google Script Editor using `alt+T+E` or by going into the `Tools` menu and clicking `Script Editor`. Once in the Google Script Editor, you must set a name for the project (we recommend "portfolioManager" for consistency) and save it.

Now that the project is prepared you need to get the project ID so `Clasp` knows where to push to. In the Google Script Editor, go into the `File` menu and select `Project properties`. Once in `Project properties` find and copy the Script ID. Once you have the script ID, create a file called `.clasp.json` in your local `portfolioManager` directory containing `{"scriptId":"<your script id>"}`.

## Pushing to Google
Once you have created a project in Google Scripts and are ready to push your work to Google you can do so using `yarn tscpush` or `npm run-script tscpush` if not using `Yarn`. 

Alternatively, if you don't wish to do it using `Yarn` or `npm` you will need to compile the `TypeScript` files into `JavaScript` before you can push to Google. To do this, run `tsc` in a terminal.

Once the `TypeScript` has been compiled into `JavaScript`, you can push the files to Google using `clasp push`. 