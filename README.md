### Importing Alma Analytics Reports to Google Sheets ###

**Purpose**

This Google Apps script allows for calling the Alma Analytics API and retrieving one or more reports, data from which is then used to populate tabs in a given Google spreadsheet.

**Setup**

1. Clone the repo and/or download `analytics_call.js` and `config.js`.

2. Modify `config.js` to include your API key from the (Ex Libris Developer Network)[https://developers.exlibrisgroup.com/] and the unique ID of the (target spreadsheet)[https://developers.google.com/sheets/api/guides/concepts#sheet_id]. You will also need to update the `reports` section of `config.js` to reflect the following:

   a. the sheet or sheet names to which you want to add data and 

   b. the path to each report (as defined within Alma Analytics).*

3. Open the your Google spreadsheet (the destination for the Alma Analytics reports), open the Script Editor, and create a new project (following the relevant instructions (here)[https://developers.google.com/apps-script/guides/sheets]). 

4. Select `New` --> `Script` from the `File` menu, and paste in the code from `analytics_call.js`. Repeat for `config.js`. 

5. Now you should be able to run the script either by selecting the `main` function from `analytics_call.js` in the Script Editor console or by setting up a (project trigger)[https://developers.google.com/apps-script/guides/triggers/].

*Note that the columns names as defined in the `columns` array of each object in `reports` are for convenience only; these will be derived at runtime from the Alma Analytics report itself.

**To Do**

* Add code to create a new menu item in the spreadsheet to call the script.

* Improve error handling: the `ui.alert` call fails on a time-based trigger.