*CataAuto*

REQUIREMENTS
-------------
Microsoft Excel 2016 or later
Scripting Runtime


INSTALLATION
-------------
Double click to open the .xlam file. Microsoft will identify this file and automatically install it as an excel add-in.


MAINTAINERS
-------------
Sean Hall (SupaSneak) - sean@oneclickretail.com, supasneak@gmail.com


GETTING STARTED:
-------------
Hotkey: Ctr + Shift + Z
How to use:
1. - Launch the user interface by using the Hotkey. 
	a. Data Entry (Doing Needs Reviews)
 		- Currently incomplete
	b. Quality Assurance (Checking NRDs)

1. - Begin every catalog by clicking “Run”

2. - When completed you will be presented with information that is simply for convenience. Press "OK" to continue. Two new columns have been created. You can ignore column B which will be titled "HIERARCHY". Column A will provide you with the errors found in the highlighted rows. Rows where only blank errors were found will not be highlighted red in Column A.

3. - Go through column A and correct the specified errors for the specified rows.
	a. The "CHECK BRAND" error simply means that this particular brand is not recognized because it is not a brand associated with any
	   products that are not marked "NRD". Verify that the brand listed is correct.
	b. The "CHECK BLANKS" error refers to any cell in that row which is blank. These are not necessarily errors but should all be filled
	   in with "UNASSIGNED". Use your own judgement to determine what cells really should be left blank (sub categories for example).
	c. The "CHECK CATEGORIES" error means that there is an inconsistency with the category hierarchy for this product. Check for any
	   any spelling errors, make sure the listings are legitimate categories, and check to make sure the category hierarchy is correct.

4. - Repeat steps 3 through 5 until you no longer have any errors. You will know you have no more errors when presented with the option to
     restore your catalog to its default format. Select "YES" to perform the catalog restore. Selecting "NO" leaves the catalog in its current
     state. The catalog can be restored at any time using the tool "Restore Catalog" under the add-in's "Advanced" tab.


Advanced Tab:
To run any tool any the advanced tab simply check the box next to the tool you would like to run and click "Run"

Below is a list of the advanced tools:
*Isolate Y-Includes - When run, this tool will separate all Y-Includes from the Z-Excludes and store them on separate sheets. If these separate sheets already exist then any Y-Include listed within the Z-Excludes will be moved over to the Y-Includes.

*Custom Hierarchy - This tool enables you to work with a catalog that has more or less than 3 category columns. When run, you will be prompted to select the category columns. Simply select each category column while holding control and then press "OK" (Whenever the tool is closed it will forget this custom hierarchy so it is important to run the catalog check immediately after using Custom Hierarchy or the hierarchy will revert to default)

*Search Zexclude Titles - This tool compiles a list of keywords using your catalog's current data and searches for those keywords within the titles of Z-Exclude products. This narrows down the amount of products that actually need to be reviewed.

*Pull NRD Brands - This tool will save you the time of manually adding valid brands to the manufacturer and brand list.

*List MFRS and Brands - This tool will prepare a list of valid manufacturer and brand relationships and place it on a new sheet.

*Change NRD Date - This tool will automatically change "NRD" to the current date.

*Fill in Manufacturers - This tool only works when a manufacturer and brand list has been created and will automatically fill in the manufacturer for the NRD rows

*Fill in Blanks - This tool will fill in all blank cells found in your currently selected range.

*Trim all Cells - This tool will trim all cells on the currently active sheet.

*Restore Catalog - This tool will restore the catalog to one single sheet.

LATEST UPDATE: 7/27/2018, Version 1.3
-------------
Major Changes
 - The "Checker(Quality Assurance)" portion of the tool is now executed in a single button instead of two.
 - Less Prompts. Message Box Prompts will no longer interrupt the program's processes (you no longer need to click "OK" in order for it to proceed. There is still a prompt to notify you when the checker is complete.)
 - Identifying keywords in Zexclude ASINs is now included in the primary execution of the tool

Minor Changes
 - Additional framework has been added to the "Cataloger(Data Entry)" portion of the tool in preparation for further work.
 - Certain dialogue

Bug Fixes
 - Zexcludes sheet will no longer interfere with catalog restoration
 - Excel function errors should no longer cause any problems
 - You will no longer be able to run the Checker on a catalog that does not contain "NRD"
 - Error Count should now more closely reflect the actual amount of errors
 - Removed redundant code