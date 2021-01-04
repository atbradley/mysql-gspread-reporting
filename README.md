# MySQL SQL Reports to Google Sheets

This was written for Brown University's [Online Course Reserves Application](http://brown.edu/go/ocra), but should be generic enough for any MySQL-backed application (and should be easy enough to make more generic for other databases).

## Setup:

1. [Create a Google service account key](http://gspread.readthedocs.org/en/latest/oauth2.html) and save the credentials in JSON format.
2. Create a new folder in Google Drive and save it with the service account you just created.
2. Create a table like this in your database (you can replace the name `reports` with anything:  

	    CREATE TABLE `reports` (  
    		`run_order` INT(11) NOT NULL,  
    		`name` VARCHAR(100) NOT NULL,  
    		`query` MEDIUMTEXT NOT NULL,  
    		`description` VARCHAR(500) NOT NULL DEFAULT '',  
    		PRIMARY KEY (`name`),  
    		UNIQUE INDEX `order` (`run_order`)  
    	);

3. Copy the file `ocra-data.conf.example.yaml` to `ocra-data.conf.yaml` and update with your database credentials, the name of your "reports" table, the ID of the new folder you created (you can copy this from the folder's URL), and the path to the service account credentials .json file.
4. Add your reports to your `reports` table. `runorder` determines the order of reports in the output spreadsheet; `name` determines the name of the sheet and will appear, along with `description` on the first page of the file as a table of contents. The `query` is simply an SQL statement that returns the data you want to save to the spreadsheet.
5. `pip install -r requirements.txt`
5. `python ocra-reporting.py`.