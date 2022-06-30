# Hermes SharePoint File Browser

Overview
------------
This utility will crawl all files and folder of a user-specified SharePoint site directory to locate all the items that 
have been created and/or modified within a user-specified number of days prior to the current date. The utility will 
output messages to the console, notifying the user that it has located a file within the time period. Once all files 
have been analyzed, the script will output a CSV file with the URL to the folder of the recently created/modified files
on the SharePoint site, the filename, whether the file is 'new' or 'modified', and the date the file was 
created/modified.

Prerequisites
-----------
First install all dependencies.

`pip3 install -r requirements.txt`

Usage
-----------
Run the script.

`python3 hermes.py`

The script will prompt the user for the following inputs:
1. SharePoint site to explore for updates. Selectable from a list of options as an integer.
   * Example: `0`
2. The number of days prior to today's date that the user would like to view the list of files that have been since received updates.
   * Example: `14`
3. Whether the user would like to search a specific directory within the selected SharePoint site or to search the entire SharePoint site.
   * `Y` = Search specific path
   * `N` = Search SharePoint site beginning at `<SITE_NAME>/Shared Documents/`
4. Whether the user would like to input the path to the directory they would like to explore, or whether they would like to enter interactive mode and browse the site directory structure and select one.
   * `Y` = Enter path to directory to be searched following the `/Documents/`prompt.
   * `N` = Begin interactive mode. Press `ENTER` once in the desired directory
5. Pfizer Active Directory username
6. Pfizer Active Directory password

The script will output messages to the console informing the user of all files that are discovered that fall within the 
timespan specified. The resultant CSV file is output to the same directory that the script was run from, and will be 
named using the SharePoint site name and dates that are within the user-specified timespan.

Alternatively, Hermes can now be run using command line arguments directly. To ensure consistency across all SharePoint
sites, ***please ensure to click the `Copy Link` button in SharePoint's user interface***. Use this URL as the path
parameter. The following are the command line arguments that are needed to be passed into Hermes to successfully pull 
the updates:
```
options:
  -h, --help   show this help message and exit
  --path PATH  The direct path to the target SharePoint directory  (STRING)
  --days DAYS  Number of days prior to today to check for updates  (INTEGER)
  --user USER  Pfizer Username                                     (STRING)
  --pwd  PWD   Password                                            (STRING)
  --mode MODE  Search mode (new, new_folder, modified, or both)    (STRING)
 ```

In order to prevent your terminal shell from misinterpreting the arguments passed, ***surround all arguments marked as 
strings with either single or double quotes***. See example below.
### Example CLI Usage:
`python3 hermes.py --path 'https://pfizer.sharepoint.com/:f:/r/sites/DMTICT-44ScratchSleepTeam/Shared%20Documents/DOLBY%20FLARE?csf=1&web=1&e=t3EHc5'
 --days 14 --user 'lubkem01' --pwd '********' --mode 'modified'`

Happy crawling!