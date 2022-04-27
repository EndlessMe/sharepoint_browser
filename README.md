# Sharepoint File Browser

Overview
------------
This utility will crawl all files and folder of the root Sharepoint site directory to locate all the items that have 
been modified within a user-specified number of days prior to the current date. The utility will output messages to
the console, notifying the user that it has located a file within the time period. Once all files have been analyzed,
the script will output a CSV file with the path to the file on the Sharepoint site as well as the last modified date.

Prerequisites
-----------
First install all dependencies.

`pip3 install -r requirements.txt`

Usage
-----------
Run the script.

`python3 sharepoint_browser.py`

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
5. Pfizer Active Directory username*
6. Pfizer Active Directory password*

*Note that the active directory credentials are required for the first time running the script only. They are 
subsequently stored inside a `.env` file, which is accessible agnostic of operating system being used.
Upon subsequent runs, the script will utilize these cached credentials and not prompt the user for inputs. 
In the event that the cached password is expired, the script will instead remove those cached credentials and re-prompt
user to authenticate. Again, these credentials will be cached and reused on subsequent runs of the script.

The script will output messages to the console informing the user of all files that are discovered that fall within the 
timespan specified. The resultant CSV file is output to the same directory that the script was run from, and will be 
named using the Sharepoint site name and dates that are within the user-specified timespan.

Happy crawling!