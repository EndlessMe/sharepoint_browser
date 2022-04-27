from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.runtime.client_request_exception import ClientRequestException
from datetime import datetime, timedelta
from dateutil.tz import tzutc
from dotenv import load_dotenv
import dateutil.parser
import pandas as pd
import getpass
import os


# Variables
sharepoint_url = 'https://pfizer.sharepoint.com/sites/'
time_now = datetime.now(tzutc()).replace(hour=0, minute=0, second=0, microsecond=0)
updated_files = []
modified_dates = []
selected_directory = ''
load_dotenv()


def authenticate_user(site_url):
    """
    :param site_url: str - URL of SharePoint site to be crawled for updates. Used for authentication.
    :return: Authenticated context object
    """
    # Take in SharePoint site URL, prompt for username and password, return OAuth token.
    try:
        if 'SP_USER' in os.environ and 'SP_PASS' in os.environ:
            # username = str(input('Enter Pfizer username: '))
            # password = getpass.getpass('Enter password: ')
            print('Using cached credentials.')
        else:
            with open('./.env', "w") as env:
                username = str(input('Enter Pfizer username: '))
                password = getpass.getpass('Enter password: ')
                env.write('SP_USER="{0}"\nSP_PASS="{1}"'.format(username, password))
            load_dotenv()
        ctx_auth = AuthenticationContext(site_url)
        # ctx_auth.acquire_token_for_user(os.environ.get('SP_USER') + '@pfizer.com', os.environ.get('SP_PASS'))
        # ctx = ClientContext(site_url, ctx_auth)
        ctx_auth.acquire_token_for_user(os.environ.get('SP_USER') + '@pfizer.com', os.environ.get('SP_PASS'))
        return ClientContext(site_url, ctx_auth)
    except IndexError:
        print('Authentication Error.')
        os.remove('./.env')


def name_to_site_mapping():
    # Keep a managed list of noted SharePoint sites that we can select from
    # Dictionary is comprised as {sharepoint_site_title: sharepoint_site_name}
    return {
        'axion': 'axion',
        'DMTI': 'DigitalDataScience',
        'DMTI ARI Team': 'DMTIRSVTeam',
        'eSource Program': 'eSourceProgram944',
        'Vol 3 4.3.1 eSource CONFORM Implementation Project Team': 'Vol34.3.1eSourceCONFORMImplementationProjectTeam',
        'R in Pfizer': 'RinPfizer'
    }


def select_folder(parent_folder, context):
    """
    :param parent_folder: str - Folder that is expanded and has its contents listed
    :param context: obj - Context object that stores auth token and traverses web location
    :return: Return function itself unless directory is selected by inputting empty string. Returns directory path.
    """
    # Used only if browse_mode=True in the process function
    # Takes in path to parent folder (string) and authenticated context object
    folder_list = []
    idx = 0
    current_location = context.web.get_folder_by_server_relative_url(parent_folder)
    # Expand folder contents
    current_location.expand(['Folders']).get().execute_query()
    for folder in current_location.folders:
        # Add server relative URLs to the folder_list array
        folder_list.append(folder.properties['ServerRelativeUrl'])
    for item in folder_list:
        # Create list of options
        print("[{0}] {1}".format(idx, item))
        idx += 1
    print("[{0}] ../".format(len(folder_list)))
    print('')
    choice = str(input('Select a folder to browse or press ENTER to search current directory: '))
    # Case when ENTER key is pressed
    if choice == "":
        print('Current directory selected.')
        return parent_folder
    # Begin the folder selection process again in the prior directory
    elif choice == str(len(folder_list)):
        dirs = parent_folder.split('/')
        dirs.pop()
        return select_folder('/'.join(dirs), context)
    # Begin the folder selection process again using the directory the user has input.
    elif int(choice):
        return select_folder(folder_list[int(choice)], context)
    # Improper inputs will just simply display a message and start the process over again rather than exiting.
    else:
        print("Select from options and try again.")
        return select_folder(parent_folder, context)


def crawl_folders(parent_folder, fn, time_delta):
    """
    :param parent_folder: str - Root Folder to start crawling for updates
    :param fn: func - function to run on files in sharepoint site
    :param time_delta: int - Representation of number of days to lookup for updates
    """
    parent_folder.expand(['Files', 'Folders']).get().execute_query()
    # First check every file's last modified date in the current directory
    for file in parent_folder.files:
        fn(file, time_delta)
    # Then repeat the process for each folder recursively
    for folder in parent_folder.folders:
        crawl_folders(folder, fn, time_delta)


# def search_sharepoint(context, doc_index):
#     request = RequestOptions(
#         "{0}/_api/search/query?querytext='contentclass:STS_Site contentclass:STS_Web sharepoint indexdocid>" +
#         str(doc_index) +
#         "'&amp;sortlist='[docid]:ascending'".format(site_url)
#     )
#     response = context.execute_request_direct(request)
#     json_response = json.loads(response.content)
#     search_results = json_response['d']['query']['PrimaryQueryResult']['RelevantResults']['Table']['Rows']['results']
#     return search_results


def find_recent_updates(f, time_delta):
    """
    :param f: obj - File to check for modified date.
    :param time_delta: int - Representation of number of days to perform search for updated files
    """
    path = f.properties['ServerRelativeUrl']
    date = f.properties['TimeLastModified']
    created_date = dateutil.parser.parse(f.properties['TimeCreated'])
    modified_date = dateutil.parser.parse(date)
    if time_now - created_date < time_delta or time_now - modified_date < time_delta:
        updated_files.append(path)
        modified_dates.append(date)
        print('Found updated file:', path, date)


def export_results(files, dates, site_name, time_delta):
    """
    Export results to CSV file.
    :param files: list - Takes in the array of discovered file paths during processing
    :param dates: list - Takes in the array of discovered updated dates during processing
    :param site_name: str - Name of the SharePoint site. Used to generate filename
    :param time_delta: int - The user-specified lookup time. Used in generating the filename.
    :return: None. Generates CSV file of results in working directory and prints informative message of location.
    """
    file_df = pd.DataFrame.from_dict({'file_path': files, 'last_modified': dates}, orient='index').transpose()
    file_name = site_name.lower() + '_modified_files_' + (str(time_now - time_delta)).split(' ')[0] + '_to_' + \
        str(time_now).split(' ')[0] + '.csv'
    print('Writing file to disk...')
    file_df.to_csv(file_name, index=False)
    print('Results are located in', file_name)


def process(site_name, time_delta, browse_path='/Shared Documents/', browse_mode=False):
    """
    Routine that handles orchestration of SharePoint browser. Completes with CSV file of results.
    :param site_name: str - SharePoint site name
    :param time_delta: int - User-specified amount of time to use for the range of days from present
    :param browse_path: str - Location to begin searching for updated files
    :param browse_mode: bool - Determines whether to initiate interactive mode
    :return: None. Calls export_results to export results to CSV file.
    """
    try:
        site_url = sharepoint_url + site_name
        time_delta = time_delta
        ctx = authenticate_user(site_url=site_url)
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()
        site_title = web.properties['Title']
        print('Authenticated into sharepoint site: ', site_title)
        folder_url = '/sites/' + site_name + browse_path
        if browse_mode:
            selected_root_folder = select_folder(folder_url, ctx)
            print('Selected root folder =', selected_root_folder)
            root_folder = ctx.web.get_folder_by_server_relative_url(selected_root_folder)
        else:
            root_folder = ctx.web.get_folder_by_server_relative_url(folder_url)
        print('Crawling files and folders...')
        crawl_folders(parent_folder=root_folder, fn=find_recent_updates, time_delta=time_delta)
        print('~~~~~~~~~~~~~~~~~~~~~~~~~~~~~')
        print('Complete!')
    except ClientRequestException:
        print('No Sharepoint site found with provided directory path value entered. Please try again.')


if __name__ == '__main__':
    key_index = 0
    print('')
    for name in name_to_site_mapping():
        print("[{0}] {1}".format(key_index, name))
        key_index += 1
    try:
        target_site = int(input("Enter number for SharePoint site you'd like to check for updates: "))
        site_name = list(name_to_site_mapping().values())[target_site]
        days = int(input("Enter desired number of days prior to today that you'd like to check for updates: "))
        lookup_days = timedelta(days=days)
        # Allow option to search a directory other than root Shared Documents directory
        search_named_folder = str(input("Would you like to search a specific directory? Enter 'N' to search this entire SharePoint site. Y/N: "))
        if search_named_folder.lower() == 'y' or search_named_folder.lower() == 'yes':
            # Give user option of browsing SharePoint site directory structure.
            browse_dirs = str(input("Do you know the path to the directory? Enter 'N' to initiate browse mode: "))
            # If user does not know path, enter interactive browse mode.
            if browse_dirs.lower() == 'n' or browse_dirs.lower() == 'no':
                process(site_name=site_name, time_delta=lookup_days, browse_mode=True)
                export_results(files=updated_files, dates=modified_dates, site_name=site_name, time_delta=lookup_days)
            # Otherwise, prompt user for path. Useful if users have direct path saved in a note somewhere.
            # SharePoint is terrible at granting the ability to directly copy the absolute path from the web.
            elif browse_dirs.lower() == 'y' or browse_dirs.lower() == 'yes':
                print('Enter path to folder. Path is case-sensitive!')
                path = "/Shared Documents/" + str(input("{0}/Documents/".format(site_name)))
                process(site_name=site_name, time_delta=lookup_days, browse_path=path)
                export_results(files=updated_files, dates=modified_dates, site_name=site_name, time_delta=lookup_days)
            else:
                print("Please enter 'Y' to search a specific directory path, or 'N' to enter interactive browse mode.")
        # Search root site directory
        elif search_named_folder.lower() == 'n' or search_named_folder.lower() == 'no':
            process(site_name=site_name, time_delta=lookup_days)
            export_results(files=updated_files, dates=modified_dates, site_name=site_name, time_delta=lookup_days)
        else:
            print('Input either "yes" or "no" and try again.')
    except AttributeError:
        print('Unable to authenticate with credentials. Please try again.')
    except IndexError:
        print('Please select number from list of options and try again.')
    except ValueError:
        print('Please enter an integer.')



