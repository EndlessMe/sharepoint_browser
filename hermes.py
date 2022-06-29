import adal
from office365.graph_client import GraphClient
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.runtime.auth.user_credential import UserCredential
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext
from office365.runtime.client_request_exception import ClientRequestException
from datetime import datetime, timedelta
from dateutil.tz import tzutc
from dotenv import load_dotenv
import dateutil.parser
import pandas as pd
import getpass
import os
import argparse


# Variables
sharepoint_url = 'https://pfizer.sharepoint.com/sites/'
tenant_id = '7a916015-20ae-4ad1-9170-eefd915e9272'
time_now = datetime.now(tzutc()).replace(hour=0, minute=0, second=0, microsecond=0)
updated_files = []
update_type = []
modified_dates = []
selected_directory = ''
load_dotenv()


def authenticate_app():
    authority_url = 'https://login.microsoftonline.com/{0}/oauth2/v2.0/token'.format(tenant_id)
    auth_ctx = adal.AuthenticationContext(authority_url)
    client_id = 'd0376a3a-fbeb-43f3-b443-c7b094f4ee76'
    secret_id = 'c878Q~Xn18uPAy~8WC66X1EuHecE3HZHHeIZnatu'
    token = auth_ctx.acquire_token_with_client_credentials(
        "https://graph.microsoft.com",
        client_id,
        secret_id
    )
    return token


def authenticate_cli_user(site_url, user, pwd):
    """
    :param: site_url: str - URL of SharePoint site to be crawled for updates. Used for authentication
    :return: Authenticated context object
    """
    try:
        user_creds = UserCredential(user + '@pfizer.com', pwd)
        return ClientContext(site_url).with_credentials(user_creds)
    except IndexError:
        print('Authentication Error.')


def authenticate_user(site_url):
    """
    :param: site_url: str - URL of SharePoint site to be crawled for updates. Used for authentication.
    :return: Authenticated context object
    """
    # Take in SharePoint site URL, prompt for username and password, return OAuth token.
    try:
        username = str(input('Enter Pfizer username: '))
        password = getpass.getpass('Enter password: ')
        # ctx_auth = AuthenticationContext(site_url)
        # ctx_auth.acquire_token_for_user(username + '@pfizer.com', password)
        user_creds = UserCredential(username + '@pfizer.com', password)
        # return ClientContext.with_user_credentials(site_url, ctx_auth)
        return ClientContext(site_url).with_credentials(user_creds)

    except IndexError:
        print('Authentication Error.')


def name_to_site_mapping():
    # Keep a managed list of noted SharePoint sites that we can select from
    # Dictionary is comprised as {sharepoint_site_title: sharepoint_site_name}
    return {
        'axion': 'axion',
        'DMTI': 'DigitalDataScience',
        'DMTI ARI Team': 'DMTIRSVTeam',
        'DMTI - Friedreich Ataxia': 'DMTIFriedreichsAtaxiaStudyTeam',
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
    elif type(int(choice)) == int:
        return select_folder(folder_list[int(choice)], context)
    # Improper inputs will just simply display a message and start the process over again rather than exiting.
    else:
        print("Select from options and try again.")
        return select_folder(parent_folder, context)


def crawl_folders(parent_folder, fn, time_delta, mode='both'):
    """
    :param parent_folder: str - Root Folder to start crawling for updates
    :param fn: func - function to run on files in sharepoint site
    :param time_delta: int - Representation of number of days to lookup for updates
    :param mode: str - Either new files only, modified files only, or both modified/new files.
    """
    parent_folder.expand(['Files', 'Folders']).get().execute_query()
    if mode == 'new_folder':
        for folder in parent_folder.folders:
            fn(folder, time_delta, mode)
        for folder in parent_folder.folders:
            crawl_folders(folder, fn, time_delta, mode)
    else:
        # First check every file's last modified date in the current directory
        for file in parent_folder.files:
            fn(file, time_delta, mode)
        # Then repeat the process for each folder recursively
        for folder in parent_folder.folders:
            crawl_folders(folder, fn, time_delta, mode)


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


def add_discovered_file_to_lists(path, file_type, date):
    updated_files.append(path)
    update_type.append(file_type)
    modified_dates.append(date)
    if file_type == 'new_folder':
        print('Found new folder:', path, date)
    else:
        print('Found {0} file:'.format(file_type), path, date)


def find_recent_updates(f, time_delta, mode='both'):
    """
    :param f: obj - File to check for modified date.
    :param time_delta: int - Representation of number of days to perform search for updated files
    :param mode: str - Which search mode to perform. New files only, modified files only, or both modified and new files.
    """
    path = f.properties['ServerRelativeUrl']
    created_date = dateutil.parser.parse(f.properties['TimeCreated'])
    modified_date = dateutil.parser.parse(f.properties['TimeLastModified'])
    search_mode = mode
    if search_mode == 'both':
        if time_now - created_date < time_delta:
            add_discovered_file_to_lists(path=path, file_type='new', date=created_date)
        elif time_now - modified_date < time_delta:
            add_discovered_file_to_lists(path=path, file_type='modified', date=modified_date)
    elif search_mode == 'new':
        if time_now - created_date < time_delta:
            add_discovered_file_to_lists(path=path, file_type=search_mode, date=created_date)
    elif search_mode == 'new_folder':
        if time_now - created_date < time_delta:
            add_discovered_file_to_lists(path=path, file_type=search_mode, date=created_date)
    elif search_mode == 'modified':
        if time_now - modified_date < time_delta:
            add_discovered_file_to_lists(path=path, file_type=search_mode, date=modified_date)


def transform_paths_to_urls(dataframe):
    df = dataframe.copy(deep=True)
    df['path'] = df['path'].map(lambda x: x.split('/'))
    df['name'] = df['path'].map(lambda x: x.pop())
    df['path'] = df['path'].map(lambda x: '/'.join(x))
    df['path'] = df['path'].map(lambda x: x.replace(' ', '%20'))
    df['path'] = 'https://pfizer.sharepoint.com' + df['path'].astype(str)
    return df[['path', 'name', 'file_type', 'date']]


def export_results(files, dates, file_type, site_name, time_delta):
    """
    Export results to CSV file.
    :param files: list - Takes in the array of discovered file paths during processing
    :param dates: list - Takes in the array of discovered updated dates during processing
    :param site_name: str - Name of the SharePoint site. Used to generate filename
    :param time_delta: int - The user-specified lookup time. Used in generating the filename.
    :return: None. Generates CSV file of results in working directory and prints informative message of location.
    """
    file_df = pd.DataFrame.from_dict({'path': files, 'file_type': file_type, 'date': dates}, orient='index').transpose()
    transformed_df = transform_paths_to_urls(file_df)
    file_name = site_name.lower() + '_modified_files_' + (str(time_now - time_delta)).split(' ')[0] + '_to_' + \
        str(time_now).split(' ')[0] + '.csv'
    print('Writing file to disk...')
    # file_df.to_csv(file_name, index=False)
    transformed_df.to_csv(file_name, index=False)
    print('Results are located in', file_name)


def process(site_name, time_delta, browse_path='/Shared Documents/', browse_mode=False, search_mode='both', cli=False, user=None, pwd=None):
    """
    Routine that handles orchestration of SharePoint browser. Completes with CSV file of results.
    :param site_name: str - SharePoint site name
    :param time_delta: int - User-specified amount of time to use for the range of days from present
    :param browse_path: str - Location to begin searching for updated files
    :param browse_mode: bool - Determines whether to initiate interactive mode
    :param search_mode: str - Determines whether to search for new files, updated files, or both.
    :param cli: boolean - Determines whether to process using CLI mode or Interactive mode
    :return: None. Calls export_results to export results to CSV file.
    """
    try:
        if not cli:
            site_url = sharepoint_url + site_name
            time_delta = time_delta
            ctx = authenticate_user(site_url=site_url)
            # ctx = authenticate_app()
            web = ctx.web
            ctx.load(web)
            # client = GraphClient(ctx)
            ctx.execute_query()
            # client.drives.get().execute_query()
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
            crawl_folders(parent_folder=root_folder, fn=find_recent_updates, time_delta=time_delta, mode=search_mode)
            print('~~~~~~~~~~~~~~~~~~~~~~~~~~~~~')
            print('Complete!')
        else:
            ctx = authenticate_cli_user(site_url=sharepoint_url+site_name, user=user, pwd=pwd)
            web = ctx.web
            ctx.load(web)
            ctx.execute_query()
            site_title = web.properties['Title']
            print('Authenticated into sharepoint site: ', site_title)
            root_folder = ctx.web.get_folder_by_server_relative_url(browse_path)
            crawl_folders(parent_folder=root_folder, fn=find_recent_updates, time_delta=time_delta, mode=search_mode)
            print('~~~~~~~~~~~~~~~~~~~~~~~~~~~~~')
            print('Complete!')
    except ClientRequestException:
        print('No Sharepoint site found with provided directory path value entered. Please try again.')


def main():
    parser = argparse.ArgumentParser(description="Command Line Interface for Hermes")
    parser.add_argument("--path", help="The direct path to the target SharePoint directory", required=False, type=str, default="")
    parser.add_argument("--days", help="Number of days prior to today to check for updates", required=False, type=int, default=14)
    parser.add_argument("--user", help="Pfizer Username (string)", required=False, type=str, default="")
    parser.add_argument("--pwd", help="Password (string)", required=False, type=str, default="")
    parser.add_argument("--mode", help="Search mode (new, new_folder, modified, or both) to use (string)", required=False, type=str, default="both")
    argument = parser.parse_args()
    cli_status = False
    if argument.path:
        cli_status = True
    try:
        if cli_status is True:
            # This assumes that the Copy Link button is being used in SharePoint
            site_name = argument.path.split('/')[6]
            copy_link = argument.path.split('?')[0]
            folder_path = '/sites/' + site_name + copy_link.split(site_name)[1].replace('%20', ' ')
            lookup_days = timedelta(days=argument.days)
            process(site_name=site_name, time_delta=lookup_days, browse_path=folder_path,
                    search_mode=argument.mode, cli=True, user=argument.user, pwd=argument.pwd)
            export_results(files=updated_files, file_type=update_type, dates=modified_dates, site_name=site_name,
                           time_delta=lookup_days)
        else:
            key_index = 0
            print('')
            for name in name_to_site_mapping():
                print("[{0}] {1}".format(key_index, name))
                key_index += 1
            target_site = int(input("Enter number for SharePoint site you'd like to check for updates: "))
            site_name = list(name_to_site_mapping().values())[target_site]
            days = int(input("Enter desired number of days prior to today that you'd like to check for updates: "))
            lookup_days = timedelta(days=days)
            print()
            print("[0] New files only")
            print("[1] New folders only")
            print("[2] Modified files only")
            print("[3] Both new and modified files")
            update_search_mode = int(input("Choose the number for the search mode you'd like to use from the options above: "))
            modes = ['new', 'new_folder', 'modified', 'both']
            # Allow option to search a directory other than root Shared Documents directory
            search_named_folder = str(input("Would you like to search a specific directory? Enter 'N' to search this entire SharePoint site. Y/N: "))
            if search_named_folder.lower() == 'y' or search_named_folder.lower() == 'yes':
                # Give user option of browsing SharePoint site directory structure.
                browse_dirs = str(input("Do you know the path to the directory? Enter 'N' to initiate browse mode: "))
                # If user does not know path, enter interactive browse mode.
                if browse_dirs.lower() == 'n' or browse_dirs.lower() == 'no':
                    process(site_name=site_name, time_delta=lookup_days, browse_mode=True, search_mode=modes[update_search_mode])
                    export_results(files=updated_files, file_type=update_type, dates=modified_dates, site_name=site_name, time_delta=lookup_days)
                # Otherwise, prompt user for path. Useful if users have direct path saved in a note somewhere.
                # SharePoint is terrible at granting the ability to directly copy the absolute path from the web.
                elif browse_dirs.lower() == 'y' or browse_dirs.lower() == 'yes':
                    print('Enter path to folder. Path is case-sensitive!')
                    path = "/Shared Documents/" + str(input("{0}/Documents/".format(site_name)))
                    process(site_name=site_name, time_delta=lookup_days, browse_path=path, search_mode=modes[update_search_mode])
                    export_results(files=updated_files, file_type=update_type, dates=modified_dates, site_name=site_name, time_delta=lookup_days)
                else:
                    print("Please enter 'Y' to search a specific directory path, or 'N' to enter interactive browse mode.")
            # Search root site directory
            elif search_named_folder.lower() == 'n' or search_named_folder.lower() == 'no':
                process(site_name=site_name, time_delta=lookup_days, search_mode=modes[update_search_mode])
                export_results(files=updated_files, file_type=update_type, dates=modified_dates, site_name=site_name, time_delta=lookup_days)
            else:
                print('Input either "yes" or "no" and try again.')
    except AttributeError:
        print('Unable to authenticate with credentials. Please try again.')
    except IndexError:
        print('Please select number from list of options and try again.')
    except ValueError as ve:
        raise('Please enter an integer.', ve)


if __name__ == '__main__':
    main()
