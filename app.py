from __future__ import print_function
import questionary
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import requests
from datetime import datetime
import collections

SCOPES = ['https://www.googleapis.com/auth/tasks.readonly', 'https://www.googleapis.com/auth/documents']

# Get Tasks Lists, along with it's metadata
def get_task_list(service_task):
    # Call the Tasks API
    results = service_task.tasklists().list(maxResults=10).execute()
    # print(results)
    # Get Tasks Lists, along with it's metadata
    items = results.get('items', [])
    return items

#Pick the tasks that are completed, and fall under the specified time-frame
def pick_tasks(items,service_task):
    # Allow users to select the task lists to move tasks from 
    todo_lists = []
    for item in items:
        todo_lists.append(item['title'])
    
    selected_lists = (
            questionary.checkbox(
                "Select List", choices=todo_lists,
            ).ask()
            or ['My List']
        )

    print(f"Moving tasks from {' and '.join(selected_lists)}.")

    #Add a date range
    min_date = questionary.text("Enter minimum date in format (yyyy-mm-dd)").ask()
    min_date = min_date+'T00:00:00.00Z'
    print(min_date)

    max_date_options = ['Now', 'Custom date']
    selected_maxDate = (
            questionary.rawselect(
                "Select maximum date of date range", choices=max_date_options,
            ).ask()
             or "do nothing"
        )
    if selected_maxDate == 'Now':
        max_date = datetime.now().strftime('%Y-%m-%dT%H:%M:%S.00Z')   

    else:    
        max_date = questionary.text("Enter maximum date in format (yyyy-mm-dd)").ask()
        max_date = max_date+'T00:00:00.00Z'
    print(max_date)

    #Filter completed tasks only
    completed_tasks = dict()
    for item in items:
        if item['title'] in selected_lists:
            task = service_task.tasks().list(tasklist=item['id'],showHidden=1,completedMin=min_date,completedMax=max_date).execute()

            # Filter tasks based on date range
            for i in task['items']:
                # We are concerned with date only, not time, hence splicing the date string to only include date
                if str(i['updated'])[0:10] in completed_tasks:
                    completed_tasks[str(i['updated'])[0:10]].append(i['title'])
                else:
                    completed_tasks[str(i['updated'])[0:10]]=[i['title']]
    
    #Sorting the tasks in ascending order of updated date
    completed_tasks = collections.OrderedDict(sorted(completed_tasks.items()))
    #print(completed_tasks)
    return completed_tasks

# Put the tasks, categorised by date, into a Google Doc 
def put_tasks(completed_tasks,service_docs ):
    # Prompt user to include google doc id, found in the link of the document: https://docs.google.com/document/d/DOCUMENT_ID/edit
    DOCUMENT_ID = questionary.text("Enter Google Doc Id").ask()
    document = service_docs.documents().get(documentId=DOCUMENT_ID).execute()
    print('The title of the document is: {}'.format(document.get('title')))

    # Format and append to doc
    for i in completed_tasks:
        date_text = '\t\t\t\t\t\t'+i+'\n\n'+'\n'.join(completed_tasks[i])+'\n-----------' +\
        '--------------------------------------------------------------------------------'+\
        '-----------------------------------\n'
        requests = [
            {
                'insertText': {
                    'location': {
                        'index': 1,
                    },
                    'text': date_text
                }
            },
        ]

        result = service_docs.documents().batchUpdate(
            documentId=DOCUMENT_ID, body={'requests': requests}).execute()


def main():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service_task = build('tasks', 'v1', credentials=creds)
    service_docs = build('docs', 'v1', credentials=creds)

    items = get_task_list(service_task)
    completed_tasks = pick_tasks(items,service_task)
    put_tasks(completed_tasks,service_docs)
    print("Completed Tasks Successfully Updated :)")





if __name__ == '__main__':
    main()