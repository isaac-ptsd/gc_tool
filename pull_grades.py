from __future__ import print_function
import json
import pickle
import os.path
import csv
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from argparse import ArgumentParser
from gooey import Gooey, GooeyParser
import pprint

###############################################
# Credentials / Authorization
###############################################

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/classroom.courses.readonly',
          'https://www.googleapis.com/auth/classroom.coursework.students',
          'https://www.googleapis.com/auth/classroom.rosters.readonly']

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

service = build('classroom', 'v1', credentials=creds)

###############################################
# Credentials / Authorization
###############################################


def student_lookup(user_id):
    """
    :param user_id:
    :return: users name: (familyName,  givenName)
    """
    user_info = service.userProfiles().get(userId=user_id).execute()
    return user_info["name"]["familyName"] + ", " + user_info["name"]["givenName"]


def get_userId_grade(json_in):
    """
    :param json_in: json object of student submissions, this is the return value of the api call to get course work
    :return: list of dictionaries: [{<userID>: <grade>}]
    """
    student_submissions = json_in["studentSubmissions"]
    return [{c["userId"]: c["assignedGrade"]} for c in student_submissions if "userId" in c and "assignedGrade" in c]


def swap_id_for_name(list_of_dicts_in):
    """
    :param list_of_dicts_in: [{<userID>: <grade>}]
    :return: [{<familyName, givenName>: <grade>}]
    """
    return [{student_lookup(k): v} for d in list_of_dicts_in for k, v in d.items()]


def to_csv(list_of_dicts_in, name_of_csv_to_create):
    """ Function that takes a list of dictionaries and creates a csv file.
    Parameter: list of dictionaries
        list_of_dicts_in; the list of dictionaries to create a csv out of
    Parameter: string
        name_of_csv_to_create; this will be the name of the resulting csv file - NOTE: include .csv
    Returns: no return value
        will create a csv file in current directory
    """
    try:
        keys = list_of_dicts_in[0].keys()
        with open(name_of_csv_to_create, 'w', newline='') as output_file:
            dict_writer = csv.DictWriter(output_file, keys)
            dict_writer.writeheader()
            dict_writer.writerows(list_of_dicts_in)
    except Exception as e:
        print(e)


# @Gooey(program_name="Fetch Grades",)
# def parse_args():
#     """ Use ArgParser to build up the arguments we will use in our script
#     Save the arguments in a default json file so that we can retrieve them
#     every time we run the script.
#     """
#     stored_args = {}
#     # get the script name without the extension & use it to build up
#     # the json filename
#     script_name = os.path.splitext(os.path.basename(__file__))[0]
#     args_file = "{}-args.json".format(script_name)
#     # Read in the prior arguments as a dictionary
#     if os.path.isfile(args_file):
#         with open(args_file) as data_file:
#             stored_args = json.load(data_file)
#     parser = GooeyParser(description='Pull grades from Google ClassRoom')
#     parser.add_argument('output_directory',
#                         action='store',
#                         widget='DirChooser',
#                         default=stored_args.get('output_directory'),
#                         help="Output directory to save csv of grades")
#     args = parser.parse_args()
#     # Store the values of the arguments so we have them next time we run
#     with open(args_file, 'w') as data_file:
#         # Using vars(args) returns the data as a dictionary
#         json.dump(vars(args), data_file)
#     return args


# if __name__ == '__main__':
#     conf = parse_args()
#     print("Running main")
#     main(conf.output_directory)
#     print("Done")

def main(write_path):
    """Shows basic usage of the Classroom API.
    Prints the names of the first 10 courses the user has access to.
    """

    # API call to get classes
    results = service.courses().list(pageSize=10).execute()
    courses = results.get('courses', [])

    if not courses:
        print('No courses found.')
    else:
        print('Writing Courses:')
        complete_name = os.path.join(write_path, "courses.txt")
        file1 = open(complete_name, "w")
        file1.write(str(courses))
        file1.close()

    # API call to get course work
    student_submissions = service.courses().courseWork().studentSubmissions().list(courseId='115992425191',
                                                                                   courseWorkId='-').execute()

    id_grade = get_userId_grade(student_submissions)
    print("get_grades: ", id_grade)
    # print("swap_id_for_name: ", swap_id_for_name(student_submissions))
    for d in get_userId_grade(student_submissions):
        for key in d:
            print(student_lookup(key))
    print(swap_id_for_name(id_grade))


    if not student_submissions:
        print('No course work found.')
    else:
        print('Writing course work:')
        complete_name = os.path.join(write_path, "grades.txt")
        file1 = open(complete_name, "w")
        file1.write(str(student_submissions))
        file1.close()


if __name__ == '__main__':
    main("c:\\Users\isaac.stoutenburgh\Documents")

