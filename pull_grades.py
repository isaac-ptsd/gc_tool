from __future__ import print_function
import json
import pickle
import os.path
import csv
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from gooey import Gooey, GooeyParser
import pprint
import pandas as pd

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
def assignment_lookup(course_id_in, id_in):
    # api call
    course_info = service.courses().courseWork().get(courseId=course_id_in, id=id_in).execute()
    return course_info["title"]


def student_lookup(user_id):
    """
    :param user_id:
    :return: users name: "familyName,  givenName"
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


@Gooey(program_name="Fetch Grades", )
def main():
    # API call to get classes
    """TODO: 1) look up students by name in template and update when found
                Goal is to preserve student ID's
             2) Pull all grades for all assignments by class. Create a separate import template for each assignment.
    """

    results = service.courses().list(pageSize=10).execute()

    # create list of course names to select from
    courses = results.get('courses')
    course_names = []
    for c in courses:
        course_names.append(c["name"])

    # gooey arguments
    parser = GooeyParser(description='Pull grades from Google ClassRoom')
    parser.add_argument('output_directory',
                        action='store',
                        widget='DirChooser',
                        help="Output directory to save csv of grades")
    parser.add_argument('course_selection',
                        action='store',
                        choices=course_names,
                        help="Choose a course to pull grade info from.")
    parser.add_argument('grade_template',
                        action='store',
                        widget='FileChooser',
                        help="Select a grade template to use")
    user_inputs = vars(parser.parse_args())

    selected_course = user_inputs['course_selection']
    slctd_crs_id = ''
    for c in courses:
        if c["name"] == selected_course:
            slctd_crs_id = c["id"]
    print("selected course id: ", slctd_crs_id)

    # API call to get course work
    student_submissions = service.courses().courseWork().studentSubmissions().list(courseId=slctd_crs_id,
                                                                                   courseWorkId='-').execute()
    print("course API call return value: ", student_submissions)

    print("CourseWorkId's: ")
    assignment_dict = [{c['courseId']: c['courseWorkId']} for c in student_submissions['studentSubmissions']]
    print(assignment_dict)
    for ass_dic in assignment_dict:
        print("ass_dic: --------------->", ass_dic)
        for k, v in ass_dic.items():
            print("ASSIGNMENT LOOKUP ------------> ", assignment_lookup(k, v))

    # print("ASSIGNMENT LOOKUP ------------> ", assignment_lookup("115992425191", "116601796515"))

    name_grade_dict_list = []
    try:
        id_grade = get_userId_grade(student_submissions)
        print("get_grades: ", id_grade)
        print("Students: ")
        for d in get_userId_grade(student_submissions):
            for key in d:
                print(" ", student_lookup(key))
        name_grade_dict_list = swap_id_for_name(id_grade)
        print("swap_id_for_name: ", swap_id_for_name(id_grade))
    except KeyError:
        print("No grades found for ", selected_course)
    except:
        print("uh-oh something broke")

    grade_template = user_inputs['grade_template']
    save_path = user_inputs['output_directory']
    df = pd.read_csv(grade_template, header=None)
    df = df.drop(df.index[7:100])
    df.columns = ['A', 'B', 'C']
    start_row_index = 7
    df.at[1, 'B'] = selected_course
    print(df)
    for d in name_grade_dict_list:
        print("PRINTING D: ", d)
        for key, value in d.items():
            df.at[start_row_index, 'B'] = key
            df.at[start_row_index, 'C'] = value
            start_row_index += 1

    output_file = os.path.join(save_path, selected_course.replace(" ", "_") + "_grade_template.csv")
    df.to_csv(output_file, header=False, index=False)


if __name__ == '__main__':
    main()
