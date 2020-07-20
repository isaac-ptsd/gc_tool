from __future__ import print_function
import pickle
import os
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from gooey import Gooey, GooeyParser
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


def find_name_location(student_name, df):
    """
    :param df:
    :param student_name: (string)
    :return: This searches the user specified grade template for a student name and returns the cell coordinates
    for the students name if found, and returns NULL if not found.

    """
    a = df.index[df['B'].str.contains(student_name, na=False)]
    if a.empty:
        return 'not found'
    elif len(a) > 1:
        return a.tolist()
    else:
        # only one value - return scalar
        return a.item()


def create_import_file(grade_template, save_path, assignment_name, course_name,  name_grade_dict_list):
    """
    :param grade_template: csv grade template from PowerSchool
    :param save_path: where to save new csv file (to be used as an import into PowerSchool)
    :param assignment_name: user selected course passed in as a parameter from the GUI
    :param name_grade_dict_list: list of dictionaries: [{<student name>: <grade>}]
    :return:
    This function will create a csv file of student grades following the grade template format
    """
    try:
        df = pd.read_csv(grade_template, header=None)
        df.columns = ['A', 'B', 'C']
        df.at[1, 'B'] = course_name
        df.at[2, 'B'] = assignment_name
        for d in name_grade_dict_list:
            for key, value in d.items():
                found_loc = find_name_location(key, df)
                if found_loc != 'not found':
                    df.at[found_loc, 'C'] = value
        bad_chars = ["/", "\\", ":", "*", "?", "\"", "<", ">", "|"]
        clean_assignment_name = ''.join(i for i in assignment_name if i not in bad_chars)
        output_file = os.path.join(save_path, clean_assignment_name + "_grade_template.csv")
        df.to_csv(output_file, header=False, index=False)
    except Exception as e:
        print("ERROR: failed to create CSV file: ", str(e), flush=True)


def student_lookup(user_id):
    """
    :param user_id:
    :return: users name: "familyName,  givenName"
    called by swap_id_for_name
    """
    # API Call
    user_info = service.userProfiles().get(userId=user_id).execute()
    # print("printing user info JSON: ", user_info)
    return user_info["name"]["familyName"] + ", " + user_info["name"]["givenName"]


def swap_student_id_for_student_name(list_of_dicts_in):
    """
    :param list_of_dicts_in: [{<userID>: <grade>}]
    :return: [{<familyName, givenName>: <grade>}]
    called by create_name_grade_dict_list(student_submissions)
    """
    return [{student_lookup(k): v} for d in list_of_dicts_in for k, v in d.items()]


def get_userId_grade(json_in, course_work_id):
    """
    :param course_work_id: the assignment we are getting names and grades for
    :param json_in: json object of student submissions, this is the return value of the api call to get course work
    :return: list of dictionaries: [{<userID>: <grade>}]
    """
    student_submissions = json_in["studentSubmissions"]
    return [{c["userId"]: c["assignedGrade"]} for c in student_submissions
            if "userId" in c
            and "assignedGrade" in c
            and c["courseWorkId"] == course_work_id]


def create_name_grade_dict_list(student_submissions, course_work_id):
    """
    :param course_work_id: the assignment we are getting names and grades for
    :param student_submissions: json result of an API call to studentSubmissions()
    :return: returns a list of dictionaries: [{<student name>: <grade>}]
    """
    try:
        id_grade = get_userId_grade(student_submissions, course_work_id)
        return swap_student_id_for_student_name(id_grade)
    except KeyError:
        return KeyError
    except Exception:
        return Exception


def selected_course_id(courses_json, selected_course):
    """
    :param courses_json: courses() api call return value
    :param selected_course: user selected course name (string value) from GUI
    :return: returns the course ID
    """
    for c in courses_json:
        if c["name"] == selected_course:
            return c["id"]


def get_all_assignments_for_course(student_submissions):
    """
    :param student_submissions: json result of an API call to studentSubmissions()
    :return: list of courseWorkId's (assignments)
    """
    list_of_assignments = [c['courseWorkId'] for c in student_submissions['studentSubmissions']
                           if "courseWorkId" in c]
    return list(set(list_of_assignments))


# def get_max_points_for_assignment(student_sumbissions):
#     # max_points = student_submissions["studentSubmissions"][1]["submissionHistory"][2]["gradeHistory"]["maxPoints"]
#     for i in student_sumbissions["studentSubmissions"]:
#         print("i:", i)
#         for k, v in i.items():
#             print("v: ", v)


@Gooey(program_name="Fetch Grades", )
def main():
    # API call to get classes
    """TODO: 1) look up students by name in template and update when found
                Goal is to preserve student ID's
    """

    # API call to get course json
    course_api_call_results = service.courses().list(pageSize=10).execute()

    # create list of course names to select from
    courses_json = course_api_call_results.get('courses')
    course_names = []
    for c in courses_json:
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

    # get user selected parameters
    selected_course = user_inputs['course_selection']
    grade_template = user_inputs['grade_template']
    save_path = user_inputs['output_directory']
    print("######################################################")
    print("\nSelected Course: ", selected_course, flush=True)

    # API call to get course work
    student_submissions = service.courses().courseWork().studentSubmissions().list(
        courseId=selected_course_id(courses_json, selected_course),
        courseWorkId='-').execute()
    course_id = selected_course_id(courses_json, selected_course)
    # cid_cwid_uid_ag = [(c['courseId'], c['courseWorkId'], c["userId"], c["assignedGrade"])
    #                    for c in student_submissions['studentSubmissions']
    #                    if "courseId" in c
    #                    and "courseWorkId" in c
    #                    and "userId" in c
    #                    and "assignedGrade" in c]

    # create import file for each assignment in selected course
    list_of_assignments = get_all_assignments_for_course(student_submissions)
    for a in list_of_assignments:
        assignment_name = assignment_lookup(course_id, a)
        name_grade_dict_list = create_name_grade_dict_list(student_submissions, a)
        print("Assignment: ", assignment_name, "  :  ", create_name_grade_dict_list(student_submissions, a), flush=True)
        create_import_file(grade_template, save_path, assignment_name, selected_course, name_grade_dict_list)


if __name__ == '__main__':
    main()
