import gspread
import email_sender

#authorized_user_filename='sheet_credentials.json'
gc = gspread.oauth()
#our_email = email_sender.our_email

#tasks = ['Water', 'Fertilize', 'Other 1', 'Other 2']
#EI = len(tasks) * 2 + 1 # column number of email info
N_TASKS = 5
EI = N_TASKS * 3 + 4 # column number of email info

def create_spreadsheet(title):
    """Create the spreadsheet file

            :param str title: title of the spreadsheet

            :rtype: a spreadsheet
        """
    return gc.create(title)

def open_spreadsheet(title):
    try:
        return gc.open(title)
    except gspread.exceptions.SpreadsheetNotFound:
        print(f'You do not have a spreadsheet titled "{title}"')
        raise

def set_up_spreadsheet(wk):

    c = 2
    # headers
    for i in range(1, N_TASKS + 1):
        # TODO: add colors
        wk.update_cell(1, c, f'Task {i} Name')
        c += 1
        wk.update_cell(1, c, f'Do Task {i} Every __ Days')
        c += 1
        wk.update_cell(1, c, f'Date of Last Task {i}')
        c += 1

    # email info cells
    wk.update_cell(1, EI - 1, 'Info of Last Email')
    email_info = [('Date', '2021-07-22T17:53+00:00'), ('ID', 'N/A'),
                  ('Got reply?', 'N/A'), ('Email Type', 'N/A'),
                  ('Task Data', 'N/A')]
    r = 2
    for tup in email_info:
        wk.update_cell(r, EI - 1, tup[0])
        wk.update_cell(r, EI, tup[1])
        r += 1

    # example plant to demonstrate; TODO: change dates to correspond to today parameter
    example = ['Example Plant', 'Water', '4', '7-21',
               'Fertilize', '28[5-8]s', '7-21', 'Turn pot', '7', '7-18']
    c = 1
    for ex in example:
        wk.update_cell(2, c, ex)
        c += 1

#def fix_spreadsheet_info(): try to correct user input, run before determining tasks
#detect if spreadsheet is messed up (ex: task headers not correct) -> raise exception


def store_dict_test(wk):
    test_dict = {'Date': '7-21', 'ID': 56, 'Replied': False}
    wk.update_cell(2, EI, str(test_dict))

def get_cell(wk, r, c):
    return wk.cell(r, c).value

def update_cell(wk, r, c, val):
    wk.update_cell(r, c, val)

def store_email_id(wk, id): #store after each send
    wk.update_cell(3, EI, id)

def store_email_date(wk, date):
    wk.update_cell(2, EI, date)

def store_got_reply(wk, got_reply): #store the true/false
    wk.update_cell(4, EI, got_reply)

def store_email_type(wk, type):
    wk.update_cell(5, EI, type)

def store_task_data(wk, task_coords): # store list of tuples as string
    wk.update_cell(6, EI, str(task_coords))

def get_email_info(wk):
    return {'date': wk.cell(2, EI).value, 'id': wk.cell(3, EI).value,
            'got_reply': wk.cell(4, EI).value, 'email_type': wk.cell(5, EI).value,
            'task_data': wk.cell(6, EI).value}


