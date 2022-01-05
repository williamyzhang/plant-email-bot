import gspread
import google.auth.exceptions
#import email_sender
import os

#authorized_user_filename='sheet_credentials.json'

#our_email = email_sender.our_email

N_TASKS = 5
EI = N_TASKS * 3 + 4 # column number of email info
home = os.path.expanduser("~")
#TODO: how will you ensure the user has gspread & credentials in the right directory?
GSPREAD_AUTH_PATH = os.path.join(home, ".config", "gspread",
                                 "authorized_user.json")

class DataSheet:
    def __init__(self, title):
        self.gc = gspread.oauth()
        self.title = title

    def create_spreadsheet(self):
        """Create the spreadsheet file

                :param str title: title of the spreadsheet

                :rtype: a spreadsheet
            """
        #global gc
        try:
            #return gc.create(title)
            self.wk = self.gc.create(self.title).get_worksheet(0)
        except google.auth.exceptions.RefreshError:
            os.remove(GSPREAD_AUTH_PATH)
            self.gc = gspread.oauth()
            self.wk = self.gc.create(self.title).get_worksheet(0)


    def open_spreadsheet(self):
        #global gc
        try:
            #return gc.open(title)
            self.wk = self.gc.open(self.title).get_worksheet(0)
        except gspread.exceptions.SpreadsheetNotFound:
            print(f'You do not have a spreadsheet titled "{self.title}"')
            raise
        except google.auth.exceptions.RefreshError:
            os.remove(GSPREAD_AUTH_PATH)
            self.gc = gspread.oauth()
            self.wk = self.gc.open(self.title).get_worksheet(0)


    def set_up_spreadsheet(self):

        c = 2
        # headers
        for i in range(1, N_TASKS + 1):
            # TODO: add colors
            self.wk.update_cell(1, c, f'Task {i} Name')
            c += 1
            self.wk.update_cell(1, c, f'Do Task {i} Every __ Days')
            c += 1
            self.wk.update_cell(1, c, f'Date of Last Task {i}')
            c += 1

        # email info cells
        self.wk.update_cell(1, EI - 1, 'Info of Last Email')
        email_info = [('Date', '2021-07-22T17:53+00:00'), ('ID', 'N/A'),
                      ('Got reply?', 'N/A'), ('Email Type', 'N/A'),
                      ('Task Data', 'N/A')]
        r = 2
        for tup in email_info:
            self.wk.update_cell(r, EI - 1, tup[0])
            self.wk.update_cell(r, EI, tup[1])
            r += 1

        # example plant to demonstrate; TODO: change dates to correspond to today parameter
        example = ['Example Plant', 'Water', '4', '7-21',
                   'Fertilize', '28[5-8]s', '7-21', 'Turn pot', '7', '7-18']
        c = 1
        for ex in example:
            self.wk.update_cell(2, c, ex)
            c += 1

    #def fix_spreadsheet_info(): try to correct user input, run before determining tasks
    #detect if spreadsheet is messed up (ex: task headers not correct) -> raise exception


    def store_dict_test(self):
        test_dict = {'Date': '7-21', 'ID': 56, 'Replied': False}
        self.wk.update_cell(2, EI, str(test_dict))

    def get_cell(self, r, c):
        return self.wk.cell(r, c).value

    def update_cell(self, r, c, val):
        self.wk.update_cell(r, c, val)

    def store_email_id(self, id): #store after each send
        self.wk.update_cell(3, EI, id)

    def store_email_date(self, date):
        self.wk.update_cell(2, EI, date)

    def store_got_reply(self, got_reply): #store the true/false
        self.wk.update_cell(4, EI, got_reply)

    def store_email_type(self, type):
        self.wk.update_cell(5, EI, type)

    def store_task_data(self, task_coords): # store list of tuples as string
        self.wk.update_cell(6, EI, str(task_coords))

    def get_email_info(self):
        return {'date': self.wk.cell(2, EI).value, 'id': self.wk.cell(3, EI).value,
                'got_reply': self.wk.cell(4, EI).value, 'email_type': self.wk.cell(5, EI).value,
                'task_data': self.wk.cell(6, EI).value}


