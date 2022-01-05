#!/Users/william/.conda/envs/plant-bot/bin/python
import time
from time import gmtime, strftime
# from timeit import default_timer as timer
import sys
from datetime import datetime, date, timezone, timedelta
import yaml
import data_sheet as ds

# import os

today = datetime.now(timezone.utc)  # we'll need this here and perhaps later
localtime = datetime.today()
LOCAL_EMAIL_INFO_FILENAME = 'temp_email_info.yaml'
ts = time.time()
utc_offset = (datetime.fromtimestamp(ts) -
              datetime.utcfromtimestamp(ts)).total_seconds() #utcoffset in secs
partial_tz = utc_offset/3600 - int(utc_offset/3600) #minutes val for UTC±X30/45
utc_str = '+' if utc_offset>0 else '' + str(int(utc_offset/3600)) + ':' +\
    str(int(60 * partial_tz)) + '0' if partial_tz==0 else ''
print(f'\n{localtime.strftime("Local time: %I:%M%p, %Y/%m/%d")};',
      f'Time zone: UTC{utc_str}')

try:
    with open("config.yaml", 'r') as stream:

        settings = yaml.safe_load(stream)
        SHEET_TITLE = settings['spreadsheet_name']
        our_email = settings['from_email']
        target_email = settings['to_email']
        initiated = settings['initiated']
        ignore_means_yes = settings['ignore_means_yes']  # TODO: implement
        #TODO: default timezone option - rather than changing w/ location, always have one
except yaml.YAMLError as exc:
    print(exc)
except FileNotFoundError:
    # print('config.yaml missing. Run create_config.py to restore')
    # raise
    settings = {'from_email': 'your_email_here@gmail.com',
                'to_email': 'email_to_send_to@example.com',
                'ignore_means_yes': False,
                'spreadsheet_name': 'Plant Task Data',
                'initiated': False}
    with open('config.yaml', 'w') as f:
        yaml.safe_dump(settings, f, default_flow_style=False, sort_keys=False)
except KeyError:
    print('config.yaml missing data')
    raise

""" # only if we're running >1 times per minute
if os.path.isfile('request_counter.yaml'):
    with open("request_counter.yaml", 'r') as stream:
        try:
            # make sure we don't exceed the google sheets quota
            counts = yaml.safe_load(stream)
            read_request_counter = counts['read_requests']
            write_request_counter = counts['write_requests']
        except yaml.YAMLError as exc:
            print(exc)
        except KeyError:
            print('request_counter.yaml missing data')
            raise
else:
    read_request_counter = 0
    write_request_counter = 0
    counts = {'read_requests': read_request_counter,
              'write_requests': write_request_counter}
    with open('request_counter.yaml', 'w') as f:
        yaml.safe_dump(counts, f, default_flow_style=False, sort_keys=False)

if read_request_counter >= 50 or write_request_counter >= 50:
    recharge_quota(60)
    read_request_counter = 0"""
read_request_counter = 0
write_request_counter = 0

# open local save - reduces sheets api usage
def read_local_info():
    try:
        with open(LOCAL_EMAIL_INFO_FILENAME, 'r') as stream:
            data = yaml.safe_load(stream)
            last_email_info = {
                'date': data['date'], 'id': data['id'],
                'got_reply': data['got_reply'], 'email_type': data['email_type'],
                'task_data': data['task_data']
            }
            # 1) Decide what to do based on info
            need_email = True
            need_sheet = True
            temp_missing = False
            if last_email_info['date'] is None:  # no date given
                need_email = True
                need_sheet = True
            else:
                last_email_date_iso = last_email_info['date'].strip()
                if (today - datetime.fromisoformat(last_email_date_iso)).days > 0:  # has 1 day passed?
                    need_email = True
                    need_sheet = True
                else:
                    need_sheet = False
                    #need_sheet = True
                    if isinstance(last_email_info['got_reply'], str):
                        last_email_str = last_email_info['got_reply'].strip().lower()
                        if last_email_str == 'true':
                            need_email = False  # nothing left to do today
                    elif last_email_info['got_reply']:
                        need_email = False  # nothing left to do today
                    else:
                        need_email = True  # have to check for replies
    except yaml.YAMLError as exc:
        print(f"Error while parsing {LOCAL_EMAIL_INFO_FILENAME}:")
        if hasattr(exc, 'problem_mark'):
            if exc.context != None:
                print('  parser says\n' + str(exc.problem_mark) + '\n  ' +
                      str(exc.problem) + ' ' + str(exc.context) +
                      '\nPlease correct data and retry.')
            else:
                print('  parser says\n' + str(exc.problem_mark) + '\n  ' +
                      str(exc.problem) + '\nPlease correct data and retry.')
        else:
            print("Something went wrong while parsing yaml file")
        raise
    except KeyError:
        print(f'{LOCAL_EMAIL_INFO_FILENAME} missing data')
        raise
    except FileNotFoundError:  # not yet created, we'll make after getting spreadsheet
        need_sheet = True
        need_email = True
        temp_missing = True
    return {'need_sheet': need_sheet, 'need_email': need_email,
            'temp_missing': temp_missing, 'email_info': last_email_info}

local_info = read_local_info()
last_email_info = local_info['email_info'] #easier to read
if local_info['need_sheet']:
#if local_info['need_email']:
    data_sheet = ds.DataSheet(SHEET_TITLE)

    if not initiated:
        # TODO： Wouldn't it be better to get the email from their authentication selection?
        settings['from_email'] = input(
            'Please enter your email which will be used to send emails and create spreadsheets: ')
        settings['to_email'] = input('Please enter the email you would like to send emails to: ')
        # TODO: check if spreadsheet already exists
        #sh = data_sheet.create_spreadsheet()
        data_sheet.create_spreadsheet()
        #wk = sh.get_worksheet(0)
        data_sheet.set_up_spreadsheet()
        print(f'Spreadsheet created with name {SHEET_TITLE}' +
              '\nPlease fill out your plant information before running the program again')
        settings['initiated'] = True
        with open('config.yaml', 'w') as f:
            yaml.safe_dump(settings, f,
                           default_flow_style=False, sort_keys=False)
        sys.exit(0)
    else:
        #sh = data_sheet.open_spreadsheet()
        data_sheet.open_spreadsheet()
        #wk = sh.get_worksheet(0)
        if local_info['temp_missing']:
            last_email_info = data_sheet.get_email_info()
            read_request_counter += 5
            with open(LOCAL_EMAIL_INFO_FILENAME, 'w') as f:  # store local info
                yaml.safe_dump(last_email_info, f,
                               default_flow_style=False, sort_keys=False)
            # now try analyzing the email_info again
            local_info = read_local_info()

need_sheet = local_info['need_sheet'] #easier to read

# case 1: last spreadsheet check <24 hrs & already got reply
if not local_info['need_email']:
    print('Nothing left to do today.')
    sys.exit(0)

import email_sender
import email_reader
#todo: right now these imports automatically oauth, but we should turn them to classes

# declare remaining vars
service = email_sender.service

today_in_tz = datetime.now()
#need_update_email_info = True  # relevant at very end

#in every email; lets us differentiate btw my email & the response in the snippet
HEADER = 'Here is what you need to do for each plant:'
HEADER1 = "I didn't get a proper response yesterday"


def recharge_quota(secs=60):  # sleeps and prints a progress bar
    print('Pausing to recharge Google Sheets quota')
    print('Progress: ', end='')
    interval = secs / 10
    for i in range(10):
        time.sleep(interval)
        print('-', end='')
    print('')

#TODO: move the functions that require data_sheet & email to their own class
#sends an email in a new thread, updates ds
def send_email(data_sheet, subject, email):
    # TODO: except non-existent email
    msg = email_sender.send_message_easy(service, target_email,
                                         subject, email)
    data_sheet.store_email_id(msg['id'])
    data_sheet.store_got_reply(False)
    # stores id when we send, and resets "got reply?" to false
    # print(msg['id'])

# replies to target's reply, updates ds
def send_reply(data_sheet, prev_reply, body):
    email_sender.reply_message(service, prev_reply['threadId'],
                               target_email, prev_reply['message_id'],
                               prev_reply['references'], prev_reply['subject'],
                               body)
    print('Responded to reply')
    data_sheet.store_got_reply(True)

# helper class for find_tasks()
def _parse_every_days_entry(task_every, c):
    sync = False
    in_season = True
    try:
        every_days = int(task_every)
    except ValueError:
        days_str = task_every  # the string that will be manipulated to get a number
        # TODO: should letters and other superfluous characters be removed?
        # TODO: also allow for fully typing out "sync"?
        # TODO: add mutually exclusive option
        if task_every[len(task_every) - 1] == 's' and c >= 6:  # sync & not task 1
            sync = True
            days_str = days_str.replace('s', '')
        if '[' in task_every and ']' in task_every:  # if month range given, check current month
            months = task_every.split('[')[1].split(']')[0].split('-')
            months = [int(months[0]), int(months[1])]
            if months[0] < months[1]:
                if today.month < months[0] or today.month > months[1]:
                    in_season = False
            else:
                if months[0] > today.month > months[1]:
                    in_season = False
            days_str = days_str.split('[')[0]
        every_days = int(days_str)
        # if it still raises exception, will be handled by find_tasks()
    return {'days': every_days, 'sync': sync, 'in_season': in_season}

#Testing purposes
def print_cells(data_sheet):
    print("printing cells")
    r = 1
    c = 2
    global read_request_counter
    data = data_sheet.get_cell(r, c)
    read_request_counter += 1
    while data is not None:
        if read_request_counter >= 55:
            recharge_quota(60)
            read_request_counter = 0
        print(f"Row {r}:")
        print("-----------")
        while data is not None:
            print(data)
            c += 1
            data = data_sheet.get_cell(r, c)
            read_request_counter += 1
        r += 1
        c = 1
        data = data_sheet.get_cell(r, c)
        read_request_counter += 1

# find what tasks need to be done today and returns coords
def find_tasks(data_sheet):
    task_coords = []

    r = 2
    c = 1
    task_counter = 0
    global read_request_counter
    while data_sheet.get_cell(r, c) is not None:  # go thru all plants until first blank line
        read_request_counter += 1
        c += 1
        while data_sheet.get_cell(r, c) is not None:  # all task names until first blank
            if read_request_counter >= 45:
                recharge_quota(60 - (task_counter / 4))
                read_request_counter = 0
                task_counter = 0
            # start = timer()
            try:  # TODO: long form, 1 task per row, sync w/ row # of task
                c += 1
                task_every = data_sheet.get_cell(r, c).strip()  # every __ days
                parsed_every = _parse_every_days_entry(task_every, c)
                if not parsed_every['in_season']:
                    c += 2
                    continue

                c += 1
                raw_date = data_sheet.get_cell(r, c).strip()  # date of last time
                #print(f'Date value in cell [{r},{c}]: {raw_date}')
                if '-' in raw_date:
                    ymd = raw_date.split('-')
                elif '/' in raw_date:
                    ymd = raw_date.split('/')
                else:
                    raise ValueError('Date improperly formatted')
                if len(ymd) == 2:
                    # TODO: add other things to fix user-inputted dates
                    # [yyyy-]mm-dd
                    # crossing 2 years (4 month buffer for really infrequent tasks)
                    if today.month < 3  and int(ymd[0]) > 9:
                        ld_year = today.year - 1 #last date year
                    else:
                        ld_year = today.year
                    last_date = date(ld_year, int(ymd[0]), int(ymd[1]))
                    # TODO: change spreadsheet dates to local tz
                    #rn they're utc b/c the epoch reply time we record is utc
                elif len(ymd) > 2:
                    if len(ymd[2]) == 4:  # mm-dd-yyyy
                        last_date = date(int(ymd[2]), int(ymd[0]), int(ymd[1]))
                        # TODO: make dd-mm-yyyy an option?
                    else: # yyyy-mm-dd
                        last_date = date(int(ymd[0]), int(ymd[1]), int(ymd[2]))
                else:
                    raise ValueError('Date improperly formatted')
                # has it been enough days since the last time task was done?
                today_date = date(today.year, today.month, today.day)
                if (today_date - last_date).days >= parsed_every['days']:
                    if not parsed_every['sync']:
                        task_coords.append((r, c - 2))
                        print(f'Adding {data_sheet.get_cell(r, 1)}: ' +
                              f'{data_sheet.get_cell(r, c - 2)} to task list')
                    else:
                        # if sync option is chosen, task 1 of this plant must
                        # also be happening today for the synced task to happen
                        if (r, 2) in task_coords:
                            task_coords.append((r, c - 2))
                            print(f'Adding {data_sheet.get_cell(r, c-2)}' +
                            f', {data_sheet.get_cell(r, 1)} to task list')

                c += 1
            except:
                print(f'Spreadsheet is not filled out correctly at cell {chr(c + 64)}{r}')
                raise
            read_request_counter += 5
            # print('Read requests: ', read_request_counter)
            task_counter += 1
            # end = timer()
            # print(f'Time for task: {end - start}') #approx 4 tasks per second

        r += 1
        c = 1

    return task_coords


# helper method, creates the task list part of email
def _create_task_list(email, task_coords):
    r = 0
    global read_request_counter
    for coord in task_coords:
        # start = timer()
        if coord[0] != r:  # since it's already ordered by row,
            r = coord[0]  # tasks can be organized by plant
            email = '\n'.join([email, f'{data_sheet.get_cell(r, 1)}: '])  # plant name
            read_request_counter += 1
        email = ''.join([email, f'{data_sheet.get_cell(r, coord[1])}, '])  # task name
        read_request_counter += 1
        # end = timer()
        # print(f'Time for coord: {end - start}') #approx 8 coords per sec
        if read_request_counter >= 54:
            recharge_quota(60)
            read_request_counter = 0

    return email

# build email from non-empty, already sorted list of coords; returns that email
def create_email(task_coords):
    email = f'''{HEADER} \n'''
    email = _create_task_list(email, task_coords)

    ending = '\nIf you complete all of these tasks, reply with "y". ' + \
             'If you complete none of them, reply with "n". If you complete some of them, ' + \
             'reply with "y except ", followed by the names of the plants for which ' + \
             'you did not complete the tasks, spelled as above and ' + \
             'separated by commas.'

    email = '\n'.join([email, ending])
    email = email.replace(', \n', '\n')  # get rid of last comma
    return email


# repeat yesterday's unanswered message, modified things to confirm; returns email
def create_ignored_email(task_coords):
    yesterday = today_in_tz - timedelta(days=1)
    email = \
        f'''{HEADER1} ({yesterday.strftime("%A, %B %d %Y")}). \
\n\nIf you complete or have completed yesterday's tasks, please reply with "y", followed by \
"today" or "yesterday" depending on when you did them. If you did only some of them, \
add "except" and then list the names of the plants whose tasks you did not complete, \
separated by commas. If you don't plan on \
doing them today, please reply with "n".
Here is a reminder of yesterday's tasks:\n'''
    email = _create_task_list(email, task_coords)

    ending = '\nThank you.'

    email = '\n'.join([email, ending])
    email = email.replace(', \n', '\n')  # get rid of last comma
    return email


# check for a reply; returns boolean - is or isn't? + relevant info
def check_reply():
    # check for a reply; only act if there is one
    if email_reader.has_reply(last_email_info['id']):
        reply = email_reader.get_reply(last_email_info['id'])  # dict
        rpl = reply['snippet'].lower()  # user's reply message
        reply_date = datetime.utcfromtimestamp(int(reply['internalDate'][0:-3]))
        return {'is_reply': True, 'rpl_snippet': rpl, 'reply': reply,
                'reply_date': reply_date,
                'task_coords': eval(last_email_info['task_data'])}
    else:
        return {'is_reply': False}


# helper class that parses and responds to user's reply
def handle_reply(data_sheet, reply, reply_date, task_coords):
    rpl = reply['snippet'].lower()  # user's reply message
    first_word = rpl.strip()[0]
    print(f'Reply received, tasks completed at {reply_date.strftime("%m/%d/%Y %I:%M%p")} UTC')
    if 'except' not in rpl:
        if first_word == 'y' or first_word[0:1] == 'ye' \
                or first_word == 'yup':

            # update wk w/ the date the user sent the reply
            update_tasks(data_sheet, task_coords, reply_date)
            send_reply(data_sheet, reply, 'Thank you, dates have been updated.')

        elif first_word == 'n' or first_word[0:1] == 'no':
            send_reply(data_sheet, reply,
                       'Thank you, another reminder will be sent tomorrow.')
            # nothing needs updating
        # else: incoherent message counts as ignore
    else:
        if HEADER.lower() in rpl:
            except_plants = rpl.split(HEADER.lower())[0]. \
                split('except')[1].strip().split(',')
        else:
            except_plants = rpl.split(HEADER1.lower())[0]. \
                split('except')[1].strip().split(',')
        b = except_plants
        # remove spaces in user inputted list
        except_plants[:] = [p.strip() for p in except_plants]
        print(b)
        update_tasks(data_sheet, task_coords, reply_date,
                     b)
        send_reply(data_sheet, reply, 'Thank you, dates have been updated.')

# updates dates of tasks in spreadsheet
def update_tasks(data_sheet, task_coords, reply_date, except_list=None):
    if except_list is None:
        except_list = ['']  # no plants to skip
    global read_request_counter
    global write_request_counter
    for coord in task_coords:
        # start = timer()
        skip_plant = False
        plant_name = data_sheet.get_cell(coord[0], 1)
        for entry in except_list:
            if plant_name.lower().strip() in entry:
                skip_plant = True
                print(f'Skipping {plant_name}')
                read_request_counter += 2
                break
        if not skip_plant:
            data_sheet.update_cell(coord[0], coord[1] + 2,
                                   reply_date.strftime("%m/%d/%Y"))
            print(f'Updating {plant_name}')
            read_request_counter += 1
            write_request_counter += 1
        # end = timer()
        # print(f'Time for update coord: {end - start}')
        # approx 2.5 updates per sec
        if read_request_counter >= 54:
            recharge_quota(60)
            read_request_counter = 0
        elif write_request_counter >= 54:
            recharge_quota(60)
            write_request_counter = 0
    print('Spreadsheet updated')

# The ordinary daily task list; stores data_sheet info
#todo: all these data_sheet stores need to go in their own function
# with all variables optional, and only storing if var [is given]:
def send_normal_email(data_sheet, task_coords):
    if len(task_coords) > 0:
        email = create_email(task_coords)
        subject = f'''{today_in_tz.strftime("%A, %B %d %Y")}'s plant tasks'''
        send_email(data_sheet, subject, email)
        # print(email)
        print('Email sent\n')
    else:  # if no tasks, no need for reply; just wait til tomorrow
        data_sheet.store_got_reply(True)
        print('No tasks to do, no email sent\n')

    # store email info
    data_sheet.store_email_date(datetime.now(timezone.utc)
                                .isoformat())
    data_sheet.store_task_data(task_coords)
    data_sheet.store_email_type('Normal')


def send_ignored_email(data_sheet, task_coords):
    if len(task_coords) > 0:
        email = create_ignored_email(task_coords)
        subject = f'''Plant task check-in: {today_in_tz.strftime("%A, %B %d %Y")}'''
        send_email(data_sheet, subject, email)
        # print(email)
        # store email info
        print('Check-in email sent')
        data_sheet.store_email_date(datetime.now(timezone.utc)
                                    .isoformat())
        data_sheet.store_email_type('Check-in')

def update_local_email_info(data_sheet):
    global read_request_counter
    global write_request_counter
    if read_request_counter >= 50 or write_request_counter >= 50:
        recharge_quota(60)
    last_email_info = data_sheet.get_email_info()
        # get email info from spreadsheet in case it's been changed
    read_request_counter += 5
    with open(LOCAL_EMAIL_INFO_FILENAME, 'w') as f:  # store local info
        yaml.safe_dump(last_email_info, f, default_flow_style=False, sort_keys=False)
    print('Local email info yaml updated')

def main():
    try:
        # 1) Gather info of last email

        """global read_request_counter
        if read_request_counter >= 50 or write_request_counter >= 50:
            recharge_quota(60)
            read_request_counter = 0


        if last_email_info['date'] is None:
            need_email = True
        else:
            last_email_date_iso = last_email_info['date'].strip()
            if last_email_date_iso.lower() == 'n/a':
                need_email = True
            else:
                rn = datetime.now(timezone.utc)  # time currently

                if (rn - datetime.fromisoformat(last_email_date_iso)).days > 0:  # has 1 day passed?
                    need_email = True
                else:
                    need_email = False"""
        # above is mainly stuff from testing

        # handling nones
        if last_email_info['got_reply'] is None:
            last_email_replied = True
        if last_email_info['email_type'] is None:
            last_email_type = 'normal'
        else:
            last_email_type = last_email_info['email_type'].strip().lower()

        # if got_reply in form of string (happens when directly from sheets)
        if isinstance(last_email_info['got_reply'], str):
            last_email_str = last_email_info['got_reply'].strip().lower()
            if last_email_str == 'true':
                last_email_replied = True
            elif last_email_str == 'false':
                last_email_replied = False
            else:
                raise ValueError('"Got Reply?" email info in improper format')

    except:
        print('Error reading email info')
        raise

    # case 2: last email <24 hrs, listen for response
    if not need_sheet:

        reply_info = check_reply()

        if reply_info['is_reply']:
            data_sheet1 = ds.DataSheet(SHEET_TITLE)  # instantiate since we haven't yet b/c need_sheet = False
            data_sheet1.open_spreadsheet()
            # reply to normal emails
            if last_email_type == 'normal':
                handle_reply(data_sheet1, reply_info['reply'],
                             reply_info['reply_date'], reply_info['task_coords'])

            # reply to check-in emails
            elif last_email_type == 'check-in':
                if 'yesterday' in reply_info['rpl_snippet']:
                    reply_date = reply_info['reply_date'] - timedelta(days=1)
                else:
                    reply_date = reply_info['reply_date']
                handle_reply(data_sheet1, reply_info['reply'], reply_date,
                             reply_info['task_coords'])
            else:
                raise ValueError('"Email Type" email data in improper format')

            update_local_email_info(data_sheet1)
        else:
            print('No reply received yet')
            #global need_update_email_info
            #need_update_email_info = False
            # do nothing if no reply, since not yet 24 hrs

    # case 3:
    # a. Email sent >24 hrs ago (or not yet), and it's just a list (normal type)
    elif last_email_type == 'normal':

        # option 1: reply is either settled or this is the first message
        global data_sheet
        if last_email_replied:
            task_coords = find_tasks(data_sheet)
            send_normal_email(data_sheet, task_coords)
            # send today's tasks, or nothing if none

        # option 2: reply still not received
        else:  # check one last time, then send ignored if none
            reply_info = check_reply()
            if reply_info['is_reply']:
                handle_reply(data_sheet, reply_info['reply'],
                             reply_info['reply_date'], reply_info['task_coords'])
                task_coords = find_tasks(data_sheet)
                send_normal_email(data_sheet, task_coords)
            else:
                print('Still no reply received, sending a check-in')
                send_ignored_email(data_sheet,eval(last_email_info['task_data']))

    # b. Email sent >24 hrs ago, and was a check-in in response to ignore
    elif last_email_type == 'check-in':
        # regardless of whether we got a response or not,
        # we're going back to our normal schedule
        reply_info = check_reply()
        if reply_info['is_reply']:
            if 'yesterday' in reply_info['rpl_snippet']:
                reply_date = reply_info['reply_date'] - timedelta(days=1)
            else:
                reply_date = reply_info['reply_date']
            handle_reply(data_sheet, reply_info['reply'], reply_date,
                         reply_info['task_coords'])
        task_coords = find_tasks(data_sheet)
        send_normal_email(data_sheet, task_coords)

    else:
        raise ValueError('"Email Type" email info in improper format')
    # last_email_info = data_sheet.get_email_info()
    # print(email_reader.has_reply(last_email_info['id']))


try:
    #print_cells(data_sheet)
    main()
    if need_sheet:
        update_local_email_info(data_sheet)
except:
    print('Something went wrong with the spreadsheet or email')
    raise


# only un-string the below script if you're running this more than once per minute
"""counts = {'read_requests': read_request_counter,
          'write_requests': write_request_counter}
with open('request_counter.yaml', 'w') as f:
    yaml.safe_dump(counts, f, default_flow_style=False, sort_keys=False)
"""
# email_reader.threads_head(5, True)
# email_reader.get_reply('')

# last_email_info = data_sheet.get_email_info()
# print(email_reader.has_reply(last_email_info['id']))


# print(data_sheet.get_cell(4, 19))
