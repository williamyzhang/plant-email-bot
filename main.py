#!/Users/william/.conda/envs/plant-bot/bin/python
import time
from timeit import default_timer as timer
import sys
from datetime import datetime, date, timezone, timedelta
import yaml

with open("config.yaml", 'r') as stream:
    try:
        settings = yaml.safe_load(stream)
        SHEET_TITLE = settings['spreadsheet_name']
        our_email = settings['from_email']
        target_email = settings['to_email']
        initiated = settings['initiated']
        ignore_means_yes = settings['ignore_means_yes'] # TODO: implement
    except yaml.YAMLError as exc:
        print(exc)
    except FileNotFoundError:
        print('config.yaml missing. Run create_config.py to restore')
        raise
    except KeyError:
        print('config.yaml missing data')
        raise

# Begin script
import data_sheet

if not initiated:
    settings['from_email'] = input(
        'Please enter your email which will be used to send emails and create spreadsheets: ')
    settings['to_email'] = input('Please enter the email you would like to send emails to: ')
    # TODO: check if spreadsheet already exists
    sh = data_sheet.create_spreadsheet(SHEET_TITLE)
    wk = sh.get_worksheet(0)
    data_sheet.set_up_spreadsheet(wk)
    print(f'Spreadsheet created with name {SHEET_TITLE}' +
          '\nPlease fill out your plant information before running the program again')
    settings['initiated'] = True
    with open('config.yaml', 'w') as f:
        yaml.safe_dump(settings, f, default_flow_style=False, sort_keys=False)
    sys.exit(0)
else:
    sh = data_sheet.open_spreadsheet(SHEET_TITLE)
    wk = sh.get_worksheet(0)

import email_sender
import email_reader

service = email_sender.service
today = datetime.now(timezone.utc)
today_in_tz = datetime.now()
HEADER = 'Here is what you need to do for each plant:'
HEADER1 = "Here is a reminder of yesterday's tasks:"

read_request_counter = 0 #make sure we don't exceed the google sheets quota
write_request_counter = 0

def send_email(subject, email):
    # TODO: except non-existent email
    msg = email_sender.send_message_easy(service, target_email,
                                         subject, email)
    data_sheet.store_email_id(wk, msg['id'])
    data_sheet.store_got_reply(wk, False)
    # stores id when we send, and resets "got reply?" to false
    # print(msg['id'])


def send_reply(prev_reply, body):  # replies to target's reply
    email_sender.reply_message(service, prev_reply['threadId'],
                               target_email, prev_reply['message_id'],
                               prev_reply['references'], prev_reply['subject'],
                               body)
    data_sheet.store_got_reply(wk, True)


def _parse_every_days_entry(task_every, c):  # helper class for find_tasks()
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


def find_tasks():  # find what tasks need to be done today
    task_coords = []

    r = 2
    c = 1
    task_counter = 0
    global read_request_counter
    while data_sheet.get_cell(wk, r, c) is not None:  # go thru all plants until first blank line
        read_request_counter += 1
        c += 1
        while data_sheet.get_cell(wk, r, c) is not None:  # all task names until first blank
            #start = timer()
            try:
                c += 1
                task_every = data_sheet.get_cell(wk, r, c).strip()  # every __ days
                parsed_every = _parse_every_days_entry(task_every, c)
                if not parsed_every['in_season']:
                    c += 2
                    continue

                c += 1
                raw_date = data_sheet.get_cell(wk, r, c).strip()  # date of last time
                if '-' in raw_date:
                    ymd = raw_date.split('-')
                elif '/' in raw_date:
                    ymd = raw_date.split('/')
                else:
                    raise ValueError('Date improperly formatted')
                if len(ymd) == 2:
                    # TODO: add other things to fix user-inputted dates
                    # yyyy-mm-dd
                    last_date = date(today.year, int(ymd[0]), int(ymd[1]))
                elif len(ymd) > 2:
                    if len(ymd[2]) == 4:  # mm-dd-yyyy
                        last_date = date(int(ymd[2]), int(ymd[0]), int(ymd[1]))
                        # TODO: make dd-mm-yyyy an option?
                    else:
                        last_date = date(int(ymd[0]), int(ymd[1]), int(ymd[2]))
                else:
                    raise ValueError('Date improperly formatted')
                # has it been enough days since the last time task was done?
                today_date = date(today.year, today.month, today.day)
                if (today_date - last_date).days >= parsed_every['days']:
                    if not parsed_every['sync']:
                        task_coords.append((r, c - 2))
                    else:
                        # if sync option is chosen, task 1 of this plant must
                        # also be happening today for the synced task to happen
                        if (r, 2) in task_coords:
                            task_coords.append((r, c - 2))

                c += 1
            except:
                print(f'Spreadsheet is not filled out correctly at cell {chr(c + 64)}{r}')
                raise
            read_request_counter += 3
            #print('Read requests: ', read_request_counter)
            task_counter += 1
            #end = timer()
            #print(f'Time for task: {end - start}') #approx 4 tasks per second
            if read_request_counter >= 45:
                recharge_quota(60 - (task_counter / 4))
                read_request_counter = 0
                task_counter = 0
        r += 1
        c = 1

    return task_coords

def recharge_quota(secs): # sleeps and prints a progress bar
    print('Pausing to recharge Google Sheets quota')
    print('Progress: ', end='')
    interval = secs / 10
    for i in range(10):
        time.sleep(interval)
        print('-', end='')
    print('')

#helper method, creates the task list part of email
def create_task_list(email, task_coords):
    r = 0
    global read_request_counter
    for coord in task_coords:
        #start = timer()
        if coord[0] != r:  # since it's already ordered by row,
            r = coord[0]  # tasks can be organized by plant
            email = '\n'.join([email, f'{data_sheet.get_cell(wk, r, 1)}: '])  # plant name
            read_request_counter += 1
        email = ''.join([email, f'{data_sheet.get_cell(wk, r, coord[1])}, '])  # task name
        read_request_counter += 1
        #end = timer()
        #print(f'Time for coord: {end - start}') #approx 8 coords per sec
        if read_request_counter >= 54:
            recharge_quota(60)
            read_request_counter = 0

    return email

def create_email(task_coords):  # build email from non-empty, already sorted list of coords
    email = f'''{HEADER} \n'''
    email = create_task_list(email, task_coords)

    ending = '\nIf you complete all of these tasks, reply with "y". ' + \
             'If you complete none of them, reply with "n". If you complete some of them, ' + \
             'reply with "y except ", followed by the names of the plants for which ' + \
             'you did not complete the tasks, spelled as above and ' + \
             'separated by commas.'

    email = '\n'.join([email, ending])
    email = email.replace(', \n', '\n')  # get rid of last comma
    return email


# repeat yesterday's unanswered message, modified things to confirm
def create_ignored_email(task_coords):
    yesterday = today - timedelta(days=1)
    email = \
        f'''I didn't get a proper response yesterday ({yesterday.strftime("%A, %B %d %Y")}). \
\n\nIf you complete or have completed yesterday's tasks, please reply with "y", followed by \
"today" or "yesterday" depending on when you did them. If you did only some of them, \
add "except" and then list the names of the plants whose tasks you did not complete, \
separated by commas. If you don't plan on \
doing them today, please reply with "n".
{HEADER1} \n'''
    email = create_task_list(email, task_coords)

    ending = '\nThank you.'

    email = '\n'.join([email, ending])
    email = email.replace(', \n', '\n')  # get rid of last comma
    return email

# helper class that parses and responds to user's reply
def handle_reply(reply, reply_date, task_coords):
    rpl = reply['snippet'].lower()  # user's reply message
    first_word = rpl.strip()[0]
    print('Reply received')
    if 'except' not in rpl:
        if first_word == 'y' or first_word[0:1] == 'ye' \
                or first_word == 'yup':

            # update wk w/ the date the user sent the reply
            update_tasks(task_coords, reply_date)
            send_reply(reply, 'Thank you, dates have been updated.')

        elif first_word == 'n' or first_word[0:1] == 'no':
            send_reply(reply,
                       'Thank you, another reminder will be sent tomorrow.')
            # nothing needs updating
        # else: incoherent message counts as ignore
    else:
        if HEADER.lower() in rpl:
            except_plants = rpl.split(HEADER.lower())[0].\
            split('except')[1].strip().split(',')
        else:
            except_plants = rpl.split(HEADER1.lower())[0]. \
                split('except')[1].strip().split(',')
        b = except_plants
        # remove spaces in user inputted list
        except_plants[:] = [p.strip() for p in except_plants]
        print(b)
        update_tasks(task_coords, reply_date,
                     b)
        send_reply(reply, 'Thank you, dates have been updated.')

def update_tasks(task_coords, reply_date, except_list=None):  # updates dates of tasks
    if except_list is None:
        except_list = [''] #no plants to skip
    global read_request_counter
    global write_request_counter
    for coord in task_coords:
        #start = timer()
        skip_plant = False
        for entry in except_list:
            if data_sheet.get_cell(wk, coord[0], 1).lower().strip() in entry:
                skip_plant = True
                print(f'Skipping {data_sheet.get_cell(wk, coord[0], 1)}')
                read_request_counter += 2
                break
        if not skip_plant:
            data_sheet.update_cell(wk, coord[0], coord[1] + 2,
                                   reply_date.strftime("%m/%d/%Y"))
            print(f'Updating {data_sheet.get_cell(wk, coord[0], 1)}')
            read_request_counter += 1
            write_request_counter += 1
        #end = timer()
        #print(f'Time for update coord: {end - start}')
        # approx 2.5 updates per sec
        if read_request_counter >= 54:
            recharge_quota(60)
            read_request_counter = 0
        elif write_request_counter >= 54:
            recharge_quota(60)
            write_request_counter = 0
    print('Spreadsheet updated')


def send_normal_email(task_coords):  # The ordinary daily task list
    if len(task_coords) > 0:
        email = create_email(task_coords)
        subject = f'''{today_in_tz.strftime("%A, %B %d %Y")}'s plant tasks'''
        send_email(subject, email)
        # print(email)
        print('Email sent')
    else:  # if no tasks, no need for reply; just wait til tomorrow
        data_sheet.store_got_reply(wk, True)
        print('No tasks to do, no email sent')

    # store email info
    data_sheet.store_email_date(wk, datetime.now(timezone.utc)
                                .isoformat())
    data_sheet.store_task_data(wk, task_coords)
    data_sheet.store_email_type(wk, 'Normal')


def send_ignored_email(task_coords):
    if len(task_coords) > 0:
        email = create_ignored_email(task_coords)
        subject = f'''Plant task check-in: {today.strftime("%A, %B %d %Y")}'''
        send_email(subject, email)
        # print(email)
        # store email info
        print('Ignored email sent')
        data_sheet.store_email_date(wk, datetime.now(timezone.utc)
                                    .isoformat())
        data_sheet.store_email_type(wk, 'Check-in')





def main():
    try:
        # 1) Gather info of last email
        last_email_info = data_sheet.get_email_info(wk)
        #global read_request_counter
        #read_request_counter += 5
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
                    need_email = False
        if last_email_info['got_reply'] is None:
            last_email_replied = 'true'
        else:
            last_email_replied = last_email_info['got_reply'].strip().lower()
        if last_email_info['email_type'] is None:
            last_email_type = 'normal'
        else:
            last_email_type = last_email_info['email_type'].strip().lower()


    except:
        print('Error reading email info')
        raise

    # 2) Act accordingly based on time passed since last email
    # a. Email sent >24 hrs ago (or not yet), and it's just a list
    if need_email and last_email_type == 'normal' or \
            need_email and last_email_type == 'n/a':

        # option 1: reply is either settled or this is the first message
        if last_email_replied == 'true' or last_email_replied == 'n/a':
            task_coords = find_tasks()
            send_normal_email(task_coords)
            # send today's tasks, or nothing if none

        # option 2: reply still not received
        elif last_email_replied == 'false':
            send_ignored_email(eval(last_email_info['task_data']))
        else:
            raise ValueError('"Got Reply?" email info in improper format')

    # b. Email sent >24 hrs ago, and was a check-in in response to ignore
    elif need_email and last_email_type == 'check-in':
        # regardless of whether we got a response or not,
        # we're going back to our normal schedule
        task_coords = find_tasks()
        send_normal_email(task_coords)

    elif need_email:
        raise ValueError('"Email Type" email info in improper format')
    # c. Email sent <24 hrs ago
    else:

        if last_email_replied == 'true':
            print('Nothing left to do today')
            #sys.exit(0)  # nothing else needed

        elif last_email_replied == 'false':
            # check for a reply; only act if there is one
            if email_reader.has_reply(last_email_info['id']):
                reply = email_reader.get_reply(last_email_info['id'])  # dict
                # print(reply)
                # print(reply[4] == 'y')
                rpl = reply['snippet'].lower()  # user's reply message
                reply_date = date.fromtimestamp(int(reply['internalDate'][0:-3]))

                # reply to normal emails
                if last_email_type == 'normal':
                    handle_reply(reply, reply_date, eval(last_email_info['task_data']))

                # reply to check-in emails
                elif last_email_type == 'check-in':
                    if 'yesterday' in rpl:
                        reply_date = reply_date - timedelta(days=1)
                    handle_reply(reply, reply_date, eval(last_email_info['task_data']))
                else:
                    raise ValueError('"Email Type" email data in improper format')

            else:
                print('No reply received')
                # do nothing, since no reply yet
        else:
            raise ValueError('"Got Reply?" email data in improper format')

    # last_email_info = data_sheet.get_email_info(wk)
    # print(email_reader.has_reply(last_email_info['id']))


try:
    main()
except:
    print('Please check that the spreadsheet is correctly filled out')
    raise
#email_reader.threads_head(5, True)
#email_reader.get_reply('')

# last_email_info = data_sheet.get_email_info(wk)
# print(email_reader.has_reply(last_email_info['id']))


# print(data_sheet.get_cell(wk, 4, 19))

# determine tasks


# turn tasks into an email message

# send email

