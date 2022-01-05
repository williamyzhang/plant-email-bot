import email_sender

our_email = email_sender.our_email
service = email_sender.service

threads = service.users().threads().list(userId=our_email).execute().get('threads', [])


def has_reply(id):
    msg = service.users().messages().get(userId=our_email, id=id).execute()

    #snippet = msg['snippet']
    send_date = int(msg['internalDate'])

    for thread in threads:
        tdata = service.users().threads().get(userId=our_email, id=thread['id']).execute()
        messages = tdata['messages']
        nmsgs = len(messages)
        threadId = tdata['id']
        if threadId == id: #look in thread of sent email
            if nmsgs > 1:
                return True
            else:
                return False
        internalDate = int(messages[nmsgs - 1]['internalDate'])
        if internalDate < send_date:
            break  # only look at messages w/ latest reply after send date of target message
                    # avoid endless loop
    return False

def get_reply(id):
    msg = service.users().messages().get(userId=our_email, id=id).execute()
    snippet = msg['snippet']
    send_date = msg['internalDate']
    return _get_reply(snippet, int(send_date))

def _get_reply(snippet, send_date): # returns reply, subject, threadid, references, message-id
    for thread in threads:
        tdata = service.users().threads().get(userId=our_email, id=thread['id']).execute()
        messages = tdata['messages'] # all messages in thread
        nmsgs = len(messages)
        threadId = tdata['id']
        internalDate = messages[nmsgs - 1]['internalDate'] # send as string
                                                        # b/c will be manipulated
        if int(internalDate) < send_date:
            break  # only look at messages w/ latest reply after send date of target message
        if nmsgs <= 1:
            continue # skip over unreplied threads
        if messages[0]['snippet'] == snippet:
            msg = messages[nmsgs - 1] # only read the last
            snippet = msg['snippet'] # ! read body rather than snippet
            subject = ''
            references = ''
            message_id = ''
            #print('Body: ', msg['payload']['body'])
            for header in msg['payload']['headers']:
                #print(header)
                # search for headers with data we want
                if header['name'].lower() == 'subject':
                    subject = header['value']
                if header['name'].lower() == 'references':
                    references = header['value']
                if header['name'].lower() == 'message-id':
                    message_id = header['value']
            return {'threadId': threadId, 'message_id': message_id,
                    'references': references, 'internalDate': internalDate,
                    'subject': subject, 'snippet': snippet}
    return {}

# preview messages in inbox
def threads_head(max_displayed=10, more=False):  # (for debugging purposes)
    n = 0
    for thread in threads:
        if n >= max_displayed:
            break
        tdata = service.users().threads().get(userId=our_email, id=thread['id']).execute()
        messages = tdata['messages']
        nmsgs = len(messages)
        if more:
            print(f"threadId: {tdata['id']}")
            for msg in messages:  # show each message in thread
                print(f"  snippet: {msg['snippet']}")
                print(f"  id: {msg['id']}, history: {msg['historyId']}")
                print(f"  date: {msg['internalDate']}")
        else:
            top = tdata['messages'][0]  # summary of thread
            print(f"snippet: {top['snippet']}\n  threadId: {top['threadId']}"
                  f"\n  # msgs: {nmsgs}, topId: {top['id']}"
                  f"\n  history: {tdata['historyId']}")

        n += 1

#threads_head(20, True)