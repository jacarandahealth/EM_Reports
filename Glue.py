import base64
import datetime
import hashlib
import json
import pickle
import string

import requests
from dateutil import parser
from dateutil.parser import parse
from html2text import html2text


# from fuzzywuzzy import fuzz


# FD API KEY: jrSuyBR81kfFcWM0nMl0
# My requester ID: 43020213962

# Pretty Print JSON:
#       loaded_json = json.loads(r.content)
#       print(json.dumps(loaded_json, sort_keys=True, indent=4))

class Glue():

    def __init__(self):
        self.fd_api_key = 'jrSuyBR81kfFcWM0nMl0'
        self.fd_domain = 'jacarandahealth'
        self.fd_password = 'x'

        self.em_eid = '1015527'
        self.em_epassword = '81ldtknyqgy6bl0'
        self.em_account_id = b'5338310569623552'
        self.MD5_PW = 'a812801c81bdf2'  # as provided under 'SendMessage' at https://www.echomobile.org/public/api

        self.fd_headers = {'Content-Type': 'application/json'}
        self.now = datetime.datetime.now()

        self.group_dict = {'ag1zfm0tc3dhbGktaHJkchMLEgtDbGllbnRHcm91cBjAmEQM': 'MIMBA',
                           'ag1zfm0tc3dhbGktaHJkchMLEgtDbGllbnRHcm91cBi_3UQM': 'BMAMA',
                           'ag1zfm0tc3dhbGktaHJkchMLEgtDbGllbnRHcm91cBi88E4M': 'BMIMBA',
                           'ag1zfm0tc3dhbGktaHJkchMLEgtDbGllbnRHcm91cBii9VEM': 'M2MIMBA',
                           'ag1zfm0tc3dhbGktaHJkchMLEgtDbGllbnRHcm91cBjEr2gM': 'MAMA2',
                           'ag1zfm0tc3dhbGktaHJkchQLEgtDbGllbnRHcm91cBiMrKQRDA': 'MAMA4',
                           'ag1zfm0tc3dhbGktaHJkchQLEgtDbGllbnRHcm91cBizo4oSDA': 'M2MAMA',
                           'ag1zfm0tc3dhbGktaHJkchQLEgtDbGllbnRHcm91cBiT_pASDA': 'MIMBA2TEST',
                           'ag1zfm0tc3dhbGktaHJkchQLEgtDbGllbnRHcm91cBiGyIUTDA': 'MAMA'
                           }

    def basic_auth_header(self, user, password):
        # encode the password to b64 and get the string of that value because that's what the EM servers expect
        user_password = '{}:{}'.format(user, password).encode('utf-8')
        base64_user_password_bytes = base64.b64encode(user_password)
        base64_user_password = base64_user_password_bytes.decode('utf-8')

        return 'Basic {}'.format(base64_user_password)

    # get UTC time now
    def get_now(self):
        return datetime.datetime.utcnow()

    # save EM run times list to pickle
    def save_em_run_times_to_pkl(self, em_run_times):
        with open('em_run_times.pkl', 'wb') as f:
            pickle.dump(em_run_times, f)

    # save FD run times list to pickle
    def save_fd_run_times_to_pkl(self, fd_run_times):
        with open('fd_run_times.pkl', 'wb') as f:
            pickle.dump(fd_run_times, f)

    # load EM run times list from pickle and return
    def load_em_run_times_from_pkl(self):
        # TODO: if pickle doesn't exist, find datetime of last EM messages from fd, create pkl, seed it with that
        with open('em_run_times.pkl', 'rb') as f:
            em_loaded_run_times = pickle.load(f)
        return em_loaded_run_times

    # load FD run times list from pickle and return
    def load_fd_run_times_from_pkl(self):
        # TODO: if pickle doesn't exist, find datetime of last EM messages from fd, create pkl, seed it with that
        with open('fd_run_times.pkl', 'rb') as f:
            loaded_run_times = pickle.load(f)
        return loaded_run_times

    # save list of phone number and tkt number values to dict
    def save_phone_dict_to_pkl(self, phones_tkts_dict):
        with open('phones_tkts.pkl', 'wb') as f:
            pickle.dump(phones_tkts_dict, f)

    # load list of phone number and tkt number values from dict
    def load_phone_dict_from_pkl(self):
        # TODO: if pickle doesn't exist, create it?
        with open('phones_tkts.pkl', 'rb') as f:
            loaded_phones_tkts = pickle.load(f)
        return loaded_phones_tkts

    # get messages from EM from the last run
    def get_messages_from_em(self, last_run_datetime):
        """Parameters
            starred (optional): Only pull starred messages int
            archived (optional): 0; pull non-archived messages, 1; pull archived messages int
            unread (optional): 0; pull all messages, 1; pull unread messages only int
            since (optional): Unix timestamp cutoff (range start, UTC, seconds) int
            until (optional): Unix timestamp cutoff (range end, UTC, seconds) int
            page (optional): If the result set exceeds the max per page, it will be paged. This parameter identifies the page the request is interested in. int
            max(optional): Max records to return (1-200). int"""

        glue = Glue()

        r = requests.get('https://www.echomobile.org/api/cms/inbox',
                         params={'max': 50,  # TODO: REMOVE THIS!!!!!
                                 'unread': 0},  # change value to 1 if duplicate messages keep coming in
                         headers={
                             'authorization': glue.basic_auth_header(glue.em_eid, glue.em_epassword),
                             'account-id': base64.b64encode(glue.em_account_id),
                             'since': str(last_run_datetime)
                             # TODO: Check this works correctly and is returning the correct number of messages
                         })
        return r.text, 'EM'

    # post messages to FD in new or existing tickets
    def post_tickets_to_fd(self, messages, source_platform, loaded_phone_dict):

        glue = Glue()
        phone_tkt_dict = {}

        data = json.loads(messages)
        for each_message in (data['ims']):
            tkt_dict = {}
            group_name = glue.group_dict[each_message['group_key']]
            phone_number = each_message['sender_phone']

            # phone_number = '254704888680'  # TODO: remove this line after testing!!!

            allow_message = glue.filter(each_message['message'])
            # pass through if False
            if allow_message == True:
                tkt_dict['phone'] = phone_number
                tkt_dict['description'] = each_message['message']
                tkt_dict['subject'] = '[{} {}] {}'.format(source_platform, each_message['sender_phone'], each_message['message'])
                tkt_dict['status'] = 2
                tkt_dict['priority'] = 1
                tkt_dict['tags'] = ['EM', group_name]
                tkt_dict['email'] = 'jpatel@jacarandahealth.org'
                tkt_dict['requester_id'] = 43000358054
                tkt_dict['group_id'] = 43000480647

                print('MESSAGE: {}'.format(each_message['message']))

                # if phone exists in phone dict, update tkt, else create new tkt
                if phone_number not in loaded_phone_dict.keys():
                    print('creating tkt')
                    r = requests.post('https://{}.freshdesk.com/api/v2/tickets'.format(glue.fd_domain),
                                      auth=(glue.fd_api_key, glue.fd_password),
                                      headers=glue.fd_headers, data=json.dumps(tkt_dict))

                    tkt_number = r.headers['Location'].split('/')[-1]
                else:
                    print('updating tkt')
                    tkt_to_update = loaded_phone_dict[phone_number]
                    update_to_tkt = {'body': tkt_dict['description']}

                    r = requests.post(
                        'https://{}.freshdesk.com/api/v2/tickets/{}/reply'.format(glue.fd_domain, tkt_to_update),
                        auth=(glue.fd_api_key, glue.fd_password),
                        headers=glue.fd_headers, data=json.dumps(update_to_tkt))

                    tkt_number = tkt_to_update  # {phone_number: tkt_number}

                if r.status_code == 200:
                    print('Ticket created/updated successfully, the response is given below' + str(r.content))
                    print('Location Header : ' + r.headers['Location'])
                elif r.status_code == 201:
                    print('Ticket created/updated successfully, the response is given below' + str(r.content))
                    print('Location Header : ' + r.headers['Location'])
                else:
                    print('FAILED to create/update ticket, errors are displayed below,')
                    response = json.loads(r.content)
                    print(response)
                    print('x-request-id : ' + r.headers['x-request-id'])
                    print('Status Code : ' + str(r.status_code))

                phone_tkt_dict[phone_number] = tkt_number  # note: this overwrites phone number!  # TODO: confirm this is ok

        return phone_tkt_dict

    # get FD tickets/conversations that have been updated since last run
    def get_fd_tkt_updates(self, last_run_datetime, tkt_num, flag):
        # Return the tickets that are new or opened & assigned to you
        # If you want to fetch all tickets remove the filter query param

        glue = Glue()
        last_run = last_run_datetime.isoformat()
        print('getting fd tkt updates')

        if flag is 'conversations':
            # this gets all conversations from a single ticket, given the ticket number
            r = requests.get(
                'https://{}.freshdesk.com/api/v2/tickets/{}?include=conversations'.format(glue.fd_domain, tkt_num),
                auth=(glue.fd_api_key, glue.fd_password))
        elif flag is 'tickets':
            # this lists the ticket number/id updated since {given date}
            r = requests.get(
                'https://{}.freshdesk.com/api/v2/tickets?updated_since={}'.format(glue.fd_domain, last_run),
                auth=(glue.fd_api_key, glue.fd_password))

        if r.status_code == 200:
            # print('Request processed successfully, the response is given below')
            pass
        else:
            print('Failed to read tickets, errors are displayed below,')
            response = json.loads(r.content)
            print(response['errors'])

            print('x-request-id : ' + r.headers['x-request-id'])
            print('Status Code : ' + str(r.status_code))

        data = json.loads(r.content)

        return data

    # get ticket numbers for updated tickets from FD
    def get_tkt_numbers(self, all_tickets):
        tkt_list = []
        for each_tkt in all_tickets:
            tkt_list.append(each_tkt['id'])

        return tkt_list

    # extract all messages from FD since last run, return list of messages
    def process_fd_conversations(self, last_run, fd_reply):
        glue = Glue()
        body_text_list = []

        # get origin platform and phone number from the subject in order to route reply correctly
        subject = (fd_reply['subject'].upper())
        last_updated = parser.parse(fd_reply['updated_at']).replace(tzinfo=None)
        # process if message originates with EM or TI and message created/updated since last run of program
        if (subject.startswith('[EM ' or subject.startswith('[TI '))) and last_run < last_updated:
            platform, phone_num = subject[subject.find('[') + 1: subject.find(']')].split(
                ' ')  # split between '[' & ']'

            for each_conversation in fd_reply['conversations']:
                # convert html to text, strip trailing whitespace, fix weird \u200b character issue
                body_text = html2text(each_conversation['body']).strip().replace(u'\u200b', '')
                body_text_list.append((phone_num, body_text))

            return platform.upper(), body_text_list
        else:
            return None, None  # return None if platform is, for example, facebook or other

    # post each conversation in message list to EM, which will send out to end-users
    def post_messages_to_em(self, messages_list):
        glue = Glue()

        for each_message in messages_list:
            phone_num = each_message[0]
            body_text = str(each_message[1])
            phone_num = '254704888680'  # TODO: remove this line after testing!!!

            # concatenate the values into a string and encode string to bytes
            concatenated_values = (glue.MD5_PW + phone_num + body_text).encode('utf-8')

            # get the MD5
            md5_digest = hashlib.md5(concatenated_values).hexdigest()

            message_data = {'phone': phone_num,
                            'message': body_text,
                            'digest': md5_digest
                            }

            message_headers = {'authorization': glue.basic_auth_header(glue.em_eid, glue.em_epassword),
                               'account-id': base64.b64encode(glue.em_account_id)
                               }

            r = requests.post('https://www.echomobile.org/api/messaging/send',
                              data=message_data,
                              headers=message_headers)

            print(r, r.text, message_data)

    # filter irrelevant messages before posting to FD
    def filter(self, incoming_message):
        allow_through = False

        words_to_filter = ['mama', 'mimba', 'm2mama', 'm2mimba', 'bmama', 'bmimba',
                           'stop', 'a', 'a.', 'b', 'b.', 'sawa', 'fine', 'ok', 'poa',
                           'asante', 'asante sana', 'thanks', 'thanx', 'thank you', 'thanks a lot', 'thanks alot',
                           'thenx', 'thenx a lot', 'thenx alot', 'asnte']

        incoming_lower = incoming_message.lower().rstrip('\n')

        # this is to filter out facility names but is not currently working reliably
        """temp_match_dict = {}
        if isinstance(incoming_lower, str) and incoming_lower not in words_to_filter and len(
                incoming_lower) > 3 and len(incoming_lower) <= 45:
            facilites_to_match_against = glue.get_sliced_list(incoming_lower)

            for each_facility in facilites_to_match_against:
                fuzz_ratio = fuzz.partial_ratio(each_facility, incoming_lower)  # get match ratio
                temp_match_dict[each_facility] = fuzz_ratio  # facility: ratio
                        
            # get highest value in temp_match_dict and return key
            # https://stackoverflow.com/questions/268272/getting-key-with-maximum-value-in-dictionary
        if len(temp_match_dict) > 0:
            best_match = max(temp_match_dict, key=temp_match_dict.get)
            print('{}, {}, {}'.format(incoming_lower, best_match, fuzz_ratio))"""

        # this allows for strings in date formats to be ignored so that they can be checked later
        exclude_punctuation = set(string.punctuation.replace('/', ''))
        message_to_process = ''.join(ch for ch in incoming_lower if ch not in exclude_punctuation)

        try:  # if incoming message is a date, filter it out
            parsed_str = parse(message_to_process, dayfirst=True)
            if isinstance(parsed_str, datetime.date):
                allow_through = False
        except ValueError:
            if isinstance(message_to_process, str) and message_to_process not in words_to_filter and 'dd/mm' not in message_to_process:
                allow_through = True
            else:
                allow_through = False

        # proceed only if allow_through is False
        return allow_through

    def get_sliced_list(self, incoming_message):

        facility_file = 'main_facility_list.csv'
        with open(facility_file) as infile:
            facility_list = infile.readlines()

        facility_list = [x.strip('\n') for x in facility_list]
        facilities_lower = [x.lower() for x in facility_list]

        incoming_first_char = incoming_message[0].lower()
        sliced_list = []

        for each in facilities_lower:
            if each[0] == incoming_first_char:
                sliced_list.append(each)

        return sliced_list


class Main():

    def __init__(self):
        pass

    def run(self):
        glue = Glue()
        # TODO: THOROUGHLY TEST "SINCE" FUNCTIONALITY

        # TODO: KEEP TRACK OF ALL PHONE NUMBER AND TICKET NUMBER COMBINATIONS IN A DICT ('PHONE': "TKT') ON DISK
        # TODO: IF PHONE PRESENT IN DICT, UPDATE TKT, ELSE CREATE TKT

        # TODO: Integrate TextIt

        # TODO: Connection timeout behaviour - erase last run from list, stop program, run again later

        # TODO: email errors or log errors and email log daily

        # get last run's timestamp
        em_run_times_list = glue.load_em_run_times_from_pkl()
        fd_run_times_list = glue.load_fd_run_times_from_pkl()

        # get this run's timestamp
        em_this_run = glue.get_now() + datetime.timedelta(seconds=0.25)  # add 1/4 second to run to prevent conflict
        em_last_run = em_run_times_list[-1]
        fd_last_run = fd_run_times_list[-1]

        em_run_times_list.append(em_this_run)  # append this run's timestamp to list and save to pkl
        glue.save_em_run_times_to_pkl(em_run_times_list)

        # em_last_run = datetime.datetime(2019, 1, 21, 7, 50, 0, 0)  # for testing

        phone_dict = glue.load_phone_dict_from_pkl()

        em_messages, source_platform = glue.get_messages_from_em(em_last_run)  # test this
        fd_phone_tkt_dict = glue.post_tickets_to_fd(em_messages, source_platform, phone_dict)  # working

        # update pkl-loaded phone/tkt dict
        for phone_num, tkt_num in fd_phone_tkt_dict.items():
            if phone_num not in phone_dict:
                phone_dict[phone_num] = tkt_num

        updated_tkts = glue.get_fd_tkt_updates(fd_last_run, '', 'tickets')  # test this
        fd_this_run = glue.get_now() + datetime.timedelta(seconds=0.25)  # add 1/4 second to run to prevent conflict
        fd_run_times_list.append(fd_this_run)
        glue.save_fd_run_times_to_pkl(fd_run_times_list)

        tkt_numbers = glue.get_tkt_numbers(updated_tkts)

        # plug updates into EM/TI
        for tkt_num in tkt_numbers:
            conversation = glue.get_fd_tkt_updates(fd_last_run, tkt_num, 'conversations')
            platform, messages_list = glue.process_fd_conversations(fd_last_run, conversation)

            if platform is not None:
                if platform == 'EM':
                    glue.post_messages_to_em(messages_list)  # working to my phone, test to others
                elif platform == 'TI':
                    pass  # TODO: TI integration here
                    # glue.post_messages_to_ti(messages_dict)"""

        # glue.post_messages_to_em(fd_reply)  # test this

        glue.save_phone_dict_to_pkl(phone_dict)

    def run2(self):
        glue = Glue()
        messages_file = 'All_Incoming_Messages.csv'
        f = open("filter_pass.csv", "a")

        with open(messages_file) as infile:
            messages_list = infile.readlines()

        # print(messages_list)
        messages_list = [x.strip('\n') for x in messages_list]

        for each_message in messages_list:
            to_filter = glue.filter(each_message)
            if to_filter == True:
                f.write('FILTERED, {}\n'.format(each_message))
                pass
            elif to_filter == False:
                f.write('PASSED, {}\n'.format(each_message))
                pass

if __name__ == '__main__':
    main = Main()
    main.run()
    # main.run2()
