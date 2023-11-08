mode = 'Dev'


import itertools,random,time,sys,configparser

def try_until(func, max_tries, sleep_time): 

    #we need to keep retrying until we get a valid combination list
    #because sometimes its not possible to pick a full valid combination
    #earlier choices and the avoid_gifting_to list can sometimes be a problem
    #this function allows repeating calls to a function
    #of course this will mask all other errors so consider when debugging calling the function without using the wrapper

    for _ in range(0,max_tries):
        try:
            return func()
        except Exception as e:
            print ('XXXXXXXXXXXXX UNABLE TO CREATE FULL SET OF MATCHES - RETRYING  xxxxxxxxxxxxxxx')
            time.sleep(sleep_time) #wait between retries

def generatePerms(participants,disallowed,forced):

    config = configparser.ConfigParser()
    config.read('config.ini')
    outputFile = config.get(mode, 'outputFile')

    #participants - a list of people participating in secret santa
    #disallowed - a list of tuples of disallowed combinations pos 0 - sender, pos 1 - recipient
    #forced - a list of tuples of combinations that have to happen pos 0 - sender, pos 1 - recipient

    combinations = list(set(itertools.product(participants,repeat=2)).difference(set(disallowed)))
    #a list of tuples containing all combinations minus the disallowed combinations

    forced_senders = [y[0] for y in forced]
    forced_recipients = [y[1] for y in forced]    

    combinations = [y for y in combinations if y[0] not in forced_senders and y[1] not in forced_recipients]
    #also remove anything where the sender or recipient is also forced by the gift_to list

    result = forced.copy() #we start our list of results, with the list of forced combinations

    partipantsToProcess = set(participants).difference(forced_senders) #our list of people that we need to find gift recipients for

    with open(outputFile,'w') as f: #create a blank output file
        f.write('')

    for participant in partipantsToProcess: #for each sender
        possibles = [y for y in combinations if y[0] == participant and y[0] != y [1]] #generate a list of possible combos
        choice = random.choice(possibles) #randomly pick one
        recipient = choice[1] 
        combinationsCopy = []

        for x in combinations: #we have selected a recipient so now we remove this recipeient from the possible combinations list
            if x[1] != recipient:
                combinationsCopy.append(x)

            combinations = combinationsCopy.copy()

        result.append(choice) #we add our selected combo to the result list
        
        with open(outputFile,'a') as f: #we add our selected combo to the output file
            f.write(f'{choice[0]} will send to {choice[1]}\n')
    
    with open(outputFile,'a') as f: #all of those forced combos from earlier, lets add them to the output file too
        for x in forced:
            f.write(f'{x[0]} will send to {x[1]}\n')

    return result



def sheetContentValidation(participants,disallowed,forced):

    #we are going to validate the spreadsheet contents
    #we want to ensure that everyone named in the gift_to and avoid_gifting_to cols
    #are actually participating

    for x in disallowed:
        if x[1] not in participants:
            print(f'The value "{x[1]}" in the avoid_gifting_to column has not been found in the particpant list - aborting')
            sys.exit() #quit if there is a problem

    for x in forced:
        if x[1] not in participants:
            print(f'The value "{x[1]}" in the gift_to column has not been found in the particpant list - aborting')
            sys.exit() #quit if there is a problem

def parseGoogleSheet():

    import gspread,pandas
    from oauth2client.service_account import ServiceAccountCredentials

    config = configparser.ConfigParser()
    config.read('config.ini')
    sheetsJson = config.get(mode, 'sheetsJson')
    sheetName = config.get(mode, 'sheetName')
    WorksheetName = config.get(mode, 'WorksheetName')    

    scope = ['https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(sheetsJson, scope)
    client = gspread.authorize(creds)
    sheet = client.open(sheetName).worksheet(WorksheetName) #Open the spreadhseet

    df = pandas.DataFrame(sheet.get_all_records())
    participants = df['name'].tolist() #create a list of the participants
    difficultPeople = df[['name','avoid_gifting_to']].query('avoid_gifting_to != ""') #dataset, people who we cannot gift to others

    disallowed = []

    x = dict(difficultPeople.values) #dictionary of people who we cannot gift to others
    for x,y in x.items(): 
        for y in y.split(","): #its a comma seperated list, lets convert it to a list of tuples of non allowable combinations pos 0 sender pos 2 recipient
            disallowed.append((x,y.strip(' ')))

    forced = []

    forced_df = df[['name','gift_to']].query('gift_to != ""') #dataset, people who are forced to gift to named other people
    x = dict(forced_df.values)
    for x,y in x.items():
        forced.append((x,y)) #a list of tuples of forced combinations pos 0 sender pos 2 recipient

    sheetContentValidation(participants,disallowed,forced) #lets do some validation before continuing

    email_lookup = df[['name','email_address']] #and while we are here lets get a name / email address look up

    return (participants,disallowed,forced,email_lookup)

def sendEmail(subject,body,recipientAddress):

    import pickle
    import base64
    import googleapiclient.discovery
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText

    config = configparser.ConfigParser()
    config.read('config.ini')
    SigName = config.get(mode, 'SigName')
    SigEmail = config.get(mode, 'SigEmail')
    SigJobTitle = config.get(mode, 'SigJobTitle')
    EmailPickle = config.get(mode, 'EmailPickle')

    sig = f'<span style="font-size:13.999999999999998pt;font-family:Arial;color:#0097bb;background-color:transparent;font-weight:700;font-style:normal;font-variant:normal;text-decoration:none;vertical-align:baseline;white-space:pre-wrap;white-space:pre-wrap">{SigName}</span></h2><h2 dir="ltr" style="line-height:1.2;margin-top:0pt;margin-bottom:0pt"><span style="font-weight:normal"><p dir="ltr" style="line-height:1.2;margin-top:0pt;margin-bottom:0pt"><span style="font-size:11pt;font-family:Arial;color:rgb(221,29,33);background-color:transparent;vertical-align:baseline;white-space:pre-wrap">{SigJobTitle}</span><span style="font-size:11pt;font-family:Arial;color:rgb(34,34,34);background-color:transparent;vertical-align:baseline;white-space:pre-wrap">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<wbr>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></p><br><p dir="ltr" style="line-height:1.2;margin-top:0pt;margin-bottom:0pt"><span style="font-size:11pt;font-family:Arial;color:rgb(89,89,89);background-color:transparent;vertical-align:baseline;white-space:pre-wrap">Shell Energy Retail Limited,</span><span style="font-size:11pt;font-family:Arial;color:rgb(89,89,89);background-color:transparent;vertical-align:baseline;white-space:pre-wrap"><br></span><span style="font-size:11pt;font-family:Arial;color:rgb(89,89,89);background-color:transparent;vertical-align:baseline;white-space:pre-wrap">Shell Energy House, Westwood Business Park,&nbsp;</span></p><p dir="ltr" style="line-height:1.2;margin-top:0pt;margin-bottom:0pt"><span style="font-size:11pt;font-family:Arial;color:rgb(89,89,89);background-color:transparent;vertical-align:baseline;white-space:pre-wrap">Westwood Way, Coventry, CV4 8HS</span></p><br><p dir="ltr" style="line-height:1.656;margin-top:0pt;margin-bottom:0pt"><span style="font-size:11pt;font-family:Arial;color:rgb(221,29,33);background-color:transparent;font-weight:700;vertical-align:baseline;white-space:pre-wrap">E</span><span style="font-size:11pt;font-family:Arial;color:rgb(34,34,34);background-color:transparent;font-weight:700;vertical-align:baseline;white-space:pre-wrap">&nbsp; </span><span style="font-size:11pt;font-family:Arial;color:rgb(34,34,34);background-color:transparent;vertical-align:baseline;white-space:pre-wrap">&nbsp;&nbsp;&nbsp;&nbsp;</span><span style="font-size:11pt;font-family:Arial;color:rgb(0,60,136);background-color:transparent;vertical-align:baseline;white-space:pre-wrap"><a href="mailto:{SigEmail}" target="_blank">{SigEmail}</a>'

    pickle_path = EmailPickle
    creds = pickle.load(open(pickle_path, 'rb'))
    service = googleapiclient.discovery.build('gmail', 'v1', credentials=creds)

    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['To'] = recipientAddress
    msgHtml = body
    msg.attach(MIMEText(msgHtml + '<p> --- <p>' +sig, 'html'))
    raw = base64.urlsafe_b64encode(msg.as_bytes())
    raw = raw.decode()
    body = {'raw': raw}

    message1 = body
    message = (
        service.users().messages().send(
            userId="me", body=message1).execute())
    print('Message Id: %s' % message['id'])

participants,disallowed,forced,email_lookup = parseGoogleSheet()
email_lookup = email_lookup.set_index('name')

perms = try_until(lambda: generatePerms(participants,disallowed,forced),5,0)

config = configparser.ConfigParser()
config.read('config.ini')
Sendmail = config.getboolean(mode, 'Sendmail')


for x in perms:
    email_address = email_lookup.loc[x[0]]['email_address']
    body = "Dear " + x[0] + ", <p> You have been randomly selected to send a Secret Santa gift to " + x[1]
    subject = "Secret Santa Information"
    
    if Sendmail:
        sendEmail(subject,body,email_address)  