from imapclient import IMAPClient
import re
import pandas as pd
import email
from email.utils import parseaddr

def extractSCL(report,from_domain):
    try:
        # Define a regular expression pattern to match SCL followed by a colon and digits
        parts = from_domain.split('@')
        domain = ""
        # Check if there are exactly two parts (username and domain)
        if len(parts) == 2:
            # Remove leading and trailing spaces from the domain part
            domain = parts[1].strip()
        else:
            domain = "No domain"

        # Use re.search to find the first match in the input string
        match = re.search('SCL:(\d+)', report)

        # If a match is found, extract and return the SCL value as an integer
        scl_value = int(match.group(1))
        if scl_value == 1:
            print(domain + " , SCL: " + str(scl_value))
        else:
            print(domain + " , SCL: " + str(scl_value) + " *********************** WARNING ******************")
    except (TypeError, KeyError) as e:
        print(f'error in function extractSCL and error is '+e)


def main():
    file=pd.read_csv('emails.csv')
    
    
    for i in range(len(file)):
        
        email_login = file.iloc[i].email
        generate_password = file.iloc[i].generate_password
        provider =file.iloc[i].provide.lower()
        if provider.lower() == 'outlook':
            provider = 'hotmail'
        providers = {
            'gmail': {'server': 'imap.gmail.com', 'folder': '[Gmail]/Spam','report':'X-Forefront-Antispam-Report'},
            'yahoo': {'server': 'imap.mail.yahoo.com', 'folder': 'Bulk','report':'X-Forefront-Antispam-Report'},
            'hotmail': {'server': 'imap-mail.outlook.com', 'folder': 'Junk','report':'X-Forefront-Antispam-Report-Untrusted'}
        }
        if provider in providers:
            settings = providers[provider]
            server = IMAPClient(settings['server'], use_uid=True)

            try:
                server.login(email_login, generate_password)
                print('logine successful '+email_login)

                folders = server.list_folders()
                select_info = server.select_folder(settings['folder'])
                messages = server.search(['ALL'])

                items = sorted(server.fetch(messages, ['RFC822']).items(), reverse=True)
                for msgid, data in items:
                    msg = email.message_from_bytes(data[b'RFC822'])
                    forefront = msg[settings['report']]
                    from_domain = msg["From"]
                    try:
                        extractSCL(forefront, from_domain)
                    except (TypeError, KeyError) as e:
                        print(f"No Scl in this from  " + from_domain)
            except Exception as e:
                print(f"Error processing {provider} emails: {e}")

            finally:
                server.logout()
        else:
            print(f"Unknown provider: {provider}")

            
    
    
   

main()
