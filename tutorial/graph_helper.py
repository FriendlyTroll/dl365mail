# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

# <FirstCodeSnippet>
import errno
import logging

from requests_oauthlib import OAuth2Session

graph_url = 'https://graph.microsoft.com/v1.0'

OUTPUT_DIRECTORY = './emls'
def get_user(token):
  graph_client = OAuth2Session(token=token)
  # Send GET to /me
  user = graph_client.get('{0}/me'.format(graph_url))
  # Return the JSON result
  return user.json()
# </FirstCodeSnippet>

# <GetCalendarSnippet>
def get_user_mails(token):
  graph_client = OAuth2Session(token=token)
  headers = {
    'Authorization': 'Bearer {0}'.format(token),
    }

  # Configure query parameters to
  # modify the results
  query_params = {
    'top': 100 # page size
    }

  # Send GET to /me/events
  #events = graph_client.get('{0}/me/events'.format(graph_url), params=query_params)

  events = graph_client.get(
    #"{0}/me/mailfolders/sentitems".format(graph_url), headers=headers, params=query_params)
    #"{0}/me/mailfolders('SentItems')/messages?select=sender,subject".format(graph_url), headers=headers, params=query_params)
    #'{0}/me/mailfolders//messages'.format(graph_url), headers=headers, params=query_params)

    # list all mail folders
    "{0}/me/mailfolders".format(graph_url), headers=headers, params=query_params)
  mail_folders = events.json()
  # for mailfolder in mail_folders["value"]:
      #print(f'{mailfolder["displayName"]}, id: {mailfolder["id"]}')
      # print(f'{mailfolder}')

     # get emails in SBOOK folder
    #'{0}/me/mailfolders/AAMkAGZkY2QwNTc2LTk5ZmMtNGY4ZS05ODk3LWUwYjkwZjQwNmZiZgAuAAAAAACH2dhbPdcWRal55qhcuA6OAQDvcMs_nI5nR6M1kM8smVKLAAAB1frLAAA=/messages'.format(graph_url), headers=headers, params=query_params)
  #msgs = events.json()["value"] # list of emails
  # next_page_url = events.json()["@odata.nextLink"] # next page of results
  # print(next_page_url)

    # get single mail
    #"{0}/me/messages/AAMkAGZkY2QwNTc2LTk5ZmMtNGY4ZS05ODk3LWUwYjkwZjQwNmZiZgBGAAAAAACH2dhbPdcWRal55qhcuA6OBwDvcMs_nI5nR6M1kM8smVKLAAAAAAEJAADvcMs_nI5nR6M1kM8smVKLAAAgoiWTAAA=/$value".format(graph_url), headers=headers, params=query_params)
  #print(events.content.decode('utf-8')) # print mime message; save this as .eml file
  
  logging.basicConfig(filename='/tmp/usermail.log', level=logging.INFO)

  for mailfolder in mail_folders["value"]:
      #if mailfolder["displayName"] in ["Inbox", "Deleted Items", "Junk Email", "Sent Items"]:
      #if mailfolder["displayName"] in ["Sent Items", "Important"]:
      logging.info(f'[*] Fetching mails from folder: >> {mailfolder["displayName"]}')
      message_list = graph_client.get(
      '{0}/me/mailfolders/{1}/messages'.format(graph_url, mailfolder["id"]), headers=headers, params=query_params)
      if message_list.json()["value"]:
          for i, email_object in enumerate(message_list.json()["value"]): # use enumerate to get index as well
              # logging.info(f'{email_object}')
              email = graph_client.get(
                  '{0}/me/messages/{1}/$value'.format(graph_url, email_object["id"]), headers = headers, params = query_params)

              email_mime = email.content # bytes object
              email_subject = email_object["subject"]
              try:
                email_from = email_object["from"]["emailAddress"]["name"]
              except KeyError as e:
                  logging.info("!!! >>> KeyError: No 'from' field in email!")
                  email_from = "Unknown"
              try:
                logging.info(f"Writing mail {email_subject}!")
                f = open(f'{OUTPUT_DIRECTORY}/{i}--{email_from}--{email_subject.replace("/", "")}.eml', 'wb')
                f.write(email_mime)
                f.close()
              except OSError as exc:
                if exc.errno == errno.ENAMETOOLONG:
                  logging.info("!!! >>> Filename too long. Truncating subject line to 150 chars")
                  f = open(f'{OUTPUT_DIRECTORY}/{i}--{email_from}--{email_subject[:150].replace("/", "")}.eml', 'wb')
                  f.write()
                  f.close()
                else:
                  logging.info(exc)
              except Exception as e:
                  logging.info(f"Couldn't write email due to {e}")
                  continue
          else:
              logging.info(f"[**] Finished with first batch. See if there are other pages.")

              # check if there is a link for next page of result
              try:
                  next_page_url = message_list.json()["@odata.nextLink"]
                  logging.info(f"[**] Next page {next_page_url}")
                  # if there is link to next page of result, use that link in next request
                  if next_page_url:
                    fetch_mails(graph_client, headers, next_page_url, query_params)
              except:
                  logging.info("Problem getting next page with emails!")
                  pass
      else:
          logging.info("[**] No emails in folder!")


def fetch_mails(graph_client, headers, next_page_url, query_params):
    # if there is link to next page of result, use that link in next request
    while next_page_url:
        message_list = graph_client.get(
            '{0}'.format(next_page_url), headers=headers)
        if message_list.json()["value"]:
            for i, email_object in enumerate(message_list.json()["value"]):
                email = graph_client.get(
                    '{0}/me/messages/{1}/$value'.format(graph_url, email_object["id"]), headers=headers,
                    params=query_params)

                email_mime = email.content  # bytes object
                email_subject = email_object["subject"]
                email_from = email_object["from"]["emailAddress"]["name"]
                try:
                    logging.info(f"Writing mail {email_subject}")
                    f = open(f'{OUTPUT_DIRECTORY}/{i}--{email_from}--{email_subject.replace("/", "")}.eml', 'wb')
                    f.write(email_mime)
                    f.close()
                except OSError as exc:
                    if exc.errno == errno.ENAMETOOLONG:
                        logging.info("!!! >>> Filename too long. Truncating subject line to 150 chars")
                        f = open(f'{OUTPUT_DIRECTORY}/{i}--{email_from}--{email_subject[:150].replace("/", "")}.eml', 'wb')
                        try:
                            f.write()
                            f.close()
                        except Exception as e:
                            logging.info(f"Couldn't write email due to {e}")
                            continue
                    else:
                        logging.info(exc)
                except Exception as e:
                    logging.info(f"Couldn't write email due to {e}")
                    continue
            logging.info(f"[**] Finished with batch. Continue to next page.")
            # check if there is a link for next page of result
            try:
                next_page_url = message_list.json()["@odata.nextLink"]
                logging.info(f"[**] Next page {next_page_url}")
            except Exception as e:
                logging.info(f'[**] No more pages with emails!')
                return


if __name__ == "__main__":
  get_user_mails()
