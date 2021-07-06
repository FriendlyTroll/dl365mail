# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

from django.shortcuts import render
from django.http import HttpResponse, HttpResponseRedirect, StreamingHttpResponse
from django.urls import reverse
from tutorial.auth_helper import get_sign_in_url, get_token_from_code, store_token, store_user, remove_user_and_token, get_token
from tutorial.graph_helper import get_user, get_user_mails
import dateutil.parser
import threading
import time

# <HomeViewSnippet>
def home(request):
  context = initialize_context(request)

  return render(request, 'tutorial/home.html', context)
# </HomeViewSnippet>

# <InitializeContextSnippet>
def initialize_context(request):
  context = {}

  # Check for any errors in the session
  error = request.session.pop('flash_error', None)

  if error != None:
    context['errors'] = []
    context['errors'].append(error)

  # Check for user in the session
  context['user'] = request.session.get('user', {'is_authenticated': False})
  return context
# </InitializeContextSnippet>

# <SignInViewSnippet>
def sign_in(request):
  # Get the sign-in URL
  sign_in_url, state = get_sign_in_url()
  # Save the expected state so we can validate in the callback
  request.session['auth_state'] = state
  # Redirect to the Azure sign-in page
  return HttpResponseRedirect(sign_in_url)
# </SignInViewSnippet>

# <SignOutViewSnippet>
def sign_out(request):
  # Clear out the user and token
  remove_user_and_token(request)

  return HttpResponseRedirect(reverse('home'))
# </SignOutViewSnippet>

# <CallbackViewSnippet>
def callback(request):
  # Get the state saved in session
  expected_state = request.session.pop('auth_state', '')
  # Make the token request
  token = get_token_from_code(request.get_full_path(), expected_state)

  # Get the user's profile
  user = get_user(token)

  # Save token and user
  store_token(request, token)
  store_user(request, user)

  return HttpResponseRedirect(reverse('home'))
# </CallbackViewSnippet>

# <CalendarViewSnippet>

def file_iterator(ffile, chunk_size=512):
    with open(ffile) as f:
        while True:
            c = f.readline(chunk_size)
            if c:
                yield c + '<br>'
                # wait a bit for the file to populate lest we break immediately
                time.sleep(0.75)
            else:
                break
    f.close()

def emails(request):
  context = initialize_context(request)
  token = get_token(request)

  # create thread to fetch the mails 
  t = threading.Thread(target=get_user_mails, args=(token,), daemon=True)
  t.start()

  # read from log file to display output on web page
  log_file = '/tmp/usermail.log'
  # wait for the file to be created
  time.sleep(1.5)
  response = StreamingHttpResponse(file_iterator(log_file))
  return response

  #if events:
  #  context['events'] = events['value']

  #print(f"Got resp: >> {events}")
  #return render(request, 'tutorial/mails.html', response)
# </CalendarViewSnippet>
