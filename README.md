This was adapted from Microsoft tutorial available [here](https://docs.microsoft.com/en-us/graph/tutorials/python)

Register your app first at Azure as described [here](https://docs.microsoft.com/en-us/graph/tutorials/python?tutorial-step=2)

Then [create oauth_settings.yml file](https://docs.microsoft.com/en-us/graph/tutorials/python?tutorial-step=3) 

You also need to create settings.yml file with your SECRET_KEY. [Example file](https://github.com/microsoftgraph/msgraph-training-pythondjangoapp/blob/main/demo/graph_tutorial/graph_tutorial/settings.py)
from MS Github page.

### Tested on Ubuntu 20.04

# Installation
## Python virtual environment
Install the python virtual environment in the project root directory

```python3 -m venv venv```

Activate the venv

```source venv/bin/activate```

Install the python requirements

```pip install -r requirements.txt```

Run the web app on your local machine
```python manage.py runserver```

Create the directory called "emls". This is where the .eml files will be written.

You should now access http://localhost:8000 and login with the user for which you want to download emails.

After successful login click "Download emails" and you should see the output in your browser.

## Docker

If you have docker installed build the image:

```docker build -t dl365mail .```

Then run it:

```docker run -it --rm -p 8000:8000 -v ~/emails:/code/emls dl365mail```

This will create a folder in your $HOME where emails will be stored. You should change the
permissions of this folder to your user.

```chown -R user emails```

Now access http://localhost:8000 as described above.
