#!/usr/bin/python3 
# execution with python3 csv-to-googleslide.py -h

# python3 csv-to-googleslide.py -f "./example/example.csv"
# python3 csv-to-googleslide.py -f "./example/example.csv" -n "name_of_presentation"

import os.path
import csv
import requests
import argparse
from googleapiclient.discovery import build
from google.oauth2 import service_account
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from oauth2client import service_account
from oauth2client.service_account import ServiceAccountCredentials
from google_auth_oauthlib.flow import Flow
import uuid
import time
import random


# for create credentials Oauth, export json key, rename json to credentials_file.json, copy in / 
# https://developers.google.com/workspace/guides/create-credentials?hl=fr#oauth-client-id

# If modifying these scopes, delete the file token.json.

# https://developers.google.com/identity/protocols/oauth2/scopes?hl=fr#slides
SCOPES = ['https://www.googleapis.com/auth/presentations']
# The ID of a sample presentation.
PRESENTATION_ID = '1EAYk18WDjIG-zp_0vLm3CsfQh_i8eXc67Jo2O9C6Vuc'


# credentials for connect drive API, application desktop, Client Oauth2.0
# credentials_file = 'credentials_file.json'

credentials_file = 'credentials_file.json'

# need to activate Google Drive API in console.cloud.google.com
# need to activate Google Sheet API in console.cloud.google.com
# need to activate Google Slide API in console.cloud.google.com


# def create_powerpoint(csv_file, sheet_name):
# main fonction 
def create_powerpoint(csv_file, presentation_name,service):
    
    #create_presentation(service, presentation_name)
    presentation_id = create_presentation(service, presentation_name) # execute and asign an id presentation 
    
    # manage csv 
    with open(csv_file, mode='r', encoding='utf-8') as f:

        reader = csv.DictReader(f) # basic fonctionnal // , skipinitialspace=True for skip space in title row
        rows = list(reader)  # read all rows into a list # for reverse 
        
        current_part = None
        
        for row in reversed(rows): # for reverse row
            
            # if new part, create slide with title             
            #if row['Intitulé partie'] != current_part:
                #current_part = row['Intitulé partie']
                #presentation_theme_slide_id = add_slide(service, presentation_id)
                #add_title_for_presentation_slide(service, presentation_id, presentation_theme_slide_id, row['Intitulé partie'])
            
            # create new slide 
            slide_id = add_slide(service, presentation_id)
            
            # add principal title on this slide # print(row['Chapter'])
            add_title_lvl1_on_slide(service, presentation_id, slide_id, row['Intitulé partie']) # end : x & y 
            add_title_lvl2_on_slide(service, presentation_id, slide_id, row['Intitulé objectif d\'apprentissage'], row['Objectif d\'apprentissage'])
            
            # juste shape for copy text
            add_shape_for_past_text(service, presentation_id, slide_id)
            
            
            # sleep randomly so as not to exceed the number of requests per minute defined by google (between 1 and 2 second)
            time.sleep(random.randint(1, 2))
            
            
            

# function for connect service to google slide 
def get_slides_service(credentials_file):
    
    creds = None
  
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_file, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('slides', 'v1', credentials=creds)
        return service 
    except HttpError as err:
        print(err)


# https://towardsdatascience.com/creating-charts-in-google-slides-with-python-896758e9bc49
# https://developers.google.com/slides/api/guides/presentations
def create_presentation(service, presentation_name):
    body = {
        'title': presentation_name
    }
    presentation = service.presentations().create(body=body).execute()

    # Return the ID of the new presentation
    return presentation['presentationId']


def add_slide(service, presentation_id):
    
    response = service.presentations().batchUpdate(presentationId=presentation_id, body={
        'requests': [{
            'createSlide': {
                'insertionIndex': 0
            }
        }]
    }).execute()
    
    slide_id=  response['replies'][0]['createSlide']['objectId']
    return slide_id 

             
# title on slide # Intitulé partie
def add_title_lvl1_on_slide(service, presentation_id, slide_id, title):
    
    # Manage reference, remove space in caractere chain
    #titleWithoutSpace = title.replace(" ", "") # methods for replace space caracter 
    #referenceObjectId = slide_id+titleWithoutSpace
    
    referenceObjectId = str(uuid.uuid4()) # create uniq string 
    
    # Create a new text box with the given title text
    requests = [
        {
            'createShape': {
                "objectId": referenceObjectId, # reference for texte box 
                'shapeType': 'TEXT_BOX',
                'elementProperties': {
                    'pageObjectId': slide_id,
                    'size': {
                        'height': {
                            'magnitude': 50,
                            'unit': 'PT'
                        },
                        'width': {
                            'magnitude': 600,
                            'unit': 'PT'
                        }
                    },
                    'transform': {
                        'scaleX': 1,
                        'scaleY': 1,
                        'translateX': 20,
                        'translateY': 20,
                        'unit': 'PT'
                    }
                }
            }
        },
        {
            'insertText': {
                'objectId': referenceObjectId, #write in text box 
                'insertionIndex': 0,
                'text': title, # insert txt
            }
        },
        {
        'updateTextStyle': {
            'objectId': referenceObjectId,
            'textRange': {
                'type': 'ALL'
            },
            'style': {
                'fontSize': {
                    'magnitude': 25,
                    'unit': 'PT'
                },
                'foregroundColor': {
                    'opaqueColor': {
                        'rgbColor': {
                            'red': 0.0,
                            'green': 0.0,
                            'blue': 0.0
                        }
                    }
                }
            },
            'fields': 'fontSize,foregroundColor'
        }
        }
    ]
    
    # Execute the requests to add the title to the slide
    body = {
        'requests': requests
    }
    response = service.presentations().batchUpdate(presentationId=presentation_id, body=body).execute()
    
    # update TextBoxId With Style 
    #text_style_update(service, presentation_id, slide_id,text_box_id)
    
    
    # Return the ID of the new text box
    text_box_id = response['replies'][0]['createShape']['objectId']
    
    return text_box_id   


# under part # Intitulé objectif d'apprentissage & Objectif d'apprentissage
def add_title_lvl2_on_slide(service, presentation_id, slide_id, text, part):
    
    # Manage reference, remove space in caractere chain
    #textWithoutSpace = text.replace(" ", "") # methods for replace space caracter 
    #referenceObjectId = slide_id+textWithoutSpace
    
    referenceObjectId = str(uuid.uuid4()) # create uniq string 
    
    # Create a new text box with the given title text
    requests = [
        {
            'createShape': {
                "objectId": referenceObjectId, # reference for texte box 
                'shapeType': 'TEXT_BOX',
                'elementProperties': {
                    'pageObjectId': slide_id,
                    'size': {
                        'height': {
                            'magnitude': 50,
                            'unit': 'PT'
                        },
                        'width': {
                            'magnitude': 600,
                            'unit': 'PT'
                        }
                    },
                    'transform': {
                        'scaleX': 1,
                        'scaleY': 1,
                        'translateX': 20,
                        'translateY': 60,
                        'unit': 'PT'
                    }
                }
            }
        },
        {
            'insertText': {
                'objectId': referenceObjectId, #write in text box 
                'insertionIndex': 0,
                'text': part+". "+text, # insert txt
            }
        },
        {
        'updateTextStyle': {
            'objectId': referenceObjectId,
            'textRange': {
                'type': 'ALL'
            },
            'style': {
                'fontSize': {
                    'magnitude': 17,
                    'unit': 'PT'
                },
                'foregroundColor': {
                    'opaqueColor': {
                        'rgbColor': {
                            'red': 0.37,
                            'green': 0.37,
                            'blue': 0.37
                        }
                    }
                }
            },
            'fields': 'fontSize,foregroundColor'
        }
        }
    ]
    
    # Execute the requests to add the title to the slide
    body = {
        'requests': requests
    }
    response = service.presentations().batchUpdate(presentationId=presentation_id, body=body).execute()
    
    # Return the ID of the new text box
    text_box_id = response['replies'][0]['createShape']['objectId']
    return text_box_id   






def add_shape_for_past_text(service, presentation_id, slide_id):
    
    # Manage reference, remove space in caractere chain
    #textWithoutSpace = text.replace(" ", "") # methods for replace space caracter 
    #referenceObjectId = slide_id+textWithoutSpace
    
    referenceObjectId = str(uuid.uuid4()) # create uniq string 
    
    # Create a new text box with the given title text
    requests = [
        {
            'createShape': {
                "objectId": referenceObjectId, # reference for texte box 
                'shapeType': 'TEXT_BOX',
                'elementProperties': {
                    'pageObjectId': slide_id,
                    'size': {
                        'height': {
                            'magnitude': 230,
                            'unit': 'PT'
                        },
                        'width': {
                            'magnitude': 650,
                            'unit': 'PT'
                        }
                    },
                    'transform': {
                        'scaleX': 1,
                        'scaleY': 1,
                        'translateX': 20,
                        'translateY': 120,
                        'unit': 'PT'
                    }
                }
            }
        },
        #{
            #'insertText': {
                #'objectId': referenceObjectId, #write in text box 
                #'insertionIndex': 0,
                #'text': part+". "+text, # insert txt
            #}
        #}
    ]
    
    # Execute the requests to add the title to the slide
    body = {
        'requests': requests
    }
    response = service.presentations().batchUpdate(presentationId=presentation_id, body=body).execute()
    
    # Return the ID of the new text box
    text_box_id = response['replies'][0]['createShape']['objectId']
    return text_box_id   

      
# for add title of presentation_slide 
         
# title on slide # Intitulé partie
def add_title_for_presentation_slide(service, presentation_id, slide_id, title):
    
    # Manage reference, remove space in caractere chain
    #titleWithoutSpace = title.replace(" ", "") # methods for replace space caracter 
    #referenceObjectId = slide_id+titleWithoutSpace
    
    referenceObjectId = str(uuid.uuid4()) # create uniq string 
    
    # Create a new text box with the given title text
    requests = [
        {
            'createShape': {
                "objectId": referenceObjectId, # reference for texte box 
                'shapeType': 'TEXT_BOX',
                'elementProperties': {
                    'pageObjectId': slide_id,
                    'size': {
                        'height': {
                            'magnitude': 50,
                            'unit': 'PT'
                        },
                        'width': {
                            'magnitude': 600,
                            'unit': 'PT'
                        }
                    },
                    'transform': {
                        'scaleX': 1,
                        'scaleY': 1,
                        'translateX': 50,
                        'translateY': 160,
                        'unit': 'PT'
                    }
                }
            }
        },
        {
            'insertText': {
                'objectId': referenceObjectId, #write in text box 
                'insertionIndex': 0,
                'text': title, # insert txt
            }
        },
        {
        'updateTextStyle': {
            'objectId': referenceObjectId,
            'textRange': {
                'type': 'ALL'
            },
            'style': {
                'fontSize': {
                    'magnitude': 55,
                    'unit': 'PT'
                },
                'foregroundColor': {
                    'opaqueColor': {
                        'rgbColor': {
                            'red': 0.0,
                            'green': 0.0,
                            'blue': 0.0
                        }
                    }
                }
            },
            'fields': 'fontSize,foregroundColor'
        }
        }
    ]
    
    # Execute the requests to add the title to the slide
    body = {
        'requests': requests
    }
    response = service.presentations().batchUpdate(presentationId=presentation_id, body=body).execute()
    
    # update TextBoxId With Style 
    #text_style_update(service, presentation_id, slide_id,text_box_id)
    
    
    # Return the ID of the new text box
    text_box_id = response['replies'][0]['createShape']['objectId']
    
    return text_box_id   


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Manage Presentation with CSV')
    parser.add_argument('-f', '--file', help='File containing the csv, see example in /example')
    parser.add_argument('-n', '--name', help='Name of google sheet', default='presentation-python-builder')
    
    args = parser.parse_args()
    
    
    


if args.file and args.name:
    service = get_slides_service(credentials_file)
    create_powerpoint(args.file,args.name,service)
elif args.name:
    print('No file specified. Use -n or --name to specify a name for google sheet.')
else:
    print('No file specified. Use -f or --file to specify a file.')


