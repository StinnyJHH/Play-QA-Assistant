import googleapiclient.discovery
import pytz
from google.oauth2 import service_account
from pprint import pprint
from datetime import datetime

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE ='credentials.json'

creds = None
creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

service = googleapiclient.discovery.build('sheets', 'v4', credentials=creds)

# ID of the spreadsheet
ss_id = '--SPREADSHEET ID--'

# Clear Analytics Sheet before performing operations
range_ = 'Analytics'
include_grid_data = False

# Clear Analytics Sheet data before performing operations
clear_values_request_body = {}

request = service.spreadsheets().values().clear(
    spreadsheetId=ss_id,
    range=range_,
    body=clear_values_request_body)
response = request.execute()

# Get sheet data
request = service.spreadsheets().get(
    spreadsheetId=ss_id,
    ranges=range_,
    includeGridData=include_grid_data)
response = request.execute()

# If charts already exist, send request to remove them
if (response["sheets"] is not None and
  response["sheets"][0] is not None and
  (response["sheets"][0].get("charts") is not None)):
  for chart in response["sheets"][0].get("charts"):
    id = chart.get("chartId")
    chart_delete_body = {
      "requests": [
        {
          "deleteEmbeddedObject": {
            "objectId": id
          }
        }
      ]
    }
    response = service.spreadsheets().batchUpdate(
      spreadsheetId=ss_id,
      body=chart_delete_body).execute()

read_range = 'Version Tracker'

# Read in data from sheet
request = service.spreadsheets().values().get(
  spreadsheetId=ss_id,
  range=read_range)
response = request.execute()
data = response.get('values', [])

devs = []
devs_prior = []
category = []
priority = []
confirmed = []
rows_to_remove = []

for row in range(len(data)):
  try:
    devs.append(data[row][1])
  except IndexError:
    devs.append("")
  try:
    devs_prior.append(data[row][4])
  except IndexError:
    devs_prior.append("")
  try:
    category.append(data[row][2])
  except IndexError:
    print("skip")
  try:
    priority.append(data[row][4])
  except IndexError:
    priority.append("")

# Dictionaries to hold values
dev_dict = {}
main_dict = {}
p0_dict = {}
p1_dict = {}
p2_dict = {}

# Storing data to dictionaries for devs bug count
for x in range(len(data)):
  currDevs = devs[x].split('/')
  currPrior = devs_prior[x]
  if(currPrior == 'P0' or currPrior == 'P1' or currPrior == 'P2'):
    for y in range (len(currDevs)):
      currDev = currDevs[y].strip()
      if(not(currDev in dev_dict)):
        dev_dict[currDev] = 1
      else:
        dev_dict.update({currDev: dev_dict.get(currDev) + 1})

# Storing data to dictionaries
for x in range(len(data)):
  # Current category
  curr = category[x]
  if curr == '':
    curr = "Other"
  # Current priority
  currPrior = priority[x]
  # Filter out bugs with 'Confirmed', 'Fixed', and 'Questions' tag
  if (currPrior == 'P0' or currPrior == 'P1' or currPrior == 'P2'):
    # First time seeing the bug category
    if(not(curr in main_dict)):
      main_dict[curr] = 1
      if(currPrior == 'P0'):
        p0_dict[curr] = 1
      elif(currPrior == 'P1'):
        p1_dict[curr] = 1
      elif(currPrior == 'P2'):
        p2_dict[curr] = 1
    else:
      main_dict.update({curr: main_dict.get(curr) + 1})
      if currPrior == 'P0':
        if(curr in p0_dict):
          p0_dict.update({curr: p0_dict.get(curr)+1})
        else:
          # First time seeing the category in P0
          p0_dict[curr] = 1
      elif currPrior == 'P1':
        if(curr in p1_dict):
          p1_dict.update({curr: p1_dict.get(curr)+1})
        else:
          # First time seeing the category in P1
          p1_dict[curr] = 1
      elif currPrior == 'P2':
        if(curr in p2_dict):
          p2_dict.update({curr: p2_dict.get(curr)+1})
        else:
          # First time seeing the category in P2
          p2_dict[curr] = 1
  elif currPrior == 'Confirmed':
    confirmed.append(data[x])
    rows_to_remove.append(x)

# Arrays to help formatting for write to Google Sheet
devs = []
main = []
p0 = []
p1 = []
p2 = []

# Format for write in Google Sheet
for key in dev_dict.keys():
  devs.append([key, dev_dict.get(key)])
for key in main_dict.keys():
  main.append([key, main_dict.get(key)])
for key in p0_dict.keys():
  p0.append([key, p0_dict.get(key)])
for key in p1_dict.keys():
  p1.append([key, p1_dict.get(key)])
for key in p2_dict.keys():
  p2.append([key, p2_dict.get(key)])
devs.sort()
main.sort()
p0.sort()
p1.sort()
p2.sort()

# Date format
EST = pytz.timezone('EST')
date = str(datetime.now(EST).strftime('%Y-%m-%d'))

# Writing confirmed bugs
if len(confirmed) != 0:
  #if(len(response) == 0 or len(response[0]) == 0 or response[0][0] != date):
  request_body = {
    'requests':[
      {
        'insertDimension':{
          'range':{
            "sheetId": --SHEET ID HERE--,
            "dimension": 'ROWS',
            "startIndex": 0,
            "endIndex": 1
          }
        }
      }
    ]
  }
  service.spreadsheets().batchUpdate(
    spreadsheetId=ss_id,
    body=request_body).execute()
  request = service.spreadsheets().values().append(
    spreadsheetId=ss_id,
    range="Archive!A1",
    valueInputOption="USER_ENTERED",
    insertDataOption="INSERT_ROWS",
    body={"values":[["Date", date],["From", "To", "Category", "Item", ":", "Note", "Reference"]]}).execute()


  request_body = {
    'requests':[
      {
        'insertDimension':{
          'range':{
            "sheetId": --SHEET ID HERE--,
            "dimension": 'ROWS',
            "startIndex": 1,
            "endIndex": 1
          }
        }
      }
    ]
  }
  service.spreadsheets().batchUpdate(
    spreadsheetId=ss_id,
    body=request_body).execute()
  request = service.spreadsheets().values().append(
    spreadsheetId=ss_id,
    range="Archive!A2",
    valueInputOption="USER_ENTERED",
    insertDataOption="INSERT_ROWS",
    body={"values":confirmed}).execute()

# Remove confirmed bug rows from Version Tracker Sheet
if len(rows_to_remove) != 0:
  request_body = {
    'requests':[]
  }

  count = 0
  for row in rows_to_remove:
    request_body.get('requests').append(
      {
        'deleteDimension':{
          'range':{
            "sheetId": --SHEET ID HERE--,
            "dimension": 'ROWS',
            "startIndex": row-count,
            "endIndex": row+1-count
          }
        }
      }
    )
    count+=1

  service.spreadsheets().batchUpdate(
    spreadsheetId=ss_id,
    body=request_body).execute()

# Writing column headers
service.spreadsheets().values().update(
  spreadsheetId=ss_id,
  range="Analytics!A1",
  valueInputOption="USER_ENTERED",
  body={"values":[["Devs", "Bug Count", "",
  "Overall Category", "Count" , "",
  "P0", "Count", "",
  "P1", "Count", "",
  "P2", "Count"]]}
  ).execute()

# Writing dev bug counts
service.spreadsheets().values().update(
  spreadsheetId=ss_id,
  range="Analytics!A2",
  valueInputOption="USER_ENTERED",
  body={"values":devs}).execute()

# Writing overall stats
service.spreadsheets().values().update(
  spreadsheetId=ss_id,
  range="Analytics!D2",
  valueInputOption="USER_ENTERED",
  body={"values":main}).execute()

# Writing P0 data
service.spreadsheets().values().update(
  spreadsheetId=ss_id,
  range="Analytics!G2",
  valueInputOption="USER_ENTERED",
  body={"values":p0}).execute()

# Writing P1 data
service.spreadsheets().values().update(
  spreadsheetId=ss_id,
  range="Analytics!J2",
  valueInputOption="USER_ENTERED",
  body={"values":p1}).execute()

#Writing P2 data
service.spreadsheets().values().update(
  spreadsheetId=ss_id,
  range="Analytics!M2",
  valueInputOption="USER_ENTERED",
  body={"values":p2}).execute()

max_len = len(main)
if max_len < len(devs):
  max_len = len(devs)

# Request body for charts
request_body = {
  "requests": [
    # Devs Chart
    {
      "addChart": {
        "chart":{
          "spec": {
            "title": "Bugs per Dev",
            "pieChart": {
              "legendPosition": "RIGHT_LEGEND",
              "threeDimensional": False,
              "domain": {
                "sourceRange": {
                  "sources": [
                    {
                      "sheetId": '--SHEET ID HERE--',
                      "startRowIndex": 0,
                      "endRowIndex": len(devs)+1,
                      "startColumnIndex": 0,
                      "endColumnIndex": 1
                    }
                  ]
                }
              },
              "series": {
                "sourceRange": {
                  "sources": [
                    {
                      "sheetId": '--SHEET ID HERE--',
                      "startRowIndex": 0,
                      "endRowIndex": len(devs)+1,
                      "startColumnIndex": 1,
                      "endColumnIndex": 2
                    }
                  ]
                }
              },
            }
          },
          "position": {
            "overlayPosition": {
              "anchorCell": {
                "sheetId": '--SHEET ID HERE--',
                "rowIndex": max_len+2,
                "columnIndex": 0
              },
              "offsetXPixels": 0,
              "offsetYPixels": 0
            }
          }
        }
      }
    },
    # Summary Chart
    {
      "addChart":{
        "chart":{
          "spec": {
            "title": "Overall Summary",
            "pieChart": {
              "legendPosition": "RIGHT_LEGEND",
              "threeDimensional": False,
              "domain": {
                "sourceRange": {
                  "sources": [
                    {
                      "sheetId": 'SHEET ID HERE',
                      "startRowIndex": 0,
                      "endRowIndex": len(main)+1,
                      "startColumnIndex": 3,
                      "endColumnIndex": 4
                    }
                  ]
                }
              },
              "series": {
                "sourceRange": {
                  "sources": [
                    {
                      "sheetId": 'SHEET ID HERE',
                      "startRowIndex": 0,
                      "endRowIndex": len(main)+1,
                      "startColumnIndex": 4,
                      "endColumnIndex": 5
                    }
                  ]
                }
              },
            }
          },
          "position": {
            "overlayPosition": {
              "anchorCell": {
                "sheetId": 'SHEET ID HERE',
                "rowIndex": max_len+2,
                "columnIndex": 7
              },
              "offsetXPixels": 0,
              "offsetYPixels": 0
            }
          }
        }
      }
    },
    # P0 Chart
    {
      "addChart":{
        "chart":{
          "spec": {
            "title": "P0 Summary",
            "pieChart": {
              "legendPosition": "RIGHT_LEGEND",
              "threeDimensional": False,
              "domain": {
                "sourceRange": {
                  "sources": [
                    {
                      "sheetId": 'SHEET ID HERE',
                      "startRowIndex": 0,
                      "endRowIndex": len(p0)+1,
                      "startColumnIndex": 6,
                      "endColumnIndex": 7
                    }
                  ]
                }
              },
              "series": {
                "sourceRange": {
                  "sources": [
                    {
                      "sheetId": 'SHEET ID HERE',
                      "startRowIndex": 0,
                      "endRowIndex": len(p0)+1,
                      "startColumnIndex": 7,
                      "endColumnIndex": 8
                    }
                  ]
                }
              },
            }
          },
          "position": {
            "overlayPosition": {
              "anchorCell": {
                "sheetId": 'SHEET ID HERE',
                "rowIndex": max_len+2,
                "columnIndex": 13
              },
              "offsetXPixels": 0,
              "offsetYPixels": 0
            }
          }
        }
      }
    },
    # P1 Chart
    {
      "addChart":{
        "chart":{
          "spec": {
            "title": "P1 Summary",
            "pieChart": {
              "legendPosition": "RIGHT_LEGEND",
              "threeDimensional": False,
              "domain": {
                "sourceRange": {
                  "sources": [
                    {
                      "sheetId": 'SHEET ID HERE',
                      "startRowIndex": 0,
                      "endRowIndex": len(p1)+1,
                      "startColumnIndex": 9,
                      "endColumnIndex": 10
                    }
                  ]
                }
              },
              "series": {
                "sourceRange": {
                  "sources": [
                    {
                      "sheetId": 'SHEET ID HERE',
                      "startRowIndex": 0,
                      "endRowIndex": len(p1)+1,
                      "startColumnIndex": 10,
                      "endColumnIndex": 11
                    }
                  ]
                }
              },
            }
          },
          "position": {
            "overlayPosition": {
              "anchorCell": {
                "sheetId": 'SHEET ID HERE',
                "rowIndex": max_len+20,
                "columnIndex": 0
              },
              "offsetXPixels": 0,
              "offsetYPixels": 0
            }
          }
        }
      }
    },
    # P2 Chart
    {
      "addChart":{
        "chart":{
          "spec": {
            "title": "P2 Summary",
            "pieChart": {
              "legendPosition": "RIGHT_LEGEND",
              "threeDimensional": False,
              "domain": {
                "sourceRange": {
                  "sources": [
                    {
                      "sheetId": 'SHEET ID HERE',
                      "startRowIndex": 0,
                      "endRowIndex": len(p2)+1,
                      "startColumnIndex": 12,
                      "endColumnIndex": 13
                    }
                  ]
                }
              },
              "series": {
                "sourceRange": {
                  "sources": [
                    {
                      "sheetId": 'SHEET ID HERE',
                      "startRowIndex": 0,
                      "endRowIndex": len(p2)+1,
                      "startColumnIndex": 13,
                      "endColumnIndex": 14
                    }
                  ]
                }
              },
            }
          },
          "position": {
            "overlayPosition": {
              "anchorCell": {
                "sheetId": 'SHEET ID HERE',
                "rowIndex": max_len+20,
                "columnIndex": 7
              },
              "offsetXPixels": 0,
              "offsetYPixels": 0
            }
          }
        }
      }
    }
  ]
}

# Request to write charts to Google Sheet
request = service.spreadsheets().batchUpdate(
  spreadsheetId=ss_id,
  body=request_body)
response = request.execute()
