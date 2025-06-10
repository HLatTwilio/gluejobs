def find_page_id(spreadsheet_id, chart_index):
    """
    Look up the chart id of a chart in a spreadsheet
    """      
    try:
    # Fetch chart information from sheets
        # service = build("slides", "v1", credentials=creds)
    
        presentation_info = slides_service.presentations().get(presentationId = spreadsheet_id).execute()
        page_id = presentation_info['slides'][chart_index]['objectId']
        return page_id
        
    except HttpError as error:
        print(F'An error occurred: {error}')

# copy the root slide in the presentation so that the same background and formatting stay the same
def duplicate_slide(presentation_id, root_page_id, new_page_id):
  # pylint: disable=maybe-no-member
  try:
    # Add a slide at index 1 using the predefined
    # 'TITLE_AND_TWO_COLUMNS' layout and the ID page_id.
    requests = [{

      "duplicateObject": {
        "objectId": root_page_id,
        "objectIds": {
          root_page_id: new_page_id
        }

      }
    } ]

    # If you wish to populate the slide with elements,
    # add element create requests here, using the page_id.

    # Execute the request.
    body = {"requests": requests}
    response = (
        slides_service.presentations()
        .batchUpdate(presentationId=presentation_id, body=body)
        .execute()
    )
    create_slide_response = response.get("replies")[0].get("createSlide")

  except HttpError as error:
    print(f"An error occurred: {error}")
    print("Slides not created")
    return error

  return response

def find_chart_id(spreadsheet_id, sheet_name, chart_index):
  """
  Look up the chart id of a chart in a spreadsheet
  """

  try:
    # Fetch chart information from sheets
    sheet = sheets_service.spreadsheets().get(
      spreadsheetId = spreadsheet_id,
      ranges = [sheet_name]).execute().get('sheets')[0]
      
    chart_id_sheets = sheet['charts'][chart_index]['chartId']

    return chart_id_sheets

  except HttpError as error:
    print(F'An error occurred: {error}')


# copy the root presentation so that the same background and formatting stay the same
def copy_presentation(presentation_id, copy_title):
  """
  Creates the copy Presentation the user has access to.
  Load pre-authorized user credentials from the environment.
  TODO(developer) - See https://developers.google.com/identity
  for guides on implementing OAuth2 for the application.
  """
  # pylint: disable=maybe-no-member
  try:
    body = {"name": copy_title}
    drive_response = (
        drive_service.files().copy(fileId=presentation_id, body=body).execute()
    )
    presentation_copy_id = drive_response.get("id")
    print(f"Duplicated presentation with ID:{presentation_copy_id}")
  except HttpError as error:
    print(f"An error occurred: {error}")
    print("Presentations  not copied")
    return error

  return presentation_copy_id

def search_file():
  """Search file in drive location

  Load pre-authorized user credentials from the environment.
  TODO(developer) - See https://developers.google.com/identity
  for guides on implementing OAuth2 for the application.
  """
  try:
    # create drive api client
    files = []
    page_token = None
    while True:
      # pylint: disable=maybe-no-member
      response = (
          drive_service.files()
          .list(
              q="mimeType contains 'presentation'",
              spaces="drive",
              fields="nextPageToken, files(id, name, permissions)",
              pageToken=page_token,
          )
          .execute()
      )
      # for file in response.get("files", []):
      #   # Process change
      #   print(f'Found file: {file.get("name")}, {file.get("id")}')
      files.extend(response.get("files", []))
      page_token = response.get("nextPageToken", None)
      if page_token is None:
        break

  except HttpError as error:
    print(f"An error occurred: {error}")
    files = None

  return files

def get_file_id(file_list, file_name):
    x = None
    for i in range(0, len(file_list)):
        if file_list[i]['name'] == file_name:
            x = file_list[i]['id']
        else:
            break
    return x


def get_file_permisson_id(file_list, file_name):
    x = None
    for i in range(0, len(file_list)):
        if file_list[i]['name'] == file_name:
            x = file_list[i]['permissions'][0]['id']
        else:
            break
    return x

def add_file_permission(file_id, permission):
    permission_ = {'role':'reader','type':'anyone'}
    drive_service.permissions().create(
                fileId=file_id, body=permission_).execute()

def create_textbox_with_text(presentation_id, page_id, element_id, values, position):
    """
    Creates the textbox with text, the user has access to.
    Load pre-authorized user credentials from the environment.
    TODO(developer) - See https://developers.google.com/identity
    for guides on implementing OAuth2 for the application.
    """
    # pylint: disable=maybe-no-member
    if position == 'R':
        translateX = 380
        translateY = 50
    elif position == 'L':
        translateX = 50
        translateY = 50
    elif position == 'M':
        translateX = 215
        translateY = 50
    else:
        translateX = 50
        translateY = 50

    try:
        # Create a new square textbox, using the supplied element ID.
        pt350 = {"magnitude": 350, "unit": "PT"}
        requests = [
            {
                "createShape": {
                    "objectId": element_id,
                    "shapeType": "TEXT_BOX",
                    "elementProperties": {
                        "pageObjectId": page_id,
                        "size": {"height": pt350, "width": pt350},
                        # the below controls the position of texts; more documentation see: https://developers.google.com/slides/api/guides/transform
                        "transform": {
                            "scaleX": 1,
                            "scaleY": 1,
                            "translateX": translateX,
                            "translateY": translateY,
                            "unit": "PT",
                        },
                    },
                }
            },
            # Insert text into the box, using the supplied element ID.
            {
                "insertText": {
                    "objectId": element_id,
                    "insertionIndex": 0,
                    "text": values,
                }
            },
            {
                "updateTextStyle": {
                    "objectId": element_id,
                    "textRange": {
                        "type": "ALL"
                    },
                    "style": {
                        "fontFamily": "Arial",
                        "fontSize": {"magnitude": 8, "unit": "PT"},
                        "foregroundColor": {
                            "opaqueColor": {
                                "rgbColor": {
                                    "red": 0.0,
                                    "green": 0.0,
                                    "blue": 0.0
                                }
                            }
                        }
                    },
                    "fields": "fontFamily,fontSize,foregroundColor",
                }
            }
        ]

        # Execute the request.
        body = {"requests": requests}
        response = (
            slides_service.presentations()
            .batchUpdate(presentationId=presentation_id, body=body)
            .execute()
        )
        create_shape_response = response.get("replies")[0].get("createShape")
        print(f"Inserted analysis with objectID: {(create_shape_response.get('objectId'))}")
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error

    return response

def get_objects_from_slide(presentation_id, slide_id):
    # Retrieve the presentation data
    presentation = slides_service.presentations().get(presentationId=presentation_id).execute()
    
    # Find the specified slide by its ID
    slide = next((slide for slide in presentation.get('slides', []) if slide.get('objectId') == slide_id), None)
    
    if not slide:
        return f"Slide with ID '{slide_id}' not found."
    
    # Collect all object IDs and types in the slide
    objects_info = {}
    for element in slide.get('pageElements', []):
        object_id = element.get('objectId')
        object_type = element.get('shape', {}).get('shapeType', '')  # Default to shape type if it's a shape
        if not object_type:  # Check for image and table
            if 'image' in element:
                object_type = 'IMAGE'
            elif 'table' in element:
                object_type = 'TABLE'
            else:
                object_type = 'UNKNOWN'  # If the object type is unknown
        
        objects_info[object_id] = object_type
    
    return objects_info

def lookup_object_ids_by_type(objects_dict, object_type):
    # Filter the dictionary based on the specified object type
    matching_ids = [object_id for object_id, obj_type in objects_dict.items() if obj_type == object_type]
    return matching_ids

def add_rows_or_columns_to_table(presentation_id, table_object_id, row_or_column, number_of_rc):

    # Create the request to insert rows into the table
    if row_or_column == 'rows':
        requests = [
    
                {
                    "insertTableRows": {
                        "tableObjectId": table_object_id,  # The table object ID
                        "insertBelow": True,  # Row index where new rows should be inserted
                        "number": number_of_rc,  # Number of rows to add
                    }
                }
            
        ]
    elif row_or_column == 'columns':
        requests = [
    
                {
                    "insertTableColumns": {
                        "tableObjectId": table_object_id,  # The table object ID
                        "insertRight": True,  # Row index where new rows should be inserted
                        "number": number_of_rc,  # Number of rows to add
                    }
                }
            
        ]

    # Send the request to Google Slides API
    body = {"requests": requests}
    response = slides_service.presentations().batchUpdate(
        presentationId=presentation_id, body=body).execute()

    return response


def duplicate_table_from_gsheets_to_gslides_wt_formatting(spreadsheet_id, worksheet_index, worksheet_name, worksheet_range, presentation_id, slide_id):
    try:
        gc = gspread.authorize(creds)
        spreadsheet = gc.open_by_url(f'https://docs.google.com/spreadsheets/d/{spreadsheet_id}')
        worksheet = spreadsheet.get_worksheet(worksheet_index)
        range_name = f'{worksheet_name}!{worksheet_range}'
        values = worksheet.get_values(worksheet_range)
        
        # Use Google Sheets API to fetch format data
        sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheet_properties = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id, ranges=range_name, fields="sheets(data(rowData(values(userEnteredFormat))))").execute()

        tab_id = get_sheet_id_by_name(spreadsheet_id, worksheet_name)

        rows = len(values)
        columns = len(values[0])

        object_dict = get_objects_from_slide(presentation_id, slide_id)
        
        table_object_id = lookup_object_ids_by_type(object_dict, 'TABLE')[0]

        rows_to_add = rows - 2
        columns_to_add = columns - 2

        add_rows_or_columns_to_table(presentation_id, table_object_id, 'rows', rows_to_add)
        add_rows_or_columns_to_table(presentation_id, table_object_id, 'columns', columns_to_add)

        requests = []
        # Insert text, alignment, background, and font style into each cell
        for r in range(rows):
            for c in range(columns):
                requests.append({
                    'insertText': {
                        'objectId': table_object_id,
                        'cellLocation': {'rowIndex': r, 'columnIndex': c},
                        'text': values[r][c]
                    }
                })
        
        # Execute the batch requests to modify the slide
        slides_service.presentations().batchUpdate(presentationId=presentation_id, body={'requests': requests}).execute()
    
        row_data = sheet_properties['sheets'][0]['data'][0]['rowData']
        backgrounds = []
        horizontal_alignments = []
        font_weights = []
        text_colors = []
        
        # Extract background colors
        for row in row_data:
            bg_row = []
            align_row = []
            weight_row = []
            color_row = []
        
            for cell in row['values']:
                user_format = cell.get('userEnteredFormat', {})
                text_format = user_format.get('textFormat', {})
                is_bold = text_format.get('bold', False)  # Default is False if not specified
        
                # Get background color
                bg_color = user_format.get('backgroundColor', {})
                
                # Get text color
                text_color = text_format.get('foregroundColor', {})
                
                if bg_color:
                    red = bg_color.get('red', 1.0)
                    green = bg_color.get('green', 1.0)
                    blue = bg_color.get('blue', 1.0)
                    bg_row.append((red, green, blue))
                else:
                    # Default white background
                    bg_row.append((1.0, 1.0, 1.0))
        
                # Extract text color values
                if text_color:
                    red = text_color.get('red', 0.0)
                    green = text_color.get('green', 0.0)
                    blue = text_color.get('blue', 0.0)
                    color_row.append((red, green, blue))
                else:
                    # Default black text
                    color_row.append((0.0, 0.0, 0.0))
                
                align_row.append(user_format.get('horizontalAlignment', 'LEFT'))
                weight_row.append(is_bold)
        
            backgrounds.append(bg_row)
            horizontal_alignments.append(align_row)
            font_weights.append(weight_row)
            text_colors.append(color_row)
            
        # Google Slides: Apply the background color to the table
        rows = len(backgrounds)
        columns = len(backgrounds[0])
        
        # Create a batch of requests to update the table cell background colors
        requests = []

        for r in range(rows):
            for c in range(columns):
                # apply font size
                try:
                    requests.append({
                        "updateTextStyle": {
                            "objectId": table_object_id,
                               "cellLocation": {
                                    "rowIndex": r,
                                    "columnIndex": c
                                },
                                "textRange": {
                                    "type": "ALL"
                                },
                            "style": {
                                "fontFamily": "Arial",
                                "fontSize": {"magnitude": 9, "unit": "PT"}
                            },
                            "fields": "fontFamily,fontSize",
                        }
                    })
                except:
                    skip
        
                # apply background color
                red, green, blue = backgrounds[r][c]  # Extract the RGB color values
                requests.append({
                    'updateTableCellProperties': {
                        'objectId': table_object_id,
                        'tableRange': {
                            'location': {'rowIndex': r, 'columnIndex': c},
                            'rowSpan': 1,
                            'columnSpan': 1
                        },
                        'tableCellProperties': {
                            'tableCellBackgroundFill': {
                                'solidFill': {
                                    'color': {
                                        'rgbColor': {
                                            'red': red,
                                            'green': green,
                                            'blue': blue
                                        }
                                    }
                                }
                            }
                        },
                        'fields': 'tableCellBackgroundFill.solidFill.color'
                    }
                })
        
                # Apply text color
                red, green, blue = text_colors[r][c]  # Extract the RGB color values
                try:
                    requests.append({
                        'updateTextStyle': {
                            'objectId': table_object_id,
                            'cellLocation': {
                                'rowIndex': r,
                                'columnIndex': c
                            },
                            'textRange': {
                                'type': 'ALL'  # Applies to all text in the cell
                            },
                            'style': {
                                "foregroundColor": {
                                    "opaqueColor": {
                                        "rgbColor": {
                                            "blue": blue,
                                            "green": green,
                                            "red": red,
                                        }
                                    }
                                }
                            },
                            'fields': 'foregroundColor'
                        }
                    })
                except:
                    skip
                
                # Apply bold styling
                is_bold = font_weights[r][c]
                try:
                    requests.append({
                        'updateTextStyle': {
                            'objectId': table_object_id,
                            'cellLocation': {'rowIndex': r, 'columnIndex': c},
                            'style': {
                                'bold': is_bold
                            },
                            'fields': 'bold'
                        }
                    })
                except:
                    skip
        
                # Apply text alignment
                alignment = {
                    'LEFT': 'START',
                    'CENTER': 'CENTER',
                    'RIGHT': 'END'
                }.get(horizontal_alignments[r][c], 'START')
        
                requests.append({
                    'updateParagraphStyle': {
                        'objectId': table_object_id,
                        'cellLocation': {'rowIndex': r, 'columnIndex': c},
                        'style': {
                            'alignment': alignment
                        },
                        'fields': 'alignment'
                    }
                })
                # apply border color 
                requests.append({
                        "updateTableBorderProperties": {
                          "objectId": table_object_id,
                          "tableBorderProperties": {
                            "tableBorderFill": {
                              "solidFill": {
                                "color": {
                                  "rgbColor": {
                                    "blue": 0.3,
                                    "green": 0.7
                                  }
                                }
                              }
                            }
                          },
                          "fields": "tableBorderFill.solidFill.color"
                        }
                      }
                    
                  })
                
                # adjust table size
                requests.append({
                        "updateTableRowProperties": {
                            "objectId": table_object_id,
                            "rowIndices": r,  # Specify which row(s) to resize
                            "tableRowProperties": {
                                "minRowHeight": {
                                    "magnitude": 3000,  # Row height in EMUs (1 inch = 914400 EMUs)
                                    "unit": "EMU"
                                }
                            },
                            "fields": "minRowHeight"
                        }
                    })
                if c == 0:
                    requests.append({
                            "updateTableColumnProperties": {
                                "objectId": table_object_id,
                                "columnIndices": c,  # Specify which column(s) to resize
                                "tableColumnProperties": {
                                    "columnWidth": {
                                        "magnitude": 1006400,  # Column width in EMUs (1 inch = 914400 EMUs)
                                        "unit": "EMU"
                                    }
                                },
                                "fields": "columnWidth"
                            }
                        })
                else:
                    requests.append({
                            "updateTableColumnProperties": {
                                "objectId": table_object_id,
                                "columnIndices": c,  # Specify which column(s) to resize
                                "tableColumnProperties": {
                                    "columnWidth": {
                                        "magnitude": 402560,  # Column width in EMUs (1 inch = 914400 EMUs)
                                        "unit": "EMU"
                                    }
                                },
                                "fields": "columnWidth"
                            }
                        })
                
                # move the table to the desired position
                requests.append({
                        "updatePageElementTransform": {
                            "objectId": table_object_id,
                            "transform": {
                                "scaleX": 1.0,  # Keep the same scale for width
                                "scaleY": 1.0,  # Keep the same scale for height
                                "translateX": 450000,  # Move X position in EMUs (1 inch = 914400 EMUs)
                                "translateY": 900000,  # Move Y position in EMUs
                                "unit": "EMU"  # Unit is in English Metric Units (EMU)
                            },
                            "applyMode": "ABSOLUTE"
                        }
                    })
        
        # Execute the batch requests to modify the slide
        slides_service.presentations().batchUpdate(presentationId=presentation_id, body={'requests': requests}).execute()

        link_request = {
            "updateTextStyle": {
                "objectId": table_object_id,
                "cellLocation": {
                    "rowIndex": 0,
                    "columnIndex": 0
                },
                "textRange": {
                    "type": "ALL"
                },
                "style": {
                    "link": {
                        "url": f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit#gid={tab_id}&range={worksheet_range}"
                    }
                },
                "fields": "link"
            }
        }
        
        # Execute the request to add the link to the entire table
        body = {'requests': [link_request]}
        
        slides_service.presentations().batchUpdate(presentationId=presentation_id, body=body).execute()
        
        print(f"Table duplicated from spreadsheet {spreadsheet_id} to presentation {presentation_id}")
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error

def read_gsheet_to_md(spreadsheet_id, worksheet_index):

    gc = gspread.authorize(creds)
    spreadsheet = gc.open_by_url(f'https://docs.google.com/spreadsheets/d/{spreadsheet_id}')
    worksheet = spreadsheet.get_worksheet(worksheet_index)
    data = worksheet.get_all_values()

    headers = data.pop(0)
    df = pd.DataFrame(data, columns=headers)

    df_md = df.to_markdown(index=False)
    return df_md

def chatgpt_analysis(md_table):
    try:
        openai_key = os.environ.get('openaikey')
        client=OpenAI(api_key=openai_key)
        
        user_prompt = """
        You are an intelligent analyst that is able to understand and analyze markdown tables containing financial data and then generate insights based on the data in the tables. 
        You learn from the insights generated from previous tables and use that knowledge to generate insights for the next table. 
        DO NOT use any insights from the past, just learn the style from them. 
        And only generate insights based on the data provided in the table for which you are generating the insight. 
        If there is no subjective insight to be generated based on data just use the numbers and create template for user to add dteails. 
        
        The table provided contains a period over period comparison/trend analysis, and the rows are product names as well as some totals and grand totals.
        
        "Comms AWS Credits" is AWS credits Twilio receive.
        "Comms BFCM & AL23 Hosting" is hosting cost which is a type of cost of goods sold that hurt profit margin.
        
        The analysis should focus on product groups that are the top drivers of gross margin increases or decreases.
        
        Focus on the line item 'Total Twilio'. All the line items above it are components of it. Summarize the contributors.
        
        Focus on the line items 'US Messaging', 'International Messaging', and 'Total Messaging'
        
        insight example 1: international messaging has lower gross margin compared to other product groups, so the fact that it's percentage of total dropped contributes to overall higher gross margin.
        insight example 2: $6M Comms AWS credits booked in Q1’24, $4M BFCM/AL23 Hosting booked in Q4’23
        
        Talk about change in profit margin in bps. Talk about absolute dollar values in millions.
        
        Return the top ten insights.

        Please return the output in plain text without any formatting.
        """
        
        content_prompt = f"""
        see below the table {md_table}
        """
        
        completion = client.chat.completions.create(
        model="gpt-4o-mini",
        # response_format={ "type": "json_object" },
        
        messages=[
            {"role": "system","content":user_prompt},
            {"role": "user", "content":content_prompt}
        ], 
        temperature = 0.1, 
        )
        
        output = str(completion.choices[0].message.content)
        
        print("Chatgpt analysis completed.")

        return output
        
        
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error

def move_file_to_folder(file_id, folder_id):
    try:
        # Retrieve the existing parents of the file
        file = drive_service.files().get(fileId=file_id, fields='parents').execute()
        previous_parents = ",".join(file.get('parents'))

        # Move the file to the new folder by updating its parents
        updated_file = drive_service.files().update(
            fileId=file_id,
            addParents=folder_id,
            # removeParents=previous_parents,
            supportsAllDrives=True,  # Include this to handle shared drives
            fields='id, parents'
        ).execute()

        print(f"File '{file_id}' moved to folder '{folder_id}'.")
        return updated_file
    except Exception as e:
        print(f"An error occurred: {e}")
        return None

def get_text_object_id_based_on_text(presentation_id, slide_object_id, first_ten_characters):
    try:
        # Get the presentation
        presentation = slides_service.presentations().get(presentationId=presentation_id).execute()

        # Dictionary to store object information (objectId, object_type, first 5 letters of text)
        object_info = []

        # Iterate through the slides in the presentation
        slides = presentation.get('slides', [])
        for slide_num, slide in enumerate(slides):
            slide_id = slide.get('objectId')
            elements = slide.get('pageElements', [])
            for element in elements:
                object_id = element.get('objectId')
                object_type = element.get('shape', {}).get('shapeType', 'Unknown')

                first_ten_letters = None 
                if 'shape' in element:
                # Check if the element contains text
                    if 'text' in element['shape']:
                        text_elements = element['shape']['text'].get('textElements', [])
                        full_text = ""
    
                        # Loop through text elements to extract text
                        for text_element in text_elements:
                            # Check if 'textRun' exists (which contains the actual text)
                            if 'textRun' in text_element:
                                full_text += text_element['textRun'].get('content', '')
                                
                # Get the first 10 letters of the text content if available
                    first_ten_letters = full_text.strip()[:10] if full_text.strip() else None
                # Add the data to the dictionary
                object_info.append({
                    'slide_id': slide_id,
                    'object_id': object_id,
                    'object_type': object_type,
                    'first_ten_letters': first_ten_letters
                })

        for i in range(0, len(object_info)):
            if object_info[i]['slide_id'] == slide_object_id:
                if object_info[i]['first_ten_letters'] == first_ten_characters:
                    return object_info[i]['object_id']

    except Exception as e:
        print(f"An error occurred: {e}")
        return None

def title_merging(presentation_id,slide_object_id, data_spreadsheet_id, spreadsheet_name_range, text_row_position, text_column_position):
  """
  Run Text merging the user has access to.
  Load pre-authorized user credentials from the environment.
  TODO(developer) - See https://developers.google.com/identity
  for guides on implementing OAuth2 for the application.
  """
  try:
      
      sheets_response = (
        sheets_service.spreadsheets()
        .values()
        .get(spreadsheetId=data_spreadsheet_id, range=spreadsheet_name_range)
        .execute()
    )
      
      title_name = sheets_response.get("values")[text_row_position][text_column_position]


      # Create the text merge (replaceAllText) requests
      # for this presentation.
      requests = [
              {
                  "replaceAllText": {
                      "containsText": {
                          "text": "{{title-input}}",
                          "matchCase": True,
                      },
                      "replaceText": title_name,
                      "pageObjectIds": slide_object_id
                  }
              }
          ]
      body = {"requests": requests}
      response = (
          slides_service.presentations()
          .batchUpdate(presentationId=presentation_id, body=body)
          .execute()
      )

  except HttpError as error:
    print(f"An error occurred: {error}")
    return error
      

def simple_text_replace(presentation_id, shape_id, replacement_text):
  """
  Run simple_text_replace the user has access to.
  Load pre-authorized user credentials from the environment.
  TODO(developer) - See https://developers.google.com/identity
  for guides on implementing OAuth2 for the application.
  """
  try:
    # Remove existing text in the shape, then insert new text.
    requests = []
    requests.append(
        {"deleteText": {"objectId": shape_id, "textRange": {"type": "ALL"}}}
    )
    requests.append(
        {
            "insertText": {
                "objectId": shape_id,
                "insertionIndex": 0,
                "text": replacement_text,
            }
        }
    )

    # Execute the requests.
    body = {"requests": requests}
    response = (
        slides_service.presentations()
        .batchUpdate(presentationId=presentation_id, body=body)
        .execute()
    )
    print(f"Replaced text in shape with ID: {shape_id}")
    return response
  except HttpError as error:
    print(f"An error occurred: {error}")
    print("Text is not merged")
    return error

def chatgpt_summary(all_analysis):
    try:
        openai_key = os.environ.get('openaikey')
        client=OpenAI(api_key=openai_key)
        
        user_prompt = """
        You are an intelligent analyst that is able to summarize financial analyses into concise and crisp short summaries to provide a high level picture of the movements of revenue and gross margin and which product groups stand out. Focus on QoQ and YoY.  
        You learn from the insights generated from previous analyses. 
        DO NOT use any insights from the past, just learn the style from them. 
        And only generate insights based on the analyses provided for which you are generating the insight. 

        The analysis should focus on product groups that are the top drivers of gross margin increases or decreases. Focus on Q3'24 vs Q2'24, and FY24 vs FY23. And when talking about changes, you must include which two periods are being compared.
        
        insight example 1: international messaging revenue up $14M from Q1'22 to Q2'23, gross profit up $6M from Q1'22 to Q2'23.

        insight example 2: FY24 revenue up $29M from Q1'22 to Q2'23, and gross profit (GP) up $7M from Q1'22 vs 5+7
        
        Talk about change in profit margin in bps. Talk about absolute dollar values in millions.
        
        Return the top ten insights.

        Please return the output in plain text without any formatting.
        """
        
        content_prompt = f"""
        see below all the analyses {all_analysis}
        """
        
        completion = client.chat.completions.create(
        model="gpt-4o-mini",
        # response_format={ "type": "json_object" },
        
        messages=[
            {"role": "system","content":user_prompt},
            {"role": "user", "content":content_prompt}
        ], 
        temperature = 0.1, 
        )
        
        output = str(completion.choices[0].message.content)
        
        print("Chatgpt analysis completed.")

        return output
        
        
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error

def move_slide_to_the_last(presentation_id, slide_object_id):
    presentation = slides_service.presentations().get(presentationId=presentation_id).execute()
    slides = presentation.get('slides', [])
    total_slides = len(slides)
    
    # Prepare the request to move the slide
    requests = [
        {
            'updateSlidesPosition': {
                'slideObjectIds': [slide_object_id],
                'insertionIndex': total_slides  # Move to the last position
            }
        }
    ]
    
    # Send the request
    body = {'requests': requests}
    response = slides_service.presentations().batchUpdate(
        presentationId=presentation_id, body=body).execute()
    
    print(f"Slide {slide_object_id} moved to the last position.")

def locate_the_only_black_line_on_first_slide(presentation_id):

    presentation = slides_service.presentations().get(presentationId=presentation_id).execute()
    
    # Iterate through the slides in the presentation
    slide = presentation.get('slides')[0]
    elements = slide.get('pageElements', [])
    for element in elements:
        if 'line' in element:
            red = element['line']['lineProperties']['lineFill']['solidFill']['color']['rgbColor']['red']
            green = element['line']['lineProperties']['lineFill']['solidFill']['color']['rgbColor']['green']
            blue = element['line']['lineProperties']['lineFill']['solidFill']['color']['rgbColor']['blue']
            if red == green == blue:
                object_id = element.get('objectId','')
                return object_id

def delete_object(presentation_id, object_id):
    # Create the request to delete the object
    requests = [
        {
            "deleteObject": {
                "objectId": object_id  # Specify the object to delete
            }
        }
    ]
    # Execute the request
    body = {'requests': requests}
    response = slides_service.presentations().batchUpdate(
        presentationId=presentation_id, body=body).execute()