#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os
import json
import pandas as pd
import io
from fnmatch import fnmatch
from office365.sharepoint.client_context import ClientContext, UserCredential
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.files.file import File

# Check if a column is unique
def is_unique(df, columns):
    # Wenn columns ein String ist, konvertiere ihn zu einer Liste
    if isinstance(columns, str):
        columns = [columns]
    
    # Wenn columns nicht angegeben wird oder eine leere Liste ist, verwende alle Spaltennamen des DataFrames
    if columns is None or not columns:
        columns = df.columns
    
    for column in columns:
        if df[column].duplicated().any():
            print('Die Spalte', column, 'ist nicht eindeutig.')
        else:
            print('Die Spalte', column, 'ist eindeutig.')

# Set index and remove old column
def set_index(df, index_column):
    df.index = df[index_column]
    df = df.drop(index_column, axis=1)
    return df

# Calculate difference of two dates
def date_diff(date1, date2):
    date1 = pd.to_datetime(date1)
    date2 = pd.to_datetime(date2)
    return date1-date2

# Read excel-files
def read_excelfile(dir, file_pattern, sheet_name=None, header=None, cols=None, nrows=None, skiprows=None):
    import io
    import json
    from fnmatch import fnmatch
    from office365.sharepoint.client_context import ClientContext, UserCredential
    from office365.sharepoint.files.file import File
    import pandas as pd
    
    # Lade SharePoint Credentials
    with open('./Keys/sharepoint-creds.json') as file:
        sharepoint_credentials = json.load(file)
        USERNAME = sharepoint_credentials['sharepoint_username']
        PASSWORD = sharepoint_credentials['sharepoint_password']

    SHAREPOINT_SITE = "https://ensoecommerce.sharepoint.com/sites/enso/myEnso_intern"

    # Authentifizierung
    ctx = ClientContext(SHAREPOINT_SITE).with_credentials(UserCredential(USERNAME, PASSWORD))
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()

    # Der Pfad des Ordners und der Datei
    dir_path = dir

    folder = web.get_folder_by_server_relative_url(dir_path)
    ctx.load(folder)
    ctx.execute_query()

    files = folder.files
    ctx.load(files)
    ctx.execute_query()

    matching_files = [f for f in files if fnmatch(f.properties['Name'], file_pattern)]
    sorted_files = sorted(matching_files, key=lambda x: x.properties['TimeLastModified'], reverse=True)

    if not sorted_files:
        print("Keine passende Datei gefunden.")
        return None

    latest_file = sorted_files[0]
    file_url = latest_file.properties["ServerRelativeUrl"]

    # Öffne Binärdatei der neuesten Datei
    response = File.open_binary(ctx, file_url)

    # Speichere Daten im BytesIO Stream
    bytes_file_obj = io.BytesIO()
    bytes_file_obj.write(response.content)
    bytes_file_obj.seek(0)

    # Lese Excel-Datei
    excel_file = pd.ExcelFile(bytes_file_obj)
    
    # Wähle das richtige Blatt, Fehlerbehandlung für nicht existierende Blätter
    if sheet_name and sheet_name not in excel_file.sheet_names:
        print(f"Fehler: Das Blatt '{sheet_name}' existiert nicht in der Datei.")
        return None
    
    if sheet_name is None:
        sheet_name = excel_file.sheet_names[0]  # Standardmäßig das erste Blatt

    # Versuche, die Daten mit den angegebenen Parametern zu lesen
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name, header=header, usecols=cols, engine='openpyxl', nrows=nrows, skiprows=skiprows)
    except ValueError as e:
        print(f"Fehler beim Lesen der Excel-Datei: {e}")
        return None

    print('Readed file:', file_url)
    return df

# Create bar chart
def plot_barchart(x, y, xlabel = 'X-Achse', ylabel = 'Y-Achse', title = None, color='#A73278'):
    import matplotlib.pyplot as plt
    plt.figure(figsize=(10, 6))
    plt.bar(x, y, color=color)

    # Achsentitel und Diagrammtitel hinzufügen
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
    if title == None:
        plt.title(f'Barchart {ylabel} nach {xlabel}')
    else:
        plt.title(title)

    # Datumsformat auf der x-Achse anpassen (optional, je nach Datenformat)
    plt.xticks(rotation=45, ha='right')

    # Diagramm anzeigen
    plt.show()

