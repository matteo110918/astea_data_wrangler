import pandas as pd
from datetime import date


# Funktion um die sturkutrell-falsche Closed Datei zu säubern
def clean_ticket_closed(df_tickets_closed):
    # Drop columns: 'actgr_id', 'bpart_id' and 15 other columns
    df_tickets_closed = df_tickets_closed.drop(columns=['actgr_id', 'bpart_id', 'call_source_id', 'cause_id', 'company state', 'created_by', 'is_credit_hold', 'is_installed', 'item_id', 'open_date', 'open_time', 'orig_install_date', 'pclass2_id', 'resolve_date', 'severity_id', 'Ticket Year', 'Request ID'])
    df_tickets_closed['ticket owner'] = '00KarLuz'
    return df_tickets_closed

# Funktion um die sturkutrell-falsche Closed Datei zu säubern
def clean_spare_parts_closed(df_spare_parts_closed):
    # Rename the columns Part Serial # in Part Serial
    df_spare_parts_closed.rename(columns={'Ticket ': 'Ticket'}, inplace=True)
    # Drop columns: 'Ticket Status', 'Call Type' and 9 other columns
    df_spare_parts_closed = df_spare_parts_closed.drop(columns=['Ticket Status', 'Call Type', 'Customer', 'Fulfilled', 'Model', 'Part Type', 'Part Desc', 'Printer Serial', 'Product Line', 'Ship Due Date', 'Ship Date'])
    
    #### Achtung, die Reihenfolge der Spalten stimmt nicht mit der Schnittstelle überein. Qty an vorletzter Stelle!!!
    
    # Replace gaps back from the next valid value in: 'actual_dt'
    df_spare_parts_closed = df_spare_parts_closed.fillna({'actual_dt': df_spare_parts_closed['actual_dt'].bfill()})
    # Replace gaps back from the next valid value in: 'actual_tm'
    df_spare_parts_closed = df_spare_parts_closed.fillna({'actual_tm': df_spare_parts_closed['actual_tm'].bfill()})
    # Replace all instances of "10-E" with "10-" in column: 'Part ID'
    df_spare_parts_closed['Part ID'] = df_spare_parts_closed['Part ID'].str.replace("10-E", "10-", regex=False)
    return df_spare_parts_closed

# Funktion die saubere Open Datei noch weiter zu säubern
def clean_spare_parts_open(df_spare_parts_open):
    # Rename the columns Part Serial # in Part Serial
    df_spare_parts_open.rename(columns={'Ticket ': 'Ticket' , 'Part Serial #': 'Part Serial'}, inplace=True)
    # Replace gaps back from the next valid value in: 'actual_dt'
    df_spare_parts_open = df_spare_parts_open.fillna({'actual_dt': df_spare_parts_open['actual_dt'].bfill()})
    # Replace gaps back from the next valid value in: 'actual_tm'
    df_spare_parts_open = df_spare_parts_open.fillna({'actual_tm': df_spare_parts_open['actual_tm'].bfill()})
    # Replace all instances of "10-E" with "10-" in column: 'Part ID'
    df_spare_parts_open['Part ID'] = df_spare_parts_open['Part ID'].str.replace("10-E", "10-", regex=False)
    return df_spare_parts_open

# Funktion, die Dataframes in ein Exel zu speichern
def save_to_excel_tickets(dataframe):
        # Search for today's date to embed it in the file name
        today = str(date.today())
        dataframe.to_excel(f'{today}_uploaded_astea_tickets.xlsx', index=False)
        print(f'{today}_uploaded_astea_tickets.xlsx was written to Excel successfully')

def save_to_excel_spares(dataframe):
    # Search for today's date to embed it in the file name
    today = str(date.today())
    dataframe.to_excel(f'{today}_uploaded_astea_spares.xlsx', index=False)
    print(f'{today}_uploaded_astea_spares.xlsx was written to Excel successfully')


### Main Skript

# Prompt File Name
is_resend = input('Do you have a resend file? (y/n)')
# Szenario Unterteilung
if is_resend.lower() == "y":
    # Prompt File Name
    tickets_closed = input('Enter the Name of the Closed Tickets File: ')
    # Einlesen der Dateien und Speicherung der Tabs in jeweilige Dataframes
    df_tickets_closed = pd.read_excel(f'{tickets_closed}', sheet_name='Ticket Summary Data')
    df_spare_parts_closed = pd.read_excel(f'{tickets_closed}', sheet_name='Ticket Parts Data')
    # Funktionen ausführen
    df_tickets_closed_clean = clean_ticket_closed(df_tickets_closed)
    df_spare_parts_closed_clean = clean_spare_parts_closed(df_spare_parts_closed)
    # Ausgabe in Exel
    save_to_excel_tickets(dataframe = df_tickets_closed_clean)
    save_to_excel_spares(dataframe = df_spare_parts_closed_clean)
else:
    # Prompt File Name
    tickets_closed = input('Enter the Name of the Closed Tickets File: ')
    tickets_open = input('Enter the Name of the Open Tickets File: ')
    # Einlesen der Dateien und Speicherung der Tabs in jeweilige Dataframes
    df_tickets_closed = pd.read_excel(f'{tickets_closed}', sheet_name='Ticket Summary Data')
    df_spare_parts_closed = pd.read_excel(f'{tickets_closed}', sheet_name='Ticket Parts Data')
    df_tickets_open = pd.read_excel(f'{tickets_open}', sheet_name='Ticket Summary Data')
    df_spare_parts_open = pd.read_excel(f'{tickets_open}', sheet_name='Ticket Parts Data')
    # Funktionen ausführen
    df_tickets_closed_clean = clean_ticket_closed(df_tickets_closed)
    df_spare_parts_closed_clean = clean_spare_parts_closed(df_spare_parts_closed)
    df_tickets_open_clean = df_tickets_open
    df_spare_parts_open_clean = clean_spare_parts_open(df_spare_parts_open)
    # Hänge zusammengehörende Tabellen untereinander zusammen
    df_tickets_appended = pd.concat([df_tickets_open_clean, df_tickets_closed_clean])
    df_spare_parts_append = pd.concat([df_spare_parts_open_clean ,df_spare_parts_closed_clean])
    # Ausgabe in Exel
    save_to_excel_tickets(dataframe = df_tickets_appended)
    save_to_excel_spares(dataframe = df_spare_parts_append)


    """
    Datei musste noch händisch angepasst werden. Und zwar scheint, dass irgendetwas mit der Close Spare Parts nicht stimmt.
    - Eine Spalte "Ticket " nimmt es mit
    - Es wurden keine Ticket Nummern von der Close Spare Part in das File geschrieben 
    """