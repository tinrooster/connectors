import PySimpleGUI as sg
import pyodbc
import pandas as pd
import os
import json
from typing import List, Optional, Dict, Any, Tuple
import logging
import datetime

# Set up logging at the top of the file
logging.basicConfig(
    filename=f'access_connector_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.log',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class Settings:
    def __init__(self, config_file: str = 'config.json'):
        self.config_file = config_file
        self.settings = {
            'db_path': '',
            'auto_sync': False,
            'sync_interval': 'Daily',
            'selected_table': '',
            'last_connection_status': False,
            'window_size': (800, 750),  # Default window size
            'table_display_rows': 12    # Number of rows to display in table
        }
        self.load_settings()

    def load_settings(self):
        """Load settings from JSON file"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    loaded_settings = json.load(f)
                    self.settings.update(loaded_settings)
        except Exception as e:
            print(f"Error loading settings: {e}")

    def save_settings(self):
        """Save settings to JSON file"""
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self.settings, f, indent=4)
            print("Settings saved successfully")
        except Exception as e:
            print(f"Error saving settings: {e}")
            sg.popup_error(f"Failed to save settings: {e}")

class AccessDatabaseManager:
    def __init__(self, db_path: str):
        self.db_path = db_path
        self.connection = None
        self._build_connection_string()

    def _build_connection_string(self):
        """Build the connection string based on the database path"""
        self.connection_string = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            f'DBQ={self.db_path};'
        )

    def connect(self) -> bool:
        """Establish a connection to the Access database."""
        try:
            if not os.path.exists(self.db_path):
                raise FileNotFoundError(f"Database file not found: {self.db_path}")
            
            self.connection = pyodbc.connect(self.connection_string)
            print("Connected to Access database.")
            return True
        except pyodbc.Error as e:
            print(f"ODBC Error connecting to database: {e}")
            sg.popup_error(f"ODBC Error: {e}")
            return False
        except Exception as e:
            print(f"Error connecting to database: {e}")
            sg.popup_error(f"Connection Error: {e}")
            return False

    def close(self):
        """Close the database connection."""
        try:
            if self.connection:
                self.connection.close()
                print("Connection closed.")
        except Exception as e:
            print(f"Error closing connection: {e}")

    def get_table_names(self) -> List[str]:
        """Get list of tables in the database"""
        try:
            if self.connection:
                cursor = self.connection.cursor()
                tables = [table.table_name for table in cursor.tables(tableType='TABLE')]
                return sorted(tables)  # Return sorted list for better UI experience
            return []
        except Exception as e:
            print(f"Error getting tables: {e}")
            sg.popup_error(f"Failed to get tables: {e}")
            return []

    def get_table_preview(self, table_name: str, limit: int = 1000) -> Tuple[List[str], List[List[Any]]]:
        """Get column names and preview data for a table"""
        try:
            logging.info(f"Getting preview for table: {table_name}")
            
            if not self.connection:
                logging.error("No database connection")
                return [], []
                
            cursor = self.connection.cursor()
            query = f"SELECT TOP {limit} * FROM [{table_name}]"
            logging.debug(f"Executing query: {query}")
            
            cursor.execute(query)
            
            # Get column names
            headers = [column[0] for column in cursor.description]
            logging.debug(f"Retrieved headers: {headers}")
            
            # Fetch all rows
            rows = cursor.fetchall()
            logging.info(f"Retrieved {len(rows)} rows")
            
            # Convert rows to list of lists and convert all values to strings
            data = [[str(cell) if cell is not None else '' for cell in row] for row in rows]
            logging.debug(f"First row of converted data: {data[0] if data else 'No data'}")
            
            return headers, data
            
        except Exception as e:
            logging.exception(f"Error getting table preview: {e}")
            sg.popup_error(f"Failed to get table preview: {e}")
            return [], []

def show_database_settings_window(settings: Settings) -> bool:
    """Show database settings dialog with improved UI and error handling"""
    
    sg.theme('SystemDefault')

    # Connection Status Indicator
    connection_status = [
        [sg.Text("Connection Status:", font=('Arial', 10, 'bold')),
         sg.Text("Not Connected", key='-CONNECTION-STATUS-', text_color='red'),
         sg.Button("Connect", key='-CONNECT-')]
    ]

    # Database Path Section
    path_section = [
        [sg.Text("Database Path:", font=('Arial', 10))],
        [sg.Input(settings.settings.get('db_path', ''), key='-DB-PATH-', size=(60, 1), enable_events=True),
         sg.FileBrowse(file_types=(("Access Files", "*.accdb;*.mdb"),))],
        [sg.Button("Test Connection", key='-TEST-CONN-', disabled=not bool(settings.settings.get('db_path', '')))]
    ]

    # Sync Settings Section
    sync_section = [
        [sg.Text("Sync Settings:", font=('Arial', 10))],
        [sg.Checkbox("Enable Auto Sync", key='-AUTO-SYNC-', 
                    default=settings.settings.get('auto_sync', False),
                    enable_events=True)],
        [sg.Text("Sync Interval:"),
         sg.Combo(['Daily', 'Weekly'], 
                default_value=settings.settings.get('sync_interval', 'Daily'),
                key='-SYNC-INTERVAL-',
                disabled=not settings.settings.get('auto_sync', False))]
    ]

    # Table Selection Section
    table_section = [
        [sg.Text("Select Table:", font=('Arial', 10))],
        [sg.Combo([], 
                 key='-TABLE-SELECT-', 
                 size=(55, 1), 
                 disabled=True, 
                 enable_events=True),
         sg.Button("Refresh", key='-REFRESH-')]
    ]

    # Table Preview Section - Reduced height
    table_preview = [
        [sg.Table(
            values=[],
            headings=['Error', 'Field', 'Row'],
            col_widths=[8, 8, 8],  # Compact initial widths
            justification='left',
            key='-TABLE-PREVIEW-',
            num_rows=8,            # Reduced from 15 to 8 rows
            alternating_row_color='#f0f0f0',
            enable_events=True,
            expand_x=True,
            expand_y=False,        # Don't expand vertically
            display_row_numbers=False,
            visible_column_map=[True]*3,
            def_col_width=8,
            auto_size_columns=False
        )]
    ]

    # Layout with all the original functionality
    layout = [
        [sg.Frame('Connection Status', connection_status, expand_x=True)],
        [sg.Frame('Database Configuration', path_section, expand_x=True)],
        [sg.Frame('Sync Options', sync_section, expand_x=True)],
        [sg.Frame('Table Selection', table_section, expand_x=True)],
        [sg.Frame('Table Preview', table_preview, expand_x=True, size=(None, 300))],  # Taller preview
        [sg.HorizontalSeparator()],
        [sg.Button("Save", key='-SAVE-', disabled=True), 
         sg.Button("Confirm", key='-CONFIRM-'),
         sg.Button("Cancel", key='-CANCEL-')]
    ]

    # Window with just size adjustments
    window = sg.Window(
        "Access Database Settings", 
        layout,
        modal=True,
        finalize=True,
        resizable=True,
        size=(750, 750),  # Adjusted height and width
        return_keyboard_events=True,
        keep_on_top=True,
        element_padding=(5, 5)
    )

    # Make the window visible in the taskbar
    window.TKroot.wm_attributes('-topmost', 0)

    db_manager = None
    last_valid_path = settings.settings.get('db_path', '')

    def update_connection_status(connected: bool, message: str = None):
        """Update connection status display"""
        if connected:
            window['-CONNECTION-STATUS-'].update("Connected", text_color='green')
            window['-TABLE-SELECT-'].update(disabled=False)
            window['-REFRESH-'].update(disabled=False)
            window['-SAVE-'].update(disabled=False)
        else:
            status_text = "Not Connected" if not message else f"Not Connected: {message}"
            window['-CONNECTION-STATUS-'].update(status_text, text_color='red')
            window['-TABLE-SELECT-'].update(disabled=True, values=[])
            window['-REFRESH-'].update(disabled=True)
            window['-SAVE-'].update(disabled=True)
            window['-TABLE-PREVIEW-'].update(values=[])

    def validate_path(path: str) -> bool:
        """Validate database path"""
        if not path:
            window['-TEST-CONN-'].update(disabled=True)
            return False
        
        if not os.path.exists(path):
            window['-TEST-CONN-'].update(disabled=True)
            return False
            
        if not path.lower().endswith(('.accdb', '.mdb')):
            window['-TEST-CONN-'].update(disabled=True)
            return False
            
        window['-TEST-CONN-'].update(disabled=False)
        return True

    def update_table_preview(window, headers, data):
        """Helper function to safely update table preview"""
        try:
            logging.info(f"Updating table preview with {len(headers)} columns and {len(data)} rows")
            
            table = window['-TABLE-PREVIEW-']
            table.update(values=data)
            
            # Configure columns with minimal initial widths
            if hasattr(table, 'Widget'):
                table.Widget.configure(displaycolumns=list(range(len(headers))))
                for idx, header in enumerate(headers):
                    table.Widget.heading(idx, text=header)
                    # Set a very compact initial width
                    table.Widget.column(idx, width=60)  # Fixed small pixel width
            
            window.refresh()
            logging.info("Table preview update completed")
            
        except Exception as e:
            logging.exception(f"Error updating table preview: {e}")
            print(f"Error updating table: {e}")

    def refresh_tables(window, db_manager):
        """Helper function to refresh tables list and preview"""
        try:
            logging.info("Starting refresh_tables function")
            
            if not db_manager:
                logging.error("No database manager instance")
                return
                
            if not db_manager.connection:
                logging.error("No active database connection")
                return
                
            # Get fresh list of tables
            logging.debug("Getting table names")
            tables = db_manager.get_table_names()
            logging.info(f"Found {len(tables)} tables: {tables}")
            
            if not tables:
                logging.warning("No tables found")
                return
            
            # Update the table dropdown
            logging.debug("Updating table dropdown")
            window['-TABLE-SELECT-'].update(values=tables)
            
            # Get currently selected table
            current_table = window['-TABLE-SELECT-'].get()
            logging.info(f"Current selected table: {current_table}")
            
            if not current_table and tables:
                current_table = tables[0]
                logging.info(f"No table selected, defaulting to: {current_table}")
                window['-TABLE-SELECT-'].update(value=current_table)
            
            if current_table:
                logging.debug(f"Getting preview data for table: {current_table}")
                headers, data = db_manager.get_table_preview(current_table)
                logging.info(f"Retrieved {len(data)} rows with {len(headers)} columns")
                logging.debug(f"Headers: {headers}")
                logging.debug(f"First row of data: {data[0] if data else 'No data'}")
                
                # Update the table preview
                table = window['-TABLE-PREVIEW-']
                logging.debug("Updating table headers")
                table.ColumnHeadings = headers
                logging.debug("Updating table data")
                table.update(values=data)
                
                logging.info(f"Successfully refreshed table {current_table}")
                
        except Exception as e:
            logging.exception(f"Error in refresh_tables: {e}")
            sg.popup_error(f"Error refreshing tables: {e}")

    def connect_and_load_table(window, db_manager, settings):
        """Helper function to connect and load the last used table"""
        if db_manager.connect():
            update_connection_status(True)
            tables = db_manager.get_table_names()
            if tables:
                # Get the last used table or default preference
                last_table = settings.settings.get('selected_table')
                preferred_table = 'afcables'  # Default preference if no last table
                
                # Update table list
                window['-TABLE-SELECT-'].update(values=tables)
                
                # Select appropriate table
                if last_table in tables:
                    selected_table = last_table
                elif preferred_table in tables:
                    selected_table = preferred_table
                else:
                    selected_table = tables[0]
                
                window['-TABLE-SELECT-'].update(value=selected_table)
                headers, data = db_manager.get_table_preview(selected_table)
                update_table_preview(window, headers, data)
                
                # Save the selected table
                settings.settings['selected_table'] = selected_table
                settings.save_settings()
            return True
        return False

    # Event Loop
    while True:
        event, values = window.read()
        
        if event in (sg.WIN_CLOSED, '-CANCEL-'):
            break

        if event == '-CONFIRM-':
            if db_manager and db_manager.connection:
                settings.settings.update({
                    'db_path': values['-DB-PATH-'],
                    'auto_sync': values['-AUTO-SYNC-'],
                    'sync_interval': values['-SYNC-INTERVAL-'],
                    'selected_table': values['-TABLE-SELECT-'],
                    'last_connection_status': True,
                    'window_size': window.size
                })
                settings.save_settings()
                break
            else:
                sg.popup_error("Please connect to database first")

        if event == '-SAVE-':
            if db_manager and db_manager.connection:
                settings.settings.update({
                    'db_path': values['-DB-PATH-'],
                    'auto_sync': values['-AUTO-SYNC-'],
                    'sync_interval': values['-SYNC-INTERVAL-'],
                    'selected_table': values['-TABLE-SELECT-'],
                    'last_connection_status': True,
                    'window_size': window.size
                })
                settings.save_settings()
                sg.popup("Settings saved successfully")

        if event == '-TABLE-SELECT-':
            if db_manager and db_manager.connection and values['-TABLE-SELECT-']:
                # Save the selected table immediately
                settings.settings['selected_table'] = values['-TABLE-SELECT-']
                settings.save_settings()
                headers, data = db_manager.get_table_preview(values['-TABLE-SELECT-'])
                update_table_preview(window, headers, data)

        if event == '-DB-PATH-':
            validate_path(values['-DB-PATH-'])

        if event == '-TEST-CONN-':
            if db_manager:
                db_manager.close()
            
            db_manager = AccessDatabaseManager(values['-DB-PATH-'])
            if db_manager.connect():
                update_connection_status(True)
                tables = db_manager.get_table_names()
                if tables:
                    window['-TABLE-SELECT-'].update(values=tables)
                    window['-TABLE-SELECT-'].update(value=tables[0])
                    headers, data = db_manager.get_table_preview(tables[0])
                    update_table_preview(window, headers, data)
                    last_valid_path = values['-DB-PATH-']
            else:
                update_connection_status(False, "Connection Failed")

        if event == '-REFRESH-':
            if db_manager and db_manager.connection:
                refresh_tables(window, db_manager)
            else:
                logging.error("Cannot refresh - no active database connection")
                sg.popup_error("Please connect to database first")

        if event == '-AUTO-SYNC-':
            window['-SYNC-INTERVAL-'].update(disabled=not values['-AUTO-SYNC-'])
            
        if event == "Save":
            if not validate_path(values['-DB-PATH-']):
                sg.popup_error("Please select a valid database file")
                continue
                
            # Save window size
            current_size = window.size
            
            settings.settings.update({
                'db_path': values['-DB-PATH-'],
                'auto_sync': values['-AUTO-SYNC-'],
                'sync_interval': values['-SYNC-INTERVAL-'],
                'selected_table': values['-TABLE-SELECT-'],
                'last_connection_status': True,
                'window_size': current_size
            })
            settings.save_settings()
            break

        if event == '-CONNECT-':
            if validate_path(values['-DB-PATH-']):
                if db_manager:
                    db_manager.close()
                
                db_manager = AccessDatabaseManager(values['-DB-PATH-'])
                connect_and_load_table(window, db_manager, settings)

    # Cleanup
    if db_manager:
        db_manager.close()
    window.close()
    
    # Return True if we have a valid database path
    return bool(last_valid_path)

if __name__ == "__main__":
    # For testing purposes
    settings = Settings()
    success = show_database_settings_window(settings)
    print(f"Settings window closed. Success: {success}")