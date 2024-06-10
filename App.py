import sys
import schedule
import time
import asyncio
import warnings
import logging
import timeit
import gspread
import sqlite3
from telegram import Bot, Update
from telegram.error import TelegramError
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox, QProgressBar, QHBoxLayout, QComboBox, QDialog, QListWidget, QListWidgetItem, QTextEdit, QTimeEdit
from PyQt5.QtCore import QThread, QTimer, QSize, Qt, QTime
from PyQt5.QtGui import QFont, QIcon
from tenacity import retry, wait_exponential, stop_after_attempt

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("bot.log"),
        logging.StreamHandler(sys.stdout)
    ]
)

# Suppress specific asyncio RuntimeError warnings
warnings.filterwarnings("ignore", category=RuntimeWarning, message=".*Event loop is closed.*")

# Initialize global variables for GUI inputs
TOKEN = ''
spreadsheet_key = ''
service_account_path = ''
category_folder = ''

# Set up the SQLite database
conn = sqlite3.connect('channels.db')
cursor = conn.cursor()

# Create a table to store channel information if it doesn't exist
cursor.execute('''
    CREATE TABLE IF NOT EXISTS channels (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        worksheet_name TEXT NOT NULL,
        channel_name TEXT NOT NULL,
        posting_time TEXT NOT NULL
    )
''')
conn.commit()

# Function to initialize the bot and Google Sheets
def initialize_bot_and_sheets():
    global bot, gc, spreadsheet
    bot = Bot(TOKEN)
    gc = gspread.service_account(filename=service_account_path)
    spreadsheet = gc.open_by_key(spreadsheet_key)

# Load existing channels from the database
worksheet_channels = {}
cursor.execute('SELECT worksheet_name, channel_name FROM channels')
for row in cursor.fetchall():
    worksheet_channels[row[0]] = row[1]

# Function to construct the caption with separators
def construct_caption(headers, data_range, separator_line):
    captions = []
    for row in data_range:
        caption_text = "\n".join([f"{header}: {info}" for header, info in zip(headers, row)])
        captions.append(caption_text)
    return f"\n{separator_line}\n".join(captions)

# Asynchronous function to send the image with a caption and pin the message
async def send_image_and_pin(channel_id, image_path, caption):
    try:
        with open(image_path, 'rb') as image_file:
            message = await bot.send_photo(chat_id=channel_id, photo=image_file, caption=caption)
        logging.info("Image has been successfully sent with caption!")

        # If pin_chat_message is truly async, use await directly
        await bot.pin_chat_message(chat_id=channel_id, message_id=message.message_id, disable_notification=True)
        logging.info("Message has been successfully pinned!")
    except TelegramError as e:
        logging.error(f"An error occurred: {e}")

# Retry decorator for functions that interact with external services
@retry(wait=wait_exponential(min=1, max=10), stop=stop_after_attempt(5))
async def fetch_worksheet(worksheet_name):
    return spreadsheet.worksheet(worksheet_name)

# Function to process each worksheet
async def process_worksheet(worksheet_name, channel_id):
    start_time = timeit.default_timer()
    logging.info(f"Processing worksheet: {worksheet_name}")

    worksheet = await fetch_worksheet(worksheet_name)

    # Fetch the headers and data
    headers = worksheet.get('C1:F1')[0]
    data_range = worksheet.get('C2:F11')

    # Fetch the category and day
    category = worksheet.acell('A2').value  # Assuming 'A2' is the category header cell
    day = worksheet.acell('B2').value  # Assuming 'B2' is the day header cell

    # Construct the image path based on category and day
    image_path = f'{category_folder}/{category}/{day.lower()}.png'

    separator_line = "-" * 30  # Define the separator line
    full_caption = construct_caption(headers, data_range, separator_line)
    await send_image_and_pin(channel_id, image_path, full_caption)

    elapsed = timeit.default_timer() - start_time
    logging.info(f"Finished processing worksheet: {worksheet_name} in {elapsed:.2f} seconds")

# Function to schedule tasks for each channel
def schedule_tasks():
    cursor.execute('SELECT worksheet_name, posting_time FROM channels')
    channels = cursor.fetchall()
    for worksheet_name, posting_time in channels:
        schedule.every().day.at(posting_time).do(
            lambda ws=worksheet_name, ch_id=worksheet_channels[worksheet_name]: asyncio.create_task(process_worksheet(ws, ch_id))
        )
    logging.info("Tasks have been scheduled.")

# Function to run scheduled tasks
async def run_scheduled_tasks():
    while True:
        schedule.run_pending()
        await asyncio.sleep(1)

# Function to start the bot and scheduling
def start_bot():
    global TOKEN, spreadsheet_key, category_folder
    TOKEN = token_entry.text()
    spreadsheet_key = spreadsheet_entry.text()
    category_folder = category_folder_entry.text()
    
    if not TOKEN or not spreadsheet_key or not service_account_path or not category_folder:
        QMessageBox.critical(None, "Error", "All fields are required!")
        return
    
    try:
        initialize_bot_and_sheets()
        schedule_tasks()
        bot_thread.start()

        timer.start(100)  # Update progress bar every 100 ms
    except Exception as e:
        logging.error(f"An error occurred: {e}")
        QMessageBox.critical(None, "Error", str(e))

# Function to update the progress bar
def update_progress_bar():
    current_value = progress_bar.value()
    new_value = (current_value + 1) % progress_bar.maximum()
    progress_bar.setValue(new_value)

# Function to stop the bot
def stop_bot():
    bot_thread.stop()
    bot_thread.wait()  # Ensure the thread is properly stopped
    timer.stop()

# Worker thread to run the bot and scheduler
class BotThread(QThread):
    def __init__(self):
        super().__init__()
        self._is_running = True

    def run(self):
        global loop
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        loop.run_until_complete(run_scheduled_tasks())

    def stop(self):
        self._is_running = False
        loop.stop()

# GUI functions
def select_service_account():
    global service_account_path
    service_account_path, _ = QFileDialog.getOpenFileName(None, "Select Service Account JSON", "", "JSON files (*.json)")
    if service_account_path:
        service_account_entry.setText(service_account_path)

def select_category_folder():
    global category_folder
    category_folder = QFileDialog.getExistingDirectory(None, "Select Category Folder")
    if category_folder:
        category_folder_entry.setText(category_folder)

def open_settings():
    settings_window = QWidget()
    settings_window.setWindowTitle("Settings")
    settings_layout = QVBoxLayout()

    settings_font = QFont('Arial', 12)

    label_setting = QLabel("This is a settings window")
    label_setting.setFont(settings_font)
    settings_layout.addWidget(label_setting)

    settings_window.setLayout(settings_layout)
    settings_window.setGeometry(100, 100, 300, 200)  # Set the size of the settings window
    settings_window.show()

def show_added_channels():
    # Create a dialog window to display added channels
    show_dialog = QDialog()
    show_dialog.setWindowTitle("Added Channels")

    layout = QVBoxLayout()

    list_widget = QListWidget()
    cursor.execute('SELECT id, worksheet_name, channel_name, posting_time FROM channels')
    for row in cursor.fetchall():
        list_item = QListWidgetItem(f"{row[1]} - {row[2]} - {row[3]}")
        list_item.setData(Qt.UserRole, row[0])  # Store the id in the item
        list_widget.addItem(list_item)

    layout.addWidget(list_widget)

    delete_button = QPushButton("Delete Selected Channel")
    layout.addWidget(delete_button)

    def delete_channel():
        selected_item = list_widget.currentItem()
        if selected_item:
            channel_id = selected_item.data(Qt.UserRole)
            cursor.execute('DELETE FROM channels WHERE id = ?', (channel_id,))
            conn.commit()
            list_widget.takeItem(list_widget.row(selected_item))
            QMessageBox.information(None, "Success", "Channel deleted successfully!")
        else:
            QMessageBox.warning(None, "Error", "No channel selected!")

    delete_button.clicked.connect(delete_channel)

    close_button = QPushButton("Close")
    close_button.clicked.connect(show_dialog.close)
    layout.addWidget(close_button)

    show_dialog.setLayout(layout)
    show_dialog.exec_()

def add_new_channel():
    # Create a dialog window to take input from the user
    input_dialog = QDialog()
    input_dialog.setWindowTitle("Add Channel")

    layout = QVBoxLayout()

    # Input fields for worksheet name, channel name, and posting time
    worksheet_label = QLabel("Worksheet Name:")
    worksheet_input = QLineEdit()
    channel_label = QLabel("Channel Name:")
    channel_input = QLineEdit()
    time_label = QLabel("Posting Time:")
    time_input = QTimeEdit()
    time_input.setDisplayFormat("HH:mm")
    time_input.setTime(QTime.currentTime())

    layout.addWidget(worksheet_label)
    layout.addWidget(worksheet_input)
    layout.addWidget(channel_label)
    layout.addWidget(channel_input)
    layout.addWidget(time_label)
    layout.addWidget(time_input)

    # Save button to store the data
    save_button = QPushButton("Save")
    layout.addWidget(save_button)

    input_dialog.setLayout(layout)

    def save_channel():
        worksheet_name = worksheet_input.text()
        channel_name = channel_input.text()
        posting_time = time_input.time().toString("HH:mm")

        if worksheet_name and channel_name:
            # Insert the new channel information into the database
            cursor.execute('''
                INSERT INTO channels (worksheet_name, channel_name, posting_time)
                VALUES (?, ?, ?)
            ''', (worksheet_name, channel_name, posting_time))
            conn.commit()

            # Update the worksheet_channels dictionary
            worksheet_channels[worksheet_name] = channel_name

            QMessageBox.information(None, "Success", "Channel added successfully!")
            input_dialog.close()
        else:
            QMessageBox.warning(None, "Error", "All fields are required!")

    save_button.clicked.connect(save_channel)
    input_dialog.exec_()

def edit_channel():
    # Create a dialog window to edit existing channels
    edit_dialog = QDialog()
    edit_dialog.setWindowTitle("Edit Channel")

    layout = QVBoxLayout()

    # Dropdown to select channel to edit
    channel_select_label = QLabel("Select Channel to Edit:")
    layout.addWidget(channel_select_label)

    channel_select = QComboBox()
    cursor.execute('SELECT id, worksheet_name, channel_name, posting_time FROM channels')
    channel_data = cursor.fetchall()
    for row in channel_data:
        channel_select.addItem(f"{row[1]} - {row[2]}", userData=row[0])

    layout.addWidget(channel_select)

    # Input fields for worksheet name, channel name, and posting time
    worksheet_label = QLabel("Worksheet Name:")
    worksheet_input = QLineEdit()
    channel_label = QLabel("Channel Name:")
    channel_input = QLineEdit()
    time_label = QLabel("Posting Time:")
    time_input = QTimeEdit()
    time_input.setDisplayFormat("HH:mm")

    layout.addWidget(worksheet_label)
    layout.addWidget(worksheet_input)
    layout.addWidget(channel_label)
    layout.addWidget(channel_input)
    layout.addWidget(time_label)
    layout.addWidget(time_input)

    old_worksheet_name = None

    def populate_fields(index):
        nonlocal old_worksheet_name
        selected_id = channel_select.currentData()
        cursor.execute('SELECT worksheet_name, channel_name, posting_time FROM channels WHERE id = ?', (selected_id,))
        row = cursor.fetchone()
        if row:
            old_worksheet_name = row[0]
            worksheet_input.setText(row[0])
            channel_input.setText(row[1])
            time_input.setTime(QTime.fromString(row[2], "HH:mm"))

    channel_select.currentIndexChanged.connect(populate_fields)

    # Save button to update the data
    save_button = QPushButton("Save")
    layout.addWidget(save_button)

    edit_dialog.setLayout(layout)

    def save_edited_channel():
        new_worksheet_name = worksheet_input.text()
        new_channel_name = channel_input.text()
        new_posting_time = time_input.time().toString("HH:mm")
        selected_id = channel_select.currentData()

        if new_worksheet_name and new_channel_name and new_posting_time and selected_id:
            # Update the channel information in the database
            cursor.execute('''
                UPDATE channels
                SET worksheet_name = ?, channel_name = ?, posting_time = ?
                WHERE id = ?
            ''', (new_worksheet_name, new_channel_name, new_posting_time, selected_id))
            conn.commit()

            # Update the worksheet_channels dictionary
            if old_worksheet_name in worksheet_channels:
                del worksheet_channels[old_worksheet_name]
            worksheet_channels[new_worksheet_name] = new_channel_name

            QMessageBox.information(None, "Success", "Channel updated successfully!")
            edit_dialog.close()
        else:
            QMessageBox.warning(None, "Error", "All fields are required!")

    save_button.clicked.connect(save_edited_channel)

    # Populate fields for the initially selected channel
    if channel_select.count() > 0:
        populate_fields(channel_select.currentIndex())

    edit_dialog.exec_()

def add_channel():
    # Create a dialog window to show options
    options_dialog = QDialog()
    options_dialog.setWindowTitle("Channel Options")

    layout = QVBoxLayout()

    show_button = QPushButton("Show Added Channels")
    show_button.clicked.connect(show_added_channels)
    layout.addWidget(show_button)

    add_button = QPushButton("Add Channel")
    add_button.clicked.connect(add_new_channel)
    layout.addWidget(add_button)

    edit_button = QPushButton("Edit Channel")
    edit_button.clicked.connect(edit_channel)
    layout.addWidget(edit_button)

    options_dialog.setLayout(layout)
    options_dialog.exec_()

# GUI setup
app = QApplication(sys.argv)
window = QWidget()
window.setWindowTitle("Telegram Bot Scheduler")
window.setGeometry(100, 100, 1200, 800)  # Set the size of the main window

main_layout = QHBoxLayout()  # Main layout to hold left and right panels

left_panel = QVBoxLayout()  # Left panel layout
right_panel = QVBoxLayout()  # Right panel layout for logs

font = QFont('Arial', 14)

# Title layout with centered title
title_layout = QHBoxLayout()
title_label = QLabel("Telegram Class Schedule Sender")
title_label.setFont(QFont('Arial', 16, QFont.Bold))
title_label.setAlignment(Qt.AlignCenter)
title_layout.addStretch()
title_layout.addWidget(title_label)
title_layout.addStretch()
left_panel.addLayout(title_layout)

# Add channel button
add_channel_button = QPushButton("Add Telegram Channel")
add_channel_button.setIcon(QIcon('add_channel.png'))  # Replace with the correct path to your plus icon
add_channel_button.setIconSize(QSize(32, 32))  # Set the icon size
add_channel_button.clicked.connect(add_channel)
left_panel.addWidget(add_channel_button)

# Settings button
settings_button = QPushButton("Settings")
settings_button.setIcon(QIcon('settings.png'))  # Replace with the correct path to your settings icon
settings_button.setIconSize(QSize(32, 32))  # Set the icon size
settings_button.clicked.connect(open_settings)
left_panel.addWidget(settings_button)

label_token = QLabel("Bot Token:")
label_token.setFont(font)
left_panel.addWidget(label_token)

token_entry = QLineEdit()
token_entry.setFont(font)
left_panel.addWidget(token_entry)

label_spreadsheet = QLabel("Spreadsheet Key:")
label_spreadsheet.setFont(font)
left_panel.addWidget(label_spreadsheet)

spreadsheet_entry = QLineEdit()
spreadsheet_entry.setFont(font)
left_panel.addWidget(spreadsheet_entry)

label_service_account = QLabel("Service Account JSON:")
label_service_account.setFont(font)
left_panel.addWidget(label_service_account)

service_account_entry = QLineEdit()
service_account_entry.setFont(font)
left_panel.addWidget(service_account_entry)

browse_button = QPushButton("Browse")
browse_button.setFont(font)
browse_button.clicked.connect(select_service_account)
left_panel.addWidget(browse_button)

label_category_folder = QLabel("Category Folder:")
label_category_folder.setFont(font)
left_panel.addWidget(label_category_folder)

category_folder_entry = QLineEdit()
category_folder_entry.setFont(font)
left_panel.addWidget(category_folder_entry)

category_browse_button = QPushButton("Browse")
category_browse_button.setFont(font)
category_browse_button.clicked.connect(select_category_folder)
left_panel.addWidget(category_browse_button)

start_button = QPushButton("Start Bot")
start_button.setFont(font)
start_button.clicked.connect(start_bot)
left_panel.addWidget(start_button)

stop_button = QPushButton("Stop Bot")
stop_button.setFont(font)
stop_button.clicked.connect(stop_bot)
left_panel.addWidget(stop_button)

progress_bar = QProgressBar()
progress_bar.setRange(0, 100)
progress_bar.setValue(0)
progress_bar.setFont(font)
left_panel.addWidget(progress_bar)

# Log output area
log_output = QTextEdit()
log_output.setFont(QFont('Courier', 10))
log_output.setReadOnly(True)

# Create a logging handler that sends log messages to the QTextEdit
class QTextEditLogger(logging.Handler):
    def __init__(self, widget):
        super().__init__()
        self.widget = widget

    def emit(self, record):
        msg = self.format(record)
        self.widget.append(msg)
        self.widget.ensureCursorVisible()

log_handler = QTextEditLogger(log_output)
log_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
logging.getLogger().addHandler(log_handler)

right_panel.addWidget(log_output)

main_layout.addLayout(left_panel, 2)  # Add left panel to main layout, takes 2/3 of the space
main_layout.addLayout(right_panel, 1)  # Add right panel to main layout, takes 1/3 of the space

window.setLayout(main_layout)

# Timer for updating the progress bar
timer = QTimer()
timer.timeout.connect(update_progress_bar)

# Initialize the bot thread
bot_thread = BotThread()

window.show()
app.aboutToQuit.connect(stop_bot)  # Ensure the bot stops when the application is closed
sys.exit(app.exec_())

# Ensure the database connection is closed when the application exits
conn.close()

