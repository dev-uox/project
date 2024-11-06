import os
import sys
import tkinter as tk
from tkinter import messagebox, Canvas, ttk, Toplevel, Label, Frame
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import datetime
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import random
import pygame

# Initialize pygame for sound effects
pygame.init()

# Global set to keep track of agents who have been congratulated
congratulated_agents = set()

# Function to get the correct path of the client_secret.json in both development and packaged modes
def get_client_secret_path():
    if hasattr(sys, '_MEIPASS'):
        # If running as a packaged .exe (PyInstaller), use the sys._MEIPASS path
        return os.path.join(sys._MEIPASS, 'client_secret.json')
    else:
        # If running the script directly, use the current directory
        return 'client_secret.json'

# Function to clean agent names and remove rows with blank names
def clean_agent_names(df, column_name):
    df[column_name] = df[column_name].str.strip().str.title()  # Removes extra spaces and capitalizes first letters
    df = df[df[column_name].notnull() & (df[column_name] != '')]  # Remove rows where the agent name is blank or NaN
    return df

# Function to authenticate and fetch data from Google Sheets, with detailed error messages
def fetch_google_sheet_data(sheet_url, worksheet_name, sheet_name):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(get_client_secret_path(), scope)
        client = gspread.authorize(creds)

        # Open the Google Sheet by URL and get the specified worksheet
        sheet = client.open_by_url(sheet_url)
        worksheet = sheet.worksheet(worksheet_name)
        headers = worksheet.row_values(1)  # Adjust if your header row is different

        # Fetch the data and convert it into a DataFrame
        data = worksheet.get_all_records(expected_headers=headers)
        df = pd.DataFrame(data)
        df.columns = df.columns.str.strip()  # Strip leading/trailing spaces from column names
        return df
    except Exception as e:
        # Display an error message in a message box and print to the console
        messagebox.showerror("Data Fetch Error", f"Error occurred while fetching data from {sheet_name}:\n{e}")
        print(f"Error occurred while fetching data from {sheet_name}: {e}")
        return None  # Return None if there is an error to handle it later

# Function to add labels on the bars
def add_labels(bars, values):
    for bar, value in zip(bars, values):
        height = bar.get_height()
        plt.text(bar.get_x() + bar.get_width() / 2, height, f'{int(value)}', ha='center', va='bottom')

# Function to create a cracker effect
def create_cracker_effect(root):
    # Load and play the congratulatory sound
    try:
        pygame.mixer.music.load("congratulations.mp3")
        pygame.mixer.music.play()
    except pygame.error:
        print("Warning: 'congratulations.mp3' not found. Skipping sound effect.")
    
    for _ in range(20):
        x = random.randint(0, root.winfo_width())
        y = random.randint(0, root.winfo_height())
        cracker_label = tk.Label(root, text="✺", font=("Helvetica", 36), fg=random.choice(["red", "orange", "yellow", "green", "blue", "purple"]), bg="#f0f4f5")
        cracker_label.place(x=x, y=y)
        root.after(2000, cracker_label.destroy)  # Remove cracker effect after 2000ms (2 seconds)

def generate_clustered_chart():
    sales_sheet_url = "https://docs.google.com/spreadsheets/d/1Ku206O26ZivlokFOiyeTV2GQ2qM-4X72-HzxjKRGcAM/edit?gid=799961172#gid=799961172"
    install_sheet_url = "https://docs.google.com/spreadsheets/d/1c9775YJINb0kixRejKbG0d_VI1tQsRrIu8x_Cof1-sY/edit?gid=0#gid=0"
    third_sheet_url = "https://docs.google.com/spreadsheets/d/1ZGvEMBOkk3WThrjTXG-DgLRWMaoV7hfTpW-elw16qqw/edit?gid=1538922778#gid=1538922778"
    
    try:
        # Fetch data for sales, installations, and the Proposal source from Google Sheets
        sales_df = fetch_google_sheet_data(sales_sheet_url, 'Form Responses 1', "Sales Sheet")
        install_df = fetch_google_sheet_data(install_sheet_url, 'Sheet1', "Installation Sheet")
        third_df = fetch_google_sheet_data(third_sheet_url, 'Form Responses 1', "Third Source Sheet")

        # Debug output to confirm data is fetched correctly
        print("Sales DataFrame:")
        print(sales_df)
        print("\nInstallations DataFrame:")
        print(install_df)
        print("\nThird Source DataFrame:")
        print(third_df)

        # If any of the data fetching fails (returns None), handle it silently and retry
        if sales_df is None or install_df is None or third_df is None:
            raise ValueError("Failed to fetch data from Google Sheets, retrying...")

        # Clean agent names in all DataFrames
        sales_df = clean_agent_names(sales_df, 'Agent Name')
        install_df = clean_agent_names(install_df, 'Agent Name')
        third_df = clean_agent_names(third_df, 'Agent Name')

        # Define the specific date range: October 13, 2024, to January 30, 2025
        start_date = datetime.datetime(2024, 10, 13)
        end_date = datetime.datetime(2025, 1, 30)

        # Apply date range filter to each DataFrame
        sales_df['Date'] = pd.to_datetime(sales_df['Date'], errors='coerce', format="%m/%d/%Y")
        install_df['Installed Date'] = pd.to_datetime(install_df['Installed Date'], errors='coerce', format="%m.%d.%Y")
        third_df['Date'] = pd.to_datetime(third_df['Date'], errors='coerce', format="%m/%d/%Y")
        
        sales_df = sales_df[(sales_df['Date'] >= start_date) & (sales_df['Date'] <= end_date)]
        install_df = install_df[(install_df['Installed Date'] >= start_date) & (install_df['Installed Date'] <= end_date)]
        third_df = third_df[(third_df['Date'] >= start_date) & (third_df['Date'] <= end_date)]

        # Debug output to check filtered data
        print("\nFiltered Sales DataFrame:")
        print(sales_df)
        print("\nFiltered Installations DataFrame:")
        print(install_df)
        print("\nFiltered Third Source DataFrame:")
        print(third_df)

        # Check if any data is available within the specified date range
        if sales_df.empty and install_df.empty and third_df.empty:
            messagebox.showinfo("No Data", f"No sales, installation, or third source data available from {start_date.strftime('%B %d, %Y')} to {end_date.strftime('%B %d, %Y')}")
            return

        # Group data by 'Agent Name'
        sales_by_agent = sales_df.groupby('Agent Name').size().reset_index(name='Total Sales')
        install_by_agent = install_df.groupby('Agent Name').size().reset_index(name='Total Installations')
        third_by_agent = third_df.groupby('Agent Name').size().reset_index(name='Third Source Data')

        # Merge the three DataFrames on 'Agent Name', fill NaN with 0
        combined_df = pd.merge(sales_by_agent, install_by_agent, on='Agent Name', how='outer').fillna(0)
        combined_df = pd.merge(combined_df, third_by_agent, on='Agent Name', how='outer').fillna(0)

        # Debug output to confirm merged DataFrame
        print("\nCombined DataFrame:")
        print(combined_df)

        # Create a figure for the clustered bar chart
        fig, ax = plt.subplots(figsize=(10, 4))  # Adjusted figure size for 1366x768 screen resolution

        # Set positions and width for the bars
        indices = np.arange(len(combined_df['Agent Name']))
        bar_width = 0.25  # Width of the bars

        # Plot Sales, Installations, and Third Source data side by side
        bars1 = ax.bar(indices, combined_df['Total Sales'], bar_width, label='Sales', color='#FF7F50')
        bars2 = ax.bar(indices + bar_width, combined_df['Total Installations'], bar_width, label='Installations', color='#32CD32')
        bars3 = ax.bar(indices + bar_width * 2, combined_df['Third Source Data'], bar_width, label='Proposal', color='#FFD700')

        # Add labels and title
        ax.set_xlabel('Agent Name', fontsize=10, fontweight='bold')
        ax.set_ylabel('Count', fontsize=10, fontweight='bold')
        ax.set_title(f'Agent Sales, Installations, and Proposals from {start_date.strftime("%B %d, %Y")} to {end_date.strftime("%B %d, %Y")}', fontsize=12, fontweight='bold')

        ax.set_xticks(indices + bar_width)
        ax.set_xticklabels(combined_df['Agent Name'], rotation=45, ha='right', fontsize=9)

        # Add numbers on top of the bars to show the values
        add_labels(bars1, combined_df['Total Sales'])
        add_labels(bars2, combined_df['Total Installations'])
        add_labels(bars3, combined_df['Third Source Data'])

        # Add a legend
        ax.legend(fontsize=9)

        # Adjust layout to avoid overlap
        plt.tight_layout()

        # Embed the chart in the Tkinter window
        canvas = FigureCanvasTkAgg(fig, master=chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(expand=True, fill=tk.BOTH)

        # Return the canvas for later removal when refreshing
        return canvas

    except Exception as e:
        print(f"Error generating chart: {e}")
        return None  # Return None if chart generation fails


# Function to refresh the chart
def refresh_chart():
    global chart_canvas  # Keep track of the current chart to remove it before refreshing
    if chart_canvas:
        chart_canvas.get_tk_widget().pack_forget()  # Remove the previous chart widget
    
    # Try generating a new chart
    chart_canvas = generate_clustered_chart()
    
    if chart_canvas is None:
        print("Error occurred. Retrying in 5 seconds...")
        root.after(5000, refresh_chart)  # Retry after 5 seconds if there was an error
    else:
        root.after(300000, refresh_chart)  # Schedule the refresh every 5 minutes

# Function to close the application
def close_app():
    root.quit()

# Create the main window
root = tk.Tk()
root.title("Progress Viewer")
root.geometry("1366x768")  # Set window size for screen resolution of 1366x768
root.configure(bg="#e6f7ff")

# Set the window icon
icon_path = os.path.join(sys._MEIPASS, 'icon.ico') if hasattr(sys, '_MEIPASS') else 'icon.ico'
root.iconbitmap(icon_path)

# Global variable to store the current chart canvas
chart_canvas = None
highlight_label = None

# Create frames for better layout organization
header_frame = tk.Frame(root, bg="#003366")
header_frame.pack(fill=tk.X, pady=10)

chart_frame = tk.Frame(root, bg="#ffffff", relief=tk.SOLID, borderwidth=1)
chart_frame.pack(expand=True, fill=tk.BOTH, padx=20, pady=(10, 5))

sidebar_frame = tk.Frame(root, bg="#f0f0f0", relief=tk.SOLID, borderwidth=1)
sidebar_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=10, pady=(10, 5))

# Create a label for instructions in the header frame
label = tk.Label(header_frame, text="Agent Progress Tracker", font=("Helvetica", 18, "bold"), bg="#003366", fg="#ffffff")
label.pack(padx=20, pady=5)

# Create a "Refresh" button in the header frame to manually refresh the chart
refresh_button = ttk.Button(header_frame, text="Refresh", command=refresh_chart)
refresh_button.pack(side=tk.RIGHT, padx=20)

# Create an "Exit" button in the header frame to close the application
exit_button = ttk.Button(header_frame, text="Exit", command=close_app)
exit_button.pack(side=tk.RIGHT, padx=10)

# Display Terms and Conditions directly in the sidebar frame
terms_frame = Frame(sidebar_frame, bg="#ffffff", relief=tk.SOLID, borderwidth=2)
terms_frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=(10, 0))

terms_label = Label(terms_frame, text="Terms and Conditions:\n\n- Only approved installations will count.\n- All installations should be done before 25th of Jan 2025.\n- ALC orders should get approved from the support leader.\n- Maintain discipline and quality. Misbehavior might have consequences.", 
                    font=("Helvetica", 12), fg="#333333", bg="#ffffff", justify="left", anchor="nw")
terms_label.pack(padx=15, pady=(15, 5))

# Start the chart refresh process
refresh_chart()

# Run the application
root.mainloop()
