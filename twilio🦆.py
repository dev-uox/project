import tkinter as tk
from tkinter import messagebox, scrolledtext
from twilio.rest import Client

# Replace these with your Twilio credentials
ACCOUNT_SID = ''
AUTH_TOKEN = ''
# Initialize Twilio client
client = Client(ACCOUNT_SID, AUTH_TOKEN)

# Default country code (US in this case)
DEFAULT_COUNTRY_CODE = '+1'

def format_phone_number(number):
    """Ensure the phone number has the correct format (E.164)"""
    if not number.startswith('+'):
        number = DEFAULT_COUNTRY_CODE + number
    return number

def lookup_phone_number():
    phone_number = phone_entry.get().strip()
    formatted_number = format_phone_number(phone_number)

    try:
        phone_info = client.lookups.v1.phone_numbers(formatted_number).fetch(type="carrier")
        
        result_text = f"Phone Number: {phone_info.phone_number}\n"
        result_text += f"National Format: {phone_info.national_format}\n"
        result_text += f"Country Code: {phone_info.country_code}\n"
        result_text += f"Carrier: {phone_info.carrier['name']}\n"
        result_text += f"Carrier Type: {phone_info.carrier['type']}\n"
        result_text += f"Mobile Country Code: {phone_info.carrier['mobile_country_code']}\n"
        result_text += f"Mobile Network Code: {phone_info.carrier['mobile_network_code']}"
    except Exception as e:
        result_text = f"Error: {str(e)}"

    result_textbox.config(state=tk.NORMAL)  # Enable editing to insert text
    result_textbox.delete(1.0, tk.END)  # Clear any previous content
    result_textbox.insert(tk.END, result_text)  # Insert new result
    result_textbox.config(state=tk.DISABLED)  # Disable editing after insertion

    result_textbox.yview(tk.END)  # Scroll to the bottom if the result is long

def clear_placeholder(event):
    """Clear placeholder text when the user clicks the entry box."""
    if phone_entry.get() == "Enter phone number":
        phone_entry.delete(0, tk.END)
        phone_entry.config(fg="black")

def add_placeholder(event):
    """Add placeholder text back if the entry box is empty."""
    if phone_entry.get() == "":
        phone_entry.insert(0, "Enter phone number")
        phone_entry.config(fg="grey")

# Create main window
root = tk.Tk()
root.title("Twilio Phone Lookup")
root.geometry("450x450")
root.resizable(False, False)
root.configure(bg="#f5f5f5")  # Light grey background

# Header Frame
header_frame = tk.Frame(root, bg="#00796b", pady=10)
header_frame.pack(fill=tk.X)

header_label = tk.Label(header_frame, text="Phone Number Lookup", font=("Arial", 16, "bold"), fg="white", bg="#00796b")
header_label.pack()

# Main Frame
main_frame = tk.Frame(root, padx=20, pady=20, bg="#f5f5f5")
main_frame.pack(padx=10, pady=10)

# Entry field with placeholder
phone_entry = tk.Entry(main_frame, font=("Arial", 12), width=30, fg="grey")
phone_entry.grid(row=0, column=0, pady=10)
phone_entry.insert(0, "Enter phone number")
phone_entry.bind("<FocusIn>", clear_placeholder)
phone_entry.bind("<FocusOut>", add_placeholder)

# Lookup button
lookup_button = tk.Button(main_frame, text="Lookup", command=lookup_phone_number, bg="#4CAF50", fg="white", font=("Arial", 12), width=15)
lookup_button.grid(row=1, column=0, pady=10)

# Scrollable Text Widget for results
result_textbox = scrolledtext.ScrolledText(main_frame, font=("Arial", 12), height=10, width=40, wrap=tk.WORD, bg="#ffffff", fg="#333333", state=tk.DISABLED, relief=tk.GROOVE, borderwidth=2)
result_textbox.grid(row=2, column=0, pady=10)

# Footer frame for copyright or info
footer_frame = tk.Frame(root, bg="#f5f5f5", pady=10)
footer_frame.pack(fill=tk.X)

footer_label = tk.Label(footer_frame, text="Powered by Twilio API", font=("Arial", 10), fg="#666666", bg="#f5f5f5")
footer_label.pack()

# Start the Tkinter event loop
root.mainloop()
