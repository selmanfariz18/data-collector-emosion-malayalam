import tkinter as tk
from tkinter import messagebox
import pandas as pd
import os
from openpyxl import load_workbook

# Function to handle submission of text and emotion
def submit_data():
    text = text_entry.get("1.0", "end-1c").strip()
    emotion = emotion_var.get()

    if text == "":
        messagebox.showerror("Input Error", "Please enter some text.")
        return

    if emotion == "":
        messagebox.showerror("Input Error", "Please select an emotion.")
        return

    # Prepare data
    data = {"Text": text, "Emotion": emotion}

    try:
        if os.path.exists("emotion_dataset.xlsx"):
            # Load the existing workbook
            book = load_workbook("emotion_dataset.xlsx")
            sheet = book.active
            # Append the new row at the end
            sheet.append([text, emotion])
            book.save("emotion_dataset.xlsx")
            book.close()
        else:
            # If file does not exist, create it with headers
            df = pd.DataFrame([data])
            df.to_excel("emotion_dataset.xlsx", index=False)

        # Clear input fields
        text_entry.delete("1.0", "end")
        emotion_var.set("")

        messagebox.showinfo("Success", "Data submitted successfully!")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Create the main window
root = tk.Tk()
root.title("Malayalam Text Emotion Dataset Collector")

# Text Entry
tk.Label(root, text="Enter Malayalam Text:").pack(pady=5)
text_entry = tk.Text(root, height=5, width=50)
text_entry.pack(pady=5)

# Emotion Dropdown
tk.Label(root, text="Select Emotion:").pack(pady=5)
emotion_var = tk.StringVar(root)
emotions = ["Happy", "Excitement", "Sad", "Sarcasm", "Humour", "Anger", "Love", "Surprise", "Abusive", "Fear"]
emotion_menu = tk.OptionMenu(root, emotion_var, *emotions)
emotion_menu.pack(pady=5)

# Submit Button
submit_button = tk.Button(root, text="Submit", command=submit_data)
submit_button.pack(pady=20)

# Run the application
root.mainloop()
