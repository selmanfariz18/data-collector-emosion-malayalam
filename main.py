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

    if emotion == "Select Emotion":  # Check if an emotion has been selected
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
        emotion_var.set("Select Emotion")  # Reset to default after submission

        messagebox.showinfo("Success", "Data submitted successfully!")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Function to handle emoji selection
def select_emotion(emotion):
    emotion_var.set(emotion)
    # Insert the emoji into the text entry area
    text_entry.insert("end", f" {emotion_dict[emotion]}")

# Create the main window
root = tk.Tk()
root.title("Malayalam Text Emotion Dataset Collector")
root.geometry("800x600")  # Adjusted size for better layout

# Styling the main window
root.configure(bg="#f0f0f0")  # Light gray background

# Text Entry
tk.Label(root, text="Enter Malayalam Text:", font=("Helvetica", 14, "bold"), bg="#f0f0f0").pack(pady=10)
text_entry = tk.Text(root, height=6, width=60, font=("Helvetica", 12))
text_entry.pack(pady=10)

# Updated Emoji Dictionary with More Colorful and Attractive Emojis
emotion_dict = {
    "Happy": "😃",
    "Excitement": "🤩",
    "Sad": "😢",
    "Sarcasm": "😏",
    "Humour": "😂",
    "Anger": "😡",
    "Love": "😍",
    "Surprise": "😲",
    "Abusive": "🤬",
    "Fear": "😨"
}

# Emotion Dropdown Menu
emotion_var = tk.StringVar(root)
emotion_var.set("Select Emotion")

emotion_label = tk.Label(root, text="Select Emotion:", font=("Helvetica", 14, "bold"), bg="#f0f0f0")
emotion_label.pack(pady=5)

# Create a list of emotion options for the dropdown
emotion_options = list(emotion_dict.keys())
emotion_menu = tk.OptionMenu(root, emotion_var, *emotion_options)
emotion_menu.config(font=("Helvetica", 12))
emotion_menu.pack(pady=10)

# Display Emoji Buttons in a Grid Layout
emoji_frame = tk.Frame(root, bg="#f0f0f0")
emoji_frame.pack(pady=20)

# Adjust number of columns in the grid (e.g., 4)
columns = 4
row = 0
col = 0

for emotion in emotion_dict:
    emoji_button = tk.Button(
        emoji_frame,
        text=emotion_dict[emotion],
        font=("Helvetica", 14),
        bg="#ffffff",
        fg="#333333",
        relief="raised",
        padx=10,
        pady=5,
        command=lambda e=emotion: text_entry.insert("end", f" {emotion_dict[e]}")
    )
    emoji_button.grid(row=row, column=col, padx=10, pady=5)
    
    col += 1
    if col >= columns:
        col = 0
        row += 1

# Submit Button with Enhanced Styling
submit_button = tk.Button(
    root,
    text="Submit",
    command=submit_data,
    font=("Helvetica", 16, "bold"),
    bg="#4CAF50",
    fg="#ffffff",
    padx=20,
    pady=10,
    relief="raised"
)
submit_button.pack(pady=20)

# Run the application
root.mainloop()
