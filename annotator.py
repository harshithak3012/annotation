import tkinter as tk
from tkinter import messagebox
import pandas as pd
import os

# File paths for input and output
input_file_path = r"C:\Users\HP\Downloads\random15.xlsx"
output_file_path = r"C:\Users\HP\Downloads\annotations.xlsx"

# Load sentences from the Excel file
def load_sentences():
    try:
        df = pd.read_excel(input_file_path)
        return df['Sentence'].tolist()  # Assuming the column name is 'Sentence'
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load sentences: {e}")
        return []

# Save annotations to Excel file
def save_annotations(sentence, sentiment, rationale):
    try:
        # Create a DataFrame for the new annotation
        new_annotation = pd.DataFrame({
            'Sentence': [sentence],
            'Sentiment': [sentiment],
            'Rationale': [rationale]
        })

        # Check if the file exists
        if os.path.exists(output_file_path):
            # Load existing data
            existing_data = pd.read_excel(output_file_path)
            # Append new annotation to existing data
            updated_data = pd.concat([existing_data, new_annotation], ignore_index=True)
        else:
            # If file doesn't exist, use the new annotation as the data
            updated_data = new_annotation

        # Save the updated data to the file
        updated_data.to_excel(output_file_path, index=False)
        # Update the success message label
        success_label.config(text="Annotation saved successfully!", fg="green")
    except Exception as e:
        success_label.config(text=f"Failed to save annotation: {e}", fg="red")

# Function to save the current annotation
def save_annotation():
    sentiment = sentiment_var.get()
    
    # Get all highlighted text ranges
    rationale_text = []
    for tag in text_widget.tag_names():
        if tag.startswith("highlight"):
            try:
                start_index = text_widget.index(f"{tag}.first")
                end_index = text_widget.index(f"{tag}.last")
                rationale_text.append(text_widget.get(start_index, end_index))
            except tk.TclError:
                pass  # Skip invalid tags

    if rationale_text:
        rationale = ', '.join(rationale_text)
        save_annotations(current_sentence, sentiment, rationale)
    else:
        success_label.config(text="Please highlight some text before saving.", fg="red")

# Highlight the selected text
def highlight_selection(event):
    try:
        # Get the current selection
        start_index = text_widget.index(tk.SEL_FIRST)
        end_index = text_widget.index(tk.SEL_LAST)

        # Create a unique tag for highlighting
        highlight_tag = f"highlight{len(text_widget.tag_names())}"
        text_widget.tag_add(highlight_tag, start_index, end_index)
        text_widget.tag_config(highlight_tag, background="yellow")
    except tk.TclError:
        pass  # No selection made

# Load the next sentence
def load_next_sentence():
    global current_sentence, sentence_list, current_index

    if current_index < len(sentence_list):
        current_sentence = sentence_list[current_index]
        sentence_label.config(text=current_sentence)
        
        # Clear the text widget and reset highlights
        text_widget.delete('1.0', tk.END)
        text_widget.insert(tk.END, current_sentence)
        text_widget.tag_delete(*text_widget.tag_names())  # Remove all highlight tags

        current_index += 1
        success_label.config(text="")  # Clear the success message when loading the next sentence
    else:
        success_label.config(text="No more sentences to annotate.", fg="blue")

# Initialize main application window
root = tk.Tk()
root.title("Sentiment Annotation Tool")

# Initialize variables
current_sentence = ""
sentence_list = load_sentences()
current_index = 0

# Sentence display
tk.Label(root, text="Sentence:").pack()
sentence_label = tk.Label(root, text="")
sentence_label.pack()

# Text widget for displaying and highlighting the sentence
text_widget = tk.Text(root, height=4, wrap=tk.WORD)
text_widget.pack()
text_widget.bind("<ButtonRelease-1>", highlight_selection)  # Bind mouse click to highlight function

# Sentiment selection
sentiment_var = tk.StringVar(value='neutral')
tk.Label(root, text="Select Sentiment:").pack()
tk.OptionMenu(root, sentiment_var, "Positive", "Negative", "Neutral").pack()

# Save and Next buttons
tk.Button(root, text="Save", command=save_annotation).pack()
tk.Button(root, text="Next", command=load_next_sentence).pack()

# Success message label
success_label = tk.Label(root, text="", font=("Arial", 10))
success_label.pack()

# Load the first sentence on startup if available
if sentence_list:
    load_next_sentence()
else:
    messagebox.showerror("Error", "No sentences found in the input file.")

root.mainloop()
