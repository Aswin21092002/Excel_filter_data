import pandas as pd
import tkinter as tk
import os
from tkinter import ttk, filedialog, messagebox
import speech_recognition as sr
import matplotlib.pyplot as plt
import seaborn as sns
from ttkthemes import ThemedTk

def load_excel():
    global df, file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    if not file_path:
        return
    try:
        df = pd.read_excel(file_path)
        column_select["values"] = list(df.columns)  # Populate dropdown with column names
        display_data(df)
        load_button.config(text=f"Loaded: {os.path.basename(file_path)}")  # Show file name on button
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load file: {e}")

def filter_data():
    if df is None:
        messagebox.showerror("Error", "No data loaded")
        return
    
    column = column_select.get()
    value = filter_entry.get()
    
    if column and value:
        filtered_df = df[df[column].astype(str).str.contains(value, case=False, na=False)]
        display_data(filtered_df)
        save_filtered_excel(filtered_df)
    else:
        messagebox.showwarning("Warning", "Select a column and enter a filter value")

def save_filtered_excel(filtered_df):
    global filtered_file_path
    filtered_file_path = "filtered_data.xlsx"
    filtered_df.to_excel(filtered_file_path, index=False)
    load_excel(filtered_file_path)  # Automatically reload filtered data
    messagebox.showinfo("Saved", f"Filtered data saved to {filtered_file_path}")

def delete_filtered_file():
    global filtered_file_path
    if os.path.exists(filtered_file_path):
        os.remove(filtered_file_path)
        messagebox.showinfo("Deleted", "Filtered data file deleted successfully")
        load_excel(file_path)  # Reload original data
    else:
        messagebox.showwarning("Warning", "No filtered data file found to delete")

def display_data(dataframe):
    for widget in frame.winfo_children():
        widget.destroy()
    
    tree = ttk.Treeview(frame, columns=list(dataframe.columns), show="headings")
    
    for col in dataframe.columns:
        tree.heading(col, text=col)
        tree.column(col, width=150)
    
    for _, row in dataframe.iterrows():
        tree.insert("", "end", values=list(row))
    
    tree.pack(fill=tk.BOTH, expand=True)

def voice_command():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        try:
            recognizer.adjust_for_ambient_noise(source)  # Adjust for background noise
            messagebox.showinfo("Voice Command", "Listening... Speak now!")
            audio = recognizer.listen(source, timeout=5, phrase_time_limit=5)  # Adjusted timing

            text = recognizer.recognize_google(audio)
            filter_entry.delete(0, tk.END)
            filter_entry.insert(0, text)
            messagebox.showinfo("Voice Command", f"Recognized: {text}")  # Show recognized text

        except sr.UnknownValueError:
            messagebox.showerror("Error", "Could not understand the voice input. Try again!")
        except sr.RequestError:
            messagebox.showerror("Error", "Internet required for speech recognition!")
        except sr.WaitTimeoutError:
            messagebox.showerror("Error", "No voice detected. Please try again!")
        except Exception as e:
            messagebox.showerror("Error", f"Voice recognition failed: {e}")



def choose_visualization():
    if df is None:
        messagebox.showerror("Error", "No data loaded")
        return
    
    def plot_selected_visualization():
        selected_option = visualization_var.get()
        numeric_columns = df.select_dtypes(include=["number"]).columns
        categorical_columns = df.select_dtypes(exclude=["number"]).columns
        
        plt.figure(figsize=(10, 6))
        
        if selected_option == "Box Plot" and len(numeric_columns) > 0:
            sns.boxplot(data=df[numeric_columns])
            plt.title("Box Plot of Numeric Columns")
        elif selected_option == "Histogram" and len(numeric_columns) > 0:
            df[numeric_columns[0]].plot(kind='hist', bins=20, color='blue', alpha=0.7)
            plt.title(f"Histogram of {numeric_columns[0]}")
        elif selected_option == "Bar Chart" and len(numeric_columns) > 0:
            df[numeric_columns].mean().plot(kind='bar', color='green')
            plt.title("Average Values of Numeric Columns")
        elif selected_option == "Pie Chart" and len(categorical_columns) > 0:
            df[categorical_columns[0]].value_counts().plot(kind='pie', autopct='%1.1f%%')
            plt.title(f"Pie Chart of {categorical_columns[0]}")
        elif selected_option == "Pairplot" and len(numeric_columns) > 1:
            sns.pairplot(df[numeric_columns])
            plt.title("Pairwise Relationships in Numeric Data")
        else:
            messagebox.showwarning("Warning", "No valid data available for the selected visualization.")
            return
        
        plt.show()
    
    vis_window = tk.Toplevel(root)
    vis_window.title("Choose Visualization")
    vis_window.geometry("300x200")
    
    visualization_var = tk.StringVar(value="Box Plot")
    
    options = ["Box Plot", "Histogram", "Bar Chart", "Pie Chart", "Pairplot"]
    for option in options:
        tk.Radiobutton(vis_window, text=option, variable=visualization_var, value=option).pack(anchor='w')
    
    tk.Button(vis_window, text="Show Visualization", command=plot_selected_visualization).pack(pady=10)

# GUI Setup
root = ThemedTk(theme="breeze")
root.title("Excel Filter App")
root.geometry("1080x720")

title_label = tk.Label(root, text="Excel Data Filter", font=("Arial", 20, "bold"))
title_label.pack(pady=10)

frame = tk.Frame(root)
frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

button_frame = tk.Frame(root)
button_frame.pack(pady=10)

button_style = {"font": ("Arial", 12, "bold"), "padx": 10, "pady": 5, "bg": "blue", "fg": "white"}

load_button = tk.Button(button_frame, text="Load Excel", command=load_excel, **button_style)
load_button.grid(row=0, column=0, padx=5)

column_select = ttk.Combobox(button_frame)
column_select.grid(row=0, column=1, padx=5)

filter_entry = tk.Entry(button_frame, font=("Arial", 12))
filter_entry.grid(row=0, column=2, padx=5)

filter_button = tk.Button(button_frame, text="Apply Filter", command=filter_data, bg="green", fg="white", font=("Arial", 12, "bold"), padx=10, pady=5)
filter_button.grid(row=0, column=3, padx=5)

voice_button = tk.Button(button_frame, text="🎙 Voice Filter", command=voice_command, bg="orange", fg="white", font=("Arial", 12, "bold"), padx=10, pady=5)
voice_button.grid(row=0, column=4, padx=5)

delete_button = tk.Button(button_frame, text="❌ Delete Filtered File", command=delete_filtered_file, bg="red", fg="white", font=("Arial", 12, "bold"), padx=10, pady=5)
delete_button.grid(row=0, column=5, padx=5)

visualize_button = tk.Button(button_frame, text="📊 Visualize Data", command=choose_visualization, bg="purple", fg="white", font=("Arial", 12, "bold"), padx=10, pady=5)
visualize_button.grid(row=0, column=6, padx=5)

root.mainloop()
