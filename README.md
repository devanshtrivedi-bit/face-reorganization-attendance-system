import tkinter as tk
from tkinter import messagebox, Toplevel, Label, Button
from PIL import Image, ImageTk
import face_recognition
import cv2
import os
import numpy as np
import pyttsx3
import datetime
import openpyxl

# --- Text-to-Speech Setup ---
engine = pyttsx3.init()
def speak(text: str):
    engine.say(text)
    engine.runAndWait()

# --- Set Paths ---
USER_NAME = os.getlogin()
DESKTOP_PATH = os.path.join("C:/Users", USER_NAME, "OneDrive/Desktop")
KNOWN_FACES_DIR = os.path.join(DESKTOP_PATH, "known_faces")
STUDENTS_FILE = os.path.join(DESKTOP_PATH, "students.txt")
EXCEL_FILE = os.path.join(DESKTOP_PATH, "attendance.xlsx")

# Ensure known_faces directory exists
os.makedirs(KNOWN_FACES_DIR, exist_ok=True)

# --- Load Known Faces ---
known_faces = {}
def load_known_faces():
    for file in os.listdir(KNOWN_FACES_DIR):
        if file.lower().endswith((".jpg", ".png")):
            path = os.path.join(KNOWN_FACES_DIR, file)
            image = face_recognition.load_image_file(path)
            encodings = face_recognition.face_encodings(image)
            if encodings:
                name = os.path.splitext(file)[0]
                known_faces[name] = encodings[0]
load_known_faces()

# --- Load Students ---
students = []
def load_students():
    if not os.path.exists(STUDENTS_FILE):
        return
    with open(STUDENTS_FILE, "r") as f:
        for line in f:
            parts = line.strip().split()
            if len(parts) >= 2:
                name = " ".join(parts[:-1])
                roll_no = parts[-1]
                students.append({"name": name, "roll_no": roll_no})
load_students()

# --- Attendance Export ---
def export_attendance(datetime_str, name, roll_no, status):
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Attendance"
        ws.append(["Date & Time", "Student Name", "Roll Number", "Status"])
    else:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
    ws.append([datetime_str, name, roll_no, status])
    wb.save(EXCEL_FILE)

# --- Face Recognition Function ---
def recognize_face():
    cap = cv2.VideoCapture(0)
    recognized_name = "Unknown"
    encodings = list(known_faces.values())
    names = list(known_faces.keys())

    while True:
        ret, frame = cap.read()
        if not ret:
            break
        small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)
        rgb_frame = cv2.cvtColor(small_frame, cv2.COLOR_BGR2RGB)
        face_locs = face_recognition.face_locations(rgb_frame)
        face_encs = face_recognition.face_encodings(rgb_frame, face_locs)
        for face_encoding in face_encs:
            matches = face_recognition.compare_faces(encodings, face_encoding)
            face_dist = face_recognition.face_distance(encodings, face_encoding)
            best_idx = np.argmin(face_dist)
            if matches[best_idx]:
                recognized_name = names[best_idx]
                break
        if recognized_name != "Unknown":
            break
    cap.release()
    return recognized_name

# --- GUI Setup ---
root = tk.Tk()
root.title("Smart Attendance System")
root.attributes('-fullscreen', True)
root.bind("<Escape>", lambda e: root.attributes('-fullscreen', False))
root.resizable(False, False)

# --- Styling ---
FONT_TITLE = ("Arial", 16, "bold")
FONT_LABEL = ("Arial", 12)
FONT_RESULT = ("Arial", 14, "bold")
BUTTON_BG = "#4caf50"
BUTTON_FG = "white"

# --- Functions ---

def clear_student_info():
    label_student_name.config(text="")
    label_image.config(image="")
    label_image.image = None
    label_message.config(text="")

def show_student_image(name):
    path = os.path.join(KNOWN_FACES_DIR, f"{name}.jpg")
    if os.path.exists(path):
        img = Image.open(path).resize((120, 120))
        img_tk = ImageTk.PhotoImage(img)
        label_image.config(image=img_tk)
        label_image.image = img_tk
    else:
        label_image.config(image="")
        label_image.image = None
        label_message.config(text="Image not found.")

def search_student():
    roll = entry_roll_no.get().strip()
    clear_student_info()
    if not roll:
        label_message.config(text="Please enter a roll number.")
        return
    student = next((s for s in students if s["roll_no"] == roll or s["roll_no"].endswith(roll)), None)
    if student:
        label_student_name.config(text=student["name"])
        speak(f"Hello {student['name']}")
        show_student_image(student["name"])
        label_message.config(text="")
    else:
        label_message.config(text="Student not found.")

def mark_attendance():
    roll = entry_roll_no.get().strip()
    status = var_status.get()
    if not roll:
        label_message.config(text="Please enter a roll number.")
        return
    student = next((s for s in students if s["roll_no"] == roll or s["roll_no"].endswith(roll)), None)
    if not student:
        label_message.config(text="Student not found.")
        return
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    export_attendance(now, student["name"], student["roll_no"], status)
    label_message.config(text=f"Attendance marked for {student['name']} as {status}.")
    speak(f"Attendance marked for {student['name']} as {status}.")
    update_dashboard()

def recognize_and_mark():
    label_message.config(text="Recognizing face, please wait...")
    root.update()
    name = recognize_face()
    if name == "Unknown":
        label_message.config(text="Face not recognized.")
        speak("Face not recognized. Please try again.")
        return
    student = next((s for s in students if s["name"].lower() == name.lower()), None)
    if student:
        entry_roll_no.delete(0, tk.END)
        entry_roll_no.insert(0, student["roll_no"])
        label_student_name.config(text=student["name"])
        show_student_image(student["name"])
        var_status.set("Present")
        mark_attendance()
    else:
        label_message.config(text="Recognized face but student details not found.")
        speak("Student details not found.")

def update_dashboard():
    label_total_students.config(text=f"Total Students: {len(students)}")
    today = datetime.datetime.now().strftime("%Y-%m-%d")
    present = absent = leave = total = 0
    if os.path.exists(EXCEL_FILE):
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] and row[0].startswith(today):
                total += 1
                status = row[3].strip().lower()
                if status == "present":
                    present += 1
                elif status == "absent":
                    absent += 1
                elif status == "leave":
                    leave += 1
    label_total_attendance.config(text=f"Total Attendance Today: {total}")
    label_present_count.config(text=f"Present: {present}")
    label_absent_count.config(text=f"Absent: {absent}")
    label_leave_count.config(text=f"Leave: {leave}")

# --- Show confirmation popup with captured photo ---
def show_photo_confirmation(image_path, name, roll_no):
    def on_yes():
        # Add to known_faces dictionary
        image = face_recognition.load_image_file(image_path)
        encodings = face_recognition.face_encodings(image)
        if not encodings:
            messagebox.showerror("Error", "No face detected in captured photo.")
            os.remove(image_path)
            popup.destroy()
            return
        known_faces[name] = encodings[0]

        # Save student info to file and list
        with open(STUDENTS_FILE, "a") as f:
            f.write(f"{name} {roll_no}\n")
        students.append({"name": name, "roll_no": roll_no})

        update_dashboard()
        messagebox.showinfo("Success", f"{name} registered successfully!")
        speak(f"{name} registered successfully!")
        entry_name.delete(0, tk.END)
        entry_roll_no_register.delete(0, tk.END)
        popup.destroy()

    def on_no():
        os.remove(image_path)
        popup.destroy()
        messagebox.showinfo("Cancelled", "Registration cancelled.")

    popup = Toplevel(root)
    popup.title("Confirm Photo")
    popup.geometry("300x350")
    popup.resizable(False, False)

    img = Image.open(image_path).resize((250, 250))
    img_tk = ImageTk.PhotoImage(img)
    label_img = Label(popup, image=img_tk)
    label_img.image = img_tk
    label_img.pack(pady=10)

    lbl = Label(popup, text="Is this your photo?", font=FONT_TITLE)
    lbl.pack(pady=5)

    btn_yes = Button(popup, text="Yes", command=on_yes, bg="#4caf50", fg="white", width=10)
    btn_yes.pack(side="left", padx=20, pady=10)

    btn_no = Button(popup, text="No", command=on_no, bg="#f44336", fg="white", width=10)
    btn_no.pack(side="right", padx=20, pady=10)

# --- Capture photo and register student ---
def register_student():
    name = entry_name.get().strip()
    roll_no = entry_roll_no_register.get().strip()
    if not name or not roll_no:
        messagebox.showwarning("Input Error", "Please enter both Name and Roll Number.")
        return
    # Capture photo
    cap = cv2.VideoCapture(0)
    messagebox.showinfo("Capture", "Press 's' to take a photo. Press 'q' to cancel.")
    while True:
        ret, frame = cap.read()
        if not ret:
            break
        cv2.imshow("Capture Photo - Press 's' to save, 'q' to quit", frame)
        key = cv2.waitKey(1)
        if key == ord('s'):
            # Save image
            filename = f"{name}.jpg"
            filepath = os.path.join(KNOWN_FACES_DIR, filename)
            cv2.imwrite(filepath, frame)
            cap.release()
            cv2.destroyAllWindows()
            show_photo_confirmation(filepath, name, roll_no)
            break
        elif key == ord('q'):
            cap.release()
            cv2.destroyAllWindows()
            break

# --- Layout ---

# Left Panel - Search & Mark Attendance
frame_left = tk.Frame(root, padx=20, pady=20)
frame_left.pack(side="left", fill="y")

tk.Label(frame_left, text="Enter Roll Number:", font=FONT_LABEL).pack(anchor="w")
entry_roll_no = tk.Entry(frame_left, font=FONT_LABEL)
entry_roll_no.pack(fill="x", pady=5)

btn_search = tk.Button(frame_left, text="Search Student", font=FONT_LABEL, bg=BUTTON_BG, fg=BUTTON_FG, command=search_student)
btn_search.pack(fill="x", pady=5)

label_student_name = tk.Label(frame_left, text="", font=FONT_TITLE)
label_student_name.pack(pady=5)

label_image = tk.Label(frame_left)
label_image.pack(pady=5)

label_message = tk.Label(frame_left, text="", font=FONT_LABEL, fg="red")
label_message.pack(pady=5)

# Status Radiobuttons
var_status = tk.StringVar(value="Present")
frame_status = tk.Frame(frame_left)
frame_status.pack(pady=5)
tk.Radiobutton(frame_status, text="Present", variable=var_status, value="Present", font=FONT_LABEL).pack(side="left")
tk.Radiobutton(frame_status, text="Absent", variable=var_status, value="Absent", font=FONT_LABEL).pack(side="left")
tk.Radiobutton(frame_status, text="Leave", variable=var_status, value="Leave", font=FONT_LABEL).pack(side="left")

btn_mark = tk.Button(frame_left, text="Mark Attendance", font=FONT_LABEL, bg=BUTTON_BG, fg=BUTTON_FG, command=mark_attendance)
btn_mark.pack(fill="x", pady=5)

btn_recognize = tk.Button(frame_left, text="Recognize & Mark Attendance", font=FONT_LABEL, bg=BUTTON_BG, fg=BUTTON_FG, command=recognize_and_mark)
btn_recognize.pack(fill="x", pady=5)

# Right Panel - Register Student
frame_right = tk.Frame(root, padx=20, pady=20)
frame_right.pack(side="right", fill="y")

tk.Label(frame_right, text="Register New Student", font=FONT_TITLE).pack(pady=10)

tk.Label(frame_right, text="Name:", font=FONT_LABEL).pack(anchor="w")
entry_name = tk.Entry(frame_right, font=FONT_LABEL)
entry_name.pack(fill="x", pady=5)

tk.Label(frame_right, text="Roll Number:", font=FONT_LABEL).pack(anchor="w")
entry_roll_no_register = tk.Entry(frame_right, font=FONT_LABEL)
entry_roll_no_register.pack(fill="x", pady=5)

btn_register = tk.Button(frame_right, text="Register Student with Photo Capture", font=FONT_LABEL, bg=BUTTON_BG, fg=BUTTON_FG, command=register_student)
btn_register.pack(fill="x", pady=15)

# Bottom Panel - Dashboard
frame_bottom = tk.Frame(root, padx=20, pady=20)
frame_bottom.pack(side="bottom", fill="x")

label_total_students = tk.Label(frame_bottom, text="Total Students: 0", font=FONT_LABEL)
label_total_students.pack(side="left", padx=10)

label_total_attendance = tk.Label(frame_bottom, text="Total Attendance Today: 0", font=FONT_LABEL)
label_total_attendance.pack(side="left", padx=10)

label_present_count = tk.Label(frame_bottom, text="Present: 0", font=FONT_LABEL, fg="green")
label_present_count.pack(side="left", padx=10)

label_absent_count = tk.Label(frame_bottom, text="Absent: 0", font=FONT_LABEL, fg="red")
label_absent_count.pack(side="left", padx=10)

label_leave_count = tk.Label(frame_bottom, text="Leave: 0", font=FONT_LABEL, fg="orange")
label_leave_count.pack(side="left", padx=10)

# Initialize dashboard stats
update_dashboard()

root.mainloop()
