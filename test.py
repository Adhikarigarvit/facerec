


import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime
from sklearn.neighbors import KNeighborsClassifier
from win32com.client import Dispatch

# Function for speech output
def speak(text):
    speak = Dispatch(("SAPI.SpVoice"))
    speak.Speak(text)

# Function to check if FACES and LABELS have the same number of samples
def validate_data(FACES, LABELS):
    if FACES.shape[0] != len(LABELS):
        print("Mismatch between faces and labels!")
        # Trimming LABELS or FACES to match the size
        min_samples = min(FACES.shape[0], len(LABELS))
        FACES = FACES[:min_samples]
        LABELS = LABELS[:min_samples]
        print(f"Adjusted FACES and LABELS to {min_samples} samples each.")
    return FACES, LABELS

# Load the pre-trained data
video = cv2.VideoCapture(0)
facedetect = cv2.CascadeClassifier('data/haarcascade_frontalface_default.xml')

with open('data/names.pkl', 'rb') as w:
    LABELS = pickle.load(w)

with open('data/faces_data.pkl', 'rb') as f:
    FACES = pickle.load(f)

print('Shape of Faces matrix --> ', FACES.shape)

# Validate FACES and LABELS consistency
FACES, LABELS = validate_data(FACES, LABELS)

# Initialize KNN classifier and fit with the face data
knn = KNeighborsClassifier(n_neighbors=5)
knn.fit(FACES, LABELS)

# Load background image for UI display
imgBackground = cv2.imread("background.png")

# Define column names for the CSV file
COL_NAMES = ['NAME', 'TIME']

# Start video capture loop
while True:
    ret, frame = video.read()
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = facedetect.detectMultiScale(gray, 1.3, 5)

    for (x, y, w, h) in faces:
        crop_img = frame[y:y+h, x:x+w, :]
        resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)
        output = knn.predict(resized_img)

        # Get the timestamp
        ts = time.time()
        date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp = datetime.fromtimestamp(ts).strftime("%H:%M:%S")

        # Check if attendance file exists
        exist = os.path.isfile(f"Attendance/Attendance_{date}.csv")

        # Draw bounding boxes and labels on the frame
        cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 0, 255), 1)
        cv2.rectangle(frame, (x, y), (x + w, y + h), (50, 50, 255), 2)
        cv2.rectangle(frame, (x, y - 40), (x + w, y), (50, 50, 255), -1)
        cv2.putText(frame, str(output[0]), (x, y - 15), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 1)

        # Prepare attendance entry
        attendance = [str(output[0]), str(timestamp)]

    # Place the video frame on the background image
    imgBackground[162:162 + 480, 55:55 + 640] = frame
    cv2.imshow("Frame", imgBackground)

    # Wait for key press events
    k = cv2.waitKey(1)

    if k == ord('o'):  # When 'o' is pressed, take attendance
        speak("Attendance Taken..")
        time.sleep(5)

        # Append or create attendance CSV file
        if exist:
            with open(f"Attendance/Attendance_{date}.csv", "a") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(attendance)
        else:
            with open(f"Attendance/Attendance_{date}.csv", "a") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(COL_NAMES)
                writer.writerow(attendance)

    if k == ord('q'):  # When 'q' is pressed, exit the program
        break

# Release the video capture and close any OpenCV windows
video.release()
cv2.destroyAllWindows()

