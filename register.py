# pyrefly: ignore [missing-import]
import cv2
import os

# Students folder
STUDENT_FOLDER = "students"

# Create folder if it doesn't exist
if not os.path.exists(STUDENT_FOLDER):
    os.makedirs(STUDENT_FOLDER)

# Ask student name
name = input("Enter Student Name: ").strip()

if not name:
    print("Name cannot be empty.")
    exit()

# Create folder for this student
student_path = os.path.join(STUDENT_FOLDER, name)

if not os.path.exists(student_path):
    os.makedirs(student_path)

# Face detector
cascade_path = "haarcascade_frontalface_default.xml"

# Fallback to OpenCV's bundled cascade file if local copy is missing or empty
if not os.path.exists(cascade_path) or os.path.getsize(cascade_path) == 0:
    cascade_path = cv2.data.haarcascades + "haarcascade_frontalface_default.xml"

face_detector = cv2.CascadeClassifier(cascade_path)

if face_detector.empty():
    print("ERROR: Face detector file not found.")
    print("Path:", cascade_path)
    exit()

# Open camera with DirectShow backend (fixes Windows MSMF error -1072875772)
def open_camera():
    for idx in [0, 1, 2]:
        for backend in [cv2.CAP_DSHOW, cv2.CAP_MSMF, cv2.CAP_ANY]:
            cap = cv2.VideoCapture(idx, backend)
            if cap.isOpened():
                ret, frame = cap.read()
                if ret and frame is not None:
                    return cap
                cap.release()
    return cv2.VideoCapture(0)

camera = open_camera()

if not camera.isOpened():
    print("Camera could not be opened.")
    exit()

sample_count = 0

print("\nCamera started.")
print("Keep your face in front of the camera.")
print("Samples will be captured automatically.")
print("Press Q to stop.\n")

while True:
    success, frame = camera.read()

    if not success:
        print("Could not read camera.")
        break

    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

    faces = face_detector.detectMultiScale(
        gray,
        scaleFactor=1.3,
        minNeighbors=5,
        minSize=(100, 100)
    )

    for (x, y, w, h) in faces:
        sample_count += 1

        # Draw rectangle
        cv2.rectangle(
            frame,
            (x, y),
            (x + w, y + h),
            (0, 255, 0),
            2
        )

        cv2.putText(
            frame,
            f"Samples: {sample_count}/30",
            (x, y - 10),
            cv2.FONT_HERSHEY_SIMPLEX,
            0.8,
            (0, 255, 0),
            2
        )

        # Save face sample
        face_image = gray[y:y + h, x:x + w]

        file_path = os.path.join(
            student_path,
            f"{sample_count}.jpg"
        )

        cv2.imwrite(file_path, face_image)

        break

    cv2.imshow("Student Registration", frame)

    # Stop after 30 samples
    if sample_count >= 30:
        break

    # Press Q to stop
    if cv2.waitKey(100) & 0xFF == ord("q"):
        break

camera.release()
cv2.destroyAllWindows()

print("\nRegistration completed!")
print(f"Student: {name}")
print(f"Samples saved: {sample_count}")
print(f"Location: {student_path}")