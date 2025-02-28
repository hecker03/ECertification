import requests
import pytesseract
import pandas as pd
import re
import cv2
import os
import smtplib
from PIL import Image, ImageDraw, ImageFont
from email.message import EmailMessage
import numpy as np
import matplotlib.pyplot as plt

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
pathofimage = "Certi.png"           # original template
updateimage = "Imageupdate"         # base name for updated images
SMTP_PORT = 587  # TLS port
SMTP_SERVER = "smtp.gmail.com"  # For Gmail
password = "espxlpmxcycpqqik"
formlink = "https://docs.google.com/spreadsheets/d/1T7-6zcZ8U4-hKTPn3T9RB7bYyu187TEflmMdnUSIsFs/edit?gid=24612527#gid=24612527"
API_KEY = "AIzaSyDtTicI-eSbEKFZlJ3QrMLt7XMBNoxheKg"
RANGE_NAME = "Sheet1!A1:G100"

def extract_sheet_id(sheet_url):
    pattern = r"/d/([a-zA-Z0-9-_]+)"
    match = re.search(pattern, sheet_url)
    return match.group(1) if match else None

SHEET_ID = extract_sheet_id(formlink)

def find_coords(y_min, y_max, lines, underline_coords):
    if lines is not None:
        for line in lines:
            # Support both flat and nested formats
            if isinstance(line, (list, np.ndarray)):
                if len(line) == 4:
                    x1, y1, x2, y2 = line
                elif len(line) > 0 and len(line[0]) == 4:
                    x1, y1, x2, y2 = line[0]
                else:
                    continue

                # Print each line's coordinates for debugging
                print(f"Line coordinates: {(x1, y1, x2, y2)}")
                # Check if both endpoints fall within the window
                if y_min <= y1 <= y_max and y_min <= y2 <= y_max:
                    underline_coords.append((x1, y1, x2, y2))
    return underline_coords

# ----- LOAD & PREPARE THE ORIGINAL IMAGE FOR LINE DETECTION -----
image = cv2.imread(pathofimage)
if image is None:
    print("Image not loaded; check the file path.")
    exit()

gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)

# Perform OCR to detect text regions
custom_config = r'--oem 3 --psm 6'
detection_data = pytesseract.image_to_data(binary, config=custom_config, output_type=pytesseract.Output.DICT)
x_center=0
# Identify key text Y-coordinates (for "to" and "upon")
award_text_y, participation_text_y, project_text_y = None, None, None
for i in range(len(detection_data["text"])):
    text = detection_data["text"][i].strip().lower()
    x_center = (detection_data["left"][i] + detection_data["width"][i])//2
    y_center = detection_data["top"][i] + detection_data["height"][i] // 2
    if "certificate" in text:
        award_text_y = y_center
        print(f"Found 'to' at y-center: {y_center} (text: {text})")
    elif "participation" in text:
        participation_text_y = y_center
        print(f"Found 'the' at y-center: {y_center} (text: {text})")

if award_text_y is None or participation_text_y is None:
    print("Could not find reference texts 'certificate' and/or 'participation' in OCR data.")

# Detect horizontal lines using Canny edge detection + HoughLinesP
edges = cv2.Canny(binary, 50, 150, apertureSize=3)
lines = cv2.HoughLinesP(edges, 1, np.pi / 180, threshold=80, minLineLength=80, maxLineGap=15)
print("Raw detected lines:", lines)

underline_coords = []
if award_text_y and participation_text_y:
    y_min, y_max = award_text_y + 10, participation_text_y - 10
    print(f"Looking for lines between y: {y_min} and {y_max}")
    underline_coords = find_coords(y_min, y_max, lines, underline_coords)

print("Filtered underline coordinates:", underline_coords)

# ----- FETCH DATA FROM GOOGLE SHEET -----
try:
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}/values/{RANGE_NAME}?key={API_KEY}"
    response = requests.get(url)
except Exception as e:
    print("Error fetching data:", e)
    exit()

data = response.json().get("values", [])
df = pd.DataFrame(data[1:], columns=data[0])

# ----- EMAIL & DRAWING LOOP -----
for index, row in df.iterrows():
    email = row['Email Address']
    Name1 = row['Student name 1'].upper()
    Name2 = row['Student name 2'].upper()
    Name3 = row['Student name 3'].upper()
    Name4 = row['Student name 4'].upper()
    Project = row['Project Name']
    # print(Name4)
    count = 3
    Names = [Name1, Name2, Name3] 
    if Name4.isalpha() or Name4 is not None or Name4 != "":
        count = 4
        Names.append(Name4)
    print(Names)
    # 1) Create a FRESH copy of the original PIL image for each iteration
    image_pil = Image.open(pathofimage).convert("RGB")
    draw = ImageDraw.Draw(image_pil)
    font = ImageFont.truetype("arial.ttf", 30)

    # 4) Prepare and send email
    msg = EmailMessage()
    msg["Subject"] = "Here is your attachment"
    semail = "shrishailya_keskar_it@moderncoe.edu.in"
    msg["From"] = semail
    msg["To"] = email
    msg.set_content("Please find the attached file.")

    if underline_coords:
        # Use the first underline for student name and the second for project name.
        x1, y1, x2, y2 = underline_coords[0]
        text_x = x1 - 140
        text_y = y1 - 70
        rect_padding = 5
        # Erase area with white rectangle
        draw.rectangle(
            [
                x1 - rect_padding,
                y1 - rect_padding,
                x2 + rect_padding,
                y2 + rect_padding
            ],
            fill="white"
        )

        # Build a list of image filenames for each updated image
        Images = []
        for i in range(1, count+1):
            # If you need to use different names, adjust this part accordingly.
            # Here, each updated image gets a unique file name
            image_pil = Image.open(pathofimage).convert("RGB")
            draw = ImageDraw.Draw(image_pil)
            Ima = f"{updateimage}{i}.png"
            Images.append(Ima)

            bbox = draw.textbbox((0, 0), text=Names[i-1], font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]

            # For demonstration, using Name1 for all iterations; update as necessary.
            draw.text((text_x, text_y), Names[i-1], fill="black", font=font)
            # draw.text((text_xp, text_yp), Project, fill="black", font=font)

            # Save updated image under a unique filename
            image_pil.save(Ima)

        # Attach each updated image to the email
        for imageupdate in Images:
            FILE_NAME = os.path.basename(imageupdate)
            with open(imageupdate, "rb") as file:
                file_data = file.read()
                file_type = imageupdate.split(".")[-1]
                maintype = "application" if file_type == "pdf" else "image"
                subtype = file_type
                msg.add_attachment(file_data, maintype=maintype, subtype=subtype, filename=FILE_NAME)
        
        # Send email
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(semail, password)
            server.send_message(msg)
            print(f"Email sent successfully to {email}") 
    else:
        print("No underline coordinates available;" )

# ----- OPTIONAL: SHOW DETECTED UNDERLINES -----
for (x1, y1, x2, y2) in underline_coords:
    cv2.line(image, (x1, y1), (x2, y2), (0, 255, 255), 2)
plt.figure(figsize=(10, 6))
plt.imshow(cv2.cvtColor(image, cv2.COLOR_BGR2RGB))
plt.title("Detected Underlines")
plt.axis("off")
plt.show()
