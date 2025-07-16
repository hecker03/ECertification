import requests
import pytesseract
import pandas as pd
import re
import cv2
import os
import smtplib
from PIL import Image, ImageDraw, ImageFont, ImageTk
from email.message import EmailMessage
import numpy as np
import matplotlib.pyplot as plt
from dotenv import load_dotenv

load_dotenv()

formlink = os.getenv("FORMLINK")
API_KEY = os.getenv("API_KEY")
password = os.getenv("EMAIL_PASSWORD")

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
pathofimage = "Certify.png"           # original template
updateimage = "Imageupdate"         # base name for updated images
SMTP_PORT =  os.getenv('SMTP_PORT') 
SMTP_SERVER = os.getenv('SMTP_SERVER')# For Gmail
RANGE_NAME = "Sheet2!A1:L100"

def extract_sheet_id(sheet_url):
    pattern = r"/d/([a-zA-Z0-9-_]+)"
    match = re.search(pattern, sheet_url)
    return match.group(1) if match else None

SHEET_ID = extract_sheet_id(formlink)

def find_coords(y_min, y_max, lines, underline_coords):
    if lines is not None :
        i = 0
        while( len(underline_coords) == 0):
            # Support both flat and nested formats
            if isinstance(lines[i], (list, np.ndarray)):
                if len(lines[i]) == 4:
                    x1, y1, x2, y2 = lines[i]
                elif len(lines[i]) > 0 and len(lines[i][0]) == 4:
                    x1, y1, x2, y2 = lines[i][0]
                else:
                    continue

                # Print each line's coordinates for debugging
                print(f"Line coordinates: {(x1, y1, x2, y2)}")
                # Check if both endpoints fall within the window
                if y_min <= y1 <= y_max and y_min <= y2 <= y_max:
                    underline_coords.append((x1, y1, x2, y2))
            i = i +1
    return underline_coords
def addbg(updated, j):
    # Open the transparent overlay image (reference size)
    overlay = Image.open(updated).convert("RGBA")

    # Open the background image
    background = Image.open("image.png").convert("RGBA")

    # Resize background to match the overlay's size
    background = background.resize(overlay.size)

    # Increase brightness (adjust factor as needed)

    # Adjust transparency of the background
    alpha_value = 175  # Range: 0 (fully transparent) to 255 (fully opaque)
    overlay.putalpha(alpha_value)

    # Blend the transparent brightened background with the overlay
    combined = Image.alpha_composite(Image.new("RGBA", overlay.size, (0, 0, 0, 0)), background)
    combined = Image.alpha_composite(combined, overlay)

    # enhancer = ImageEnhance.Brightness(combined)
    # combined= enhancer.enhance(0.5)  # Increase brightness (1.0 = original, >1.0 = brighter)
    # Show or save the result
    # combined.show()  # Display the final image
    output=f"output{j}.png"
    combined.save(output)  # Save the image
    return output


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
y_center =0
# Identify key text Y-coordinates (for "to" and "upon")
award_text_y, participation_text_y, project_text_y = None, None, None
for i in range(len(detection_data["text"])):
    text = detection_data["text"][i].strip().lower()
    x_center = (detection_data["left"][i] + detection_data["width"][i])//2
    y_center = detection_data["top"][i] + detection_data["height"][i] // 2
    if "certificate" in text:
        award_text_y = y_center
        print(f"Found 'to' at y-center: {y_center} (text: {text})")
    elif "organized" in text:
        participation_text_y = y_center
        print(f"Found 'the' at y-center: {y_center} (text: {text})")
    # elif "prinicpal" in text:  # replaced "upon" with "course"
    #     project_text_y = y_center
    #     print(f"Found 'course' at y-center: {y_center} (text: {text})")

if award_text_y is None or participation_text_y is None: 
# or project_text_y is None:
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
underlines = underline_coords[:]
del underline_coords[:]
if participation_text_y :
    y_min, y_max = participation_text_y -10 , participation_text_y + 100
    print(f"Looking for lines between y: {y_min} and {y_max}")
    underline_coords = find_coords(y_min, y_max, lines, underline_coords)
underlines.extend(underline_coords)
print("Filtered underline coordinates:", underlines)

# ----- FETCH DATA FROM GOOGLE SHEET -----
try:
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}/values/{RANGE_NAME}?key={API_KEY}"
    response = requests.get(url)
    print(response)
except Exception as e:
    print("Error fetching data:", e)
    exit()

data = response.json().get("values", [])
print(data)
if len(data) > 1:
    df = pd.DataFrame(data[1:], columns=data[0])
else:
    print("No data found in the Google Sheet.")
    exit()
print("Column names:", df.columns)

# ----- EMAIL & DRAWING LOOP -----
for index, row in df.iterrows():
    email = row['Email Address']
    Name1 = row['Student name 1'].upper()
    Name2 = row['Student name 2'].upper()
    Name3 = row['Student name 3'].upper()
    Name4 = row['Student name 4'].upper()
    Project = row['Project Name']
    Rank = row['Rank']  # Assuming you have added a 'Rank' column in your Google Sheet
    Year = row['Year']
    # print(Name4)
    count = 3
    Names = [Name1, Name2, Name3] 
    if len(Name4.split(",")) == 2:
        count = 5
        Names.append(Name4.split(",")[0])
        Names.append(Name4.split(",")[1])
    elif Name4.strip() != "":
        count = 4
        Names.append(Name4)
    print(Names)
    # 1) Create a FRESH copy of the original PIL image for each iteration
    image_pil = Image.open(pathofimage).convert("RGB")
    draw = ImageDraw.Draw(image_pil)
    font = ImageFont.truetype("comic.ttf", 27)  # Adjust size as needed

    # 4) Prepare and send email
    msg = EmailMessage()
    semail = os.getenv('SEMAIL')
    msg["From"] = semail
    msg["To"] = email
    if Year == "SE":
        if Rank == "1":
            msg["Subject"] = "Congratulations on 1st Rank!"
            msg.set_content("Dear Team Leader, \nCongratulations on achieving the 1st rank in the PBL Project Competition. Please find the attached certificates.")
            pathofimage = "SE1st_rank.png"  # Use a different template for 1st rank
        elif Rank == "2":
            msg["Subject"] = "Congratulations on 2nd Rank!"
            msg.set_content("Dear Team Leader, \nCongratulations on achieving the 2nd rank in the PBL Project Competition. Please find the attached certificates.")
            pathofimage = "SE2nd_rank.png"  # Use a different template for 2nd rank
        elif Rank == "3":
            msg["Subject"] = "Congratulations on 3rd Rank!"
            msg.set_content("Dear Team Leader, \nCongratulations on achieving the 3rd rank in the PBL Project Competition. Please find the attached certificates.")
            pathofimage = "SE3rd_rank.png"  # Use a different template for 3rd rank
        else:
            msg["Subject"] = "Here is your attachment"
            msg.set_content("Dear Team Leader,\nAll the members from your team have received E-Certificate on the basis of Participating in the PBL Project Competition. Please find the attached images.")
            pathofimage = "SEParticipants.png"  # Use the default template for participation
    elif Year == "FY":
        if Rank == "1":
            msg["Subject"] = "Congratulations on 1st Rank!"
            msg.set_content("Dear Team Leader, \nCongratulations on achieving the 1st rank in the PBL Project Competition. Please find the attached certificates.")
            pathofimage = "FY1st_rank.png"  # Use a different template for 1st rank
        elif Rank == "2":
            msg["Subject"] = "Congratulations on 2nd Rank!"
            msg.set_content("Dear Team Leader, \nCongratulations on achieving the 2nd rank in the PBL Project Competition. Please find the attached certificates.")
            pathofimage = "FY2nd_rank.png"  # Use a different template for 2nd rank
        elif Rank == "3":
            msg["Subject"] = "Congratulations on 3rd Rank!"
            msg.set_content("Dear Team Leader, \nCongratulations on achieving the 3rd rank in the PBL Project Competition. Please find the attached certificates.")
            pathofimage = "FY3rd_rank.png"  # Use a different template for 3rd rank
        else:
            msg["Subject"] = "Here is your attachment"
            msg.set_content("Dear Team Leader, \nAll the members from your team have received E-Certificate on the basis of Participating in the PBL Project Competition. Please find the attached images.")
            pathofimage = "FYParticipants.png"  # Use the default template for participation

    if underlines:
        # Use the first underline for student name and the second for project name.
        x1, y1, x2, y2 = underlines[0]
        text_x = x1 + 10
        text_y = y1 
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
        xp1, yp1, xp2, yp2 = underlines[1]
        text_xp = x_center + 150  
        text_yp = yp1 - 20
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

            bbox = draw.textbbox((0, 0), text=Names[i-1], font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]

            # For demonstration, using Name1 for all iterations; update as necessary.
            draw.text((text_x, text_y), Names[i-1], fill="black", font=font)
            draw.text((text_xp, text_yp), Project, fill="black", font=font)
            # Save updated image under a unique filename
            image_pil.save(Ima)
            final = addbg(Ima, i)
            print(final)
            Images.append(final)

        # Attach each updated image to the email
        for updates in Images:
            FILE_NAME = os.path.basename(updates)
            with open(updates, "rb") as file:
                file_data = file.read()
                file_type = updates.split(".")[-1]
                maintype = "application" if file_type == "pdf" else "image"
                subtype = file_type
                msg.add_attachment(file_data, maintype=maintype, subtype=subtype, filename=FILE_NAME)
        
        # Send email
        if SMTP_PORT and SMTP_SERVER:
            with smtplib.SMTP(SMTP_SERVER, int(SMTP_PORT)) as server:
                server.starttls()
                if semail and password:
                    server.login(semail, password=password)
                server.send_message(msg)
                print(f"Email sent successfully to {email}") 
    else:
        print("No underline coordinates available;")

# ----- OPTIONAL: SHOW DETECTED UNDERLINES -----
for (x1, y1, x2, y2) in underlines:
    cv2.line(image, (x1, y1), (x2, y2), (0, 255, 255), 2)
plt.figure(figsize=(10, 6))
plt.imshow(cv2.cvtColor(image, cv2.COLOR_BGR2RGB))
plt.title("Detected Underlines")
plt.axis("off")
plt.show()
