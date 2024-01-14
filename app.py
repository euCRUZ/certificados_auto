# Get data from spreadsheet
import openpyxl
from PIL import Image, ImageDraw, ImageFont

# Open spreadsheets
workbook_students = openpyxl.load_workbook('spreadsheet_students.xlsx')
sheet_students = workbook_students['Sheet1']

for index, row in enumerate(sheet_students.iter_rows(min_row=2)):
    # each cell that contains the information you need
    class_name = row[0].value
    participant_name = row[1].value
    participation_type = row[2].value
    workload = row[5].value

    start_date = row[3].value
    end_date = row[4].value

    issue_date = row[6].value

# Transfer spreadsheet data to certificate image
font_name = ImageFont.truetype('./fonts/tahomabd.ttf',90)
font_general = ImageFont.truetype('./fonts/tahoma.ttf',70)
font_date = ImageFont.truetype('./fonts/tahoma.ttf',45,)

image = Image.open('./default_certification.jpg')
draw = ImageDraw.Draw(image)

draw.text((1040,827), participant_name, fill='black', font=font_name)
draw.text((1080,960), class_name, fill='black', font=font_general)
draw.text((1455,1075), participation_type, fill='black', font=font_general)
draw.text((1500, 1192), str(workload), fill='black', font=font_general)

draw.text((770, 1795), start_date, fill='black', font=font_date)
draw.text((770, 1945), end_date, fill='black', font=font_date)

image.save(f'./certifications/{index}_{participant_name}_certification.png')
