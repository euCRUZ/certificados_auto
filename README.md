<h1>Certificate Generator Project</h1>
<h2>Overview<br></h2>
This Python project utilizes the openpyxl library to extract data from a spreadsheet and then creates personalized certificates by overlaying the information onto a predefined certificate template image using the PIL (Pillow) library. The certificates are saved as individual PNG files in a specified directory.

<h2>Requirements<br></h2>
Make sure you have the following libraries installed before running the script:<br>
<br>
openpyxl<br>
Pillow<br>
You can install them using the following commands:<br>
<br>
bash<br>
Copy code<br>
pip install openpyxl<br>
pip install Pillow<br>
<br>
<h2>Project Structure<br></h2>h2>

spreadsheet_students.xlsx: Excel spreadsheet containing student information.<br>
default_certification.jpg: Template image for the certificates.<br>
fonts/: Directory containing Tahoma and Tahoma Bold fonts.<br>
certificate_generator.py: Python script to read data from the spreadsheet and generate certificates.<br>
<br>Usage<br>
Replace the sample data in the spreadsheet_students.xlsx file with the actual student information.
Ensure that the template image (default_certification.jpg) and font files in the fonts/ directory are suitable for your certificates.
Run the certificate_generator.py script.
Certificate Output
Generated certificates will be saved in the certifications/ directory with filenames in the format: <index>_<participant_name>_certification.png.

<h2>Note</h2>h2>
The script assumes that the relevant information is present in the spreadsheet columns (class name, participant name, participation type, workload, start date, end date, issue date).
Feel free to customize the script and template image to better suit your requirements.
