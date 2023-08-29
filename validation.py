import os
import datetime
import json



# Path to the parent directory containing student folders
parent_folder_path = "Students"

# ... (previous part of your code)

# Define the student data
student_data = {
    "PPP001": "Mohamed Hasir",
    "PPP002": "Ganesh Kumar R",
    "PPP003": "Deepa N",
    "PPP004": "Nt. Nallathayammal",
    "PPP005": "Prasanth Govindaraj",
    "PPP006": "Murali T",
    "PPP007": "LEEMAN THOMAS",
    "PPP008": "Vimal Nadarajan",
    "PPP009": "Saravanan Selvam",
    "PPP010": "Srinivasan SR",
    "PPP011": "David Raj",
    "PPP012": "Yogesh Kumar JG",
    "PPP013": "Aravindhan Selvaraj",
    "PPP014": "Naveen Bromiyo A R",
    "PPP016": "Madhan Karthick",
    "PPP017": "Pavithra Selvaraj",
    "PPP018": "Sindhu Laheri Uthaya Surian",
    "PPP019": "Nalina Athinamilagi",
    "PPP020": "Nithya Naveen",
    "PPF001": "Ranjitha",
    "PPF002": "Suganthi Ramaraj",
    "PPF004": "Swathipriya",
    "PPF005": "Jumana",
    "PPF006": "Indira Priyadharshini",
    "PPF007": "Riyas ahamed J"
}

# ... (continue with the code I provided earlier)

# ... (rest of your code)


# Load the expected files configuration
with open("expected_files_config.json", "r") as config_file:
    expected_files_config = json.load(config_file)

def validate_week_folder(week_folder_path, expected_files):
    files_in_folder = os.listdir(week_folder_path)
    
    files_in_folder_stripped = [f.strip() for f in files_in_folder]
    
    present_files = [file for file in expected_files if any(file.lower() == f.lower() for f in files_in_folder_stripped)]
    missing_files = [file for file in expected_files if not any(file.lower() == f.lower() for f in files_in_folder_stripped)]
    
    return present_files, missing_files

# Specific week you want to check
specific_week = "Week01"

# Get current date and time
current_datetime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# Create a report table with Bootstrap styling, additional CSS, and JavaScript for freezing table header
report_table = f"""
<!DOCTYPE html>
<html>
<head>
    <title>Weekly Report</title>
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    <style>
        .fixed-header {{
            position: sticky;
            top: 0;
            background-color: white;
            z-index: 1;
        }}
        .container {{
            max-width: 100%;
            max-height: 80vh;
            overflow-y: auto;
        }}
        .table-container {{
            max-height: calc(100vh - 200px); /* Adjust the value based on your needs */
            overflow-y: auto;
        }}
    </style>
    <script>
        window.addEventListener("DOMContentLoaded", function () {{
            var table = document.querySelector(".table");
            var thead = table.querySelector("thead");

            table.addEventListener("scroll", function () {{
                thead.style.transform = "translateY(" + table.scrollTop + "px)";
            }});
        }});
    </script>
</head>
<body>
    <div class="fixed-header py-3 px-4 bg-light">
        <div class="d-flex justify-content-between">
            <h1 class="mb-0">Weekly Report</h1>
            <div class="text-right">
                <p class="mb-0">Generated: {current_datetime}</p>
            </div>
        </div>
        <h2 class="mb-0">Week: {specific_week}</h2>
    </div>
    <div class="container mt-5">
        <div class="table-responsive table-container">
            <table class="table table-bordered">
                <thead class="thead-dark">
                    <tr>
                        <th>Student ID</th>
                        <th>Student Name</th>
                        <th>Week</th>
                        <th>Completed Task</th>
                        <th>Pending Task</th>
                        <th>Completion Status</th>
                        
                    </tr>
                </thead>
                <tbody>
"""

for student_id, student_name in student_data.items():
    student_folder_path = os.path.join(parent_folder_path, f"{student_id} - {student_name}")
    week_folder_name = specific_week
    week_folder_path = os.path.join(student_folder_path, week_folder_name)

    if os.path.exists(student_folder_path) and os.path.isdir(student_folder_path) and os.path.exists(week_folder_path) and os.path.isdir(week_folder_path):
        expected_files = expected_files_config.get(week_folder_name, {}).get(student_id, [])
        
        present_files, missing_files = validate_week_folder(week_folder_path, expected_files)
        
        present_files_str = ', '.join(present_files)
        missing_files_str = ', '.join(missing_files)

        if len(present_files) == len(expected_files):
            completion_status = "Completed"
        else:
            completion_status = "Pending"
        report_table += f"""
                    <tr>
                        <td>{student_id}</td>
                        <td>{student_name}</td>
                        <td>{week_folder_name}</td>
                        <td>{present_files_str}</td>
                        <td>{missing_files_str}</td>
                        <td>{completion_status}</td>
                    </tr>
        """
    else:
        report_table += f"""
                    <tr>
                        <td>{student_id}</td>
                        <td>{student_name}</td>
                        <td>{week_folder_name}</td>
                        <td colspan="2">Folder or data not found</td>
                    </tr>
        """

report_table += """
                </tbody>
            </table>
        </div>
    </div>
</body>
</html>
"""

# Save report to a file
report_filename = f"{specific_week}_report.html"
with open(report_filename, "w") as report_file:
    report_file.write(report_table)

# Print a message indicating HTML report generation
print(f"HTML report generated: {report_filename}")


#testing
