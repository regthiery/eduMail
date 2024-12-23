import os
import subprocess
from unidecode import unidecode
import argparse
import configparser

# Function to handle command-line arguments
def get_config_path():
    parser = argparse.ArgumentParser(description="Read a configuration file")
    parser.add_argument('-c', '--config', type=str, required=True, help="Path to the configuration file")
    args = parser.parse_args()
    return args.config

if __name__ == '__main__':
    # Read the specified configuration file
    config_path = get_config_path()

    # Check if the configuration file exists
    if not os.path.exists(config_path):
        print(f"Error: Configuration file '{config_path}' does not exist.")
        exit(1)

    # Create a configparser object to read the configuration file
    config = configparser.ConfigParser()

    # Load the configuration file
    config.read(config_path)

    # Read configuration values
    student_info_file = config.get('paths', 'students_info_file')
    email_template_file = config.get('paths', 'email_template_file')
    pdf_folder = config.get('paths', 'pdf_folder')
    comments_folder = config.get('paths', 'comments_folder')
    
    # Email subject
    email_subject = config.get('email', 'subject')
    cc_email = config.get('email', 'cc_email')

    # Reading coefficients from the configuration
    coefficients = [float(c) for c in config.get('notes', 'coefficients').split(',')]
    notes_count = config.getint('notes', 'notes_count')

    # Validating the number of coefficients
    if len(coefficients) != notes_count:
        raise ValueError("The number of coefficients does not match the specified number of notes.")

    # Checking if the sum of coefficients equals 1
    if round(sum(coefficients), 2) != 1.0:
        raise ValueError("The sum of the coefficients must equal 1.")

    # Check configuration options for including comment and Google Drive link
    include_comment = config.getboolean('options', 'include_comment')
    include_google_drive_link = config.getboolean('options', 'include_google_drive_link')


    # Function to escape special characters for AppleScript
    def escape_applescript_text(text):
        return (
            text.replace("\\", "\\\\")  # Escape backslashes
            .replace('"', '\\"')       # Escape double quotes
            .replace("\n", "\\n")      # Escape newlines
        )

    # Read the student information file
    with open(student_info_file, 'r', encoding='utf-8') as info_file:
        student_lines = info_file.readlines()

    # Read the email template
    with open(email_template_file, 'r', encoding='utf-8') as template_file:
        email_template = template_file.read()

    for index, line in enumerate(student_lines, start=1):
        line = line.strip()
        if line and not line.startswith('#'):  # Skip empty lines or comments
            # Split the line into columns
            columns = line.split('\t')

            # Extract necessary information
            student_id = columns[0]
            lastname = unidecode(columns[1].replace(' ', '_'))
            firstname = unidecode(columns[2].replace(' ', '_'))
            student_number = columns[3]

            # Extract and compute scores based on the number of notes and their coefficients
            scores = [float(columns[4 + i]) for i in range(notes_count)]
            average_score = round(sum(score * coef for score, coef in zip(scores, coefficients)), 2)

            if include_comment:
                comment_file = columns[4 + notes_count]
            if include_google_drive_link:                
                google_drive_link = columns[4 + notes_count]

            email = columns[-1]

            if include_comment:
                # Read the content of the comment file
                comment_file_path = os.path.join(os.path.dirname(comments_folder), comment_file)
                try:
                    with open(comment_file_path, 'r', encoding='utf-8') as comment_file_content:
                        comment_content = comment_file_content.read().strip()
                except FileNotFoundError:
                    print(f"Comment file not found: {comment_file_path}")
                    comment_content = "[Content not available]"

            # Attachment name
            attachment_name = f"{student_id}-{firstname}_{lastname}.pdf"
            attachment_path = os.path.join(pdf_folder, attachment_name)

            # Check if the attachment exists
            if not os.path.exists(attachment_path):
                print(f"Attachment not found: {attachment_path}")
                continue

            # Fill in the email body
            email_body = email_template.replace('<firstname>', firstname)
            email_body = email_body.replace('<lastname>', lastname)

            # Replace placeholders for individual scores
            for i, score in enumerate(scores, start=1):
                email_body = email_body.replace(f'<score{i}>', str(score))

            # Replace placeholder for the average score
                email_body = email_body.replace('<average_score>', str(average_score))

            # Replace placeholder for comment if enabled
            if include_comment:
                email_body = email_body.replace('<comment>', comment_content)
            else:
                email_body = email_body.replace('<comment>', '')

            # Replace placeholder for Google Drive link if enabled
            if include_google_drive_link:
                email_body = email_body.replace('<google_drive_link>', google_drive_link)
            else:
                email_body = email_body.replace('<google_drive_link>', '')

            # Escape text for AppleScript
            escaped_email_body = escape_applescript_text(email_body)

            # Prepare the AppleScript to open Mail
            applescript = f'''
tell application "Mail"
    set newMessage to make new outgoing message with properties {{subject:"{email_subject}", content:"{escaped_email_body}", visible:true}}
    tell newMessage
        if "{email}" is not equal to "" then
            make new to recipient at end of to recipients with properties {{address:"{email}"}}
        end if    
        if "{attachment_path}" is not equal to "" then
            make new attachment with properties {{file name:"{attachment_path}"}} at after the last paragraph
        end if    
        if "{cc_email}" is not equal to "" then
            make new cc recipient at end of cc recipients with properties {{address:"{cc_email}"}}
        end if    
    end tell
end tell
'''

            # Execute the AppleScript
            subprocess.run(["osascript", "-e", applescript])
            print(f"Email body prepared successfully for {firstname} {lastname}.")

