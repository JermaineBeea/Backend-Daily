import os

# Define directory and file paths
INVITED_NAMES_PATH = os.path.abspath('mail merge/Names/inivited_names.txt')
STARTING_LETTER_PATH = os.path.abspath('mail merge/letters/starting_letter.docx')
READY_TO_SEND_DIR = os.path.abspath('ReadyToSend')

# Ensure the "ReadyToSend" directory exists
os.makedirs(READY_TO_SEND_DIR, exist_ok=True)

def load_names(file_path):
    """Load names from file."""
    with open(file_path, mode='r') as file:
        return [line.strip() for line in file.readlines()]


def load_template(file_path):
    """Load template content."""
    with open(file_path, mode='r') as file:
        return file.read()


def write_letter(name, content):
    """Save letter for each name."""
    modified_letter_path = os.path.join(READY_TO_SEND_DIR, f'{name}.docx')
    with open(modified_letter_path, mode='w') as file:
        file.write(content)


def create_letter():
    """Generate letters for each name."""
    try:
        names = load_names(INVITED_NAMES_PATH)
        template_content = load_template(STARTING_LETTER_PATH)

        for name in names:
            personalized_content = template_content.replace('[name]', name)
            write_letter(name, personalized_content)

        print("Letters created successfully in the 'ReadyToSend' folder.")
    except FileNotFoundError as e:
        print(f"Error: {e}")


if __name__ == "__main__":

    create_letter()
