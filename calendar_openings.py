from datetime import datetime
import os
import webbrowser

def generate_calendar_times_txt():
    from calendar_functions import get_availability_text  # Make sure this is imported correctly

    try:
        # Get available times
        availability = get_availability_text()

        # Build output string
        timestamp = datetime.now().strftime("%A, %B %d at %I:%M %p")
        content = f"{availability}\n\nLast updated: {timestamp}"

        # Write to text file
        file_path = os.path.join(os.getcwd(), "calendar_times.txt")
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(content)

        # Open the file in the system's default text editor
        os.startfile(file_path) if os.name == "nt" else webbrowser.open(file_path)

        print(f"✅ Calendar times saved and opened for editing: {file_path}")

    except Exception as e:
        print(f"❌ Failed to generate calendar times: {e}")

def load_calendar_text_from_file(file_path="calendar_times.txt"):
    """Reads availability text from a local text file, excluding 'Last updated' line."""
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            lines = f.readlines()
            if lines and lines[-1].lower().startswith("last updated:"):
                lines = lines[:-1]
            return "".join(lines).strip()
    except FileNotFoundError:
        return "⚠️ Availability file not found. Please click 'Update Calendar Times' first."
