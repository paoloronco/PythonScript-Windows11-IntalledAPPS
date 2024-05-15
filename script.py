import subprocess
import os
import time

def get_desktop_path():
    return os.path.join(os.path.expanduser('~'), 'Desktop')

def get_installed_apps():
    try:
        # Run the command to get the list of installed apps
        result = subprocess.run(['wmic', 'product', 'get', 'Name'], capture_output=True, text=True, timeout=10)

        # Extract app names from the output string
        installed_apps = result.stdout.split('\n')
        
        # Remove the first line containing the header
        installed_apps = installed_apps[1:]

        # Remove any empty elements
        installed_apps = [app.strip() for app in installed_apps if app.strip()]

        return installed_apps

    except subprocess.TimeoutExpired:
        print("Timeout expired while executing the command.")
        return None
    except Exception as e:
        print(f"An error occurred: {e}")
        return None

def save_to_file(apps, filename):
    try:
        with open(filename, 'w') as file:
            for app in apps:
                file.write(app + '\n')
        print(f"The file {filename} has been successfully created.")
    except Exception as e:
        print(f"An error occurred while saving the file: {e}")

def main():
    # Get the desktop path
    desktop_path = get_desktop_path()

    # Output file name
    output_file = os.path.join(desktop_path, 'installed_apps.txt')

    # Add a delay before getting the list of installed apps
    time.sleep(2)  # 2 seconds delay

    # Get the list of installed apps
    installed_apps = get_installed_apps()

    if installed_apps:
        # Save the list of installed apps to a text file
        save_to_file(installed_apps, output_file)
    else:
        print("Unable to get the list of installed apps.")

if __name__ == "__main__":
    main()
