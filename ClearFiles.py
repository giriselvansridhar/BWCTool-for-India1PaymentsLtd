import os
import time

def clear_files_animation(folder):
    files = os.listdir(folder)
    for file_name in files:
        file_path = os.path.join(folder, file_name)
        if os.path.isfile(file_path):
            os.remove(file_path)
            print(f"Deleting: {file_path}")
            time.sleep(0.1)  # Adjust sleep time for animation effect

def clear_files():
    folders_to_clear = [
        'BWCFiles\\Completed',
        'BWCFiles\\Ready',
        'BWCFiles\\Error'
    ]
    
    for folder in folders_to_clear:
        try:
            print(f"Clearing files in: {folder}")
            clear_files_animation(folder)
            print(f"Finished clearing files in: {folder}\n")
            
        except FileNotFoundError:
            print(f"Folder '{folder}' not found. Skipping.")

if __name__ == "__main__":
    print("Starting file clearing process...\n")
    clear_files()
    print("File clearing process completed.")
