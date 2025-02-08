import subprocess
import sys
import os
import time

def print_progress_bar(current, total, bar_length=40):
    """
    Prints a progress bar to the console with green-filled progress.
    The bar is updated in-place.
    """
    percent = float(current) / total
    filled_length = int(round(bar_length * percent))
    # ANSI escape for green is \033[92m, reset is \033[0m
    bar = "\033[92m" + "#" * filled_length + "\033[0m" + "-" * (bar_length - filled_length)
    sys.stdout.write(f"\rProgress: |{bar}| {percent*100:.0f}%")
    sys.stdout.flush()

def run_script(script_path):
    """
    Runs the given Python script using the current interpreter.
    If the script fails (non-zero exit code), the program exits.
    """
    print(f"\nRunning {script_path}...")
    result = subprocess.run([sys.executable, script_path], capture_output=True, text=True)
    print(result.stdout)
    if result.returncode != 0:
        print(f"Error running {script_path}:")
        print(result.stderr)
        sys.exit(result.returncode)
    else:
        print(f"Finished running {script_path}.\n")

def main():
    start_time = time.time()

    directory_url = os.path.join(os.getcwd(), "scrapers") + os.sep
    
    # List your scraper and conversion scripts in the order you want them to run.
    tasks = [
        directory_url + "uesp_banners_scraper.py",
        directory_url + "uesp_esoplus_scraper.py",
        directory_url + "uesp_literature_scraper.py",
        directory_url + "uesp_maps_scraper.py",
        directory_url + "uesp_music_boxes_scraper.py",
        directory_url + "uesp_paintings_scraper.py",
        directory_url + "uesp_tapestries_scraper.py",
        directory_url + "uesp_maps_scraper.py",
        "data_excel_to_lua.py"
    ]
    
    total_tasks = len(tasks)
    
    # Run each script sequentially.
    for i, script in enumerate(tasks, start=1):
        run_script(script)
        print_progress_bar(i, total_tasks)
        # Optional: pause a bit to let you see the progress bar update.
        time.sleep(0.5)
    
    end_time = time.time()
    total_run_time = end_time - start_time
    minutes, seconds = divmod(total_run_time, 60)
    print(f"\n\nAll scripts completed successfully in {int(minutes)} minutes {int(seconds)} seconds.")

if __name__ == "__main__":
    main()
