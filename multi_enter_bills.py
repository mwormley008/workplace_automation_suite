import subprocess

def run_script(script_path):
    try:
        subprocess.run(["python", script_path], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error running script: {e}")
        # You can handle the error here if needed

if __name__ == "__main__":
    # Replace the paths with the actual paths to your scripts
    script1_path = r"C:\Users\Michael\Desktop\python-work\invoice_input.py"
    script2_path = r"C:\Users\Michael\Desktop\python-work\invoice_input_abc.py"
    script3_path = r"C:\Users\Michael\Desktop\python-work\bill_entry_and_payment.py"
    
    run_script(script1_path)
    run_script(script2_path)
    run_script(script3_path)

    print("All scripts have been executed.")
