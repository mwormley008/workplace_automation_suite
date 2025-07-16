import json, re, os

def parse_ahk_script(file_path):
    # Dictionary to store company details
    companies = {}

    try:
        # Read the AHK script
        with open(file_path, 'r', encoding='utf-8', errors='replace') as file:
            script_contents = file.read()

            # Check if the script is being read correctly
            print(f"Script contents:\n{script_contents[:500]}...")  # print the first 500 characters for review

            # Regular expression to match the company details in the script
            # This regex assumes the company details are written in a very specific format.
            pattern = re.compile(r':\*:[\w\/]+::.*?{Delete}(.*?){.*?{Home}(.*?){.*?{Home}(.*?){', re.DOTALL)

            # Find all matches of the pattern
            matches = pattern.findall(script_contents)

            # Check if any matches are found
            if matches:
                print(f"Found {len(matches)} matches.")
                for match in matches:
                    company_name = match[0].strip()
                    address = match[1].strip() + ", " + match[2].strip()  # combining parts of the address
                    print(f"Extracted - {company_name}: {address}")  # print extracted details
                    companies[company_name] = address  # add to dictionary
            else:
                print("No matches found in the script.")

    except Exception as e:
        print(f"An error occurred: {e}")

    return companies

def save_to_json(data, file_path):
    with open(file_path, 'w', encoding='utf-8') as file:
        json.dump(data, file, ensure_ascii=False, indent=4)

def main():
    # specify the path to your AHK script
    ahk_script_path = r"C:\Users\Michael\Desktop\python-work\company_shortcuts.txt"
    json_output_path = r"company_shortcuts.json"  # replace with desired output path

    companies = parse_ahk_script(ahk_script_path)

    save_to_json(companies, json_output_path)


    print("Final extracted data:")
    for company, address in companies.items():
        print(f"{company}: {address}")

if __name__ == "__main__":
    main()