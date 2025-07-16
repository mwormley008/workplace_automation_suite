import os
import csv
from bs4 import BeautifulSoup

def extract_cards_from_html(filepath):
    with open(filepath, "r", encoding="utf-8") as file:
        soup = BeautifulSoup(file, "html.parser")

    cards = []
    round_tables = soup.select("table.round")

    for round_table in round_tables:
        rows = round_table.find_all("tr")
        categories = []

        # First row = categories
        category_cells = rows[0].find_all("td", class_="category")
        for cell in category_cells:
            cat_name = cell.find("td", class_="category_name")
            categories.append(cat_name.text.strip() if cat_name else "Unknown")

        # Loop through clues
        for row in rows[1:]:
            clues = row.find_all("td", class_="clue")
            for i, clue in enumerate(clues):
                q_text = clue.find("td", class_="clue_text")
                if not q_text or not q_text.get("id"):
                    continue

                q_id = q_text["id"]
                a_id = q_id + "_r"
                a_cell = soup.find("td", {"id": a_id})
                a_text = a_cell.find("em", class_="correct_response") if a_cell else None

                if a_text:
                    category = categories[i] if i < len(categories) else "Unknown"
                    question = q_text.get_text(separator=" ", strip=True)
                    answer = a_text.get_text(separator=" ", strip=True)
                    cards.append([category, question, answer])
    
    return cards


# Target CSV file
csv_filename = "jeopardy_anki_with_categories.csv"
file_exists = os.path.exists(csv_filename)

# Open once, append everything
with open(csv_filename, "a", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    if not file_exists:
        writer.writerow(["Category", "Question", "Answer"])  # Write header if needed

    total_new_cards = 0

    # Loop through all .html files
    for filename in os.listdir():
        if filename.endswith(".html"):
            print(f"Processing {filename}...")
            cards = extract_cards_from_html(filename)
            writer.writerows(cards)
            total_new_cards += len(cards)

print(f"Appended {total_new_cards} total cards to '{csv_filename}'")
