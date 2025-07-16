import webbrowser
import time
import urllib.parse

# Load missed answers from a text file (1 per line)
with open("missed_questions.txt", "r", encoding="utf-8") as f:
    topics = [line.strip() for line in f if line.strip()]

# Open Wikipedia pages
for topic in topics:
    query = urllib.parse.quote(topic)
    url = f"https://en.wikipedia.org/wiki/{query}"
    print(f"Opening: {url}")
    webbrowser.open_new_tab(url)
    time.sleep(1)  # Delay to avoid overwhelming your browser
