import requests
import urllib.parse
import os
from bs4 import BeautifulSoup

# Folder to store downloaded images
img_folder = "wiki_images"
os.makedirs(img_folder, exist_ok=True)

# Load missed topics
with open("missed_questions.txt", "r", encoding="utf-8") as f:
    topics = [line.strip() for line in f if line.strip()]

for topic in topics:
    search_query = urllib.parse.quote(topic)
    url = f"https://en.wikipedia.org/wiki/{search_query}"
    print(f"Searching for: {topic} ({url})")

    try:
        response = requests.get(url)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")

        # Try to grab the first infobox image
        infobox = soup.find("table", class_="infobox")
        if infobox:
            img = infobox.find("img")
            if img and "src" in img.attrs:
                img_url = "https:" + img["src"]
                ext = os.path.splitext(img_url)[1].split("?")[0] or ".jpg"
                img_path = os.path.join(img_folder, f"{topic.replace(' ', '_')}{ext}")

                img_data = requests.get(img_url).content
                with open(img_path, "wb") as out:
                    out.write(img_data)
                print(f"✓ Saved image for '{topic}'")
            else:
                print(f"⚠️ No image found for '{topic}'")
        else:
            print(f"⚠️ No infobox for '{topic}'")
    except Exception as e:
        print(f"❌ Error fetching '{topic}': {e}")
