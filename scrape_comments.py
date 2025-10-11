from googleapiclient.discovery import build
import pandas as pd
from config import API_KEY

# Konfigurasi
VIDEO_ID = "-K9QTJZb1fE"  # Ganti jika perlu
MAX_COMMENTS = 2000

def scrape_comments(video_id, max_comments):
    youtube = build("youtube", "v3", developerKey=API_KEY)

    comments = []
    next_page = None

    while len(comments) < max_comments:
        request = youtube.commentThreads().list(
            part="snippet",
            videoId=video_id,
            maxResults=100,
            pageToken=next_page,
            order="relevance"
        )
        response = request.execute()

        for item in response["items"]:
            comment = item["snippet"]["topLevelComment"]["snippet"]
            comments.append({
                "author": comment.get("authorDisplayName"),
                "text": comment.get("textDisplay")
            })

            if len(comments) >= max_comments:
                break

        next_page = response.get("nextPageToken")
        if not next_page:
            break

    return comments

if __name__ == "__main__":
    print("Mengambil komentar...")
    data = scrape_comments(VIDEO_ID, MAX_COMMENTS)

    df = pd.DataFrame(data)
    df.to_excel("youtube_comments.xlsx", index=False)
    print("✅ Data berhasil disimpan ke youtube_comments.xlsx")
