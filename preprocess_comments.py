import pandas as pd
import re
import emoji
from googletrans import Translator
import nltk
from nltk.corpus import stopwords

# Download stopwords jika belum pernah
nltk.download('stopwords')

# Bahasa Inggris dan Indonesia
stop_words_en = set(stopwords.words('english'))
stop_words_id = set(stopwords.words('indonesian'))

translator = Translator()

def clean_text(text):
    if not isinstance(text, str):
        return ""

    # 1. Lowercase
    text = text.lower()

    # 2. Hapus mention seperti @username
    text = re.sub(r'@\w+', '', text)

    # 3. Hapus emoji
    text = emoji.replace_emoji(text, replace='')

    # 4. Hapus simbol unik / karakter non-alfanumerik
    text = re.sub(r'[^a-zA-Z0-9\s]', '', text)

    # 5. Hapus stopwords
    words = text.split()
    words = [w for w in words if w not in stop_words_en and w not in stop_words_id]
    text = ' '.join(words)

    # 6. Translate ke bahasa Indonesia
    try:
        translated = translator.translate(text, dest='id')
        text = translated.text
    except:
        pass  # jika gagal translate, biarkan teks aslinya

    return text

def preprocess_excel(input_file, output_file):
    df = pd.read_excel(input_file)

    if 'text' not in df.columns:
        raise ValueError("Kolom 'text' tidak ditemukan dalam file Excel!")

    df['cleaned_text'] = df['text'].apply(clean_text)
    df.to_excel(output_file, index=False)
    print(f"✅ Preprocessing selesai! Disimpan ke {output_file}")

if __name__ == "__main__":
    preprocess_excel("youtube_comments.xlsx", "youtube_comments_cleaned.xlsx")
