import pandas as pd
import re
import emoji
import nltk
from nltk.corpus import stopwords

# Download stopwords jika belum pernah
try:
    nltk.data.find('corpora/stopwords')
except nltk.downloader.DownloadError:
    nltk.download('stopwords')

# Bahasa Inggris dan Indonesia
stop_words_en = set(stopwords.words('english'))
stop_words_id = set(stopwords.words('indonesian'))

def clean_text(text):
    if not isinstance(text, str):
        return ""

    # 1. Lowercase
    text = text.lower()

    # 2. Hapus mention seperti @username
    text = re.sub(r'@\w+', '', text)

    # 3. Hapus emoji
    # Menghapus emoji
    text = emoji.replace_emoji(text, replace=' ')

    # 4. Hapus simbol unik / karakter non-alfanumerik
    # Hanya menyisakan huruf, angka, dan spasi
    text = re.sub(r'[^a-zA-Z0-9\s]', '', text)

    # 5. Hapus stopwords
    words = text.split()
    words = [w for w in words if w not in stop_words_en and w not in stop_words_id]
    text = ' '.join(words)

    # 6. Menghapus spasi ganda
    text = re.sub(r'\s+', ' ', text).strip()

    return text

def preprocess_excel(input_file, output_file):
    try:
        df = pd.read_excel(input_file)
    except FileNotFoundError:
        print(f"❌ Error: File '{input_file}' tidak ditemukan!")
        return

    if 'text' not in df.columns:
        raise ValueError("Kolom 'text' tidak ditemukan dalam file Excel! Pastikan kolom yang berisi teks komentar bernama 'text'.")

    # Terapkan fungsi cleaning
    df['cleaned_text'] = df['text'].apply(clean_text)
    
    # Simpan hasilnya
    df.to_excel(output_file, index=False)
    print(f"✅ Preprocessing selesai! Disimpan ke {output_file}")

if __name__ == "__main__":
    # Ganti 'youtube_comments.xlsx' dengan nama file input Anda jika berbeda
    # Pastikan file tersebut ada di direktori yang sama
    preprocess_excel("youtube_comments.xlsx", "youtube_comments_cleaned.xlsx")