import pandas as pd
import re
import emoji
import nltk
from nltk.corpus import stopwords, wordnet
from nltk.stem import WordNetLemmatizer
from wordcloud import WordCloud
import matplotlib.pyplot as plt
from sklearn.feature_extraction.text import CountVectorizer

# Download resources jika belum
nltk.download('stopwords')
nltk.download('wordnet')
nltk.download('omw-1.4')

# Stopwords Inggris
stop_words = set(stopwords.words('english'))

# Tambahan stopwords khusus (nama artis & kata umum)
custom_stopwords = {"alex", "rex", "orange", "county", "music", "song", "songs", "listen", "listening", "br", "oh", "39", "com", "www", "http", "https", "like", "just", "get", "got", "im", "ive", "u", "ur", "dont", "cant", "wanna", "gonna", "k9qtjzb1fe", "youtube", "video", "videos", "subscribe", "channel", "new", "one", "make", "time", "see", "good", "know", "really", "people", "think", "way", "want", "much", "still", "take", "back", "come", "first", "last", "never", "p", "amp", "oh", "href"}
stop_words = stop_words.union(custom_stopwords)

lemmatizer = WordNetLemmatizer()

def clean_text(text):
    if not isinstance(text, str):
        return ""

    # 1. lowercase
    text = text.lower()

    # 2. hapus mention
    text = re.sub(r'@\w+', '', text)

    # 3. hapus emoji
    text = emoji.replace_emoji(text, replace=' ')

    # 4. hapus simbol & karakter non-alfanumerik
    text = re.sub(r'[^a-zA-Z0-9\s]', ' ', text)

    # 5. hapus stopwords & lemmatization
    words = [
        lemmatizer.lemmatize(word)
        for word in text.split()
        if word not in stop_words
    ]
    text = ' '.join(words)

    # 6. hapus spasi ganda
    text = re.sub(r'\s+', ' ', text).strip()

    return text

def generate_top_ngrams(text_series, n, top_n=20):
    """
    Mengambil top N n-gram dari kolom teks bersih.
    """
    vectorizer = CountVectorizer(ngram_range=(n, n))
    X = vectorizer.fit_transform(text_series)
    freqs = zip(vectorizer.get_feature_names_out(), X.sum(axis=0).tolist()[0])
    sorted_freqs = sorted(freqs, key=lambda x: x[1], reverse=True)
    return sorted_freqs[:top_n]

def create_wordcloud_from_phrases(phrases, output_filename):
    """
    Membuat wordcloud dari daftar n-gram.
    """
    if not phrases:
        print(f"Tidak ada data untuk {output_filename}")
        return

    text_data = ' '.join(['_'.join(p[0].split()) for p in phrases])

    wordcloud = WordCloud(width=1000, height=600, background_color='white').generate(text_data)
    plt.figure(figsize=(10, 6))
    plt.imshow(wordcloud, interpolation='bilinear')
    plt.axis("off")
    plt.tight_layout()
    plt.savefig(output_filename)
    print(f"✅ WordCloud disimpan sebagai {output_filename}")

def preprocess_excel(input_file, output_file):
    df = pd.read_excel(input_file)

    if 'text' not in df.columns:
        raise ValueError("Kolom 'text' tidak ditemukan dalam file Excel!")

    # Bersihkan teks
    df['cleaned_text'] = df['text'].apply(clean_text)

    # Hapus duplikat
    df.drop_duplicates(subset=['cleaned_text'], inplace=True)

    # Simpan ke Excel
    df.to_excel(output_file, index=False)
    print(f"✅ Preprocessing selesai! Disimpan ke {output_file}")

    # Ambil teks bersih
    cleaned_series = df['cleaned_text'].dropna()

    # 🔹 Tampilkan Bigram
    print("\n🔹 Top 20 Bigram:")
    bigrams = generate_top_ngrams(cleaned_series, 2, top_n=20)
    for phrase, count in bigrams:
        print(f"{phrase}: {count}")

    # 🔹 Tampilkan Trigram
    print("\n🔹 Top 20 Trigram:")
    trigrams = generate_top_ngrams(cleaned_series, 3, top_n=20)
    for phrase, count in trigrams:
        print(f"{phrase}: {count}")

    # 🔹 Tampilkan WordCloud utama (unigram)
    all_text = ' '.join(cleaned_series)
    main_wc = WordCloud(width=1000, height=600, background_color='white').generate(all_text)
    plt.figure(figsize=(12, 7))
    plt.imshow(main_wc, interpolation='bilinear')
    plt.axis("off")
    plt.title("WordCloud Komentar (Unigram)")
    plt.show()

    # 🔹 Simpan WordCloud Bigram & Trigram
    create_wordcloud_from_phrases(bigrams, "bigram_wordcloud.png")
    create_wordcloud_from_phrases(trigrams, "trigram_wordcloud.png")

if __name__ == "__main__":
    preprocess_excel("youtube_comments.xlsx", "youtube_comments_cleaned.xlsx")
