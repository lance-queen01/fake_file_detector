import pandas as pd
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from wordcloud import WordCloud
import matplotlib.pyplot as plt

def analyze_and_visualize_wordcloud(file_path, categories):
    # Load the dataset
    df = pd.read_excel(file_path)

    # Function to preprocess text data
    def preprocess_text(text):
        # Tokenize the text
        tokens = word_tokenize(text)

        # Remove stop words and non-alphabetic tokens
        stop_words = set(stopwords.words('english'))
        tokens = [word.lower() for word in tokens if word.isalpha() and word.lower() not in stop_words]

        return tokens

    # Analyze and visualize words for each category
    for category in categories:
        # Filter rows with non-"no violation" reviews in the specified category
        relevant_rows = df[df[category].apply(lambda x: str(x).lower() != "no violation")]

        if not relevant_rows.empty:
            # Combine the Header and Body for the specified category
            category_text = ' '.join(relevant_rows['Body'])

            # Preprocess the text
            tokens = preprocess_text(category_text)

            if tokens:
                # Join the tokens into a single string for word cloud generation
                text_for_wordcloud = ' '.join(tokens)

                # Generate word cloud
                wordcloud = WordCloud(width=800, height=400, background_color='white').generate(text_for_wordcloud)

                # Plot the word cloud
                plt.figure(figsize=(10, 5))
                plt.imshow(wordcloud, interpolation='bilinear')
                plt.title(f'Word Cloud for {category}')
                plt.axis('off')
                plt.show()
            else:
                print(f'No meaningful tokens found for {category}.')
        else:
            print(f'No reviews found for {category}.')


