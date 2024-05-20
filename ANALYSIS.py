# -------------------------------------------
# WORD FREQUENCY TREND ANALYSIS
# -------------------------------------------

import os
import pandas as pd
from zipfile import ZipFile
from xml.etree.ElementTree import XML
from datetime import datetime
import re

def extract_text_between_keywords(docx_file):
    extracted_text = []

    # Open the .docx file as a zip archive
    with ZipFile(docx_file) as zipf:
        # Read the content of the "word/document.xml" file
        with zipf.open("word/document.xml") as xml_file:
            xml_content = xml_file.read()

    # Parse the XML content
    tree = XML(xml_content)

    # Flags to indicate whether to start/stop extracting text
    start_extraction = False

    # Find all text elements (w:t) in the XML tree
    for elem in tree.iter():
        if elem.tag.endswith('t'):
            text = elem.text
            # Check if the text contains the keywords "Body", "Load-date", and "End of document"
            if text:
                if "Body" in text:
                    # Start extracting text after the "Body" keyword
                    start_extraction = True
                elif "End of document" in text:
                    # Stop extracting text when reaching the "End of document" keyword
                    break
                elif start_extraction:
                    # Split words with capital letters in the middle
                    words = re.findall(r'[A-Z][^A-Z]*', text)
                    # Join the split words with a space
                    extracted_text.extend(words)
                elif "Load-Date" in text:
                    # Stop extracting text before the "Load-date" keyword
                    start_extraction = False

    return ' '.join(extracted_text)

def extract_date_between_keywords(docx_file):
    # Open the .docx file as a zip archive
    with ZipFile(docx_file) as zipf:
        # Read the content of the "word/document.xml" file
        with zipf.open("word/document.xml") as xml_file:
            xml_content = xml_file.read()

    # Parse the XML content
    tree = XML(xml_content)

    # Flag to indicate whether to start extracting date text
    start_extraction = False
    date_text = ""

    # Find all text elements (w:t) in the XML tree
    for elem in tree.iter():
        if elem.tag.endswith('t'):
            text = elem.text
            if start_extraction and text:
                # Try to extract date pattern from text
                match = re.search(r"\b(?:\d{1,2}\s+\w+\s+\d{4}|\w+\s+\d{1,2},?\s+\d{4}|\d{1,2}/\d{1,2}/\d{4})\b", text)
                if match:
                    date_text = match.group()
                    break
            # Check if the text contains the "Load-date" keyword
            if text and "Load-Date" in text:
                start_extraction = True
            # Stop extracting text when reaching the "End of document" keyword
            elif text and "End of document" in text:
                break

    return date_text

def parse_date(date_text):
    try:
        # Parse the date text using datetime.strptime
        date_obj = datetime.strptime(date_text, "%B %d, %Y")
        month = date_obj.strftime("%B")
        year = date_obj.strftime("%Y")
        return month, year
    except ValueError:
        # Handle cases where date parsing fails
        return None, None

def extract_month_year(date_text):
    try:
        # Parse the date text using datetime.strptime
        date_obj = datetime.strptime(date_text, "%B %d, %Y")
        month_year = date_obj.strftime("%B %Y")
        return month_year
    except ValueError:
        # Handle cases where date parsing fails
        return None

def process_docx_files(folder_path):
    data = []
    for filename in os.listdir(folder_path):
        if filename.endswith(".docx"):
            docx_file = os.path.join(folder_path, filename)
            extracted_text = extract_text_between_keywords(docx_file)
            extracted_date = extract_date_between_keywords(docx_file)
            month, year = parse_date(extracted_date)
            month_year = extract_month_year(extracted_date)  # Extract month and year together
            data.append({'Filename': filename, 'Extracted_Text': extracted_text, 'Date': extracted_date, 'Month': month, 'Year': year, 'Month_Year': month_year})  # Include month_year in the data
    return pd.DataFrame(data)

# Parsing text into dataframes by month and concatenating them all to one large dataframes
folder_path1 = ".../07-07-2023 to 06-08-2023"
folder_path2 = ".../07-08-2023 to 06-09-2023"
folder_path3 = ".../07-09-2023 to 06-10-2023"
folder_path4 = ".../07-10-2023 to 07-11-2023"
folder_path5 = ".../08-11-2023 to 07-12 2023"
folder_path6 = ".../08-12-2023 to 07-01-2024"
folder_path7 = ".../08-01-2024 to 07-02-2024"
folder_path8 = ".../08-02-2024 to 07-03-2024"
folder_path9 = ".../08-03-2024 to 07-04-2024"
folder_path10 = ".../08-04-2024 to 20-04-2024"

jul23 = process_docx_files(folder_path1)
aug23 = process_docx_files(folder_path2)
sep23 = process_docx_files(folder_path3)
oct23 = process_docx_files(folder_path4)
nov23 = process_docx_files(folder_path5)
dec23 = process_docx_files(folder_path6)
jan24 = process_docx_files(folder_path7)
feb24 = process_docx_files(folder_path8)
mar24 = process_docx_files(folder_path9)
apr24 = process_docx_files(folder_path10)

dfs = [jul23, aug23, sep23, oct23, nov23, dec23, jan24, feb24, mar24]
data = pd.concat(dfs, ignore_index=True)

# Print the combined DataFrame
print(data)

# -------------------------------------------
# PREPROCESSING
# -------------------------------------------

# Cleaning text to remove superfluous phrases, special characters, punctuation.
def clean_text(text):
    # Remove "Speech to text transcript:"
    text = re.sub(r'Speech to text transcript:', '', text)
    # Remove strange pattern of special characters that used to indicate formatting
    text = re.sub(r'&#\d+;', '', text)
    # Remove URLs
    text = re.sub(r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+', '', text)
    # Remove special characters and numbers
    text = re.sub(r'[^A-Za-z\s]', '', text)
    # Remove additional white spaces
    text = re.sub(r'\s+', ' ', text).strip()
    return text

# Case normalisation
def case_normalization(text):
    return text.lower()

# Tokenisation
def tokenization(text):
    return text.split()

# Remove stopwords
stopwordlist = set(['i','me','my','myself','we','our','ours','ourselves','you','your','yours','yourself','yourselves','he','him','his','himself','she','her','hers','herself','it','its','itself','they','them','their','theirs','themselves','what','which','who','whom','this','that','these','those','am','is','are','was','were','be','been','being','have','has','had','having','do','does','did','doing','a','an','the','and','but','if','or','because','as','until','while','of','at','by','for','with','about','against','between','into','through','during','before','after','above','below','to','from','up','down','in','out','on','off','over','under','again','further','then','once','here','there','when','where','why','how','all','any','both','each','few','more','most','other','some','such','no','nor','not','only','own','same','so','than','too','very','s','t','can','will','just','don','should','now'])
def remove_stopwords(tokens):
    return [token for token in tokens if token not in stopwordlist]

# Lemmatisation
import nltk
from nltk.stem import PorterStemmer
from nltk.stem import WordNetLemmatizer
from nltk.tokenize import word_tokenize

# Initialize lemmatizer
lemmatizer = WordNetLemmatizer()

def lemmatize(tokens):
    # Perform lemmatization; returns shortened versions of words.
    return [lemmatizer.lemmatize(token) for token in tokens]

# Preprocessing: Apply all functions
data['Preprocessed'] = data['Extracted_Text'].apply(lambda x: lemmatize(remove_stopwords(tokenization(case_normalization(clean_text(x))))))
print(data)

# ---------------------------------------------------------------------
# PROPORTIONS OF 'TABOO' VOCABULARY OVER TOTAL WORD COUNT OVER TIME (%)
# ---------------------------------------------------------------------
from collections import Counter
import matplotlib.pyplot as plt

# List of 'taboo' words
word_list = ['palestine','genocide', 'ethnic cleansing', 'occupied territory', 'refugee camp', 'zionist', 'zionism']

# Function to count frequencies of specific words in a document and calculate proportions
def count_specific_words(document):
    total_words = len(document)
    
    # Check if total_words is not zero to avoid division by zero error
    if total_words != 0:
        word_counts = Counter(document)
        word_proportions = {word: word_counts[word] / total_words for word in word_list}
    else:
        # If total_words is zero, return a dictionary of zeros
        word_proportions = {word: 0 for word in word_list}
    
    return word_proportions

# Count frequencies of specific words in each document and calculate proportions
word_proportions_df = data['Preprocessed'].apply(count_specific_words).apply(pd.Series)

# Concatenate word proportions with original DataFrame
data_with_word_proportions = pd.concat([data, word_proportions_df], axis=1)

# Extract month and year from the 'Date' column and combine them into 'Month_Year'
data_with_word_proportions['Month_Year'] = pd.to_datetime(data_with_word_proportions['Date']).dt.to_period('M')

# Aggregate proportion counts by month and year
word_proportions_over_time = data_with_word_proportions.groupby('Month_Year')[word_list].mean()

# Plot the changes in proportions over time for each specific word
plt.figure(figsize=(10, 6))
for word in word_list:
    plt.plot(word_proportions_over_time.index.astype(str), word_proportions_over_time[word], label=word)
plt.xlabel('Month-Year')
plt.ylabel('Proportion')
plt.title('Average Proportions of Taboo Words Used over Total Word Count in A Document')
plt.legend()
plt.xticks(rotation=45)
plt.show()

import pandas as pd
import matplotlib.pyplot as plt

# ------------------------------------------------------
# ARTICLES CONTAINING TABOO VOCABULARY OVER TIME (%)
# ------------------------------------------------------
# List of 'taboo' words
word_list = ['palestine', 'genocide', 'ethnic cleansing', 'occupied territory', 'refugee camp', 'zionist', 'zionism']

# Initialize a dictionary to store the proportions for each word
word_presence_over_time = {}

# Extract month and year from the 'Date' column and combine them into 'Month_Year'
data_with_word_proportions['Month_Year'] = pd.to_datetime(data_with_word_proportions['Date']).dt.to_period('M')

# Iterate over each word in word_list
for word in word_list:
    # Filter the DataFrame to include only articles containing the current word in 'Preprocessed' text
    data_with_word_proportions[word] = data_with_word_proportions['Preprocessed'].apply(lambda x: 1 if word in x else 0)

    # Group by Month_Year and calculate the proportion of articles containing the current word
    proportions = data_with_word_proportions.groupby('Month_Year')[word].mean()

    # Store the proportions in the dictionary
    word_presence_over_time[word] = proportions

# Plot the proportions for each word
plt.figure(figsize=(10, 6))
for word, proportions in word_presence_over_time.items():
    # Ensure the proportions DataFrame is sorted by the categorical order of 'Month_Year' before plotting
    proportions = proportions.sort_index()
    plt.plot(proportions.index.to_timestamp(), proportions.values, label=word)

plt.xlabel('Month, Year')
plt.ylabel('Proportion of Documents')
plt.title('Proportions of Documents Containing Taboo Words per Month-Year')
plt.legend()
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()


#---------------------------------
# TOTAL COUNT OF RELEVANT ARTICLES
#---------------------------------
article_counts_over_time = data_with_word_proportions.groupby('Month_Year').size()

# Plot the total number of articles over time
plt.figure(figsize=(10, 6))
plt.plot(article_counts_over_time.index.astype(str), article_counts_over_time.values, label='Total Articles', color='blue', linestyle='-')
plt.xlabel('Month-Year')
plt.ylabel('Number of Articles')
plt.title('Total Number of Relevant Articles over Time')
plt.legend()
plt.xticks(rotation=45)
plt.show()


# ---------------------
# TOPIC MODELLING (LDA)
# ---------------------
import pandas as pd
import gensim
from gensim import corpora, models
import scipy # Needs scipy v1.12.0 and not v1.13.0

# Create a dictionary and corpus needed for Topic Modeling
dictionary = corpora.Dictionary(data['Preprocessed']) 
corpus = [dictionary.doc2bow(text) for text in data['Preprocessed']] 

# LDA model
lda_model = gensim.models.ldamodel.LdaModel(corpus=corpus,
                                           id2word=dictionary,
                                           num_topics=10, 
                                           random_state=100,
                                           update_every=1,
                                           chunksize=100,
                                           passes=10,
                                           alpha='auto',
                                           per_word_topics=True)

# Print the topics found by the LDA model
topics = lda_model.print_topics(num_words=10)
for topic in topics:
    print(topic)

# Clearer format
for topic_num, topic in topics:
    # Parse the topic output to show only words
    print("Topic #{}:".format(topic_num + 1) + " " + "".join([word.split("*")[1].replace('"', '') for word in topic.split("+")]))

# Most illustrative documents of relevant topics (#3, #4, #6, #8, #9)

# ------------------------------------------------------------
# TOPIC 3 - israel said state hamas u m russian news also iran
# ------------------------------------------------------------

# Analyze topic distribution for each document
doc_topics = [lda_model.get_document_topics(item) for item in corpus]

# Determine the top document for Topic 3
topic_id = 2  # Topics are zero-indexed, so Topic 4 is ID 3
top_doc_for_topic_3, max_contribution = None, 0

for i, doc_distribution in enumerate(doc_topics):
    for topic_num, contrib in doc_distribution:
        if topic_num == topic_id and contrib > max_contribution:
            max_contribution = contrib
            top_doc_for_topic_3 = i

# Print results
if top_doc_for_topic_3 is not None:
    print(f"Top document for Topic 3 is Document #{top_doc_for_topic_3 + 1} with a contribution of {max_contribution:.4f}:")
    print(data.iloc[top_doc_for_topic_3]['Filename'])
else:
    print("No document has Topic 3 as the predominant topic.")

# --------------------------------------------------------------------
# TOPIC 4 - israel gaza hamas israeli u hostage say ceasefire b people
# --------------------------------------------------------------------

# Analyze topic distribution for each document
doc_topics = [lda_model.get_document_topics(item) for item in corpus]

# Determine the top document for Topic 3
topic_id = 3  # Topics are zero-indexed, so Topic 4 is ID 3
top_doc_for_topic_4, max_contribution = None, 0

for i, doc_distribution in enumerate(doc_topics):
    for topic_num, contrib in doc_distribution:
        if topic_num == topic_id and contrib > max_contribution:
            max_contribution = contrib
            top_doc_for_topic_4 = i

# Print results
if top_doc_for_topic_3 is not None:
    print(f"Top document for Topic 4 is Document #{top_doc_for_topic_4 + 1} with a contribution of {max_contribution:.4f}:")
    print(data.iloc[top_doc_for_topic_4]['Filename'])
else:
    print("No document has Topic 4 as the predominant topic.")

# --------------------------------------------------------------------
# TOPIC 6 - gaza israel u n hamas b people say hospital israeli
# --------------------------------------------------------------------

# Analyze topic distribution for each document
doc_topics = [lda_model.get_document_topics(item) for item in corpus]

# Determine the top document for Topic 3
topic_id = 5  # Topics are zero-indexed, so Topic 4 is ID 3
top_doc_for_topic_6, max_contribution = None, 0

for i, doc_distribution in enumerate(doc_topics):
    for topic_num, contrib in doc_distribution:
        if topic_num == topic_id and contrib > max_contribution:
            max_contribution = contrib
            top_doc_for_topic_6 = i

# Print results
if top_doc_for_topic_6 is not None:
    print(f"Top document for Topic 6 is Document #{top_doc_for_topic_6 + 1} with a contribution of {max_contribution:.4f}:")
    print(data.iloc[top_doc_for_topic_6]['Filename'])
else:
    print("No document has Topic 6 as the predominant topic.")

# -----------------------------------------------------------------------------
# TOPIC 8 - israel palestinian hamas u said b march resolution israeli security
# -----------------------------------------------------------------------------

# Analyze topic distribution for each document
doc_topics = [lda_model.get_document_topics(item) for item in corpus]

# Determine the top document for Topic 3
topic_id = 7  # Topics are zero-indexed, so Topic 4 is ID 3
top_doc_for_topic_8, max_contribution = None, 0

for i, doc_distribution in enumerate(doc_topics):
    for topic_num, contrib in doc_distribution:
        if topic_num == topic_id and contrib > max_contribution:
            max_contribution = contrib
            top_doc_for_topic_8 = i

# Print results
if top_doc_for_topic_8 is not None:
    print(f"Top document for Topic 8 is Document #{top_doc_for_topic_8 + 1} with a contribution of {max_contribution:.4f}:")
    print(data.iloc[top_doc_for_topic_8]['Filename'])
else:
    print("No document has Topic 8 as the predominant topic.")

# -----------------------------------------------------------------------------
# TOPIC 9 - al gaza israeli hamas said israel also medium palestinian u
# -----------------------------------------------------------------------------

# Analyze topic distribution for each document
doc_topics = [lda_model.get_document_topics(item) for item in corpus]

# Determine the top document for Topic 3
topic_id = 8  # Topics are zero-indexed, so Topic 4 is ID 3
top_doc_for_topic_9, max_contribution = None, 0

for i, doc_distribution in enumerate(doc_topics):
    for topic_num, contrib in doc_distribution:
        if topic_num == topic_id and contrib > max_contribution:
            max_contribution = contrib
            top_doc_for_topic_9 = i

# Print results
if top_doc_for_topic_9 is not None:
    print(f"Top document for Topic 9 is Document #{top_doc_for_topic_9 + 1} with a contribution of {max_contribution:.4f}:")
    print(data.iloc[top_doc_for_topic_9]['Filename'])
else:
    print("No document has Topic 9 as the predominant topic.")

# ---------------------------
# TOPIC MODELLING - BERTOPIC
# ---------------------------
import pandas as pd
import warnings
from IPython.display import display
warnings.filterwarnings("ignore")

from bertopic import BERTopic

def analyze_topics(texts):

    model = BERTopic(verbose=True,embedding_model='paraphrase-MiniLM-L3-v2', min_topic_size= 7)
    topics, _ = model.fit_transform(texts)
    
    freq = model.get_topic_info()
    print("Number of topics: {}".format( len(freq)))
    display(freq.head(20))
    return model,topics,freq

# Convert each list of words into a single string
data['Preprocessed_Bertopic'] = data['Preprocessed'].apply(lambda x: ' '.join(x))

# Discover topics
model, topics, freq = analyze_topics(data['Preprocessed_Bertopic'])
data['Topic'] = topics

# Print out topics
for topic_id, topic_info in model.get_topic_info().iterrows():
    print(f"Topic {topic_id}:")
    print(f"  Count: {topic_info['Count']}")
    print(f"  Name: {topic_info['Name']}")
    print(f"  Representation: {topic_info['Representation']}")
    print("  Representative Document:")
    if topic_info['Representative_Docs']:
        print(f"    {topic_info['Representative_Docs'][0]}")
    else:
        print("    No representative document found.")
    print()

# Print topics (shorter)
for topic_id, topic_info in model.get_topic_info().iterrows():
    print(f"Topic {topic_id}:")
    print(f"  Count: {topic_info['Count']}")
    print(f"  Name: {topic_info['Name']}")
    print(f"  Representation: {topic_info['Representation']}")
    print()

# Visualisation
import plotly.io as pio
import plotly.express as px
pio.renderers.default = "browser"
import plotly.graph_objects as go

bertopic_visualize = model.visualize_topics()
fig1 = go.Figure(bertopic_visualize)
fig1.show()

bertopic_visualize_bars = model.visualize_barchart()
fig2 = go.Figure(bertopic_visualize_bars)
fig2.show()

# Finding each corresponding filename of a representative document for first 10 (largest) topics

# Specify the specific value for data['Preprocessed_Bertopic']
specific_value = "u n r know staff involved accused helping hamas hamas well u n secretary general antonio guterres say immediately dismissed one died two currently identified little concrete known actual allegation heard report israeli intelligence israeli military said passed information say show active participation people th october attack use u n r o d facility vehicle mark regev senior adviser israel prime minister benjamin netanyahu told b b c last week released israeli hostage said held home u n w r e member timing allegation debated mean see israel israel friday international court justice u n top court need prevent genocide gaza israel israeli government say actually un released news time try bury c j ruling headline either way obviously serious serious allegation treated voice support u n say look mission employ people total handful bad apple seen tarnish entire reputation important work mission b b c mark lowen meanwhile diplomatic effort free hostage still held hamas c expected meet official israel egypt qatar undisclosed location france coming day expected work towards securing release hostage held gaza western government consider hamas spoke garcia baskin lead negotiator release israeli soldier gilad shalit held hostage hamas he also middle east director international community organisation asked latest round talk could help reach agreement think team fact talk taking place highest level possible highest level people intelligence community good sign think need move away background noise report kind agreement another kind understand thats part negotiating process part psychological warfare reality hostage israeli hostage gaza believed le alive every day god risk life israeli mounting military campaign south gaza rafah border crossing rapid city rafah last place attacked israel hamas want ceasefire u also one significant number palestinian prisoner released israel difficult negotiation gap quite large difficult negotiation somebody taken part similar negotiation hamas would talking would would strategy involved try bring hostage release firstly recognise israel hamas egyptian intelligence qatari government excited still bring message red line willing accept willing accept mass end war israeli withdrawal israel palestinian prisoner released brazil youre want want get back hostage willing read palestinian president significance somehow medias head c attraction qatari find middle ground make work side time understand first point view even hostage released virtually nothing prevents moving war one everyones re load date january end document"

# Filter the DataFrame based on the specific value
matching_rows = data[data['Preprocessed_Bertopic'] == specific_value]

# Access the 'Filename' column for the matching rows
filenames = matching_rows['Filename']

# Set max_colwidth option to None to display full filenames
pd.set_option('display.max_colwidth', None)

# Print the filenames
print(filenames)