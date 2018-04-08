import win32com.client
from nltk.tokenize import word_tokenize
from nltk.tokenize import sent_tokenize
from nltk.tokenize import RegexpTokenizer
from nltk import pos_tag
from nltk import FreqDist
from nltk import NaiveBayesClassifier
from nltk import classify
from nltk.corpus import movie_reviews
from nltk.classify.scikitlearn import SklearnClassifier
import random
import pickle
from MessageFinder import MessageFinder

# Started from previously titled 'EmailParse_v08'

# Sentiment analysis on customer feedback e-mails
## Phase 1:
### 1) Use Python to access Outlook folder with feedback e-mails
### 2) Figure out how to iterate over e-mails
### 3) Select only the Feedback portion of the e-mail
###### 4) Classify as positive, negative or neutral
### 5) Output numbers

## Phase 2:
### 1) Gather Date information
### 2) Gather regions informtion
### 3) Output numbers by week / region

## Phase 3:
### 1) Bucket feedback in a feature based on if exact text in list:
#### a) Spoiled / Rotten
#### b) Packaging
#### c) Thawed
#### d) Delivery
#### e) Selection
### 2) Output numbers
### 3) Store these numbers in Excel to build dashboard

## Phase 4:
### 1) Add trending words / themes by week
### 2) Output numbers
### 3) Update dashboard

## Phase 5:
### 1) Implement machine learning

messages = MessageFinder('carriaus@amazon.com', 'Inbox', 'AmazonFresh Feedback')

# Get the latest message
message = messages.GetLast()
body_content = message.body

# Tokenize the E-mail to find the final location of "Feedback:" to find the true body of the e-mail
tokenizer = RegexpTokenizer('\s+',gaps=True)
body_list_full = tokenizer.tokenize(body_content)

feedback_loc = 0
counter = 1
for word in body_list_full:
    if word == "Feedback:":
        feedback_loc = counter
        counter = counter + 1
    else:
        counter = counter + 1

body_list = body_list_full[feedback_loc:]

# Add POS tagging
body_tagged = pos_tag(body_list)

# Get all of the movie reviews in a word list pos/neg pair
documents = [(list(movie_reviews.words(fileid)), category)
              for category in movie_reviews.categories()
              for fileid in movie_reviews.fileids(category)]

random.shuffle(documents)

# Normalize the words to lowercase and get the frequency of each
all_words = []
for w in movie_reviews.words():
    all_words.append(w.lower())

all_words = FreqDist(all_words)

# Getting the 3000 most common words
word_features = list(all_words.keys())[:3000]

# Declaring a function that takes one argument, document, puts the words in a set (gets every word)
# then indicates true or false depending on if it can be found in word_features
def find_features(document):
    words = set(document)
    features = {}
    for w in word_features:
        features[w] = (w in words)

    return features

# For every review, in each category, call the return the result of the find_features function
# on the review and the category
featureset = [(find_features(rev), category) for (rev, category) in documents]


training_set = featureset[:1900]
testing_set = featureset[1900:]

classifier = NaiveBayesClassifier.train(training_set)


print("Naive Bayes Algorithm accuracy:", (classify.accuracy(classifier, testing_set))*100)
classifier.show_most_informative_features(15)



##print(body_tagged)

##print(body_list)







### Opening saved pickle classifier
##classifier_f = open("naivebayes_test.pickle", "rb")
##classifier = pickle.load(classifier_f)
##classifier.close()

### Creating new pickle classifier
##save_classifier = open("naivebayes_test.pickle", "wb")
##pickle.dump(classifier, save_classifier)
##save_classifier.close()





"""
from nltk.corpus import wordnet

syns = wordnet.synsets("program")

print(syns)

print(syns[0])

# just the word
print(syns[0].lemmas()[0].name())

# definition
print(syns[0].definition())

# examples
print(syns[0].examples())

# getting list of synonyms and antonyms for a word
synonyms = []
antonyms = []

for syn in wordnet.synsets("good"):
    for l in syn.lemmas():
        synonyms.append(l.name())
        if l.antonyms():
            antonyms.append(l.antonyms()[0].name())

print(set(synonyms))
print(set(antonyms))

# Check similarities
w1 = wordnet.synset('ship.n.01')
w2 = wordnet.synset('boat.n.01')
print(w1.wup_similarity(w2))

"""

"""

COULD BE USED FOR FINDING SPECIFIC THEMES

from nltk.stem import WordNetLemmatizer

lemmatizer = WordNetLemmatizer()

print(lemmatizer.lemmatize("cats"))
print(lemmatizer.lemmatize("better"))
print(lemmatizer.lemmatize("better", pos="a"))
print(lemmatizer.lemmatize("run"))
print(lemmatizer.lemmatize("run", "v"))

"""



##for message in messages:
##    message = messages.GetLast()
##    body_content = message.body
##    print(body_content)
##    counter = counter + 1
##    if counter == 10:
##        break


##print(msg.SenderName)
##print(msg.SenderEmailAddress)
##print(msg.SentOn)
##print(msg.To)
##print(msg.CC)
##print(msg.BCC)
##print(msg.Subject)
##print(msg.Body)

##body_paragraph = "".join([" "+i if not i.startswith("'") and i not in string.punctuation else i for i in body_list]).strip()

"""
WordPunctTokenizer: will seperate all punctuation
RegexpTokenizer: Is something you can use if you don't want split up contractions
    from nltk.tokenize import RegexpTokenizer
    tokenizer = RegexpTokenizer("[\w'+]")
    tokenizer.tokenize("can't is a contraction")
    ["can't", "is", "a", "contraction"]

    tokenizer = RegexpTokenizer('\s+',gaps=True)
    ["can't", "is", "a", "contraction"]

Frequency Distribution: Doesn't work: https://www.youtube.com/watch?v=zi16nl82AMA&list=PLQVvvaa0QuDf2JswnfiGkliBInZnIC4HL&index=11
    all_words = []
    for word in body_paragraph.words():
        all_words.append(w.lower())

    all_words = nltk.FreqDist(all_words)
"""
