import win32com.client
from nltk.tokenize import word_tokenize
from nltk.tokenize import sent_tokenize
from nltk.tokenize import RegexpTokenizer
from nltk import FreqDist
from nltk import NaiveBayesClassifier
from nltk import classify
from nltk.corpus import movie_reviews
import random
import pickle
from MessageFinder import MessageFinder

### Started from previously titled 'EmailParse_v08'

# Opening pickled classifier
classifier_f = open("20180408_Words-Movie_Reviews_Classifier.pickle", "rb")
classifier = pickle.load(classifier_f)
classifier_f.close()

# Opening pickled word features
word_features_f = open("20180408_Words-Movie_Reviews_Features.pickle", "rb")
word_features = pickle.load(word_features_f)
word_features_f.close()

# Find the 'features' or words in the email that match words we classified as positive or negative
def find_features(word_list):
    features = {}
    for w in word_features:
        features[w] = (w in word_list)
    return features

# Calls MessageFinder function from the MessageFinder script, returns a message object for specified folder
messages = MessageFinder('carriaus@amazon.com', 'Inbox', 'AmazonFresh Feedback')

count = 0

for message in messages:
    count = count + 1
    # Get the latest message
    message = messages.GetNext()
    body_content = message.body
    body_content = message

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

    feats = find_features(body_list)
    print(body_content)
    print(classifier.classify(feats))

    if count == 25:
        break

