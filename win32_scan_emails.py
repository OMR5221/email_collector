import win32com.client
import nltk
from nltk.sentiment.vader import SentimentIntensityAnalyzer


def print_score(scores):
    for k in sorted(scores):
        print('{0}: {1}, '.format(k, scores[k]), end="\n")


def get_scores(sia, message_text):
    return sia.polarity_scores(message_text)




session = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = session.GetDefaultFolder(6)
print(len(inbox.Items))
message = inbox.Items

body_sender = ''

sia = SentimentIntensityAnalyzer()

for msg in message:
    try:
        body_sender = msg.Sender
    except:
        body_sender = ''

    print(body_sender, msg.Subject, msg.Body)
    scores = get_scores(sia, msg.Body)
    print_score(scores)
