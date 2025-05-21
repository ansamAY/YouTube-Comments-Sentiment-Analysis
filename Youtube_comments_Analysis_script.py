import pandas as pd
import requests
import time

df = pd.read_excel("youtube-comments682d077ab7f43-BkryqKjptb0.xlsx")

comments_column = "Comment"

#add your key and endpoint values
api_key =""
endpoint = ""

sentiment_url = endpoint + "text/analytics/v3.1/sentiment"
keyphrase_url = endpoint + "text/analytics/v3.1/keyPhrases"

headers = {
    "Ocp-Apim-Subscription-Key": api_key,
    "Content-Type": "application/json"
}


sentiments = []
key_phrases_list = []

for i, comment in enumerate(df[comments_column]):
    if pd.isna(comment) or str(comment).strip() == "":
        sentiments.append("Empty")
        key_phrases_list.append("")
        continue

    doc = {
        "documents": [
            {
                "id": str(i),
                "language": "ar", 
                "text": str(comment)
            }
        ]
    }

    
    sent_resp = requests.post(sentiment_url, headers=headers, json=doc)
    if sent_resp.status_code == 200:
        sentiment = sent_resp.json()["documents"][0]["sentiment"]
    else:
        sentiment = "Error"

   
    key_resp = requests.post(keyphrase_url, headers=headers, json=doc)
    if key_resp.status_code == 200:
        key_phrases = key_resp.json()["documents"][0]["keyPhrases"]
        key_phrases = ", ".join(key_phrases)
    else:
        key_phrases = "Error"

    sentiments.append(sentiment)
    key_phrases_list.append(key_phrases)

    time.sleep(0.5) 


df["Sentiment"] = sentiments
df["Key Phrases"] = key_phrases_list

df.to_excel("analyzed_comments_with_keywords.xlsx", index=False)

print("Done! File saved as 'analyzed_comments_with_keywords.xlsx'")


