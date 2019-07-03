
import re 
from textblob import TextBlob 
from dbfread import DBF
import csv
import xlsxwriter

class TwitterClient(object): 

    def dbf_to_xlsx(self, fileName):
        i=0
        fetched_tweets = DBF(fileName, encoding='latin-1')
        xlsxFile = fileName[:-4]+'.xlsx'
        xbook = xlsxwriter.Workbook(xlsxFile )
        xsheet = xbook.add_worksheet('Tweets')
        bold = xbook.add_format({'bold': True})
        b=fetched_tweets.field_names

        xsheet.write_row(0,0,b, bold)
        row=1
        try:
            for tweet in fetched_tweets:
                a=[]
                a.append(tweet['Airline'])
                a.append(tweet['UserName'])
                a.append(tweet['FormalName'])
                a.append(str(tweet['TimeStamp']))
                a.append(tweet['Text'])
                xsheet.write_row(row,0 , a)
                
                i=i+1
                row=row+1
        except AttributeError:
            # counters is not a dictionary, ignore and move on
            pass
        xbook.close()

    def save_to_xlsx(self, fileName, data, headers):
        i=0
        xbook = xlsxwriter.Workbook(str(fileName+'.xlsx'))
        xsheet = xbook.add_worksheet('Tweets')
        bold = xbook.add_format({'bold': True})
        print('save to xlsx')
        

        xsheet.write_row(0,0,headers, bold)
        row=1
        for tweet in data:
            a=[]
            
            a.append(tweet['Airline'])
            a.append(tweet['UserName'])
            a.append(tweet['FormalName'])
            a.append(str(tweet['TimeStamp']))
            a.append(tweet['Text'])
            a.append(tweet['SentimentPolarity'])
            a.append(tweet['Sentiment'])
           
            xsheet.write_row(row,0 , a)
            i=i+1
            row=row+1
        print('saved tweets with sentiments = ' , i)
        xbook.close()




    def clean_tweet(self, tweet): 
        ''' 
        Utility function to clean tweet text by removing links, special characters 
        using simple regex statements. 
        '''
        return ' '.join(re.sub("(@[A-Za-z0-9]+)|([^0-9A-Za-z \t]) |(\w+:\/\/\S+)", " ", tweet).split()) 
  
    def get_tweet_sentiment(self, tweet): 
        analysis = TextBlob(self.clean_tweet(tweet)) 
        return analysis.sentiment.polarity
        
    def get_tweets_dbf(self, fileName): 
        tweets = []
        i=0

        
        fetched_tweets = DBF(fileName,  encoding='latin-1')
        
        print('field_names')
        print(fetched_tweets.field_names)
        # parsing tweets one by one 
        for tweet in fetched_tweets: 
            # empty dictionary to store required params of a tweet 
            parsed_tweet = {} 
            parsed_tweet['Airline'] = ''
            parsed_tweet['UserName'] = ''
            parsed_tweet['FormalName'] = ''
            parsed_tweet['TimeStamp'] = ''
            parsed_tweet['Text'] = ''
            if tweet !=None:
                # saving text of tweet 
                if tweet['Airline'] !=None:
                    parsed_tweet['Airline'] = tweet['Airline']
                if tweet['UserName'] !=None:
                    parsed_tweet['UserName'] = tweet['UserName']
                if tweet['FormalName'] !=None:
                    parsed_tweet['FormalName'] = tweet['FormalName']
                if tweet['TimeStamp'] !=None:
                    parsed_tweet['TimeStamp'] = tweet['TimeStamp']
                if tweet['Text'] !=None:
                    parsed_tweet['Text'] = tweet['Text']

            
            # saving sentiment of tweet 
            parsed_tweet['SentimentPolarity'] = self.get_tweet_sentiment(tweet['Text']) 
            if parsed_tweet['SentimentPolarity'] > 0: 
                parsed_tweet['Sentiment'] = 'positive'
            elif parsed_tweet['SentimentPolarity'] == 0: 
                parsed_tweet['Sentiment'] = 'neutral'
            else: 
                parsed_tweet['Sentiment']= 'negative'
            
            tweets.append(parsed_tweet) 
           
            i=i+1
            print('sentiments i = ', i)
        # return parsed tweets 
        return tweets 
  


def main(): 
    # creating object of TwitterClient Class 
    api = TwitterClient() 
    # api.dbf_to_xlsx('tweets.dbf')

    # calling function to get tweets 
    tweets = api.get_tweets_dbf('tweets.dbf') 
    print('================================')
    print(tweets)
    # with open("sentiments.csv",'wb') as resultFile:
    #     wr = csv.writer(resultFile, dialect='excel')
    #     for tweet in tweets:
    #         wr.writerow([tweet])
    #     # wr.writerows([tweets])
    


    headers=['Airline', 'UserName', 'FormalName', 'TimeStamp', 'Text', 'SentimentPolarity', 'Sentiment']
    api.save_to_xlsx('tweets-sentiments', tweets, headers)
    # picking positive tweets from tweets 
    ptweets = [tweet for tweet in tweets if tweet['Sentiment'] == 'positive'] 
    # percentage of positive tweets 
    print("Positive tweets percentage: {} %".format(100*len(ptweets)/len(tweets))) 
    # picking negative tweets from tweets 
    ntweets = [tweet for tweet in tweets if tweet['Sentiment'] == 'negative'] 
    # percentage of negative tweets 
    print("Negative tweets percentage: {} %".format(100*len(ntweets)/len(tweets))) 
    # percentage of neutral tweets 
    print("Neutral tweets percentage: {} %".format(100*(len(tweets) - len(ntweets) - len(ptweets))/len(tweets))) 
    print('total tweets:', len(tweets))
    print('positive tweets:', len(ptweets))
    print('negative tweets:', len(ntweets))
    print('neutral tweets', len(tweets) - len(ntweets) - len(ptweets))

    # printing first 5 positive tweets 
    print("\n\nPositive tweets:") 
    for tweet in ptweets[:10]: 
        print(tweet['Text']) 
  
    # printing first 5 negative tweets 
    print("\n\nNegative tweets:") 
    for tweet in ntweets[:10]: 
        print(tweet['Text']) 
  
if __name__ == "__main__": 
    # calling main function 
    main() 
