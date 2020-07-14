"""
See the blog post https://www.pyxll.com/blog/a-real-time-twitter-feed-in-excel/

This version has been updated to use the pandas DataFrame type to return
all the data in a single call.
"""
from pyxll import RTD, xl_func
from tweepy.streaming import StreamListener
from tweepy import OAuthHandler
from tweepy import Stream
import pandas as pd
import threading
import logging
import json

_log = logging.getLogger(__name__)

# User credentials to access Twitter API
access_token = "YOUR ACCESS TOKEN"
access_token_secret = "YOUR ACCESS TOKEN SECRET"
consumer_key = "YOUR CONSUMER KEY"
consumer_secret = "YOUR CONSUMER KEY SECRET"


class TwitterListener(StreamListener):
    """tweepy.StreamListener that notifies multiple subscribers when
    new tweets are received and keeps a buffer of the last 100 tweets
    received.
    """
    __max_size = 100

    def __init__(self, phrases):
        self.__lock = threading.RLock()
        self.__connected = False
        self.__phrases = phrases
        self.__subscriptions = set()
        self.__tweets = []

    def connect(self):
        """Listen for tweets in a background thread"""
        with self.__lock:
            if self.__connected:
                raise AssertionError("Already connected")

            _log.info("Connecting twitter listener for [%s]" % ", ".join(self.__phrases))
            auth = OAuthHandler(consumer_key, consumer_secret)
            auth.set_access_token(access_token, access_token_secret)
            self.__stream = Stream(auth, listener=self)
            self.__stream.filter(track=self.__phrases, is_async=True)
            self.__connected = True

    def disconnect(self):
        """Disconnect from the twitter stream and remove from the cache of listeners."""
        with self.__lock:
            if not self.__connected:
                raise AssertionError("Not connected")

            _log.info("Disconnecting twitter listener for [%s]" % ", ".join(self.__phrases))
            self.__stream.disconnect()
            self.__connected = False

    def subscribe(self, subscriber):
        """Add a subscriber that will be notified when new tweets are received"""
        with self.__lock:
            self.__subscriptions.add(subscriber)
            if not self.__connected:
                self.connect()

    def unsubscribe(self, subscriber):
        """Remove subscriber added previously.
        When there are no more subscribers the listener is stopped.
        """
        with self.__lock:
            self.__subscriptions.remove(subscriber)
            if not self.__subscriptions and self.__connected:
                self.disconnect()

    def on_data(self, data):
        data = json.loads(data)
        if data:
            with self.__lock:
                self.__tweets.insert(0, data)
                self.__tweets = self.__tweets[:self.__max_size]
                for subscriber in self.__subscriptions:
                    try:
                        subscriber.on_data(data)
                    except:
                        _log.error("Error calling subscriber", exc_info=True)
        return True

    def on_error(self, status):
        with self.__lock:
            for subscriber in self.__subscriptions:
                try:
                    subscriber.on_error(status)
                except:
                    _log.error("Error calling subscriber", exc_info=True)

    def get_dataframe(self, columns):
        """Build a DataFrame from the current tweets"""
        with self.__lock:
            # Check if we have any tweets
            if not self.__tweets:
                df_dicts = [{c: "No tweets :(" for c in columns}]
                return pd.DataFrame(df_dicts, columns=columns)

            # Construct a list of dicts from all the tweets and the columns
            df_dicts = []
            for tweet in self.__tweets:
                df_dict = {}
                for col in columns:
                    value = tweet
                    for key in col.split("."):
                        if not isinstance(value, dict):
                            value = ""
                            break
                        value = value.get(key, {})
                    df_dict[col] = value
                df_dicts.append(df_dict)

            # Build a DataFrame from the dictionaries
            return pd.DataFrame(df_dicts, columns=columns)


class TwitterRTD(RTD):

    def __init__(self, phrases, columns):
        self.__listener = TwitterListener(phrases)
        self.__columns = columns

        initial_value = self.__listener.get_dataframe(self.__columns)
        super().__init__(value=initial_value)

    def connect(self):
        self.__listener.subscribe(self)

    def disconnect(self):
        self.__listener.unsubscribe(self)

    def on_data(self, data):
        self.value = self.__listener.get_dataframe(self.__columns)

    def on_error(self, status):
        self.value = f"#ERROR {status}"


@xl_func("string[], string[]: rtd")
def twitter_listen(phrases, columns):
    return TwitterRTD(phrases, columns)


@xl_func("object: dataframe<index=False, columns=False>")
def df_explode(df, n=100):
    return df.head(n)


if __name__ == '__main__':
    import time
    logging.basicConfig(level=logging.INFO)

    class TestSubscriber(object):
        """simple subscriber that just prints tweets as they arrive"""

        def on_error(self, status):
            print(f"Error: {status}")

        def on_data(self, data):
            print(f"{data['user']['name']}: {data['text']}")

    # Create the subscriber that will print out the tweets as they arrive
    subscriber = TestSubscriber()

    # Create the twitter listener and subscribe to updates
    listener = TwitterListener(["python", "excel"])
    listener.subscribe(subscriber)

    # listen for 10 seconds then stop
    time.sleep(10)
    listener.unsubscribe(subscriber)
