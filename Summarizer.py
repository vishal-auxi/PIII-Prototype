from sumy.parsers.html import HtmlParser
from sumy.parsers.plaintext import PlaintextParser
from sumy.nlp.tokenizers import Tokenizer
from sumy.summarizers.lex_rank import LexRankSummarizer
from sumy.summarizers.lsa import LsaSummarizer as Summarizer
from sumy.nlp.stemmers import Stemmer
from sumy.utils import get_stop_words

from nltk.tokenize import sent_tokenize

LANGUAGE = "english"
SENTENCES_COUNT = 5


def summarize_lsa(document, sentences_count=SENTENCES_COUNT):
    parser = PlaintextParser.from_string(document, Tokenizer(LANGUAGE))

    # parser = HtmlParser.from_url(url, Tokenizer(LANGUAGE))
    # parser = PlaintextParser.from_file("covid.txt", Tokenizer(LANGUAGE))

    stemmer = Stemmer(LANGUAGE)

    summarizer = Summarizer(stemmer)
    summarizer.stop_words = get_stop_words(LANGUAGE)

    result = summarizer(parser.document, sentences_count)
    summary = [str(i) for i in list(result)]
    return summary


def summarize_lex_rank(document, sentences_count=SENTENCES_COUNT):
    summarizer = LexRankSummarizer()

    parser = PlaintextParser.from_string(document, Tokenizer(LANGUAGE))

    summary = summarizer(parser.document, sentences_count)

    for sentence in summary:
        print(sentence)


def sentence_count(document):
    return len(sent_tokenize(document))
