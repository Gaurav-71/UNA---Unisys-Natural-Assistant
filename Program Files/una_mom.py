import os
import speech_recognition as sr
import ffmpeg
import subprocess
import soundfile as sf
from punctuator import Punctuator

import language_tool_python
from moviepy.editor import *

#Summary
from nutshell.algorithms.information_retrieval import ClassicalIR
from nutshell.algorithms.ranking import TextRank
from nutshell.algorithms.similarity import BM25Plus
from nutshell.model import Summarizer
from nutshell.preprocessing.cleaner import NLTKCleaner
from nutshell.preprocessing.preprocessor import TextPreProcessor
from nutshell.preprocessing.tokenizer import NLTKTokenizer
from nutshell.utils import load_corpus, construct_sentences_from_ranking
import nltk

#Keywords
from nutshell.algorithms.information_retrieval import ClassicalIR
from nutshell.model import KeywordExtractor
from nutshell.preprocessing.cleaner import NLTKCleaner
from nutshell.preprocessing.preprocessor import TextPreProcessor
from nutshell.preprocessing.tokenizer import NLTKTokenizer
from nutshell.utils import load_corpus

import sys


def main(fileName):
    fileName, fileExt = fileName.split('.')
    print(fileName, fileExt)
    # os.system(f"ffmpeg -i C:\\Users\\Chinmay\\CSE\\pyCode\\Unysis\\UNA-Unisys-Natural-Assistant-main\\UNA-Unisys-Natural-Assistant-main\\{fileName}.{fileExt} -ab 160k -ac 2 -ar 44100 -vn C:\\Users\\Chinmay\\CSE\\pyCode\\Unysis\\UNA-Unisys-Natural-Assistant-main\\UNA-Unisys-Natural-Assistant-main\\{fileName}.wav")

    # ipFile = ffmpeg.input(fileName + fileExt)
    # opFile = ffmpeg.output(ipFile, fileName + ".wav")
    
    clip = AudioFileClip(f"{fileName}.{fileExt}")
    clip.write_audiofile(f"{fileName}.wav",codec='pcm_s16le')

    f = sf.SoundFile(f'{fileName}.wav')
    audio_dur = len(f) / f.samplerate


    r = sr.Recognizer()
    text = ""
    rec_dur = 25

    with sr.AudioFile(f'{fileName}.wav') as source:
        for x in range(0, int(audio_dur / rec_dur)):
            audio = r.record(source, duration = rec_dur) 
            try:
                new_txt = r.recognize_google(audio)
                text = text + new_txt
            except:
                pass
            
        audio = r.record(source, duration = (audio_dur - int(audio_dur/rec_dur)))
        try:
            new_txt = r.recognize_google(audio)
            text = text + new_txt
        except:
            pass

        print("Done")


    p = Punctuator('Demo-Europarl-EN.pcl')
    text = p.punctuate(text)

    tool = language_tool_python.LanguageTool('en-US') 

    matches = tool.check(text)
    print(len(matches))

    for lab in range(len(matches)):
        print(lab)
        print(matches[lab].ruleId, matches[lab].replacements)

    text_new = tool.correct(text)

    print(text_new)



    nltk.download('punkt')
    nltk.download('stopwords')

    preprocessor = TextPreProcessor(NLTKTokenizer(), NLTKCleaner())
    similarity_algorithm = BM25Plus()
    ranker = TextRank()
    ir = ClassicalIR()

    # Text Summarization
    model = Summarizer(preprocessor, similarity_algorithm, ranker, ir)
    summarised_content = model.summarise(text_new, reduction_ratio=0.80, preserve_order=True)

    print("\n --- Summarized Text ---\n")
    print(construct_sentences_from_ranking(summarised_content))
    
    with open(f"{fileName}.txt", "w+") as file:
        file.write(construct_sentences_from_ranking(summarised_content))


    # Text Keyword Extraction
    preprocessor = TextPreProcessor(NLTKTokenizer(), NLTKCleaner(skip_stemming=True))
    keyword_extractor = KeywordExtractor(preprocessor, ClassicalIR())
    keywords = keyword_extractor.extract_keywords(text, count=10, raw=False)


    print("\n --- Keywords ---\n")
    print(keywords)

print(sys.argv[1])
main(sys.argv[1])