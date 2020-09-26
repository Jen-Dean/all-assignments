import os
import re

word_count = []
sentence_count = []

p_txt = os.path.join("PyParagraph/Resources/paragraph_2.txt")

with open(p_txt) as txtfile:
    pg = txtfile.read()

#Import a text file filled with a paragraph of your choosing.

#Assess the passage for each of the following:

#Approximate word count
words = re.split(r'[;,\s]\s*', pg)
count_words = len(words)
#print(count_words)

#Approximate sentence count
sentences = pg.split(".")
count_sentences = len(sentences)
#print(count_sentences)

#Approximate letter count (per word)
for word in words:
    word_count.append(len(word))
Average = (sum(word_count) / len(word_count))
Round_Av_word = (round(Average, 1))

#Average sentence length (in words)
for sentence in sentences:
    sentence_count.append(len(sentence))
AverageS = (sum(sentence_count)/ len(sentence_count))
Round_Av_Sent = (round(AverageS, 1))

print("Paragraph Analysis")
print("-----------------")
print(f"Approximate Word Count: {count_words}")
print(f"Approximate Sentence Count: {count_sentences}")
print(f"Average Letter Count: {Round_Av_word}")
print(f"Average Sentence Length: {Round_Av_Sent}")