# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import spacy
import docx2txt
import pandas as pd
import numpy as np

# Load English tokenizer, tagger, parser, NER and word vectors
nlp = spacy.load('en_core_web_sm')


my_text = docx2txt.process("test.docx")

my_text = my_text.replace('\xa0', ' ')
# Process whole documents
'''
text = (u'AutoAlliance (Thailand) Co., Ltd. (“AAT”) — a 50/50 joint'
        u'venture between Ford and Mazda that owns and operates a manufacturing'        
        u'plant in Rayong, Thailand. AAT produces Ford and Mazda products for '
        u'domestic and export sales. Changan Ford Automobile Corporation, Ltd. '
        u'(“CAF”) — a 50/50 joint venture between Ford and Chongqing Changan '
        u'Automobile Co., Ltd. (“Changan”). CAF currently operates five assembly '
        u'plants, an engine plant, and a transmission plant in China where it '
        u'produces and distributes an expanding variety of Ford passenger vehicle models.'
        u'Changan Ford Mazda Engine Company, Ltd. (“CFME”) — a joint venture among '
        u'Ford (25% partner), Mazda (25% partner), and Changan (50% partner).  '
        u'CFME is located in Nanjing, and produces engines for Ford and Mazda '
        u'vehicles manufactured in China.Ford Otomotiv Sanayi Anonim Sirketi '
        u'(“Ford Otosan”) — a joint venture in Turkey among Ford (41% partner), '
        u'the Koc Group of Turkey (41% partner), and public investors (18%) that '
        u'is a major supplier to us of the Transit, Transit Custom, and Transit Courier '
        u'commercial vehicles and is our sole distributor of Ford vehicles in Turkey. '
        u'Ford Otosan also makes the Cargo truck for the Turkish and export markets, '
        u'and certain engines and transmissions, most of which are under license from us. '  
        u'The joint venture owns two plants, a parts distribution depot, a product '
        u'development center, and a new research and development center in Turkey.'
        u'Getrag Ford Transmissions GmbH (“GFT”) — a 50/50 joint venture with '
        u'Getrag International GmbH, a German company. GFT operates plants in Halewood, '
        u'England; Cologne, Germany; Bordeaux, France; and Kechnex, Slovakia to produce, '
        u'among other things, manual transmissions for our Europe business unit. '
        u'JMC — a publicly-traded company in China with Ford (32% shareholder) and '
        u'Jiangling Holdings, Ltd. (41% shareholder) as its controlling shareholders.  '
        u'Jiangling Holdings, Ltd. is a 50/50 joint venture between Changan and Jiangling Motors Company Group. '
        u'The public investors in JMC own 27% of its total outstanding shares.  '
        u'JMC assembles Ford Transit, Ford Everest, Ford engines, and non-Ford vehicles and engines for distribution in China and '
        u'in other export markets. JMC operates two assembly plants and one engine plant in Nanchang. '
        u'In 2015, JMC opened a new plant in Taiyuan to assemble heavy duty trucks and engines.')
'''

doc = nlp(my_text)


entity_text = []
entity_label = []
entity_seten = []

for entity in doc.ents:
    if entity.label_ == 'ORG':
        for sentence in doc.sents:
            if entity.start_char >= sentence.start_char and entity.end_char <= sentence.end_char:    
                entity_text.append(entity.text)
                entity_label.append(entity.label_)
                entity_seten.append(sentence.text)

df = pd.DataFrame({'name':entity_text, 'label':entity_label, 'sentence':entity_seten})

df.to_excel('word_parse_result.xlsx')
 
    




