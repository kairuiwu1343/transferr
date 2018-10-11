# -*- coding: utf-8 -*-
"""
Created on Tue Oct  9 03:48:41 2018

@author: Karry
"""
import spacy
import pandas as pd
import numpy as np
from pdfminer.pdfparser import PDFParser,PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextBoxHorizontal,LAParams
from pdfminer.pdfinterp import PDFTextExtractionNotAllowed


path = r'sr17-2017-form-10-k.pdf'
def parse():
    results = ''
    fp = open(path, 'rb') 

    praser = PDFParser(fp)

    doc = PDFDocument()

    praser.set_document(doc)
    doc.set_parser(praser)


    doc.initialize()


    if not doc.is_extractable:
        raise PDFTextExtractionNotAllowed
    else:

        rsrcmgr = PDFResourceManager()

        laparams = LAParams()
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)

        interpreter = PDFPageInterpreter(rsrcmgr, device)


        for page in doc.get_pages(): 
            interpreter.process_page(page)

            layout = device.get_result()
            
            for x in layout:
                if (isinstance(x, LTTextBoxHorizontal)):
                        results += x.get_text()
    return results
                     
                        

data = parse()

nlp = spacy.load('en_core_web_sm')
doc = nlp(data)


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

df.to_excel('pdf_result.xlsx')