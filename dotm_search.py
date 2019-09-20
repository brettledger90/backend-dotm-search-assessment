#!/usr/bin/env python
# -*- coding: utf-8 -*-
import zipfile 
import os 
import argparse
"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
def create_parser():
    parser = argparse.ArgumentParser()
    parser.add_argument("-d","--dir", help="directory_name")
    parser.add_argument("text", help="Text Search For")
    
    return parser

__author__ = "???"
def paths_to_filez(directory):
    nothing_fancy = []
    for root,directories,files in os.walk(directory):
        for file_name in files:
            path = os.path.join(root,file_name)
            nothing_fancy.append(path)
    return nothing_fancy


def searchin_stuff(file,text):
    document = zipfile.ZipFile(file)
    matcharoos=0
    with document.open("word/document.xml") as opening_stuff:
        for line in opening_stuff:
            temp1 = line.find(text)
            if temp1 != -1:
                cool_junk=line[temp1-40:temp1+40]
                matcharoos += 1
                print(cool_junk)
    
    return matcharoos


        

def main():
    parser = create_parser()
    args= parser.parse_args()
    pathz= paths_to_filez(args.dir)
    lil_counter=0
    counting_matches=0
    for file in pathz:
        if file.endswith('.dotm'):
            lil_counter += 1
            lil_matches= searchin_stuff(file,args.text)
            if lil_matches > 0:
                counting_matches += 1
                print(file)
        else:
            continue

    print("Total Files with Match : "+ str(counting_matches) + " Total number of files searched : " + str(lil_counter))
    





if __name__ == '__main__':
    main()
