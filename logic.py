#!/usr/bin/env python

from collections import Counter
import matplotlib.pyplot as plt
from pylab import axes, pie, figure, title,  savefig
from matplotlib.backends.backend_pdf import PdfPages
import datetime
import numpy as np
import argparse
import glob
import os
import pandas as pd
import csv
import zipfile
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak, PageTemplate
from reportlab.lib import colors
from reportlab.lib.units import cm, inch
from reportlab.lib.pagesizes import A3, A4, landscape, portrait, letter, cm
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER, TA_JUSTIFY
from reportlab.pdfgen import canvas
from netaddr import IPNetwork, IPAddress
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Frame as fs
from functools import partial
import reportlab
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import registerFontFamily
from reportlab.pdfbase.ttfonts import TTFont
import StringIO
import itertools


cwd = os.getcwd()
csv_cwd = os.path.join(cwd, "csv/" )
zip_cwd = os.path.join(cwd, "zipped/" )
logo_cwd = os.path.join(cwd, "logo.png" )

# From the excel file generate a dictionary 
def generate_ip_blocks(ipblock):
    keys = []
    values = []
    file_name = os.path.join(cwd,ipblock)
    ip_allocations = pd.read_excel(file_name).to_dict()
    addresses = ip_allocations.items()[0][1]
    institutions = ip_allocations.items()[1][1]
    for key in addresses:
        if addresses.keys() == institutions.keys():
            keys.append(str(institutions[key]))
            values.append(str(addresses[key]))
    dictionary = dict(itertools.izip(values,keys))
    return dictionary

# Unzip files and put them in a new folder
def extract_file(zip_files_path, reportspath):
    extension = ".zip"
    os.chdir(zip_files_path)
    val= zip_files_path.split('/')[1]
    print zip_files_path 
    print val
    for item in os.listdir(zip_files_path): # loop through items in dir
        if item.endswith(extension): # check for ".zip" extension
           
            file_name = os.path.abspath(item) # get full path of files
            zip_ref = zipfile.ZipFile(file_name) # create zipfile object
            zip_ref.extractall(reportspath) # extract file to dir
            zip_ref.close() # close file
            #os.remove(file_name) # delete zipped file

# Combine all the CSV files into one file
def combined_files(zip_files_path, reportspath):
    extract_file(zip_files_path,  reportspath)
    read_files = glob.glob(os.path.join(reportspath,"*.csv"))
    with open(reportspath + "\combined_files.csv", "w") as outfile:
        #print read_files
        for filename in read_files:
            attack_name = filename.split("-")[3]
            with open(filename) as infile:
                for line in infile:
                    lineElement = line.split(',')
                    outfile.write('{},{},{}\n'.format(lineElement[0],lineElement[1], attack_name))
            os.remove(filename)

    with open(reportspath +'\combined_files.csv', 'rb') as inp, open(reportspath +'\AllFiles.csv', 'wb') as out:
        writer = csv.writer(out)
        for row in csv.reader(inp):
            if row[0] != "timestamp":
                writer.writerow(row)
    os.remove(reportspath +'\combined_files.csv')

# Gnerate a unique pdf report
def renu_report(ipblock, zip_files_path, reportspath):
    ips = []
    attack = []
    allattacks = []
    with  open(reportspath + '\AllFiles.csv', 'rb') as infile:
        reader = csv.reader(infile)
        mydict1 = dict((rows[0],rows[1]) for rows in reader)
    with  open(reportspath + '\AllFiles.csv', 'rb') as infile:
        reader = csv.reader(infile)
        mydict2 = dict((rows[0],rows[2]) for rows in reader)
    for key in mydict1:
         ips.append(mydict1[key])
    for key in mydict2:
         attack.append(mydict2[key])
    allattacks = zip(attack,ips)
    newl = []
    for tup in  allattacks:
        if tup in newl:
            pass
        else:
            newl.append(tup)
    counts = dict(Counter(elem[0] for elem in newl))
    with PdfPages(reportspath + '\RENU REPORT.pdf') as pdf:

        fig = plt.figure()
        width = 0.35
        ind = np.arange(len(counts))
        plt.bar( range(len(counts)), list(counts.values()), width=width)
        plt.xticks(range(len(counts) + int(0.15)) , list(counts.keys()))
        plt.title('Percentage of hosts affected by same Vulnerability', bbox={'facecolor':'0.5', 'pad':3})
        plt.ylabel('Number of Hosts')
        plt.xlabel('Name of Vulnerability')
        fig.autofmt_xdate()
        pdf.savefig()
        plt.close()


        plt.rc('text', usetex=False)
        fig2 = plt.figure(1, figsize=(6,6))
        ax = axes([0.1, 0.1, 0.8, 0.8])
        # The slices will be ordered and plotted counter-clockwise.
        labels = list(counts.keys())
        fracs = list(counts.values())
        plt.pie(fracs, labels=labels,
                        autopct='%1.1f%%', shadow=True, startangle=90)
        plt.title('Percentage of hosts affected by same Vulnerability', bbox={'facecolor':'0.5', 'pad':3})
        pdf.savefig(fig2)  # or you can pass a Figure object to pdf.savefig
        plt.close()


def company_reports(ipblock, zip_files_path, reportspath):
    with open(reportspath + '\AllFiles.csv', 'rb') as bigfile:
        for row in csv.reader(bigfile):
            for k in generate_ip_blocks(ipblock):
                ipadr =  str(k).replace(".", "/")
                if IPAddress(str(row[1])) in IPNetwork(ipadr.replace("/", ".", 3)):
                    if os.path.isfile(reportspath + "\\" + str(generate_ip_blocks(ipblock)[k]) + '.csv'):
                        writer = csv.writer(open(reportspath + "\\" + str(generate_ip_blocks(ipblock)[k]) + '.csv', 'ab'))
                    else:
                        writer = csv.writer(open(reportspath + "\\" + str(generate_ip_blocks(ipblock)[k]) + '.csv', 'wb'))
                        writer.writerow(["TIMESTAMP", "HOST IPADDRESS", "NAME OF ATTACK"])
                    writer.writerow(row)

def generatepdf(ipblock, zip_files_path, reportspath):
     # combined_files(zip_files_path, reportspath)
     # renu_report(ipblock, zip_files_path, reportspath)
     # company_reports(ipblock, zip_files_path, reportspath)
     # os.remove(reportspath + '\AllFiles.csv')
     readfiles = glob.glob(os.path.join(reportspath,"*.csv"))
     for filename in readfiles:
         if "combined_files" in filename:
             os.remove(reportspath +'\combined_files.csv')
         filena, file_extension = os.path.splitext(filename)
         filenam = str(filena.split("\\")[-1])

         # Content.
         line1 = "Vulnerability Report for "
         content =  line1 + filenam
         content2 = line1 + filenam + " continued ... "

         def logo(canvas, doc):
             logo_img = 'F:\\flask\logo.png'
             canvas.drawImage(logo_img, 300, 730)
             canvas.drawString(100,715,content)
         def logo2(canvas, doc):
             logo_img = 'F:\\flask\logo.png'
             canvas.drawImage(logo_img, 300, 730)
             canvas.drawString(100,715,content2)

         pdfReportPages = str(filena) + '.pdf'
         doc = SimpleDocTemplate(pdfReportPages, pagesize=letter)

         elements = []

         col1 = "TIMESTAMP"
         col2 = "HOST"
         col3 = "NAME OF VULNERABILITY"
         data = [[col1, col2, col3]]
         with open(filename) as infile:
             for line in infile:
                 reader = csv.reader(infile)
                 data.extend(list(reader))
         tableThatSplitsOverPages = Table(data, [5 * cm, 6 * cm, 5* cm], repeatRows=1)
         tableThatSplitsOverPages.hAlign = 'LEFT'
         tblStyle = TableStyle([('TEXTCOLOR',(0,0),(-1,-1),colors.black),
                                ('VALIGN',(0,0),(-1,-1),'TOP'),
                                ('LINEBELOW',(0,0),(-1,-1),1,colors.black),
                                ('INNERGRID', (0,0), (-1,-1), 1.5, colors.black),
                                ('BOX', (0,0), (-1,-1), 1.5, colors.black)])
         tblStyle.add('BACKGROUND',(0,0),(-1,-1),colors.lightblue)
         tblStyle.add('BACKGROUND',(0,1),(-1,-1),colors.white)
         tableThatSplitsOverPages.setStyle(tblStyle)
         elements.append(tableThatSplitsOverPages)

         doc.topMargin=1.3* inch
         doc.build(elements, onFirstPage=logo, onLaterPages=logo2)
         del data[:]
         os.remove(filename)

def generate_pdf_reports(ipblock, zip_files_path,reportspath):

    generatepdf(ipblock, zip_files_path, reportspath)
