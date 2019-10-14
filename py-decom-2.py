# Copyright (C) 2019 John Edward Davidson <john.davidson.ctr@hotmail.co.uk>

import os, requests, re, easygui as gui, html5lib, io, datetime, shutil, webbrowser
from bs4 import BeautifulSoup

ppath = "C:\OGA\Programmes"
os.makedirs(ppath, exist_ok=True)
os.chdir(ppath)

rpath = "C:\OGA\Reports"
os.makedirs(rpath, exist_ok=True)

wpath = "C:\OGA\Wells"
os.makedirs(wpath, exist_ok=True)

def getpdfs():
    """Parameters
       ----------
       None

       Returns
       -------
       filename
           PDF file downloaded from OGA site and written to hard drive
       """

    url = ('https://www.gov.uk/guidance/oil-and-gas-decommissioning'
           '-of-offshore-installations-and-pipelines#table-of-draft'
           '-decommissioning-programmes-under-consideration')

    res = requests.get(url)
    res.raise_for_status()
    page = requests.get(url).content
    soup = BeautifulSoup(page, 'html5lib')

    # Find all links on the page
    links = soup.find_all('a', href=True)
    for link in links:
        link_url = requests.compat.urljoin(url, link['href'])
        if link_url.endswith('.pdf'):
            # Download pdf files and write to Programmes directory
            filename = link_url.split('/')[-1]
            with open(filename, 'wb') as pdffile:
                pdffile.write(requests.get(link_url).content)
                print(filename)

from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfpage import PDFPage
import chardet

def extract_text_from_pdf(pdf_path):
    """Parameters
       ----------
       pdf_path
           passed PDF Programme filename in convert_files function
       Returns
       -------
       filename
           decFile object - PDF file converted to .txt
       """

    resource_manager = PDFResourceManager()
    fake_file_handle = io.StringIO()
    converter = TextConverter(resource_manager, fake_file_handle)
    page_interpreter = PDFPageInterpreter(resource_manager, converter)

    with open(pdf_path, 'rb') as fh:
        for page in PDFPage.get_pages(fh,
                                      caching=True,
                                      check_extractable=True):
            page_interpreter.process_page(page)

        text = fake_file_handle.getvalue()

    # Close open handle
    converter.close()
    fake_file_handle.close()

    if text:
        #Write returned text to file
        decFile = open(pdf_path + '.txt', 'w', encoding='utf-8')
        decFile.write(text)
        decFile.close()

def convert_files():
    """Parameters
       ----------
       None

       Returns
       -------
       extract_text_from_pdf(filename)
           filename PDF object converted to .txt
       exception, UnicodeEncodeError or TypeError
           Filename of PDF file that could not be converted
       """

    for filename in os.listdir('.'):
        txt_file = filename + '.txt'
        if filename.endswith('.pdf') and txt_file not in os.listdir('.'):
            try:
                print('Converting ' + filename)
                extract_text_from_pdf(filename)
            except (UnicodeEncodeError, TypeError) as error:
                print('Could not convert file {0}'.format(filename))
                pass
            else:
                extract_text_from_pdf(filename)

def term_search():
    """Parameters
       ----------
       None
           
       Returns
       -------
       fileList
           A list of all the Programme files with matching serach terms
       """

    global fileList
    fileList = []
    srtitle = "py-decom"
    usearch = gui.enterbox("Enter a Search Term: ", srtitle)
    if usearch:
        pass
    else:
        rootmenu()
    fsearch = usearch.lower()
    fileDate = datetime.datetime.now()
    fileName = fileDate.strftime("%Y%m%d%H%M%S")
    docName = fileDate.strftime("%d-%m-%Y %H:%M:%S")
    searchName = ('Decom-Search-' + fileName + '.txt')
    searchfile = open(searchName, 'w', encoding='utf-8')
    searchfile.write('Decom Search for: ' + usearch + '\n\n' + docName + '\n\n')
    searchfile.write('Your search term appeared in these Programmes:\n\n')
    for filename in os.listdir('.'):
        if filename.endswith('.txt'):
            with open(filename, 'r', encoding='utf-8') as reader:
                line = reader.readline().lower()
                if fsearch in line:                
                    searchfile.write(filename + '\n\n')
                    fileList.append(filename)
    searchfile.close()
    shutil.move(searchName, rpath)
    getFile = os.path.join(rpath, searchName)
    webbrowser.get().open(getFile)
    return fileList

def open_pdfs():
    """Parameters
       ----------
       None
           
       Returns
       -------
       os.startfile(pdfFile)
           Opens each Programme that matches search criteria from term_search
       """

    # Remove .txt extension so that pdfFile object can be used to open files
    pdfList = ([item[:-4] for item in fileList])
    for pdfFile in pdfList:
        os.startfile(pdfFile)

def wellsearch():
    """Parameters
       ----------
       None
           
       Returns
       -------
       webbrowser.get().open(wgetFile)
           Opens the text file showing well search results
       """

    wfileList = []
    # Regex configuration TODO - combine these
    wellregex1 = re.compile(r'\d{1,3}/\d{1,2}-\w{1,4}')
    wellregex2 = re.compile(r'\d{1,3}/\d{1,2}\w-\w{1,4}')
    wellregex3 = re.compile(r'\d{1,3}/\d{1,2}[a-z]')
    wfileDate = datetime.datetime.now()
    wfileName = wfileDate.strftime("%Y%m%d%H%M%S")
    wdocName = wfileDate.strftime("%d-%m-%Y %H:%M:%S")
    wsearchName = ('Well-Search-' + wfileName + '.txt')
    wsearchfile = open(wsearchName, 'w', encoding='utf-8')
    wsearchfile.write('Decom Wells Search ' + '\n\n' + wdocName + '\n\n')
    wsearchfile.write('Well search on current Programmes returned the following:\n\n')
    for filename in os.listdir('.'):
        if filename.endswith('.txt'):
            with open(filename, 'r', encoding = 'utf-8') as reader:
                line = reader.readline()
                wmatch1 = re.findall(wellregex1, line)
                wfileList.append(wmatch1)
                wmatch2 = re.findall(wellregex2, line)
                wfileList.append(wmatch2)
                wmatch3 = re.findall(wellregex3, line)
                wfileList.append(wmatch3)
    # Flatten wfileList
    well_flat_list = [val for sublist in wfileList for val in sublist]
    # Remove duplicates from well_flat_list
    well_remd_list = list(dict.fromkeys(well_flat_list))
    well_remd_list.sort()
    well_count = len(well_remd_list)
    well_sort_list = '\n'.join(well_remd_list)
    wsearchfile.write('Approximately ' + str(well_count) + ' Wells are under consideration or have been decommissioned.\n\n')
    wsearchfile.write(well_sort_list)
    wsearchfile.close()
    shutil.move(wsearchName, wpath)
    wgetFile = os.path.join(wpath, wsearchName)
    webbrowser.get().open(wgetFile)

#GUI Configuration
def resource_path(relative_path):
    """Parameters
       ----------
       relative_path
           Image file
           
       Returns
       -------
       os.path.join link to temporary pyinstaller sys._MEIPASS file
       """

    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

def rootmenu():
    """Parameters
       ----------
       None
           
       Returns
       -------
       rselect
           Gui object allowing multiple choices contained in schoice list
       getpdfs, convert_files....
           Function call aligned with schoice 1-8 objects
       """

    #image = "C:\\RealPython\\python-decom-2\\py-decom.png"
    decimg = resource_path('py-decom.png')
    startmsg = ""
    schoice1 = 'Download Programmes'
    schoice2 = 'Convert Files'
    schoice3 = 'Search Programmes'
    schoice4 = 'Open Files'
    schoice5 = 'Well Search'
    schoice6 = 'Readme'
    schoice7 = 'Licence'
    schoice8 = 'Exit'
    schoice = [schoice1, schoice2, schoice3, schoice4, schoice5, schoice6, schoice7, schoice8]
    rselect = gui.buttonbox(startmsg, image=decimg, choices=schoice)
    # Use rootmenu as cyclic event loop
    if rselect == schoice1:
        getpdfs()
        rootmenu()
    if rselect == schoice2:
        convert_files()
        rootmenu()
    if rselect == schoice3:
        term_search()
        rootmenu()
    if rselect == schoice4:
        open_pdfs()
        rootmenu()
    if rselect == schoice5:
        wellsearch()
        rootmenu()
    if rselect == schoice6:
        readme()
        rootmenu()
    if rselect == schoice7:
        rlicence()
        rootmenu()
    if rselect == schoice8:
        exit()

def readme():
    """Parameters
       ----------
       None
           
       Returns
       -------
       readmepg
           Gui object containing README information
       """

    rdmsg = "py-decom Readme"
    rdtitle = "py-decom"
    rdtext = """
    py-decom

    What this program does:

    1. Looks at all of the approved and under consideration 
       Decommissioning [PDF] programmes on the OGA site.
    2. Saves them all to hard drive.
    3. Converts them to a fully searchable format.
    4. Searches them with user input search terms and outputs 
       a results file.
    5. Opens all the search files on user request.
    6. Searches all the downloaded files for recorded well numbers.

    Instructions:

    1. Download all the current Decom Programmes using the 
       'Download Programmes' button.
    2. Convert the files to searchable [.txt] format using the 
       'Convert Files' button.
    3. Search the Programme Files with term of your choice 
       using 'Search Programmes'.
    4. Open the returned files with the 'Open Files' button 
       [will open all associated search PDFs].
    5. Use 'Well Search' to return a list of all Well 
       references in the Programmes.

    The program will save and time/date-stamp all term & well 
    searches in a .txt format.

    All Programmes and converted files are stored in C:\OGA\Programmes
    All Term Search data is stored in C:\OGA\Reports
    All Well Search data is stored in C:\OGA\Wells

    The program is written with an easygui interface.

    For support contact john.davidson.ctr@hotmail.co.uk"""
    readmepg = gui.codebox(rdmsg, rdtitle, rdtext)

def rlicence():
    """Parameters
       ----------
       None
           
       Returns
       -------
       rlicpg
           Gui object containing licence information
       """

    rlicmsg = """
    py-decom version 0.1 [MVP]
    Copyright (c) 2019, John Edward Davidson
    All rights reserved"""
    rlictitle = "py-decom"
    rlictxt = """
    Redistribution and use in source and binary forms, with 
    or without modification, are permitted provided that the 
    following conditions are met:

    1. Redistributions of source code must retain the above 
       copyright notice, this list of conditions and the following 
       disclaimer.

    2. Redistributions in binary form must reproduce the above 
       copyright notice, this list of conditions and the 
       following disclaimer in the documentation and/or other 
       materials provided with the distribution.

    THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDER AND CONTRIBUTOR 
    "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT 
    LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS 
    FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE 
    COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, 
    INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, 
    BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; 
    LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER 
    CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT 
    LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN 
    ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF 
    THE POSSIBILITY OF SUCH DAMAGE.
    
    Contains information provided by the OGA.
    Please refer to the OGA Open User Licence which can be accessed here:
    https://www.ogauthority.co.uk/media/5850/oga-open-user-licence_210619v2.pdf"""
    rlicpg = gui.codebox(rlicmsg, rlictitle, rlictxt)
    
rootmenu()