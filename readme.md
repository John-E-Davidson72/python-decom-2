Program for downloading and searching OGA Decom programmes.
Rel. 0.0 MVP for testing and development.

py-decom

What this program does:

1. Finds all of the approved and under consideration Decommissioning programmes on the OGA site.
2. Saves them all to your hard drive.
3. Converts them to a fully searchable format.
4. Searches them with user input search terms and outputs a results file.
5. Opens all the search files on user request.
6. Searches all the downloaded files for recorded well numbers.

How to install:

This software is distributed using pyInstaller packaging.
All libraries and dependencies have been packaged into a single EXE file in dist folder.

How to set up the dev environment:

See requirements.txt

Example Usage: 

1. Download all the current Decom Programmes using the 'Download Programmes' button.
2. Convert the files to searchable .txt format using the 'Convert Files' button.
3. Search the Programme Files using 'Search Programmes'.
4. Open the returned files with the 'Open Files' button (will open all associated search PDFs).
5. Use 'Well Search' to return a list of all Well references in the Programmes.

The program will save and time / date-stamp all term & well searches in a .txt format.

All PDF Programmes and converted .txt files are stored in C:\OGA\Programmes
All Term Search .txt data is stored in C:\OGA\Reports
All Well Search .txt data is stored in C:\OGA\Wells

Converted files are available as a Zip File - this speeds up the conversion process as only
the new / updated Programmes that have been added to the OGA site need to be converted.
An up-to-date file can be made available on request.

Users will need to Unzip the files to the C:\OGA\Programmes folder.

This program is written with an easygui interface.

For support contact john.davidson.ctr@hotmail.co.uk








