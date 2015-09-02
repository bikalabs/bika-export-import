# bika-export-import
Scripts for exporting/importing bika.lims field values to/from a zip file.  The zip file contains setupdata.xlsx, and a collection of files which match the contents of File and Image fields.

> This requires a recent openpyxl; however bika.lims still pins openpyxl to 1.5.8 in setup.py.  This is no longer required and before using these scripts, this version pin should be removed, and buildout re-run.

## export

    $ bin/client1 run export_bika_setup.py --help
    usage: interpreter [-h] [-s SITEPATH] [-u USERNAME] [-o OUTPUTFILE]

    Export bika_setup into an Open XML (XLSX) workbook

    optional arguments:
      -h, --help     show this help message and exit
      -s SITEPATH    full path to site root (default: Plone)
      -u USERNAME    zope admin username (default: admin)
      -o OUTPUTFILE  output zip file name (default: SITEPATH.zip)

    This script is meant to be run with zopepy or bin/instance. See
    http://docs.plone.org/develop/plone/misc/commandline.html for details.

## import

    $ bin/client1 run import_bika_setup.py --help

    usage: interpreter [-h] -s SITEPATH -i INPUTFILE [-u USERNAME] [-t TITLE]
                       [-l LANGUAGE] [-p PROFILES]
    
    Import bika setupdata created by export_bika_setup.py
    
    optional arguments:
      -h, --help    show this help message and exit
      -s SITEPATH   full path to Plone site root. Site will be created if it does
                    not already exist.
      -i INPUTFILE  input zip file, created by the export script.
      -u USERNAME   zope admin username (default: admin)
      -t TITLE      If a new Plone site is created, this specifies the site Title.
      -l LANGUAGE   If a new Plone site is created, this is the site language.
                    (default: en)
      -p PROFILES   If a new Plone site is created, this option may be used to
                    specify additional profiles to be activated.
    
    This script is meant to be run with zopepy or bin/instance. See
    http://docs.plone.org/develop/plone/misc/commandline.html for details.
