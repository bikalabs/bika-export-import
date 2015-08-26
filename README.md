# bika-export-import
Scripts for exporting/importing bika.lims field values to/from a zip file

## export

campbell@campbell-np600:~/Plone/zeocluster$ bin/client1 run export_bika_setup.py --help
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
