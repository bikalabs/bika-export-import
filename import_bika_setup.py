import argparse
import os
import sys
import tempfile
import zipfile
import shutil
import pdb
import traceback
import sys

from Products.CMFPlone.factory import _DEFAULT_PROFILE
from Products.CMFPlone.factory import addPloneSite
from AccessControl.SecurityManagement import newSecurityManager
from bika.lims.catalog import getCatalog
from Products.Archetypes import Field
from Products.CMFCore.interfaces import ITypeInformation
from Products.CMFCore.utils import getToolByName
import openpyxl


def excepthook(typ, value, tb):
    traceback.print_exception(typ, value, tb)
    pdb.pm()


sys.excepthook = excepthook

# If creating a new Plone site:
default_profiles = [
    'plonetheme.classic:default',
    'plonetheme.sunburst:default',
    'plone.app.caching:default',
    'bika.lims:default',
]


class Main:
    def __init__(self, args):
        self.args = args

    def __call__(self):
        """Export entire bika site
        """
        # pose as user
        self.user = app.acl_users.getUserById(self.args.username)
        newSecurityManager(None, self.user)
        # get or create portal object
        try:
            self.portal = app.unrestrictedTraverse(self.args.sitepath)
        except KeyError:
            profiles = default_profiles
            if self.args.profiles:
                profiles += list(self.args.profiles)
            addPloneSite(
                app,
                self.args.sitepath,
                title=self.args.title,
                profile_id=_DEFAULT_PROFILE,
                extension_ids=profiles,
                setup_content=True,
                default_language=self.args.language
            )
            self.portal = app.unrestrictedTraverse(self.args.sitepath)
        # Extract zipfile
        self.tempdir = tempfile.mkdtemp()
        zf = zipfile.ZipFile(self.args.inputfile, 'r')
        zf.extractall(self.tempdir)
        # Open workbook
        self.wb = openpyxl.load_workbook(
            os.path.join(self.tempdir, 'setupdata.xlsx'))
        # Import
        self.import_laboratory()
        self.import_bika_setup()
        for portal_type in self.get_portal_type_sheet_names():
            self.import_portal_type(portal_type)
        # Remove tempdir
        shutil.rmtree(self.tempdir)

    def get_portal_type_sheet_names(self):
        typestool = getToolByName(self.portal, 'portal_types')
        for sheetname in self.wb.get_sheet_names():
            import pdb;pdb.set_trace()

    def mutate(self, field, value):
        return value

    def import_laboratory(self):
        instance = self.portal.bika_setup.laboratory
        schema = instance.schema
        ws = self.wb['Laboratory']
        for row in ws.rows:
            field = schema[row[0].value]
            value = self.mutate(field, row[1].value)
            import pdb;pdb.set_trace()
            field.set(instance, value)

    def import_bika_setup(self):
        instance = self.portal.bika_setup
        schema = instance.schema
        ws = self.wb['BikaSetup']
        for row in ws.rows:
            field = schema[row[0].value]
            value = self.mutate(field, row[1].value)
            import pdb;pdb.set_trace()
            field.set(instance, value)

    def import_portal_type(self):
        import pdb;

        pdb.set_trace()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Import bika setupdata created by export_bika_setup.py',
        epilog='This script is meant to be run with zopepy or bin/instance. See'
               ' http://docs.plone.org/develop/plone/misc/commandline.html for'
               ' details.'
    )
    parser.add_argument(
        '-s',
        dest='sitepath',
        required=True,
        help='full path to Plone site root.  Site will be created if it does'
             ' not already exist.')
    parser.add_argument(
        '-i',
        dest='inputfile',
        required=True,
        help='input zip file, created by the export script.')
    parser.add_argument(
        '-u',
        dest='username',
        default='admin',
        help='zope admin username (default: admin)')
    parser.add_argument(
        '-t',
        dest='title',
        help='If a new Plone site is created, this specifies the site Title.'),
    parser.add_argument(
        '-l',
        dest='language',
        default='en',
        help='If a new Plone site is created, this is the site language.'
             ' (default: en)')
    parser.add_argument(
        '-p',
        dest='profiles',
        action='append',
        help='If a new Plone site is created, this option may be used to'
             ' specify additional profiles to be activated.'),
    args, unknown = parser.parse_known_args()

    main = Main(args)
    main()
