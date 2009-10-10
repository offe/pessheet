from distutils.core import setup
import py2exe
import sys
import shutil
import os

from pessheet import getApplicationName, getApplicationVersion

appName = getApplicationName()
appVer = getApplicationVersion()
distdir = '%s_%s.win32' % (appName, appVer)

origIsSystemDLL = py2exe.build_exe.isSystemDLL
def isSystemDLL(pathname):
    if os.path.basename(pathname).lower() in ['msvcp71.dll']:
        return 0
    return origIsSystemDLL(pathname)
py2exe.build_exe.isSystemDLL = isSystemDLL

# no arguments
if len(sys.argv) == 1:
    sys.argv.append("py2exe")

def compile(pyName):
    OPTIONS = {"py2exe": \
               {"compressed": 1,
                "optimize": 0,
                "bundle_files": 1,
                'dist_dir': distdir,
               }
              }
    ZIPFILE = None

    console = pyName.endswith('.py')
    #console = True

    appinfo = {'script': pyName, 'icon_resources': [(0, 'resources/pessheet.ico')]}

    setup_args = dict(
            options=OPTIONS,
            zipfile=ZIPFILE,
    )

    if console:
        setup_args['console'] = [appinfo]
    else:
        setup_args['windows'] = [appinfo]
    setup(**setup_args)
    os.remove('%s/w9xpopen.exe' % distdir)
    shutil.copytree('resources', '%s/resources' % distdir)
    shutil.copytree('examples', '%s/examples' % distdir)

try:
    shutil.rmtree(distdir)
except WindowsError, e:
    pass

compile('pessheet.pyw')
shutil.rmtree('build')

