from distutils.core import setup, Extension

module1 = Extension('orkoswmi',
                    libraries = ['kernel32', 'user32', 'gdi32', 'winspool', 'comdlg32', 'advapi32', 'shell32', 'ole32', 'oleaut32', 'uuid', 'odbc32', 'odbccp32'],
                    include_dirs = [],
                    library_dirs = [],
                    sources = ['orkoswmi.cpp'])

setup (name = 'orkoswmi',
       version = '0.1',
       description = 'This package is designed to allow native WMI calls from within python',
       ext_modules = [module1])