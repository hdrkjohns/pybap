import sys
from setuptools import setup, find_packages

if(sys.version_info < (3,0)):
    execfile('pybap/version.py')
else:
    exec(open('pybap/version.py').read())

setup(
    name = "pybap",
    version = __version__,
    packages = find_packages('.',exclude=["*tests"]),
    scripts = [],
    install_requires = [ \
            'numpy' \
            ,'pandas' \
            ,'shapely' \
            ,'python-docx'
            ],
    extras_require={ \
    },
    package_data = {
            '': ['*.txt', '*.rst'], \
            # And include any *.msg files found in the 'hello' package, too:
            'pybap': ['static/*','templates/*'], \
    },
    include_package_data=True,
    dependency_links = [ \
                'http://pygame.org/download.shtml' \
                , 'http://www.wxpython.org/download.php' \
                ],
    # metadata for upload to PyPI
    author = "kevin johns",
    author_email = "kevin.johns@hdrinc.com",
    description = "contains a set of utilities to assist with BAP data",
    license = "PSF",
    keywords = "bap pybap",
    url = None,   # project home page, if any

    # could also include long_description, download_url, classifiers, etc.
)
