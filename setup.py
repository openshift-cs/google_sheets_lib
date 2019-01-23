#!/usr/bin/env python
# -*- coding: utf-8 -*-

import io
import os

from setuptools import setup, find_packages

# Package meta-data.
NAME = 'google_sheets_lib'
DESCRIPTION = 'An opinionated library for accessing Google Sheets documents'
SOURCE_URL = 'https://github.com/openshift-cs/google_sheets_lib'
DOCS_URL = 'https://google_sheets_lib.readthedocs.io'
EMAIL = 'wgordon@redhat.com'
AUTHOR = 'Will Gordon'
REQUIRES_PYTHON = '>=3.6.0'
VERSION = None

root_dir = os.path.abspath(os.path.dirname(__file__))

# Import the README and use it as the long-description.
# Note: this will only work if 'README.md' is present in your MANIFEST.in file!
try:
    with io.open(os.path.join(root_dir, 'README.md'), encoding='utf-8') as f:
        long_description = '\n' + f.read()
except FileNotFoundError:
    long_description = DESCRIPTION

about = {}
if not VERSION:
    with open(os.path.join(root_dir, NAME, '__version__.py')) as version_file:
        exec(version_file.read(), about)
else:
    about['__version__'] = VERSION


setup(
    name=NAME,
    version=about['__version__'],
    description=DESCRIPTION,
    long_description=long_description,
    long_description_content_type='text/markdown',
    author=AUTHOR,
    author_email=EMAIL,
    python_requires=REQUIRES_PYTHON,
    url=SOURCE_URL,
    packages=find_packages(exclude=('tests',)),
    install_requires=[
        'oauth2client',
        'pygsheets>=2'
    ],
    setup_requires=['pytest-runner'],
    tests_require=['pytest', 'pytest-dependency', 'pytest-flake8'],
    include_package_data=True,
    license='MIT',
    classifiers=[
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3 :: Only',
        'Programming Language :: Python :: Implementation :: CPython',
    ],
    project_urls={
        'Documentation': DOCS_URL,
        'Source': SOURCE_URL,
    },
    command_options={
        'build_sphinx': {
            'project': ('setup.py', NAME),
            'version': ('setup.py', about['__version__']),
            'release': ('setup.py', about['__version__'])
        }
    },
)
