# -*- coding: utf-8 -*-
"""A setuptools based setup module.
See:
https://packaging.python.org/en/latest/distributing.html
https://github.com/pypa/sampleproject
"""

# Always prefer setuptools over distutils
from setuptools import setup, find_packages
# To use a consistent encoding
from codecs import open
from os import path

here = path.abspath(path.dirname(__file__))

# Get the long description from the README file
with open(path.join(here, 'README.rst'), encoding='utf-8') as f:
    long_description = f.read()

setup(
    name='xml2xlsx',
    version='0.2a',
    description='XML to XLSX converter',
    long_description=long_description,
    url='https://github.com/marrog/xml2xlsx',
    author='Piotr Kaczyński',
    author_email='pkaczyns@gmail.com',
    license='MIT',
    classifiers=[
        # How mature is this project? Common values are
        #   3 - Alpha
        #   4 - Beta
        #   5 - Production/Stable
        'Development Status :: 3 - Alpha',
        # Indicate who your project is intended for
        'Intended Audience :: Developers',
        'Topic :: Software Development :: Build Tools',
        # Pick your license as you wish (should match "license" above)
        'License :: OSI Approved :: MIT License',
        # Specify the Python versions you support here. In particular, ensure
        # that you indicate whether you support Python 2, Python 3 or both.
        'Programming Language :: Python :: 2.7',
    ],
    keywords='xml lxml xlsx development',
    packages=['xml2xlsx'],
    install_requires=['lxml', 'openpyxl'],
    # pip to create the appropriate form of executable for the target platform.
    entry_points={
        'console_scripts': ['xml2xlsx=xml2xlsx.command_line:main'],
    },
    zip_safe=False,
)
