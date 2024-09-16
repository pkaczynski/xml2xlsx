# -*- coding: utf-8 -*-
from codecs import open
from os import path

from setuptools import setup

here = path.abspath(path.dirname(__file__))

with open(path.join(here, 'README.rst'), encoding='utf-8') as f:
    long_description = f.read()

setup(
    name='xml2xlsx',
    version='1.0.2',
    description='XML to XLSX converter',
    long_description=long_description,
    url='https://github.com/marrog/xml2xlsx',
    author='Piotr Kaczyński',
    author_email='pkaczyns@gmail.com',
    license='MIT',
    classifiers=[
        'Development Status :: 5 - Production/Stable',
        'Intended Audience :: Developers',
        'Topic :: Software Development :: Build Tools',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.5',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
    ],
    keywords='xml lxml xlsx development',
    packages=['xml2xlsx'],
    install_requires=['lxml>=3.6', 'openpyxl>=2.5.0', 'six>=1.10'],
    test_requires=['nose', 'tox', 'coverage'],
    entry_points={
        'console_scripts': ['xml2xlsx=xml2xlsx.command_line:main'],
    },
    zip_safe=False,
)
