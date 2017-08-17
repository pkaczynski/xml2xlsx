# -*- coding: utf-8 -*-
from setuptools import setup
from codecs import open
from os import path

here = path.abspath(path.dirname(__file__))

with open(path.join(here, 'README.rst'), encoding='utf-8') as f:
    long_description = f.read()

setup(
    name='xml2xlsx',
    version='1.0.0',
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
    ],
    keywords='xml lxml xlsx development',
    packages=['xml2xlsx'],
    install_requires=['lxml>=3.6', 'openpyxl>=2.4.7,<2.5', 'six>=1.10'],
    entry_points={
        'console_scripts': ['xml2xlsx=xml2xlsx.command_line:main'],
    },
    zip_safe=False,
)
