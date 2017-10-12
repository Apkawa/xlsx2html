# -*- coding: utf-8 -*-
# !/usr/bin/env python
import os
from setuptools import setup, find_packages

__doc__ = """A simple export from xlsx format to html tables with keep cell formatting"""

project_name = 'xlsx2html'

version = '0.1.4'

ROOT = os.path.dirname(__file__)


def read(fname):
    return open(os.path.join(ROOT, fname)).read()


try:
    import pypandoc

    long_description = pypandoc.convert('README.md', 'rst')
except ImportError:
    long_description = read('README.md')

setup(
    name=project_name,
    version=version,
    description=__doc__,
    long_description=long_description,
    url="https://github.com/Apkawa/xlsx2html",
    author="Arkadii Ivanov",
    author_email='apkawa@gmail.com',
    packages=[package for package in find_packages() if package.startswith(project_name)],
    install_requires=['six', 'openpyxl>=2.4.8,<3', 'babel>=2.3.4,<3'],
    python_requires='>=2.7, !=3.2.*, <4',
    zip_safe=False,
    include_package_data=True,
    license='MIT',
    keywords='converter xlsx html',
    classifiers=[
        'Development Status :: 3 - Alpha',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3.3',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
        'Programming Language :: Python :: 3.6',
        'Topic :: Text Processing',
        'Topic :: Text Processing :: Markup :: HTML',
    ],
)
