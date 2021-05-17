# -*- coding: utf-8 -*-

from setuptools import setup, find_namespace_packages
from codecs import open
import os
import re

root_dir = os.path.abspath(os.path.dirname(__file__))

def filetext(path):
    with open(os.path.join(root_dir, path), "r", encoding="utf-8") as fi:
        text = fi.read()
    return text

def getversion(text):
    m = re.search("__version__\\s+=\\s+'([^']+)'", text)
    if m:
        return m.group(1)
    raise ValueError("バージョン番号がありません")
        
#
#
#
package_name = "python-xlsx-xtended"

version = getversion(filetext("xlsxx/__init__.py"))
license = filetext("LICENSE")
requirements = [x.strip() for x in filetext("REQUIREMENTS.txt").splitlines()]
test_requirements = [x.strip() for x in filetext("TEST-REQUIREMENTS.txt").splitlines()]
long_description = filetext("README.rst") + "\n" + filetext("HISTORY.rst")


setup(
    name=package_name,
    version=version,
    
    packages=find_namespace_packages(exclude=['tests', 'tests.*']),
    package_data={'docxx': ['templates/*.xml', 'templates/*.docx']},
    
    license=license,
    
    install_requires=requirements,
    tests_require=test_requirements,
    test_suite="tests",
    
    author='Goro Sakata',
    author_email='gorosakata@ya.ru',
    url='',
    
    description='Create and update Microsoft Excel .xlsx files.',
    long_description=long_description,
    keywords='xlsx office openxml word',

    classifiers=[
        'Development Status :: 1 - Planning',
        'Environment :: Console',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
        'Programming Language :: Python',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
        'Programming Language :: Python :: 3.6',
        'Topic :: Office/Business :: Office Suites',
        'Topic :: Software Development :: Libraries :: Python Modules',
    ],
)




