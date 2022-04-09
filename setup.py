from distutils.core import setup
import py2exe
import pandas
import xlsxwriter
from setuptools import setup

packages = ["pandas", "xlsxwriter", "tkinter", "os", "sys"]
setup(
    windows=[
        {
            "script": "main.py",
            'icon_resources': [(0, 'launcher.ico')],
            "dest_base": "Konvertor"
        }
    ],
    options={"py2exe":
                 {"packages": packages},
             },
    name='TabDelimitedTextToExcel',
    version='1.0',
    packages=[],
    url='https://github.com/andjelkadzida/TabDelimitedTextToExcel',
    license='',
    author='Andjelka Dzida',
    author_email='andjelkadzida@gmail.com',
    description='Convertor for xyz, qtt and lst extenstions to Excel files.'
)
