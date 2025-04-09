# -*- coding: utf-8 -*-
"""
Created on Wed Apr  9 16:01:43 2025

@author: BENJAMIN
"""

from setuptools import setup

APP = ['app_test_V2.py']
OPTIONS = {
    'argv_emulation': True,
    'packages': [],  # Si tu utilises des packages externes
}

setup(
    app=APP,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
