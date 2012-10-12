#! /usr/bin/env python

from setuptools import setup, find_packages
import sys, os
import glob

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
SCRIPTS = glob.glob(os.path.join(BASE_DIR, 'scripts', '*.py'))

version = '0.1'

setup(name='GooPy',
      version=version,
      description="A python package manipulating Google docs",
      long_description="""\
A python package for manipulating Google docs""",
      classifiers=[], # Get strings from http://pypi.python.org/pypi?%3Aaction=list_classifiers
      keywords='Productiviy, Google documents',
      author='Jamie Oaks',
      author_email='joaks1@gmail.com',
      url='',
      license='GPL',
      packages=find_packages(exclude=['ez_setup', 'examples', 'tests']),
      scripts=SCRIPTS,
      include_package_data=True,
      zip_safe=True,
      test_suite="goopy.test",
      install_requires=[
          'gdata'
      ],
      entry_points="""
      # -*- Entry points: -*-
      """,
      )
