#! /usr/bin/env python

import os

TEST_SUPPORT_DIR = os.path.abspath(os.path.dirname(__file__))
TEST_DIR = os.path.abspath(os.path.dirname(TEST_SUPPORT_DIR))
TEST_DATA_DIR = os.path.join(TEST_DIR, "data")
TEST_OUTPUT_DIR = os.path.join(TEST_DIR, "output")
PACKAGE_DIR = os.path.abspath(os.path.dirname(TEST_DIR))
BASE_DIR = os.path.abspath(os.path.dirname(PACKAGE_DIR))
SCRIPTS_DIR = os.path.join(BASE_DIR, "scripts")

def data_path(filename=""):
    return os.path.join(TEST_DATA_DIR, filename)

def data_stream(filename):
    return open(data_path(filename), 'rU')

def scripts_path(filename=""):
    return os.path.join(SCRIPTS_DIR, filename)

def test_path(filename=""):
    return os.path.join(TEST_DIR, filename)

def output_path(filename=""):
    return os.path.join(TEST_OUTPUT_DIR, filename)

def output_stream(filename):
    return open(output_path(filename), 'w')
