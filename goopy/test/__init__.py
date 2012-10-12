#! /usr/bin/env python

import unittest
import os
import sys
import glob
from goopy.test.support import package_paths

def get_test_file_names():
    test_file_paths = glob.glob(os.path.join(package_paths.TEST_DIR, "test_*"))
    tests = []
    for path in test_file_paths:
        test_file_name = os.path.basename(path).split(".")[0]
        test = ".".join(["goopy.test", test_file_name])
        sys.stdout.write("%s\n" % test)
        tests.append(test)
    return tests

def get_unittest_suite():
    test_file_names = get_test_file_names()
    tests = unittest.defaultTestLoader.loadTestsFromNames(test_file_names)
    suite = unittest.TestSuite(tests)
    return suite

def run():
    test_runner = unittest.TextTestRunner()
    test_runner.run(get_unittest_suite())

if __name__ == '__main__':
    run()
