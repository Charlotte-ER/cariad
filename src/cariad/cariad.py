'''Main module. Run cariad from the cli.'''

import argparse
import os
import sys

import globals

def main():
    target, reference = parse_cli_args()

    # Usage checks

    # Update global variables
    globals.target_directory = target
    globals.reference_spreadsheet = reference


def parse_cli_args():
    '''Parse command line arguments, rejecting invalid usage.'''
    parser = argparse.ArgumentParser()
    parser.add_argument('target', type=str)
    parser.add_argument('reference', type=str)
    args, unknown = parser.parse_known_args()

    target = args.target
    reference = args.reference

    # Handle spaces in target filepath, indicated by unknown arguments
    if unknown:
        unknown.insert(0,reference)
        reference = unknown.pop(-1)
        unknown.insert(0,target)
        target = ' '.join(unknown)

    if not os.path.isdir(target):
        raise sys.exit('Target directory does not exist')

    if not os.path.isfile(os.path.join(target, reference)):
        raise sys.exit('Reference spreadsheet does not exist in target directory')

    return target, reference


if __name__ == '__main__':
    main()