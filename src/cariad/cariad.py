'''Main module. Launch cariad from the cli.'''

import argparse
import os
import sys

import run
import variables

def main():
    variables.target_directory, variables.reference_spreadsheet = parse_cli_args()
    run.run_cariad(variables.target_directory, variables.reference_spreadsheet, 'cli')


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

    reference_path = os.path.join(target, reference)

    if not os.path.isfile(reference_path):
        raise sys.exit('Reference spreadsheet does not exist in target directory')

    return target, reference_path


if __name__ == '__main__':
    main()