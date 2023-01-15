import os
import pytest
import sys

sys.path.append('../src/cariad')
from cariad import parse_cli_args

# Testing: cariad.parse_cli_args()
valid_directory_example = os.getcwd()

def test_no_args():
    sys.argv = ['cariad.py']
    with pytest.raises(SystemExit):
        assert parse_cli_args()


def test_only_one_arg():
    sys.argv = ['cariad.py', 'reference.xlsx']
    with pytest.raises(SystemExit):
        assert parse_cli_args()


def test_args_in_wrong_order():
    sys.argv = ['cariad.py', 'reference.xlsx', valid_directory_example]
    with pytest.raises(SystemExit):
        assert parse_cli_args()


def test_nonexistent_target_directory():
    sys.argv = ['cariad.py', f'{valid_directory_example}\\invalid_subfolder', 'nosuchfile.xlsx']
    with pytest.raises(SystemExit):
        assert parse_cli_args()


def test_nonexistent_reference_file():
    sys.argv = ['cariad.py', valid_directory_example, 'nosuchfile.xlsx']
    with pytest.raises(SystemExit):
        assert parse_cli_args()


def test_target_directory_includes_space():
    ...

