import pytest
import os

pytest.main(['-s', 'test_case.py', '--alluredir', 'temp'])
os.system('allure generate ./temp -o ./report --clean')
