from setuptools import setup, find_packages

setup(
    name='google_spreadsheets',
    version='0.1.0',
    packages=find_packages(),
    install_requires=[
        'requests==2.31.0',
        'oauth2client==4.1.3',
        'google-api-python-client==2.104.0',
        'httplib2==0.22.0',
        'pyasn1',
    ]
)
