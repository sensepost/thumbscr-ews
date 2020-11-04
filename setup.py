import os

from setuptools import setup

# here - where we are.
here = os.path.abspath(os.path.dirname(__file__))

# read the package requirements for install_requires
with open(os.path.join(here, 'requirements.txt'), 'r') as f:
    requirements = f.readlines()

# setup!
setup(
    name='thumbscr-ews',
    description='A wrapper around the amazing exchangelib to do some common EWS operations.',
    license='GPL v3',
    packages=['thumbscrews'],
    install_requires=requirements,
    python_requires='>=3.6',
    classifiers=[
        'Operating System :: OS Independent',
        'Natural Language :: English',
        'Programming Language :: Python :: 3 :: Only',
    ],
    entry_points={
        'console_scripts': [
            'thumbscr-ews = thumbscrews.cli:cli',
        ],
    },
)
