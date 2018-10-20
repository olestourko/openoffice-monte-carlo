from setuptools import setup, find_packages

setup(
    name='open-office-monte-carlo',
    version='1.0',
    description='Monte Carlo simulation for OpenOffice calc',
    packages=find_packages(exclude=["tests*"]),
    entry_points={
        'console_scripts': [
            'monte-carlo = src.monte_carlo:main'
        ]
    }
)