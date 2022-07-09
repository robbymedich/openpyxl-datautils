""" Use setup.cfg for main package configuration """
from setuptools import setup, find_packages


setup(
    package_dir={"": "src"},
    packages=find_packages(where="src"),
    python_requires=">=3.6",
)
