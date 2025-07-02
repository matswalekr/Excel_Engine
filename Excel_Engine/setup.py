from setuptools import setup, find_packages

setup(
    name="Excel_Engine",  # Replace with your module name
    version="0.1.0",
    packages=find_packages(),
    install_requires=[
        "pandas",
        "pycel",
        "openpyxl"
    ],
    description="A Python module to interact with wrds to query financial data",
    author="Mats Walker",
    author_email="matswalker2@gmail.com",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
)