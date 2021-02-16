import setuptools
from macrobond import __version__
from macrobond import __name__

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name=__name__,
    version=__version__,
    author="Robin Sedman",
    author_email="robin.sedman@seb.se",
    description="Pandas wrapper for the Macrobond API",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://sebgroup.com",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
)