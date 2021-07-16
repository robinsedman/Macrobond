import setuptools
from macrobond import __version__
from macrobond import __name__

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name=__name__,
    version=__version__,
    author="Robin Sedman",
    author_email="robin.sedman@gmail.com",
    description="Pandas wrapper for the Macrobond API",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/robinsedman/Macrobond",
    packages=setuptools.find_packages(),
    classifiers=["Programming Language :: Python :: 3",
                 "Programming Language :: Python :: 3.7",
                 "Programming Language :: Python :: 3.8",
                 "Intended Audience :: End Users/Desktop",
                 "Intended Audience :: Financial and Insurance Industry",
                 "Natural Language :: English",
                 "Development Status :: 4 - Beta",
                 "Environment :: Win32 (MS Windows)",
                 "License :: OSI Approved :: MIT License",
                 "Operating System :: Microsoft :: Windows"],
)