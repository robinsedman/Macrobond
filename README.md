# Macrobond
[![Build Status](http://img.shields.io/travis/badges/badgerbadgerbadger.svg?style=flat-square)](https://travis-ci.org/badges/badgerbadgerbadger)
[![Coverage Status](http://img.shields.io/coveralls/badges/badgerbadgerbadger.svg?style=flat-square)](https://coveralls.io/r/badges/badgerbadgerbadger) 

A Pandas wrapper for the Macrobond API. Some of the methods require a Data+ license (https://www.macrobond.com/desktop-solutions)

Documentation from Macrobond can be found here:
* https://help.macrobond.com/technical-information/the-macrobond-api-for-python/
* https://help.macrobond.com/technical-information/common-metadata/

# Installation
The [macrobond-api-constants](https://pypi.org/project/macrobond-api-constants/) package is required:

    pip install macrobond_api_constants
    
Then simply install the latest version Python package index (PyPI) as follows:
    
    pip install --upgrade macrobond

# A few examples

    from macrobond import c_macrobond
    mb = c_macrobond.Macrobond()

### Fetch single time series
    
	df = mb.FetchOneSeries(ticker='usnaac0057')

### Fetch several time series

    df = mb.FetchSeries(ticker_list=['usnaac0057', 'senaac0067'])

### Time series and first revision

    df = mb.FetchOneSeriesWithRevisions(ticker='usnaac0057')

### Fetch several time series in the same currency (default: USD)

    df = mb.CreateUnifiedSeriesRequst(ticker_list=['usnaac0057', 'senaac0067'])

### Get all tickers for a concept

    ticker_list = mb.CreateSearchQuery(concept_filter='gdp_total')

#### The method above has a few more possibilities:
* Possible to narrow to specific regions/countries
* Possible to create a free text search query

### Get summary of several time series

    df = mb.m_get_full_info(ticker_list=['usnaac0057', 'senaac0067'])

### Get the concept of a time series

    short_concept, long_concept = mb.m_series_concept(ticker='usnaac0057')

### Extract frequency for time series

    freq = mb.m_get_frequency(ticker='usnaac0057')

### Get title/description for a time series

    title = mb.m_get_title(ticker='usnaac0057')

### Get previous/next release date if it exist

    next_release = mb.m_release_date(ticker='usnaac0057', date_option='Next')

### Extract metadata for a series

    currency = mb.m_get_metadata(ticker='usnaac0057', option='Currency')

# Disclaimer
Kindly note that this is an unofficial wrapper for the Macrobond API and the underlying structure could be subject to change at any point in time.

# LICENSE
MIT license. See the LICENSE file for details.

# RESPONSIBILITIES

The author of this software is not responsible for any indirect damages (foreseeable or unforeseeable), such as, if necessary, loss or alteration of or fraudulent access to data, accidental transmission of viruses or of any other harmful element, loss of profits or opportunities, the cost of replacement goods and services or the attitude and behavior of a third party.
