import pandas as pd
import pytz
import datetime as dt
import numpy as np
import win32com.client
import macrobond_api_constants.SeriesFrequency
from typing import Tuple

'''
All the different macrobond constants that exist

import macrobond_api_constants.CalendarDateMode
import macrobond_api_constants.CalendarMergeMode
import macrobond_api_constants.MetadataValueType
import macrobond_api_constants.SeriesMissingValueMethod
import macrobond_api_constants.SeriesPartialPeriodsMethod
import macrobond_api_constants.SeriesToHigherFrequencyMethod
import macrobond_api_constants.SeriesToLowerFrequencyMethod
import macrobond_api_constants.SeriesWeekdays
'''


class Macrobond:
	def __init__(self):
		# Initiate win32com connection to macrobond
		c = win32com.client.Dispatch('Macrobond.Connection')

		# Create connection to the database
		db = c.Database

		# Create default attribute with all regions
		region_map, region_map_inverse = self.f_region_map()
		region_list = list(region_map.keys())

		# Save MacrobondDatabase attribute
		self.mbdb = db
		self.region_list_all = region_list

	def FetchOneSeries(self, ticker: str) -> pd.DataFrame:
		"""
		Fetch One timeseries from Macrobond
		In this method we gather a few examples of attributes that could be extracted
		"""
		series = self.mbdb.FetchOneSeries(ticker)

		# Assert all is well
		try:
			assert series.IsError is False
		except AssertionError:
			print(f'Error: {series.ErrorMessage}')
			return pd.DataFrame()

		# Some text info for the series
		name: str = series.Name
		title: str = series.Title

		# Tuple with True / False
		forecast_flags = series.ForecastFlags

		# Typical observations per year
		obs_per_year: float = series.TypicalObservationCountPerYear

		# Timeseries frequency
		freq: macrobond_api_constants.SeriesFrequency = series.Frequency

		# Convert some dates
		p_start_dates = pd.to_datetime([date.strftime('%Y-%m-%d') for date in series.DatesAtStartOfPeriod])
		p_end_dates = pd.to_datetime([date.strftime('%Y-%m-%d') for date in series.DatesAtEndOfPeriod])

		# Extract start and end date
		start_date = series.StartDate.date()
		end_date = series.EndDate.date()

		# Dates above contain additional info (time & timezone)
		start_date_time = series.EndDate.time()
		end_date_time = series.EndDate.time()
		start_date_tz_info = series.StartDate.tzinfo
		end_date_tz_info = series.EndDate.tzinfo

		# Generate pd.DataFrame that we return
		df = pd.DataFrame(series.Values, index=p_end_dates)
		df.columns = [series.Name]

		return df

	def FetchSeries(self, ticker_list: [str]) -> pd.DataFrame:
		"""
		Fetch several series and return a dataframe
		"""

		# Assert type
		try:
			assert type(ticker_list) is list
		except AssertionError as ae:
			print(f'Input must be a list')
			raise ae

		# Fetch all the series
		series = self.mbdb.FetchSeries(ticker_list)

		# Convert it to a pd.DataFrame
		df = self.m_series_tuple_to_df(ticker_list=ticker_list, series=series)

		return df

	def FetchOneSeriesWithRevisions(self, ticker: str) -> pd.DataFrame:
		"""
		We only care about the original series & first revision in this function
		"""

		# Download series
		series = self.mbdb.FetchOneSeriesWithRevisions(ticker)

		# Check that revisions exist
		try:
			assert series.HasRevisions
		except AssertionError:
			print(f'No revisions exist for {ticker}. Error message: {series.ErrorMessage}')
			return pd.DataFrame()

		# Check that not everything in revision 1 is nan
		try:
			assert not np.all(np.isnan(series.GetNthRelease(1).Values))
		except AssertionError:
			print(f'No revisions exist for {ticker}. Error message: {series.ErrorMessage}')
			return pd.DataFrame()

		# Extract series
		s0 = series.GetNthRelease(0)
		s1 = series.GetNthRelease(1)

		# Convert series to pd.Series. Original and 1st revision
		x0 = self.f_unpack_series(s0)
		x1 = self.f_unpack_series(s1)

		# Create list of tuples with column names
		column_list = [(s0.Name, 'Rev0'), (s1.Name, 'Rev1')]

		# Pre-allocate pd.DataFrame()
		df = pd.DataFrame()

		# Insert values and index
		df[column_list[0]] = x0.values
		df[column_list[1]] = x1.values
		df['index'] = x0.index

		# Set index of df
		df = df.set_index(keys='index', drop=True, append=False)

		# Convert to two-level columns
		df.columns = pd.MultiIndex.from_tuples(df.columns)

		# Check how many revisions we have (not used at the moment)
		n = 0
		while True:
			x = pd.Series(series.GetNthRelease(n).Values)

			# If all in the series is nan then we assume we should break
			if np.all(np.isnan(x)):
				print(f'We have: {n - 1} revisions')
				break
			n += 1

		'''
		Functions / attributes that could be interesting to further develop

		series.GetCompleteHistory(): Get list of all revisions.
		series.StoresRevisions: If True, then the series stores revision history. This can be True while HasRevisions is False if no revisions have yet been recorded.
		series.TimeOfLastRevision: The timestamp of the last revision.
		'''

		return df

	def CreateUnifiedSeriesRequst(self, ticker_list: list, **kwargs) -> pd.DataFrame:
		"""
		Function that e.g. can extract several series in one currency
		This can be expanded at time where we limit series to start date and end date etc
		https://help.macrobond.com/technical-information/the-macrobond-api-for-python/#iseriesrequest
		"""

		# Define currency for the request
		# Currency codes that is used in Macrobond: https://www.macrobond.com/currency-list/
		currency = 'USD'

		# Extract kwargs
		for key, val in kwargs.items():
			if key.lower() == 'currency':
				currency = kwargs.get('Currency')
			else:
				raise KeyError(f'Kwargs key: {key} not defined')

		# Create the request
		req = self.mbdb.CreateUnifiedSeriesRequest()

		# Add all tickers to the request
		for tick in ticker_list:
			req.AddSeries(tick)

		# Add currency for the request
		req.Currency = currency

		# Finally fetch the data
		series = self.mbdb.FetchSeries(req)

		# Convert it to pd.DataFrame
		df = self.m_series_tuple_to_df(ticker_list=ticker_list, series=series)

		return df

	def CreateSearchQuery(self, concept_filter: str = 'gdp_total', entity_type_filter: str = 'TimeSeries',
						  **kwargs) -> list:
		"""
		Function to define a query (concept) and then find series that match the concept and return those tickers

		Details on how to create narrow search queries:
		https://help.macrobond.com/tutorials-training/2-finding-data/finding-data-in-search/search-terms-for-more-accurate-results/
		"""

		# Region list, kwargs key: RegionList
		region_map, region_map_inverse = self.f_region_map()
		region_list = list(region_map.keys())

		# Season adjustment, kwargs key: SeasonAdj
		season_adj_tf = False

		# Frequency of time series, kwargs key: Frequency
		frequency_tf = False
		frequency = ''  # E.g. 'weekly'

		# Include Discontinued Series, kwargs key: IncludeDiscontinued
		include_discontinued_tf = False

		# Free text search, kwargs key: FreeText
		free_text_search_tf = False
		free_text_query = ''

		# Extract kwargs
		for key, val in kwargs.items():
			if key.lower() == 'regionlist':
				region_list: list = kwargs.get('RegionList')
			elif key.lower() == 'frequency':
				frequency: str = kwargs.get('Frequency')
				frequency_tf = True
			elif key.lower() == 'seasonadj':
				season_adj_tf: bool = kwargs.get('SeasonAdj')
			elif key.lower() == 'includediscontinued':
				include_discontinued_tf: bool = kwargs.get('IncludeDiscontinued')
			elif key.lower() == 'freetext':
				free_text_search_tf = True
				free_text_query: str = kwargs.get('FreeText')
			else:
				raise KeyError(f'Kwargs key: {key} not defined')

		# Define a search query
		query = self.mbdb.CreateSearchQuery()

		'''
		Entity Type Filter:
		A single string or a vector of strings identifying what type of entities to search for.
		The options are TimeSeries, Release, Source, Index, Security, Region, RegionKey, Exchange and Issuer.

		Typically you want to set this to TimeSeries. If not specified, the search will be made for several entity types.
		'''
		query.SetEntityTypeFilter(entity_type_filter)

		'''
		The RegionKey defines a "concept" (see "Concept & Category" database view in the Macrobond application)
		Example of concepts: gdp_total_sa, markit_prev_manu_pmi, markit_prev_serv_pmi
		'''

		# Either free text search or concept search
		if free_text_search_tf:
			query.Text = free_text_query
		else:
			query.AddAttributeValueFilter("RegionKey", concept_filter)

		'''
        Region filter to only include time series with region chosen
        Example regions in: self.f_region_map()
        '''

		# Region filter
		query.AddAttributeValueFilter("Region", region_list)

		# Frequency filter
		if frequency_tf:
			query.AddAttributeValueFilter("Frequency", frequency)

		# Optionally set seasonlity filter
		if season_adj_tf:
			query.AddAttributeFilter("SeasonAdj")

		# Include discontinued series or not
		query.IncludeDiscontinued = include_discontinued_tf

		# Do the search
		search_result = self.mbdb.Search(query)

		# Extract any entities we found
		result = search_result.Entities

		# Check if result is truncated
		if search_result.isTruncated:
			print(f'Search was truncated. Truncated at {len(result)} entries.')

		# Extract the tickers (names)
		tickers = [s.Name for s in result]

		return tickers

	def m_series_tuple_to_df(self, ticker_list: list, series) -> pd.DataFrame:
		"""
		Method just to convert series request to a pd.DataFrame
		:param ticker_list: list
		:param series:  (<COMObject FetchSeries>, ..., <COMObject FetchSeries>)
		"""

		# First create a series with ALL dates
		date_list = list()
		for i, ticker in enumerate(ticker_list):
			dates = series[i].DatesAtEndOfPeriod

			# Convert all dates to ordinal
			date_list.extend([dt.datetime.toordinal(date) for date in dates])

		# Get unique dates and sort them
		date_list_unique = list(set(date_list))
		date_list_unique.sort(reverse=False)

		# Generate dates from ordinal format to pd.DatetimeIndex
		p_dates = pd.to_datetime([dt.date.fromordinal(date) for date in date_list_unique])

		# Pre-allocate DataFrame
		df = pd.DataFrame(index=p_dates, columns=ticker_list)

		for i, ticker in enumerate(ticker_list):
			unpacked_series = self.f_unpack_series(series[i])

			df[ticker] = unpacked_series

		return df

	def m_get_full_info(self, ticker_list: list) -> pd.DataFrame:
		"""
		Method to gather some data of a list of tickers in a pd.DataFrame
		"""

		next_rel = [self.m_release_date(ticker=tick, date_option='next') for tick in ticker_list]
		prev_rel = [self.m_release_date(ticker=tick, date_option='previous') for tick in ticker_list]
		freq = [self.m_get_frequency(ticker=tick) for tick in ticker_list]
		title = [self.m_get_title(ticker=tick) for tick in ticker_list]
		concept = [self.m_series_concept(ticker=tick) for tick in ticker_list]
		region_short = [self.m_get_metadata(ticker=tick, option='Region') for tick in ticker_list]
		region_full = [self.m_get_region(reg) for reg in region_short]
		currency = [self.m_get_metadata(ticker=tick, option='Currency') for tick in ticker_list]
		database = [self.m_get_metadata(ticker=tick, option='Database') for tick in ticker_list]
		release = [self.m_get_metadata(ticker=tick, option='Release') for tick in ticker_list]

		# Define dictionary then convert it to a dataframe
		d = {'Description': title, 'NextRelease': next_rel, 'PreviousRelease': prev_rel, 'Frequency': freq,
			 'RegionLong': region_full, 'RegionShort': region_short, 'Currency': currency, 'Ticker': ticker_list,
			 'Source': database, 'Release': release, 'Concept': concept}
		df = pd.DataFrame(d)

		return df

	def m_series_concept(self, ticker: str) -> Tuple[str, str]:
		"""
		Method to find the concept of a given series
		"""

		# Get series of the ticker
		s = self.mbdb.FetchOneSeries(ticker)

		# Get the "concept" this series is associated with
		region_key = s.Metadata.GetFirstValue("RegionKey")
		if region_key is None:
			short_concept = ''
			long_concept = ''
		else:
			# Also get some information about this metadata value
			keyMetaInfo = self.mbdb.GetMetadataInformation("RegionKey")
			long_concept = keyMetaInfo.GetValuePresentationText(region_key)
			short_concept = region_key

		return short_concept, long_concept

	def m_discontinued(self, ticker: str):
		"""
		Method with the sole purpose to check if a series has been discontinued or not
		"""

		# Get the series
		series = self.mbdb.FetchOneSeries(ticker)

		# Get the state
		series_state = series.Metadata.GetFirstValue('EntityState')

		if series_state == 0:
			discontinued_tf = False
		elif series_state == 4:
			discontinued_tf = True
		else:
			discontinued_tf = None

		return discontinued_tf

	def m_get_replacement_ticker(self, ticker: str) -> str:
		"""
		Method that returns a replacement ticker for a discontinued series if it exists
		"""
		# Get the series of the old ticker
		series = self.mbdb.FetchOneSeries(ticker)

		# Get replacement comment and ticker
		replacement_comment = series.Metadata.GetFirstValue('EntityDiscontinuedComment')
		replacement_ticker_tuple = series.Metadata.GetValues('EntityDiscontinuedReplacements')

		if replacement_comment is None and len(replacement_ticker_tuple) == 0:
			# No replacement found
			replacement_ticker = ''
		else:
			# Check that we only get one ticker back then print info
			assert len(replacement_ticker_tuple) == 1
			print(replacement_comment)
			replacement_ticker = replacement_ticker_tuple[0]

		return replacement_ticker

	def m_get_frequency(self, ticker: str) -> str:
		"""
		Method with the sole purpose of returning the frequency of a series
		"""

		series = self.mbdb.FetchOneSeries(ticker)

		# freq_int: macrobond_api_constants.SeriesFrequency = series.Frequency

		freq_str = series.Metadata.GetFirstValue('Frequency')

		return freq_str

	def m_get_title(self, ticker: str):
		"""
		Get title of a series if it exist
		:param ticker: str
		:return:
		"""

		s = self.mbdb.FetchOneSeries(ticker)

		# Assert all is well
		try:
			assert s.IsError is False
		except AssertionError:
			print(f'Error: {s.ErrorMessage}')
			return None

		return s.Title

	def m_release_date(self, ticker: str, date_option: str, full_date_format_tf: bool = False) -> dt.datetime:
		"""
		Enter Macrobond series name and get next or previous release date back, if it exists otherwise return 1900-01-01
		"""
		date_option_list = ['next', 'previous']

		if date_option.lower() not in date_option_list:
			raise KeyError(f"Invalid date option. Expected: {date_option_list}")

		if date_option.lower() == 'next':
			mb_date_option = 'NextReleaseEventTime'
		else:
			mb_date_option = 'LastReleaseEventTime'

		# Get the series & metadata
		s = self.mbdb.FetchOneSeries(ticker)
		m = s.Metadata.GetFirstValue('Release')

		if m is not None:
			r = self.mbdb.FetchOneEntity(m)
			date = r.MetaData.GetFirstValue(mb_date_option)

			if date is not None:
				if full_date_format_tf:
					x = dt.datetime.combine(date=date.date(),
											time=date.time(),
											tzinfo=date.tzinfo)
				else:
					x = date.date()
			else:
				if full_date_format_tf:
					x = dt.datetime(1900, 1, 1, 0, 0, 0, 0, pytz.UTC)
				else:
					x = dt.date.fromisoformat('1900-01-01')
		else:
			if full_date_format_tf:
				x = dt.datetime(1900, 1, 1, 0, 0, 0, 0, pytz.UTC)
			else:
				x = dt.date.fromisoformat('1900-01-01')

		return x

	def m_get_metadata(self, ticker: str, option: str):
		"""
		Enter macrobond series & option and get some metadata back
		"""
		# This method could be further developed by help of:
		# https://help.macrobond.com/technical-information/the-macrobond-api-for-python/#imetadata
		# https://help.macrobond.com/technical-information/common-metadata/

		# We name a few options of metadata that we (possibly) can extract otherwise we get None back
		'''
		metadata_options = ['Region',
		                    'Currency',
		                    'DisplayUnit',
		                    'Class',
		                    'ForecastCutoffDate',
		                    'MaturityDate',
		                    'MaturityDays',
		                    'RateMethod',
		                    'IHCategory',
		                    'IHInfo',
		                    'LastModifiedTimeStamp',
		                    'LastReleaseEventTime',
		                    'NextReleaseEventTime',
		                    'Database',
		                    'EntityState',
		                    'EntityType',
		                    'Frequency',
		                    'FullDescription',
		                    'OriginalCurrency',
		                    'OriginalEndDate',
		                    'OriginalFrequency',
		                    'OriginalStartDate']
		'''

		# Description of few of the above
		'''
		EntityState: Tells if a series is active or discontinued.
		OriginalCurrency: The currency of a series before any transformations or calculations. This is set when you have requested the series using a unified series request, which can involve automatic currency conversion.
		OriginalEndDate: The end date of a series before any transformations or calculations. This is set when you have requested the series using a unified series request, which can involve transformations like frequency conversion.
		OriginalFrequency: The frequency of a series before any transformations or calculations. This is set when you have requested the series using a unified series request, which can involve transformations like frequency conversion.
		OriginalStartDate: The start date of a series before any transformations or calculations. This is set when you have requested the series using a unified series request, which can involve transformations like frequency conversion.
		'''

		# Get the series
		s = self.mbdb.FetchOneSeries(ticker)

		# Then get the metadata option that we want to look at
		m = s.Metadata.GetFirstValue(option)

		# Return data
		if m is not None:
			return m
		else:
			return None

	def m_get_region(self, region: str, short_input: bool = True) -> str:
		"""
		Convert either short region name to long or otherwise
		From short region name to long, e.g. f_get_region(region='br', short_input=True) -> 'Brazil'
		From long region name to short, e.g. f_get_region(region='Brazil', short_input=False) -> 'br'
		"""
		reg_map, reg_map_inverse = self.f_region_map()

		if short_input:
			output_region = reg_map.get(region)
		else:
			output_region = reg_map_inverse.get(region)

		return output_region

	@staticmethod
	def f_unpack_series(series) -> pd.Series:
		"""
		Function used to simply unpack timeseries
		"""
		# Convert dates
		p_end_dates = pd.to_datetime([date.strftime('%Y-%m-%d') for date in series.DatesAtEndOfPeriod])

		return pd.Series(series.Values, index=p_end_dates)

	@staticmethod
	def f_create_bbg_ticker(bbg_ticker: [str], **kwargs) -> [str]:
		"""
		Function which sole purpose is to format Bloomberg ticker as macrobond tickers

		Example input:
		bbg_ticker = ['MXWO Index', 'MXWO000G index', 'ECSURPUS index']
		kwargs = {'BBG_Fields': ['PX_LAST', 'PX_OPEN', '']}
		"""

		bbg_field_tf = False
		bbg_fields = []

		try:
			assert type(bbg_ticker) is list
		except AssertionError as ae:
			print(f"Ticker input must be a list. Error: {ae}")
			return []

		# Extract kwargs
		for key, val in kwargs.items():
			if key.lower() == 'bbg_fields':
				bbg_field_tf = True
				bbg_fields: list = kwargs.get('BBG_Fields')
			else:
				raise KeyError(f'Kwargs key: {key} not defined')

		# If we use fields then make sure lengths are correct
		if bbg_field_tf:
			try:
				assert len(bbg_ticker) == len(bbg_fields)
			except AssertionError as ae:
				print(f'Length of bbg_ticker must be the same as bbg_field. Error: {ae}')
				return []

		# Pre-allocate output array
		macrobond_tickers = []

		if bbg_fields:
			for ticker, field in zip(bbg_ticker, bbg_fields):

				# If a specific field entry is empty then we dont add that one (some series dont need a field)
				if len(field) == 0:
					macrobond_tickers.append(f'ih:bl:{ticker.lower()}')
				else:
					macrobond_tickers.append(f'ih:bl:{ticker.lower()}:{field.lower()}')

		else:
			for ticker in bbg_ticker:
				macrobond_tickers.append(f'ih:bl:{ticker.lower()}')

		return macrobond_tickers

	@staticmethod
	def f_region_map():
		"""
		List of regions with code and description as a dictionary
		Region shortnames can be on link found below
		https://www.macrobond.com/region-list/
		Based on two leter ISO 3166 codes
		"""
		region_map = {'asia': 'Asia',
					  'asiapjp': 'Asia + Japan',
					  'asiaxjp': 'Asia ex Japan',
					  'asxmc': 'Asia ex Mainland China',
					  'apac': 'Asia Pacific',
					  'apacxjp': 'Asia Pacific ex Japan',
					  'au': 'Australia',
					  'br': 'Brazil',
					  'ca': 'Canada',
					  'cn': 'China',
					  'dk': 'Denmark',
					  'devasia': 'Developing Asia',
					  'dvmkts': 'Developed Markets',
					  'emkts': 'Emerging Markets',
					  'eueu': 'EU',
					  'eu': 'Euro Area',
					  'europe': 'Europe',
					  'fi': 'Finland',
					  'fr': 'France',
					  'de': 'Germany',
					  'hk': 'Hong Kong',
					  'in': 'India',
					  'it': 'Italy',
					  'jp': 'Japan',
					  'latam': 'Latin America',
					  'mfivasia': 'Major Five Asia',
					  'nordic': 'Nordic Countries',
					  'noram': 'North America',
					  'no': 'Norway',
					  'opec': 'OPEC Members',
					  'sg': 'Singapore',
					  'za': 'South Africa',
					  'kr': 'South Korea',
					  'es': 'Spain',
					  'se': 'Sweden',
					  'tw': 'Taiwan',
					  'gb': 'United Kingdom',
					  'us': 'United States'}

		# We may invert the dictionary since entries are unique
		region_map_inverse = {v: k for k, v in region_map.items()}

		return region_map, region_map_inverse
