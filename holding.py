# coding=utf-8
# 
# This file is used to parse the diversified income fund excel files from
# trustee, read the necessary fields and save into a csv file for
# reconciliation with Advent Geneva.
#

from xlrd import open_workbook
from xlrd.xldate import xldate_as_datetime
from trustee.utility import logger, get_input_directory, get_datemode
from DIF.open_holding import read_bond_fields
from DIF.utility import retrieve_or_create
from bochk.open_bochk import retrieve_date_from_filename
import re



class BadAccountingTreatment(Exception):
	pass

class UnrecognizedHoldingLine(Exception):
	pass



def read_holding(ws, port_values):
	"""
	Copied from DIF.open_holding.py, read_holding() function, with modifications
	to accomodate the difference between the trustee NAV file and DIF NAV file.

	The structure of holdings data is

	Section (Debt Securities - USD):
		sub section (Held to Maturity xxx):
			holding1
			holding2
			...

		sub section (Available for Sales xxx):
			...

		Total (end of section)
	Section
		sub section:

		...

	This version of read_holding() function does not care about equity holdings,
	it reads bond holdings only.
	"""
	row = 0
	while (row < ws.nrows):
			
		# search the first column
		cell_value = ws.cell_value(row, 0)

		if isinstance(cell_value, str) and 'debt securities' in cell_value.lower():
			logger.debug('read_holding(): bond section: {0}'.format(cell_value))

			currency = read_currency(cell_value)
			fields, n = read_bond_fields(ws, row)	# read the bond field names
			row = read_section(ws, row+n, fields, 'bond', currency, port_values)

		else:
			row = row + 1



def read_section(ws, row, fields, asset_class, currency, port_values):
	"""
	Copied from DIF.open_holding, read_section() function, with some changes,
	e.g., trustee has 'HTM' and 'AFS', but not trading.

	asset_class: either 'bond' or 'equity', this is because later the
		other functions will use these keys to retrieve holdings.

	Return the row number after the whole section.
	"""
	holding = retrieve_or_create(port_values, asset_class)

	while (row < ws.nrows):
		cell_value = ws.cell_value(row, 0)
		if not isinstance(cell_value, str):
			row = row + 1
			continue

		# a subsection looks like (i) Held to Maturity (Transfer from ...)
		if sub_section_begins(cell_value):
			if 'held to maturity' in cell_value.lower():
				accounting_treatment = 'HTM'
				
			elif 'available for sale' in cell_value.lower():
				accounting_treatment = 'AFS'

			else:
				logger.error('read_section(): invalid accounting treament at row {0}: {1}'.
								format(row, cell_value))
				raise BadAccountingTreatment()

			row = read_sub_section(ws, row+1, accounting_treatment, fields, asset_class, currency, holding)

		elif section_ends(cell_value):
			break

		else:
			row = row + 1

	return row



def read_sub_section(ws, row, accounting_treatment, fields, asset_class, currency, holding):
	"""
	Copied from DIF.open_holding, read_sub_section() function, changes:

	1. Reading of the sub section won't stop at a blank line, because there
		maybe blank lines within a subsection. The reading stops at the beginning
		of the next subsection or the end of the parenet section.

	2. Return the row number where the reading stops.
	"""
	while (row < ws.nrows):
		cell_value = ws.cell_value(row, 0)
		if not isinstance(cell_value, str) or cell_value.strip() == '':
			row = row + 1
			continue

		if sub_section_begins(cell_value) or section_ends(cell_value):
			break

		m = re.search('\([A-Za-z0-9]+\)', cell_value)
		if m is None:
			logger.error('read_sub_section(): unrecognized line at row {0}'.format(row))
			raise UnrecognizedHoldingLine()

		security_id = m.group(0)[1:-1].strip()
		if len(security_id) != 12:
			logger.warning('read_sub_section(): security id {0} at row {1} is not ISIN'.
							format(security_id, row))

		security = {}
		security['isin'] = security_id
		security['name'] = cell_value[len(security_id)+2:].strip()
		security['currency'] = currency
		security['accounting_treatment'] = accounting_treatment

		column = 2	# now start reading fields in column C
		bond_valid = True
		for field in fields:
			cell_value = ws.cell_value(row, column)
			if field == 'par_amount' and \
				(cell_value == 0 or isinstance(cell_value,str) and cell_value.strip() == ''):
				bond_valid = False
				break	# ignore this holding
			
			if field in ['coupon_start_date', 'maturity_date']:
				try:
					cell_value = xldate_as_datetime(cell_value, get_datemode())
				except:
					logger.warning('read_sub_section(): invalid date value {0} at row {1}'.
						format(cell_value, row))

			security[field] = cell_value
			column = column + 1
		# end of for loop

		if bond_valid:
			holding.append(security)

		row = row + 1
	# end of while

	return row



def read_currency(cell_value):
	"""
	Copied from DIF.open_holding.py, read_currency() function, add mapping
	for 'HK$'.

	Read the currency from the cell containing a section start, such as
	'V. Debt Securities - US$  (債務票據- 美元)',
	'V. Debt Securities - SGD  (債務票據- 星加坡元)',
	'X. Equities - USD (股票-美元)'

	From the above strings, the function return USD, SGD, USD
	"""
	temp_list = cell_value.split('-')
	token = temp_list[1]
	temp_list = token.split('(')
	currency = str.strip(temp_list[0])

	if currency == 'US$':
		return 'USD'
	elif currency == 'HK$':
		return 'HKD'

	return currency



def sub_section_begins(cell_value):
	m = re.search('\([iv]+\)([A-Za-z\s]+)\(.*\)', cell_value)
	if m is None:
		return False
	else:
		return True



def section_ends(cell_value):
	if cell_value.lower().startswith('total'):
		return True
	else:
		return False




if __name__ == '__main__':
	import argparse
	parser = argparse.ArgumentParser(description='Read trustee NAV file and create csv output for Geneva reconciliation.')
	parser.add_argument('nav_file')
	args = parser.parse_args()

	import os, sys
	input_file = os.path.join(get_input_directory(), args.nav_file)
	if not os.path.exists(input_file):
		print('{0} does not exist'.format(input_file))
		sys.exit(1)

	wb = open_workbook(filename=input_file)
	ws = wb.sheet_by_name('Portfolio Val.')
	port_values = {}
	port_values['date'] = retrieve_date_from_filename(args.nav_file)
	print(port_values['date'])
	read_holding(ws, port_values)
