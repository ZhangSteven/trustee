# coding=utf-8
# 
# This file is used to parse the diversified income fund excel files from
# trustee, read the necessary fields and save into a csv file for
# reconciliation with Advent Geneva.
#

from xlrd import open_workbook
from xlrd.xldate import xldate_as_datetime
from trustee.utility import logger, get_input_directory, get_datemode, \
						get_output_directory
from DIF.open_holding import read_bond_fields
from DIF.utility import retrieve_or_create
from DIF.open_dif import convert_datetime_to_string
from bochk.open_bochk import retrieve_date_from_filename
from datetime import datetime
import re, csv



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
	port_values['portfolio_id'], row = read_portfolio_id(ws, 0)
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
		
			if field in ['par_amount', 'average_cost', 'amortized_cost', 'book_cost', 
							'interest_bought', 'amortized_value', 'accrued_interest',
							'amortized_gain_loss', 'fx_gain_loss'] \
				and not isinstance(cell_value, float):
				try:
					cell_value = float(cell_value)
				except:
					logger.error('read_sub_section(): data value {0} at row {1} should be float'.
							format(cell_value, row))
					raise ValueError()

			security[field] = cell_value
			column = column + 1
		# end of for loop

		if bond_valid:
			holding.append(security)

		row = row + 1
	# end of while

	return row



def read_portfolio_id(ws, row):
	p_map = {
		'CLI HK BR Trust Fund (Capital) (Sub-Fund-Bond)':'12732',
		'CLI HK BR Trust Fund (Capital)':'12857',
		'CLI HK BR (CLASS A-HK) TRUST FUND (SUB-FUND-BOND)':'12229',
		'CLI HK BR (Class A-HK) Trust Fund - Sub Fund I':'12734',
		'CLI HK BR (Class A-HK) Trust Fund (Sub-Fund-Trading Bond)':'12528',
		'CLI HK BR (CLASS A-HK) TRUST FUND':'11490',
		'CLI MACAU BR (Class A-MC) TRUST FUND (SUB-FUND-BOND)':'12366',
		'CLT-CLI Macau BR (Class A-MC) Trust Fund':'12298',
		'CLI HK BR (Class G-HK) Trust Fund (Sub-Fund-Bond)':'12630',
		'CLI HK BR (Class G-HK) Trust Fund':'12341',
		'CLI MACAU BR (Class G-MC) TRUST FUND (SUB-FUND-BOND)':'12548',
		'CLT-CLI Macau BR (Class G-MC) Trust Fund':'12726',
		'CLI Overseas Trust Fund (Capital) (Sub-Fund-Bond)':'12733',
		'CLI Overseas Trust Fund (Capital)':'12856'
	}

	while (row < ws.nrows):
		cell_value = ws.cell_value(row, 0)
		if isinstance(cell_value, str) and \
			cell_value.strip().lower().startswith('fund name'):
			m = re.search('\(基金名稱\) :(.*)中國人壽', cell_value)
			if m is None:
				logger.error('read_portfolio_id(): failed to get fund name at row {0}, {1}'.
								format(row, cell_value))

			try:
				return p_map[m.group(0)[8:-4].strip()], row
			except KeyError:
				logger.error('read_portfolio_id(): fund name {0} does not find a match'.
								format(m.group(0)[8:-4].strip()))
				raise InvalidFundName()

		row = row + 1





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



def get_custodian(portfolio_id):
	c_map = {
		'12229':'BOCHK',
		'12630':'BOCHK',
		'12366':'BOCHK',
		'12548':'JPM',
		'12528':'BOCHK',
		'12732':'BOCHK',
		'12733':'BOCHK',
		'12734':'BOCHK'
	}

	return c_map[portfolio_id]



def filter_maturity(bond_holding):
	"""
	Filter out bond with maturity earlier than 2016-12-31.
	"""
	new_holding = []
	for bond in bond_holding:
		if not isinstance(bond['maturity_date'], datetime) or \
			bond['maturity_date'] > datetime(2016,12,31):
			new_holding.append(bond)

	return new_holding



def merge_lots(bond_holding):
	"""
	The bond holding may contain multiple lots for the same bond, the function
	merge all of them into one position.
	"""
	merged_positions = []
	merged_isin_list = []
	for bond in bond_holding:
		if bond['isin'] in merged_isin_list:	# a merged position exists
			merge_position(find_position(merged_positions, bond['isin']), bond)
		else:
			merged_isin_list.append(bond['isin'])
			merged_positions.append(bond)

	return merged_positions



def find_position(positions, isin):
	for position in positions:
		if position['isin'] == isin:
			return position

	return None



def merge_position(p1, p2):
	"""
	Merge position p2 into p1. Note that the merged position's accounting
	treatment, trade day FX will be p1's.
	"""
	# try:
	p1['average_cost'] = (p1['par_amount']*p1['average_cost'] + p2['par_amount']*p2['average_cost']) / (p1['par_amount']+p2['par_amount'])
	p1['par_amount'] = p1['par_amount'] + p2['par_amount']
	# p1['amortized_cost'] = p1['amortized_cost'] + p2['amortized_cost']
	# p1['book_cost'] = p1['book_cost'] + p2['book_cost']
	# p1['interest_bought'] = p1['interest_bought'] + p2['interest_bought']
	# p1['amortized_value'] = p1['amortized_value'] + p2['amortized_value']
	# p1['accrued_interest'] = p1['accrued_interest'] + p2['accrued_interest']
	# p1['amortized_gain_loss'] = p1['amortized_gain_loss'] + p2['amortized_gain_loss']
	# p1['fx_gain_loss'] = p1['fx_gain_loss'] + p2['fx_gain_loss']
	# except:
	# 	print(p1['isin'], p1['par_amount'], p2['isin'], p2['par_amount'])
	# 	import sys
	# 	sys.exit(1)



def write_bond_holding_csv(port_values, output_dir=get_output_directory()):
	"""
	Copied from DIF.open_holding.py, write_htm_holding_csv() function, with
	slight modifications.
	"""
	date = convert_datetime_to_string(port_values['date'])
	portfolio_id = port_values['portfolio_id']
	holding_file = join(output_dir, portfolio_id+'_'+date+'_trustee_nav.csv')
		
	with open(holding_file, 'w', newline='') as csvfile:
		file_writer = csv.writer(csvfile, delimiter='|')
		bond_holding = port_values['bond']

		# pick all fields that HTM bond have
		fields = ['name', 'currency', 'accounting_treatment', 'par_amount', 
					'is_listed', 'listed_location', 'fx_on_trade_day', 
					'coupon_rate', 'coupon_start_date', 'maturity_date', 
					'average_cost', 'amortized_cost', 'book_cost', 
					'interest_bought', 'amortized_value', 'accrued_interest', 
					'amortized_gain_loss', 'fx_gain_loss']

		file_writer.writerow(['portfolio', 'date', 'custodian', 'geneva_investment_id', 
								'isin', 'bloomberg_figi'] + fields)
		
		for bond in bond_holding:

			row = [portfolio_id, date, get_custodian(portfolio_id),
					bond['isin']+' HTM', '', '']
			
			for fld in fields:
				try:
					item = bond[fld]
					if fld in ['coupon_start_date', 'maturity_date'] and isinstance(item, datetime):
						item = convert_datetime_to_string(item)
				except KeyError:
					item = ''

				row.append(item)

			file_writer.writerow(row)




if __name__ == '__main__':
	import argparse, sys, glob
	from os.path import join, isdir, exists
	parser = argparse.ArgumentParser(description='Read trustee NAV file and create csv output for Geneva reconciliation.')
	parser.add_argument('--folder', help='folder containing multiple NAV files', required=False)
	parser.add_argument('--file', help='input NAV file', required=False)
	args = parser.parse_args()

	if not args.file is None:
		file = join(get_input_directory(), args.file)
		if not exists(file):
			print('{0} does not exist'.format(file))
			sys.exit(1)

		files = [file]

	elif not args.folder is None:
		folder = join(get_input_directory(), args.folder)
		if not exists(folder) or not isdir(folder):
			print('{0} is not a valid directory'.format(folder))
			sys.exit(1)

		files = glob.glob(folder+'\\*.xls*')

	else:
		print('Please provide either --file or --folder input')
		sys.exit(1)

	for input_file in files:
		filename = input_file.split('\\')[-1]	# the file name without path

		# focus on the bond portfolios (excluding trading bond) first
		if 'sub fund i' in filename.lower() or \
			'bond' in filename.lower() and not 'trading bond' in filename.lower():
			print('read file {0}'.format(filename))
			port_values = {}
			port_values['date'] = retrieve_date_from_filename(filename)
			wb = open_workbook(filename=input_file)
			ws = wb.sheet_by_name('Portfolio Val.')
			read_holding(ws, port_values)
			port_values['bond'] = merge_lots(filter_maturity(port_values['bond']))
			write_bond_holding_csv(port_values)
