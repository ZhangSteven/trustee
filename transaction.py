# coding=utf-8
# 
# From a trustee NAV excel file, from the transaction worksheet, get
# trade information from it.
#

from xlrd import open_workbook
from xlrd.xldate import xldate_as_datetime
from trustee.utility import logger, get_input_directory, get_datemode, \
						get_output_directory, retrieve_or_create
from trustee.holding import get_security_id_map
from jpm.open_jpm import is_blank_line
from DIF.open_dif import convert_datetime_to_string
from bochk.open_bochk import retrieve_date_from_filename
from datetime import datetime
import re, csv



class BadAccountingTreatment(Exception):
	pass

class UnrecognizedHoldingLine(Exception):
	pass

class InvalidFundName(Exception):
	pass



def read_transaction(ws, port_values):
	"""
	Transaction data is grouped into 3 parts, namely Purchase, Sale and
	Foreign Currency Commitment, as follows:

	Purchase:
	I. Fixed Deposit
	II. Debt Securities
	III. Equities

	Sale:
	I. Debt Securities
	II. Equities

	Foreign Currency Commitment
	"""
	port_values['portfolio_id'], row = get_portfolio_id(ws)
	while (row < ws.nrows):
			
		if is_purchase_section(ws, row):
			logger.debug('read_transaction(): purchase section at row: {0}'.format(row))
			row = read_section(ws, row, 'buy', port_values)
			
		elif is_sale_section(ws, row):
			logger.debug('read_transaction(): sale section at row: {0}'.format(row))
			row = read_section(ws, row, 'sell', port_values)
			
		elif is_fx_section(ws, row):
			logger.debug('read_transaction(): FX section at row: {0}'.format(row))
			row = read_FX(ws, row, port_values)
			break	# the last section

		else:
			row = row + 1



def read_section(ws, row, action, port_values):
	"""
	Read a Puchase or Sale section.

	Assumption: 
	
	1. A Purchase or Sale section always consist of an equity sub section,
	even if it is empty.

	2. The equity sub section is the last sub section of a Purchase or Sale
	section.
	"""
	while (row < ws.nrows):
			
		# search the first column
		cell_value = ws.cell_value(row, 0)

		if isinstance(cell_value, str) and 'fixed deposit' in cell_value.lower():
			logger.debug('read_section(): fixed desposition sub section at row: {0}'.format(row))
			row = read_fixed_deposit(ws, row, port_values)
		elif isinstance(cell_value, str) and 'debt securities' in cell_value.lower():
			logger.debug('read_section(): bond sub section at row: {0}'.format(row))
			row = read_bond_section(ws, row, action, port_values)
		elif isinstance(cell_value, str) and 'equities' in cell_value.lower():
			logger.debug('read_section(): equity sub section at row: {0}'.format(row))
			row = read_equity_section(ws, row, action, port_values)
			break
		else:
			row = row + 1

	return row



def read_bond_section(ws, row, action, port_values):
	"""
	Read a bond sub section 
	"""
	transactions = retrieve_or_create(port_values, 'bond_transactions')
	fields, row = get_bond_fields(ws, row)
	while (row < ws.nrows):
		cell_value = ws.cell_value(row, 0)
		if isinstance(cell_value, str):
			cell_value = cell_value.strip()

		if is_blank_line(ws, row):
			row = row + 1
			continue

		m = re.search('[iv]+\.\s*equities', cell_value.lower())
		if not m is None:
			break

		t = {}
		t['action'] = action
		t['portfolio_id'] = port_values['portfolio_id']
		i = 0
		for fld in fields:
			cell_value = ws.cell_value(row, i)
			if isinstance(cell_value, str):
				cell_value = cell_value.strip()

			i = i + 1
			if fld == 'empty' or fld == 'currency' and i == 1:
				continue

			if fld in ['trade date', 'value date']:
				try:
					t[fld] = xldate_as_datetime(cell_value, 0)
				except:
					print(row, i, cell_value)
					import sys
					sys.exit(1)

			elif fld == 'description':
				t['security_id'], t['description'] = get_id_description(cell_value)

			elif fld == 'reference code' and isinstance(cell_value, float):
				t[fld] = str(int(cell_value))
			else:
				t[fld] = cell_value

		transactions.append(t)
		row = row + 1

	return row



def is_purchase_section(ws, row):
	cell_value = get_cell_value(ws, row)
	if isinstance(cell_value, str) and cell_value.lower().startswith('purchase'):
		return True
	else:
		return False



def is_sale_section(ws, row):
	cell_value = get_cell_value(ws, row)
	if isinstance(cell_value, str) and cell_value.lower().startswith('sale'):
		return True
	else:
		return False



def is_fx_section(ws, row):
	cell_value = get_cell_value(ws, row)
	if isinstance(cell_value, str) and cell_value.lower().startswith('foreign currency'):
		return True
	else:
		return False



def get_cell_value(ws, row):
	"""
	"""
	cell_value = ws.cell_value(row, 0)
	if isinstance(cell_value, str) and cell_value.strip() == '':
		return ws.cell_value(row, 1)
	else:
		return cell_value



def get_bond_fields(ws, row):
	"""
	Read the bond transaction fields from row, then return the fields
	and the next row number after reading the fields.

	We hardcoded the fields here, but it is better to read it each time,
	otherwise when the field order changes, this function will lead to
	incorrrect results.
	"""
	fields = ['currency', 'empty', 'trade date', 'empty', 'value date', 
			'empty', 'empty', 'description', 'empty', 'empty',
			'empty', 'empty', 'empty', 'empty', 'reference code',
			'empty', 'broker', 'empty', 'empty', 'currency',
			'empty', 'amount', 'empty', 'empty', 'empty',
			'price', 'empty', 'empty', 'cost', 'empty',
			'empty', 'bought interest', 'empty', 'empty', 'fx rate',
			'empty', 'empty', 'cost HKD equivalent', 'empty', 'empty',
			'empty', 'bought interest HKD equivalent', 'empty',
			'effective yield']

	while (row < ws.nrows):
			
		# search the first column
		cell_value = ws.cell_value(row, 0)
		if cell_value == 'CCY':
			return fields, row+1
		
		row = row + 1



def get_id_description(cell_value):
	"""
	The security's description contains both the security id and description,
	separate them into two parts.
	"""
	tokens = cell_value.split()
	return tokens[0], cell_value[len(tokens[0]):].strip()



def read_equity_section(ws, row, action, port_values):
	return row+1	# dummy only, needs implementation



def read_fixed_deposit(ws, row, port_values):
	return row+1



def read_FX(ws, row, port_values):
	return row+1



def get_report_name(ws):
	"""
	Each worksheet in the NAV file is a report, get the report name.
	"""
	col = 0
	while col < ws.ncols:
		cell_value = ws.cell_value(2, col)
		if isinstance(cell_value, str):
			if cell_value.strip() != '':
				return cell_value.strip()

		col = col + 1

	return ''



def map_portfolio_id(fund_name):
	p_map = {
		'CLT-CLI HK BR (Class A-HK) Trust Fund (Sub-Fund-Bond)':'12229',
		'CLT-CLI Macau BR (Class A-MC)Trust Fund (Sub-Fund-Bond)':'12366',
		'CLT-CLI HK BR (Class A-HK)Trust Fund(Sub-Fund-Trading Bond)':'12528',
		'CLT-CLI Macau BR (Class G-MC)Trust Fund (Sub-Fund-Bond)':'12548',
		'CLI HK BR (Class G-HK) Trust Fund (Sub-Fund-Bond)':'12630',
		'CLT-CLI HK BR Trust Fund (Capital) (Sub-Fund-Bond) ':'12732',
		'CLT-CLI Overseas Trust Fund (Capital) (Sub-Fund-Bond)':'12733',
		'CLT-CLI HK BR (Class A-HK) Trust Fund - Sub Fund I':'12734'
	}
	return p_map[fund_name]



def get_portfolio_id(ws):
	"""
	Find the fund name in a worksheet then map it to a portfolio id.
	"""
	row = 0
	while row < ws.nrows:
		cell_value = ws.cell_value(row, 0)
		if isinstance(cell_value, str) and cell_value.lower().startswith('fund name:'):
			return map_portfolio_id(cell_value[10:].strip()), row

		row = row + 1



def read_file(filename):
	"""
	Open a NAV file, search for the sheet that contains the transactions,
	then read and return the list of transactions.
	"""
	wb = open_workbook(filename=filename)
	for ws in wb.sheets():
		if get_report_name(ws) == 'SECURITIES TRANSACTION IMPLEMENTED':
			port_values = {}
			read_transaction(ws, port_values)
			return port_values['bond_transactions']



def accumulate_transactions(total_transactions, transactions):
	for t in transactions:
		total_transactions.append(t)



def map_security_id(transactions):
	"""
	Map trustee security code to investment code used in Geneva.
	"""
	i_map = get_security_id_map()	
	for t in transactions:
		if t['security_id'] in i_map:
			t['security_id'] = i_map[t['security_id']]

	return transactions



def write_simple_transaction_csv(filename, transactions, output_dir=get_output_directory()):
	with open(join(output_dir, filename), 'w', newline='') as csvfile:
		file_writer = csv.writer(csvfile, delimiter='|')

		# pick all fields that HTM bond have
		fields = ['portfolio_id', 'action', 'security_id', 'description', 
					'trade date', 'value date', 'currency', 'reference code',
					'amount', 'price', 'fx rate']
		file_writer.writerow(fields)
		
		for t in transactions:
			row = []
			for fld in fields:
				try:
					item = t[fld]
					if fld in ['trade date', 'value date']:
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

	transactions = []
	for input_file in files:
		accumulate_transactions(transactions, read_file(input_file))

	write_simple_transaction_csv('trades.csv', map_security_id(transactions))