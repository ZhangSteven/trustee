# coding=utf-8
# 
# Read the Geneva positions from Geneva exported local position
# appraisal report.



from small_program.read_file import read_file
from trustee.utility import get_output_directory, get_input_directory
from DIF.open_dif import convert_datetime_to_string
from datetime import datetime
import re



class InvalidDateString(Exception):
	pass



def read_line(ws, row, fields):
	position = {}
	i = 0
	for fld in fields:
		cell_value = ws.cell_value(row, i)
		if isinstance(cell_value, str):
			cell_value = cell_value.strip()

		if fld == 'Portfolio' and isinstance(cell_value, float):
			cell_value = str(int(cell_value))

		elif fld == 'Description':
			m = re.search('\d{2}/\d{2}/\d{2}', cell_value)
			if m is None:
				position['MaturityDate'] = ''
			else:
				position['MaturityDate'] = get_maturity_date(m.group(0))

		position[fld] = cell_value
		i = i + 1

	return position



def get_maturity_date(date_string):
	"""
	Get date from string 'mm/dd/yy'
	"""
	tokens = date_string.split('/')
	try:
		return datetime(int(tokens[2])+2000, int(tokens[0]), int(tokens[1]))
	except:
		logger.warning('get_maturity_date(): unable to convert date string {0}'.format(date_string))
		return ''



def filter_maturity(holding):
	"""
	Filter out all cash positions, and bond with maturity earlier than 
	2016-12-31.
	"""
	new_holding = []
	for position in holding:
		if position['Group1'] == 'Cash and Equivalents' or \
			position['Group2'] == 'Cash and Equivalents' or \
			isinstance(position['MaturityDate'], datetime) and \
			position['MaturityDate'] < datetime(2017,1,1):
			
			continue

		new_holding.append(position)

	return new_holding




def write_bond_holding_csv(holding, filename, output_dir=get_output_directory()):
	with open(join(output_dir, filename+'_output.csv'), 'w', newline='') as csvfile:
		file_writer = csv.writer(csvfile, delimiter='|')
	
		# pick all fields that HTM bond have
		fields = ['Portfolio', 'InvestID', 'ExtendedDescription', 'Quantity', 'UnitCost', 'MaturityDate']

		file_writer.writerow(fields)
		
		for bond in holding:
			row = []
			
			for fld in fields:
				try:
					item = bond[fld]
					if fld == 'MaturityDate' and isinstance(item, datetime):
						item = convert_datetime_to_string(item)
				except KeyError:
					item = ''

				row.append(item)

			file_writer.writerow(row)



if __name__ == '__main__':
	import argparse, sys, glob, csv
	from os.path import join, isdir, exists
	parser = argparse.ArgumentParser(description='Read Geneva position file and create csv output.')
	parser.add_argument('--folder', help='folder containing multiple position files', required=False)
	parser.add_argument('--file', help='input position file', required=False)
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
		holding, row_in_error = read_file(input_file, read_line)

		filename = input_file.split('\\')[-1]	# the file name without path
		print('read file {0}'.format(filename))
		if len(row_in_error) > 0:
			print('some rows in error')

		write_bond_holding_csv(filter_maturity(holding), filename.split('.')[0])

