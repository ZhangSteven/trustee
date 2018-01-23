# coding=utf-8
# 
# Create TSCF uplad files for the Bloomberg AIM.
# 
# File 1 is for two data fields -- "Yield at Cost" and "Purchase Cost" 
# -- for each position in trustee fixed income portfolios in Bloomberg AIM.
#
# File 1 是具有下列格式的CSV文件。除前两行header row之外，从第三行
# 开始每一行代表一个data field，
# 
# Upload Method,INCREMENTAL,,,,
# Field Id,Security Id Type,Security Id,Account Code,Numeric Value,Char Value
# CD021,4,XS1556937891,12229,7.889,7.889
# CD022,4,XS1556937891,12229,98.89,98.89
# CD023,4,XS1556937891,12229,1289,1289
# 
# 
# 为实现上述目的，我们需要读取下列文件：
#
#	1. Geneva local position apprasail：对于每一个组合我们读取最新的
# 		holding，这样可以知道有哪些Position。样本见
# 		samples/12229_local_appraisal_sample1.xls
# 
#	2. Jones holding：提供了所有HTM组合中每一只债券的Yield at Cost, 
# 		以及Purchase Cost, 这里我们假设截止Jones创建此文件为止，一只
# 		债券即使分配到多个组合，它在每一个组合中的Yield at Cost 以及
# 		Purchase Cost都相同。但是这并不能保证将来也是如此。本文件在 
# 		samples/Jones Holding 2017.12.20.xlsx
#
# 
# File 2 is for another data field -- "days between last year end and
# maturity" -- for each security in trustee fixed income portfolios in
# Bloomberg AIM. Because a bond (except a few) has "days to maturity"
# data field in Bloomberg, therefore combining the two we can calculate
# how many dates between today and the end of last year. File 2 needs to
# be uploaded per security at the beginning of each year.
# 
# 



from small_program.read_file import read_file
from trustee.holding import get_security_id_map
from trustee.utility import get_output_directory, get_geneva_input_directory, \
							get_current_directory
from trustee.geneva import read_line
from trustee.quick_holding import is_cash_position, is_AFS_position
from datetime import date, timedelta
from os.path import join, isfile
from os import listdir
import logging, csv



logger = logging.getLogger(__name__)



class PositionNotFound(Exception):
	pass



def read_line_jones(ws, row, fields):
	position = {}
	for i in range(22):
		cell_value = ws.cell_value(row, i)
		if isinstance(cell_value, str):
			cell_value = cell_value.strip()

		if i == 1:
			position['ISIN'] = cell_value

		elif i == 20:
			position['Purchase Cost'] = cell_value

		elif i == 21:
			position['Yield at Cost'] = cell_value
	# end for loop

	return position



def update_position(geneva_holding, jones_holding):
	"""
	For each position in Geneva holding, find its Yield at Cost
	and Purchase Cost from Jones Holdings, and update that position
	with these two fields.
	"""
	for position in geneva_holding:
		if is_cash_position(position):
			continue

		try:
			p = find_jones_position(jones_holding, get_ISIN_from_investID(position['InvestID']))
			position['Yield at Cost'] = p['Yield at Cost']
			position['Purchase Cost'] = p['Purchase Cost']

		except PositionNotFound:
			pass

		if position['MaturityDate'] == '':
			position['Maturity to Last Year End'] = 0
		else:
			position['Maturity to Last Year End'] = get_days_maturity_LYE(position['MaturityDate'])



def get_days_maturity_LYE(maturity_date):
	"""
	Workout the number of days between the maturity date (of type
	datetime.datetime) and the last year end (LYE).
	"""
	m = date(maturity_date.year, maturity_date.month, maturity_date.day)
	lye = date(date.today().year-1, 12, 31)
	return (m - lye).days



def get_ISIN_from_investID(geneva_invest_id):
	return geneva_invest_id.split()[0]



def find_jones_position(jones_holding, isin):
	for position in jones_holding:
		if position['ISIN'] == isin:
			return position

	print('{0} not found in jones holding'.format(isin))
	raise PositionNotFound()



def get_filename():
	return 'f3321tscf.yield_at_cost.inc'



def get_portfolio_code(geneva_holding):
	return geneva_holding[0]['Portfolio']



def write_upload_csv(geneva_holding, output_dir=get_output_directory()):
	with open(join(output_dir, get_filename()), 'w', newline='') as csvfile:

		file_writer = csv.writer(csvfile, delimiter=',')
		file_writer.writerow(['Upload Method','INCREMENTAL','','','',''])
		file_writer.writerow(['Field Id','Security Id Type','Security Id','Account Code',
								'Numeric Value','Char Value'])

		for position in geneva_holding:
			if is_cash_position(position):
				continue

			# print(position['InvestID'])
			row1 = ['CD021','4',get_ISIN_from_investID(position['InvestID']),
					position['Portfolio'],position['Yield at Cost'],position['Yield at Cost']]
			row2 = ['CD022','4',get_ISIN_from_investID(position['InvestID']),
					position['Portfolio'],position['Purchase Cost'],position['Purchase Cost']]
			
			file_writer.writerow(row1)
			file_writer.writerow(row2)



def get_holding_from_files(input_dir=get_geneva_input_directory()):
	"""
	Based on all local position appraisal files in a directory, then
	read Geneva holdings from them.
	"""
	file_list = [join(input_dir, f) for f in listdir(input_dir) \
					if isfile(join(input_dir, f)) and f.split('.')[-1] == 'xlsx']

	geneva_holding = []
	for file in file_list:
		holding, row_in_error = read_file(file, read_line)
		geneva_holding = geneva_holding + holding

	return geneva_holding



def consolidate_security(position_holding):
	"""
	When there are multiple positions of the same security in the 
	position_holding, we just leave one position in the consolidated 
	holding. This is because when we create the upload file for 
	"maturity date to last year end", it's security level, not 
	position level.
	"""
	holding = []
	for position in position_holding:
		if is_cash_position(position):
				continue

		if not has_position(holding, position):
			holding.append(position)

	return holding



def has_position(holding, position):
	for p in holding:
		if get_ISIN_from_investID(p['InvestID']) == \
			get_ISIN_from_investID(position['InvestID']):
			
			return True

	return False



def write_upload_csv_maturity(holding, output_dir=get_output_directory()):
	with open(join(output_dir, 'f3321tscf.maturity_to_lye.inc'), 'w', newline='') as csvfile:

		file_writer = csv.writer(csvfile, delimiter=',')
		file_writer.writerow(['Upload Method','INCREMENTAL','','','',''])
		file_writer.writerow(['Field Id','Security Id Type','Security Id','Account Code',
								'Numeric Value','Char Value'])

		for position in holding:
			if is_cash_position(position):
				continue

			# print(position['InvestID'])
			row = ['CD023','4',get_ISIN_from_investID(position['InvestID']),
					'', position['Maturity to Last Year End'],
					position['Maturity to Last Year End']]

			file_writer.writerow(row)




if __name__ == '__main__':
	import logging.config
	logging.config.fileConfig('logging.config', disable_existing_loggers=False)

	input_file = join(get_current_directory(), 'samples', 'Jones Holding 2017.12.20.xlsx')
	jones_holding, row_in_error = read_file(input_file, read_line_jones)

	holding = get_holding_from_files()
	update_position(holding, jones_holding)
	write_upload_csv(holding)

	consolidated_holding = consolidate_security(holding)
	write_upload_csv_maturity(consolidated_holding)