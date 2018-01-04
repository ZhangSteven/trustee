# coding=utf-8
# 
# Create TSCF uplad file for the two data fields -- "Yield at Cost"
# and "Purchase Cost" -- for the HTM portfolios in Bloomberg AIM.
#
# 目标是生成两个具有下列格式的CSV文件（除前两行header row之外，从第三行
# 开始每一行代表一个Position）：
#
# 第一个文件，上传一个组合中每一个Position的Yield at Cost,
# 
# Upload Method,INCREMENTAL,,,,
# Field Id,Security Id Type,Security Id,Account Code,Numeric Value,Char Value
# CD021,4,XS1556937891,12229,7.889,7.889
# 
# 第二个文件，上传一个组合中每一个Position的Purchase Cost,
# 
# Upload Method,INCREMENTAL,,,,
# Field Id,Security Id Type,Security Id,Account Code,Numeric Value,Char Value
# CD022,4,XS1556937891,12229,98.89,98.89
# 
# 
# 为实现上述目的，对于每一个组合我们需要读取两个文件：
#
#	1. Geneva local position apprasail：读取最新的holding，知道每一个
# 		组合内有哪些Position。样本见
# 		samples/12229_local_appraisal_sample1.xls
# 
#	2. Jones holding：所有HTM组合中每一只债券的Yield at Cost, 以及
# 		Purchase Cost, 这里我们假设截至Jones创建此文件为止，一只债
# 		券即使分配到多个组合，它在每一个组合中的Yield at Cost 以及
# 		Purchase Cost都相同。本文件在 
# 		samples/Jones Holding 2017.12.20.xlsx
#



from small_program.read_file import read_file
from trustee.holding import get_security_id_map
from trustee.utility import get_output_directory, get_input_directory, \
							get_current_directory
from trustee.geneva import read_line
from trustee.quick_holding import is_cash_position, is_AFS_position
import logging
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
		if is_cash_position(position) or is_AFS_position(position):
			continue

		try:
			p = find_jones_position(jones_holding, get_ISIN_from_investID(position['InvestID']))
			position['Yield at Cost'] = p['Yield at Cost']
			position['Purchase Cost'] = p['Purchase Cost']

		except PositionNotFound:
			pass



def get_ISIN_from_investID(geneva_invest_id):
	return geneva_invest_id.split()[0]



def find_jones_position(jones_holding, isin):
	for position in jones_holding:
		if position['ISIN'] == isin:
			return position

	print('{0} not found in jones holding'.format(isin))
	raise PositionNotFound()



def get_filename(portfolio_code):
	return 'f3321tscf.' + portfolio_code + '.inc'



def get_portfolio_code(geneva_holding):
	return geneva_holding[0]['Portfolio']



def write_upload_csv(geneva_holding, output_dir=get_output_directory()):
	with open(join(output_dir, get_filename(get_portfolio_code(geneva_holding))), 
					'w', newline='') as csvfile:

		file_writer = csv.writer(csvfile, delimiter=',')
		file_writer.writerow(['Upload Method','INCREMENTAL','','','',''])
		file_writer.writerow(['Field Id','Security Id Type','Security Id','Account Code',
								'Numeric Value','Char Value'])

		for position in geneva_holding:
			if is_cash_position(position) or is_AFS_position(position):
				continue

			# print(position['InvestID'])
			row1 = ['CD021','4',get_ISIN_from_investID(position['InvestID']),
					position['Portfolio'],position['Yield at Cost'],position['Yield at Cost']]
			row2 = ['CD022','4',get_ISIN_from_investID(position['InvestID']),
					position['Portfolio'],position['Purchase Cost'],position['Purchase Cost']]
			
			file_writer.writerow(row1)
			file_writer.writerow(row2)




if __name__ == '__main__':
	import argparse, sys, glob, csv
	from os.path import join, isdir, exists
	import logging.config
	logging.config.fileConfig('logging.config', disable_existing_loggers=False)

	parser = argparse.ArgumentParser(description='Create TSCF upload \
										file based on Geneva local position \
										and Jones Au\'s file.')

	parser.add_argument('--geneva', help='geneva local position appraisal file', required=True)
	args = parser.parse_args()

	input_file = join(get_current_directory(), 'samples', 'Jones Holding 2017.12.20.xlsx')
	jones_holding, row_in_error = read_file(input_file, read_line_jones)

	input_file = join(get_input_directory(), args.geneva)
	geneva_holding, row_in_error = read_file(input_file, read_line)

	update_position(geneva_holding, jones_holding)
	write_upload_csv(geneva_holding)
