# coding=utf-8
# 
# Read holdings from the trustee file. For a sample, see 
# samples/new_12229.xlsx.
#
# 本程序是一个赶工的产物，因为时间不够。实际上，程序读取的 holding 文件来
# 自于2017年以后的Trustee报表中holding的部分，只是将其中的HTM部分Copy&Paste
# 到一个新的文件中，去除了其中的section header的部分，使其便于读取。
#
# 程序目的是生成一个具有下列格式的CSV文件（每一行代表一个Position）：
#
# 列数：1				2			3			4			8
# 数值：Currency		Accounting	Quantity	Identifier	Amortized Cost
#
#		11					17
#		Market Value Local	Portfolio Code
#
# 其中：Market Value Local = Quantity * Amortized Cost / 100
#
# 除上述列以外，其他列均为空。另外，文件没有Header，第一行就代表第一个Position，
# 
# 为实现上述目的，对于每一个组合我们需要读取两个文件：
#
#	1. Geneva local position apprasail：读取最新的Quantity，样本见
# 		samples/12229_local_appraisal_sample1.xls
#	2. Trustee holding：含有上月末的Amortized Cost，样本见
# 		samples/new_12229.xlsx
#
# 读取完成后，将二者合一，将最新的Quantity和上月末的Amortized Cost合在最终的
# CSV文档中。



from small_program.read_file import read_file
from trustee.holding import get_security_id_map
from trustee.utility import get_output_directory, get_input_directory
from trustee.geneva import read_line
import logging
logger = logging.getLogger(__name__)



class PositionNotFound(Exception):
	pass



def read_line_trustee(ws, row, fields):
	position = {}
	for i in range(16):
		cell_value = ws.cell_value(row, i)
		if isinstance(cell_value, str):
			cell_value = cell_value.strip()

		if i == 0:
			position['Identifier'] = get_identifier(cell_value)

		elif i == 15:
			position['Amortized Cost'] = cell_value
	# end for loop

	return position



def get_identifier(description):
	"""
	Create position identifier based on trustee position description. Normally
	the description is 'ISIN code + description', such as:

	HK0000134780 FarEast Horizon5.75%

	In this case, the identifier will be 'HK0000134780 HTM'

	However, sometimes trustee uses other code instead of ISIN, such as:

	HSBCFN13014 NEW WORLD 6%

	In this case, we need to transform the code 'HSBCFN13014' to ISIN code
	first, then create the identifier.
	"""
	code = description.split()[0]
	try:
		return get_security_id_map()[code] + ' HTM'
	except KeyError:
		return code + ' HTM'



def update_amortized_cost(geneva_holding, trustee_holding):
	for position in trustee_holding:
		try:
			p = find_geneva_position(geneva_holding, position['Identifier'])
			p['Amortized Cost'] = position['Amortized Cost']

		except PositionNotFound:
			pass



def find_geneva_position(geneva_holding, invest_id):
	for position in geneva_holding:
		if position['InvestID'] == invest_id:
			return position

	print('{0} not found in geneva holding'.format(invest_id))
	raise PositionNotFound()



def get_filename(portfolio_code):
	return 'f3321Custom.gw1.local_HTM.' + portfolio_code + '.inc'



def get_portfolio_code(geneva_holding):
	return geneva_holding[0]['Portfolio']



def is_cash_position(geneva_position):
	if geneva_position['Group2'] == 'Cash and Equivalents':
		return True
	else:
		return False



def is_AFS_position(geneva_position):
	"""
	Tell whether a geneva position is an available for sale (AFS) position.
	"""
	if 'HTM' in geneva_position['InvestID']:
		return False
	else:
		return True



def add_double_quote(string_item):
	return '\"'+string_item+'\"'



def write_upload_csv(geneva_holding, output_dir=get_output_directory()):
	with open(join(output_dir, get_filename(get_portfolio_code(geneva_holding))), 
					'w', newline='') as csvfile:

		# note there is requirement for the upload file that all non numerical
		# fields are double quoted. So we use the quoting parameter.
		# For more quoting information, see this link:
		#
		# https://pymotw.com/2/csv/
		#
		file_writer = csv.writer(csvfile, delimiter=',', quoting=csv.QUOTE_NONNUMERIC)
	
		for position in geneva_holding:
			if is_cash_position(position) or is_AFS_position(position):
				continue

			# print(position['InvestID'])
			row = []
			for i in range(17):
				if i == 0:	# currency
					item = position['Group1']
				elif i == 1:
					item = 'Held to Maturity'
				elif i == 2:
					item = position['Quantity']
				elif i == 3:	
					item = position['InvestID']
				elif i == 5: # description (optional)
					item = position['ExtendedDescription']
				elif i == 7:
					item = position['Amortized Cost']
				elif i == 10: # market value local
					item = position['Amortized Cost']*position['Quantity']/100.0
				elif i == 16:
					item = position['Portfolio']
				else:
					item = None

				# if isinstance(item, str) and item != '':
				# 	item = add_double_quote(item)

				row.append(item)

			file_writer.writerow(row)




if __name__ == '__main__':
	import argparse, sys, glob, csv
	from os.path import join, isdir, exists
	import logging.config
	logging.config.fileConfig('logging.config', disable_existing_loggers=False)

	parser = argparse.ArgumentParser(description='Create custom amortized cost \
										upload file based on Geneva local position \
										and trustee file.')
	parser.add_argument('--trustee', help='trustee file', required=True)
	parser.add_argument('--geneva', help='geneva local position appraisal file', required=True)
	args = parser.parse_args()

	input_file = join(get_input_directory(), args.trustee)
	trustee_holding, row_in_error = read_file(input_file, read_line_trustee, starting_row=2)

	input_file = join(get_input_directory(), args.geneva)
	geneva_holding, row_in_error = read_file(input_file, read_line)

	update_amortized_cost(geneva_holding, trustee_holding)
	write_upload_csv(geneva_holding)
