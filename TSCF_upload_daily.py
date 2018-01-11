# coding=utf-8
# 
# Create TSCF uplad file for the data fields -- "Days to LYE" 
# -- days to the last year end for those securities which don't have
# a "days to maturity" data field.
#
# 目标是生成具有下列格式的CSV文件。除前两行header row之外，从第三行
# 开始每一行代表一个data field，
# 
# Upload Method,INCREMENTAL,,,,
# Field Id,Security Id Type,Security Id,Account Code,Numeric Value,Char Value
# CD024,4,HK0000337607,,35,35
# CD024,4,USG8116KAB82,,35,35
# CD024,4,XS1499209861,,35,35
# 
# 
# 上述文件的目的是为某几只债券提供一个特定的data field (CD024)，即从
# 去年底（last year end）至今日有多少天。因为这些债券没有 "days to
# maturity" 这个data field，所以需要借助这个CD024来计算从去年底至今日
# 的interest income。
# 
# 这些债券的列表见 "bond list.config" 文件。
# 
# Create another TSCF upload file, similar to the one above, but for
# Exchange rate upload (CD025, exchange rate from bond currency to
# portfolio currency HKD).
# 


from trustee.utility import get_output_directory, get_input_directory, \
							get_current_directory, get_exchange_file
from trustee.TSCF_upload import consolidate_security, get_ISIN_from_investID, \
							get_holding_from_files
from trustee.sftp import upload
from datetime import date, timedelta
from os.path import join
import logging, csv, configparser
logger = logging.getLogger(__name__)



class InvalidCurrency(Exception):
	pass

class ExchangeRateNotFound(Exception):
	pass



def get_days_since_LYE():
	"""
	Workout the number of days since the last year end (LYE) to today.
	"""
	lye = date(date.today().year-1, 12, 31)
	return (date.today() - lye).days



def get_bond_list():
	"""
	Read the bond list configuration file 
	"""
	cfg = configparser.ConfigParser()
	cfg.read(join(get_current_directory(), 'bond list.config'))

	return cfg['info']['bond_list'].split(',')



def get_exchange_rate(currency_description):
	if "exchange" not in get_exchange_rate.__dict__:
		get_exchange_rate.exchange = configparser.ConfigParser()
		get_exchange_rate.exchange.read(get_exchange_file())

	c_map = {
		'United States Dollar':'USD',
		'Chinese Renminbi Yuan':'CNY'
	}

	if currency_description == 'Hong Kong Dollar':
		return 1.0

	else:
		try:
			currency = c_map[currency_description]
			return get_exchange_rate.exchange['Exchange'][currency+'HKD']

		except KeyError:
			logger.exception('get_exchange_rate(): currency {0} has problem'.
								format(currency_description))
			raise ExchangeRateNotFound()



def date_to_string():
	"""
	Convert today's date to string, say 2018-1-9, it is converted to
	string '20180109'.
	"""
	t = date.today()
	return str(t.year*10000 + t.month*100 + t.day)



def get_lye_file_name():
	return 'f3321tscf.days_since_lye.' + date_to_string() + '.inc'



def get_exc_file_name():
	return 'f3321tscf.exchange_rate.' + date_to_string() + '.inc'



def write_upload_csv_lye(bond_list=get_bond_list(), output_dir=get_output_directory()):
	"""
	Create the "days since last year end" upload file.
	"""
	upload_file = join(output_dir, get_lye_file_name())
	with open(upload_file, 'w', newline='') as csvfile:

		file_writer = csv.writer(csvfile, delimiter=',')
		file_writer.writerow(['Upload Method','INCREMENTAL','','','',''])
		file_writer.writerow(['Field Id','Security Id Type','Security Id','Account Code',
								'Numeric Value','Char Value'])

		days_since_lye = get_days_since_LYE()
		for isin in bond_list:
			row = ['CD024','4',isin,'',days_since_lye,days_since_lye]
			file_writer.writerow(row)

	return upload_file



def write_upload_csv_exc(holding, output_dir=get_output_directory()):
	"""
	Create the "exchange rate" upload file.
	"""
	upload_file = join(output_dir, get_exc_file_name())
	with open(upload_file, 'w', newline='') as csvfile:

		file_writer = csv.writer(csvfile, delimiter=',')
		file_writer.writerow(['Upload Method','INCREMENTAL','','','',''])
		file_writer.writerow(['Field Id','Security Id Type','Security Id','Account Code',
								'Numeric Value','Char Value'])

		for security in holding:
			row = ['CD025','4',get_ISIN_from_investID(security['InvestID']),
					'',get_exchange_rate(security['Group1']),
					get_exchange_rate(security['Group1'])]

			file_writer.writerow(row)
		# end of for loop

	return upload_file



if __name__ == '__main__':
	# Testing code here.
	# 
	# Actual code to do periodic upload is in do_upload.py
	# 
	import logging.config
	logging.config.fileConfig('logging.config', disable_existing_loggers=False)

	logger.info('start to create TSCF upload file.')
	upload_file_lye = write_upload_csv_lye()
	upload_file_exc = write_upload_csv_exc(consolidate_security(get_holding_from_files()))