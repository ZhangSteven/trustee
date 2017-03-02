# coding=utf-8
# 
import configparser, os
from config_logging.file_logger import get_file_logger



# initialize config object
if not 'config' in globals():
	config = configparser.ConfigParser()
	config.read('trustee.config')



def get_current_directory():
	"""
	Get the absolute path to the directory where this module is in.

	This piece of code comes from:

	http://stackoverflow.com/questions/3430372/how-to-get-full-path-of-current-files-directory-in-python
	"""
	return os.path.dirname(os.path.abspath(__file__))



# initialize logger
if not 'logger' in globals():
	logger = get_file_logger(os.path.join(get_current_directory(), 'trustee.log'), 
								config['logging']['log_level'])



def get_datemode():
	# for xlrd package, data mode for windows is 0.
	return 0



def get_input_directory():
	"""
	Where the input files reside.
	"""
	global config
	if config['data']['input'].strip() == '':
		return get_current_directory()

	return config['data']['input']



def retrieve_or_create(port_values, key):
	"""
	retrieve or create the holding objects (list of dictionary) from the 
	port_values object, the holding place for all items in the portfolio.
	"""

	if key in port_values:	# key exists, retrieve
		holding = port_values[key]	
	else:					# key doesn't exist, create
		if key in ['bond', 'equity', 'cash_accounts', 'expense']:
			holding = []
		else:
			# not implemented yet
			logger.error('retrieve_or_create(): invalid key: {0}'.format(key))
			raise ValueError('invalid_key')

		port_values[key] = holding

	return holding
