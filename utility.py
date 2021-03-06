# coding=utf-8
# 
import configparser, os
# from config_logging.file_logger import get_file_logger



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



# # initialize logger
# if not 'logger' in globals():
# 	logger = get_file_logger(os.path.join(get_current_directory(), 'trustee.log'), 
# 								config['logging']['log_level'])



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



def get_geneva_input_directory():
	"""
	Where the input files reside.
	"""
	global config
	if config['data']['geneva_input'].strip() == '':
		return get_current_directory()

	return config['data']['geneva_input']




def get_output_directory():
	"""
	Where to put the output csv files.
	"""
	global config
	if config['data']['output'].strip() == '':
		return get_current_directory()

	return config['data']['output']



def get_exchange_file():
	global config
	return config['data']['exchange_file']



def retrieve_or_create(port_values, key):
	if not key in port_values:
		# print('create')
		port_values[key] = []

	# print('retrieve')
	return port_values[key]