# coding=utf-8
# 
# Scheduled job: Generate "exchange rate for bond local currency to
# portfolio currency" data field for all securities and upload to 
# Bloomberg AIM.
# 


from trustee.TSCF_upload_daily import write_upload_csv_exc
from trustee.TSCF_upload import consolidate_security, \
									get_holding_from_files
from trustee.utility import get_exchange_file
from trustee.sftp import upload
from os.path import isfile, getmtime
from datetime import datetime, timedelta
import logging

logger = logging.getLogger(__name__)



def exchange_file_exists():
	if isfile(get_exchange_file()):
		return True
	else:
		return False



def modified_within(minutes):
	"""
	Compare the last modified time of the exchange file to the
	time now, see whether it's within X minutes (the input).
	"""
	if datetime.now() - datetime.fromtimestamp(getmtime(get_exchange_file())) \
		> timedelta(minutes=minutes):

		return False

	return True



if __name__ == '__main__':

	import argparse
	parser = argparse.ArgumentParser(description='Create and upload the exchange \
										rate file.')
	parser.add_argument('--minutes', help='if exchange file modified within this \
										amount of time, the upload will be triggered', 
										required=True)
	args = parser.parse_args()

	import logging.config
	logging.config.fileConfig('logging.config', disable_existing_loggers=False)

	# Try finding the last modified time of the exchange rate file,
	# if the file exists and the last modified time is within X minutes
	# from now, the do nothing. Othereise do upload.
	# 
	logger.info('program starts')
	if exchange_file_exists() and modified_within(int(args.minutes)):

		logger.info('start to upload EXC file.')
		result = upload([write_upload_csv_exc(consolidate_security(get_holding_from_files()))])
		if len(result['pass']) == 1:
			logger.info('upload OK: {0}'.format(result['pass'][0]))
		else:
			logger.error('upload failed: {0}'.format(result['fail'][0]))