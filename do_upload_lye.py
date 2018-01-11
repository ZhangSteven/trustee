# coding=utf-8
# 
# Scheduled job: Generate "days since last year end" data field for
# a few securities and upload to Bloomberg AIM.
# 

from trustee.TSCF_upload_daily import write_upload_csv_lye
from trustee.sftp import upload
import logging

logger = logging.getLogger(__name__)



if __name__ == '__main__':
	import logging.config
	logging.config.fileConfig('logging.config', disable_existing_loggers=False)

	logger.info('start to upload LYE file.')
	result = upload([write_upload_csv_lye()])
	if len(result['pass']) == 1:
		logger.info('upload OK: {0}'.format(result['pass'][0]))
	else:
		logger.error('upload failed: {0}'.format(result['fail'][0]))