# coding=utf-8
# 
# Upload files to a sftp site. The code relies on WinSCP program to
# do the FTP work.
# 
# SFTP parameters like site IP address, user name and password are
# set in the sftp.config file.
# 
# There are two known issues:
# 
# 1. For a SFTP site, please use WinSCP to connect at least once
# 	so that the server's host key can be verified.
# 
# 2. There cannot be any spaces in file name or directory name in
# 	the upload file path.
# 

import time, os
from os.path import join
from subprocess import run, TimeoutExpired, CalledProcessError
from trustee.utility import get_current_directory
import logging

logger = logging.getLogger(__name__)



# initialized only once when this module is first imported by others
if not 'config' in globals():
	import configparser
	config = configparser.ConfigParser()
	config.read(join(get_current_directory(), 'sftp.config'))



def upload(file_list):
	"""
	Call winscp.com to execute the sftp upload job.
	"""
	winscp_script, winscp_log = create_winscp_files(file_list)
	try:
		args = [get_winscp_path(), '/script={0}'.format(winscp_script), \
				'/log={0}'.format(winscp_log)]

		result = run(args, timeout=get_timeout(), check=True)
	except TimeoutExpired:
		logger.error('upload(): timeout {0} expired'.format(get_timeout()))
	except CalledProcessError:
		logger.error('upload(): upload job did not complete successfully')
	except:
		logger.error('upload(): some other error occurred')
		logger.exception('upload():')

	result = {}
	result['pass'] = read_log(winscp_log)
	result['fail'] = get_fail_list(file_list, result['pass'])
	return result



def create_winscp_files(file_list):
	time_string = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
	tokens = time_string.split()[1].split(':')
	suffix = time_string.split()[0] + 'T' + tokens[0] + tokens[1] + tokens[2]
	
	return create_winscp_script(file_list, suffix), \
			create_winscp_log(suffix)



def create_winscp_script(file_list, suffix, directory=None):
	"""
	Create a script file to be loaded by WinSCP.com, the file
	contains instructions of ftp, like:

	open sftp://demo:password@test.rebex.net/
	cd pub/example
	get ConsoleClient.png
	exit

	"""
	if directory is None:
		directory = get_winscp_script_directory()

	script_file = join(directory, 'run-sftp_{0}.txt'.format(suffix))
	with open(script_file, 'w') as f:
		f.write('open sftp://{0}:{1}@{2}\n'.format(get_sftp_user(), \
				get_sftp_password(), get_sftp_server()))

		for file in file_list:
			f.write('put {0}\n'.format(file))

		f.write('exit')

	return script_file



def create_winscp_log(suffix, directory=None):
	"""
	Create an empty log file for winscp.com to write log messages to,
	otherwise there will be an error.
	"""
	if directory is None:
		directory = get_winscp_log_directory()

	log_file = join(directory, 'log_{0}.txt'.format(suffix))
	with open(log_file, 'w') as f:	# just create an empty file
		pass

	return log_file



def read_log(winscp_log):
	"""
	Look for successful transfer records in the winscp logfile, then report
	which files are successfully transferred, and the date and time those
	transfers are completed.

	The successful transfer records are in the following format (get and put)

	> 2016-12-29 17:20:40.652 Transfer done: '<file full path>' [xxxx]

	The starting symbol can be '>', '<', '.', '!', depending on the type of
	the record.

	If it is a get, then 'file full path' will be the remote directory's
	file path. If it is a put, then 'file full path' will be the local
	directory's file path.
	"""
	successful_list = []
	with open(winscp_log) as f:
		for line in f:
			tokens = line.split()
			if len(tokens) < 6:
				continue

			if tokens[3] == 'Transfer' and tokens[4] == 'done:':
				successful_list.append(tokens[5][1:-1])

	return successful_list



def get_fail_list(file_list, pass_list):
	fail_list = []
	d = {key:0 for key in file_list}
	for file in pass_list:
		try:
			d[file] = d[file] + 1
		except KeyError:
			logger.error('get_fail_list(): {0} not in file list, but in pass list'.
							format(file))
			pass

	for file, value in d.items():
		if value == 0:
			fail_list.append(file)

	return fail_list



def get_winscp_path():
	global config
	return config['winscp']['application']



def get_winscp_script_directory():
	global config
	return config['winscp']['script_dir']



def get_winscp_log_directory():
	global config
	return config['winscp']['log_dir']



def get_timeout():
	global config
	return float(config['sftp']['timeout'])



def get_sftp_server():
	global config
	return config['sftp']['server']



def get_sftp_user():
	global config
	return config['sftp']['username']



def get_sftp_password():
	global config
	return config['sftp']['password']




if __name__ == '__main__':
	"""
	For testing only.
	"""
	file_list = [join(get_current_directory(), 'samples', 'upload_sample1.txt'), 
					join(get_current_directory(), 'samples', 'upload_sample2.txt')]

	result = upload(file_list)
	print(result)