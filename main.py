#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import xlrd
import xlwt
import re
from datetime import datetime






"""
	sheet-> Upload videos
"""
UPLOADSHEET = "upload_list"
FILEINDEX = 0
FILEDIR = 1
FILENAME = 2
TITLE = 3
DESCRIPTION = 4
KEYWORD = 5
PRIVACYSTATUS = 6

"""
	sheet-> Upload caption
"""

CAPTIONSHEET = "caption_list"
CAPTIONINDEX = 0
CAPTIONFILEDIR = 1
CAPTIONFILENAME = 2
CAPTIONLANG = 3
CAPTIONNAME = 4
CAPTIONVIDEOID = 5


EXCELFILE = "upload_list.xls"
wb = xlrd.open_workbook(EXCELFILE)

CAPTIONEXTENSION = {'.srt'}




def read_xlm(task_flag):
	
	all_data = []
	if task_flag == '1':
		print('---------------------------------Import list of video info from excel file----------------------------\n\n')

		sheetstruc = wb.sheet_by_name(UPLOADSHEET)
		for row in range(1, sheetstruc.nrows):

			idx_ = sheetstruc.cell_value(row,FILEINDEX)
			filedir_ = sheetstruc.cell_value(row,FILEDIR)
			filename_ = sheetstruc.cell_value(row,FILENAME)
			title_ = sheetstruc.cell_value(row,TITLE)
			desc_ = sheetstruc.cell_value(row,DESCRIPTION)
			keyw_ = sheetstruc.cell_value(row,KEYWORD)
			priv_ = sheetstruc.cell_value(row,PRIVACYSTATUS)
			
			all_data.append({'row':row,
				'id': idx_,
				'file_dir': filedir_,
				'filename': filename_,
				'title': title_,
				'description': desc_,
				'keyword': keyw_,
				'privacy_status': priv_})

	elif task_flag == '2':
		print('---------------------------------Import list of transcipt info from excel file----------------------------\n\n')
	
		sheetstruc = wb.sheet_by_name(CAPTIONSHEET)
		for row in range(1, sheetstruc.nrows):

			idx_ = sheetstruc.cell_value(row,CAPTIONINDEX)
			filedir_ = sheetstruc.cell_value(row,CAPTIONFILEDIR)
			filename_ = sheetstruc.cell_value(row,CAPTIONFILENAME)
			lang_ = sheetstruc.cell_value(row,CAPTIONLANG)
			name_ = sheetstruc.cell_value(row,CAPTIONNAME)
			videoid_ = sheetstruc.cell_value(row,CAPTIONVIDEOID)
			videoid_ = re.sub(r'\S+be/', '', videoid_)
			
			all_data.append({'row':row,
				'id': idx_,
				'file_dir': filedir_,
				'filename': filename_,
				'lang': lang_,
				'name': name_,
				'videoid': videoid_})


	else:
		print("wrong task flag")
		exit()

		
	

	return(all_data)


def upload_video(videos):

	for video in videos:
		filename_template = os.path.join(video['file_dir'],video['filename'])
		upload_command_template = 'python upload_video.py --file='+ filename_template  
		if video['title'] != "":
			upload_command_template = upload_command_template + " --title=" + str((video['title'])) 
		if video['description'] != "":
			upload_command_template = upload_command_template + " --description=" +str(video['description']) 
		if video['keyword'] != "":		
			upload_command_template = upload_command_template + " --keywords=" +str(video['keyword']) 
		if video['privacy_status'] != "":
			upload_command_template = upload_command_template + " --privacyStatus=" +str(video['privacy_status'])
		#print(type(video["keyword"]))
		print("---------------------------------start uploading " + video['filename'] + "---------------------------------")
		print(upload_command_template)
		os.system(upload_command_template)
		print('-------------------------------------------------------------------------------------------------------\n')



def upload_transcript(transcripts):

	for transcript in transcripts:

		if transcript['filename'] == '':
			print 'no filename specified for transcript at excel id: ' + transcript['id']
			continue

		filename_template = os.path.join(transcript['file_dir'],transcript['filename'])

		for e_ext in CAPTIONEXTENSION:
			if e_ext in filename_template and os.path.isfile(filename_template):
				newfile = filename_template.replace(e_ext, '.bin')
				tmp = filename_template
				os.rename(filename_template, newfile)
				filename_template = newfile


		upload_command_template = 'python upload_caption.py --videoid='+ str(transcript['videoid'])  + ' --file='+filename_template

		if transcript['name'] != "":
			upload_command_template = upload_command_template + " --name=" + str((transcript['name'])) 
		if transcript['lang'] != "":
			upload_command_template = upload_command_template + " --language=" +str(transcript['lang']) 
		
		upload_command_template = upload_command_template + ' --action=upload'
		print("---------------------------------start uploading " + transcript['filename'] + "---------------------------------")
		print(upload_command_template)
		os.system(upload_command_template)
		print('-------------------------------------------------------------------------------------------------------\n')
		os.rename(filename_template,tmp)

def output_video_list(title,youtube_id):

	output_name = 'uploaded_video_link.txt' 
	file = open(output_name,'a')
	file.write( datetime.now().strftime('%Y-%m-%d %H:%M:%S') +' , '+title+' , '+'https://youtu.be/'+youtube_id+'\n')
	file.close()


#python captions.py --videoid='<video_id>' --name='<name>' --file='<file>' --language='<language>' --action='action'

def main():
	flag = 0
	global file
	while(flag==0):
		command = raw_input("enter [1-2]\n1.Upload video\n2.Upload Caption\n")
		if command == '1':
			print ('the Upload video task is chosen')
			all_videos = read_xlm(command)
			upload_video(all_videos)
			flag = 1
		elif command == '2':
			print ('the Upload Caption task is chosen')
			all_transcripts = read_xlm(command)
			upload_transcript(all_transcripts)
			flag = 1
		else:
			print ('wrong command, try again!!!!')




	




if __name__ == '__main__':
	try:
		main()
	except KeyboardInterrupt:
		logging.warn("\n\nCTRL-C detected, shutting down....")
		sys.exit(ExitCode.OK)



