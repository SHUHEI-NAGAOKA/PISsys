import os
import requests
import shutil

destination_folder = '/Users/nagaokashuuhei/Desktop/sys_practice/PISsys'

#ダウンロードフォルダのパスを変数に格納
download_folder = os.path.join(os.path.expanduser('~'), 'Downloads')

# if not os.path.exists(download_folder):
#     print(f"エラー: ダウンロードフォルダ '{download_folder}' が見つかりません。")

# files = []
# for entry in os.listdir(download_folder):
# 	full_path = os.path.join(download_folder, entry)
# 	if os.path.isfile(full_path):
# 		files.append(full_path)

# if not files:
# 	print('ダウンロードフォルダにファイルが見つかりませんでした')
# paitiants_list = ['岡本','手塚']

# files = []
# for entry in os.listdir(download_folder):
# 	full_path = os.path.join(download_folder, entry)
# 	if os.path.isfile(full_path):
# 		files.append(full_path)
# 		print(full_path)

# for paitiant in reversed(paitiants_list):
# 	latest_file = max(files, key=os.path.getmtime)
# 	file_name = os.path.basename(latest_file)
# 	source_path = latest_file
# 	destination_path = os.path.join(destination_folder, file_name)

# 	try:
# 		shutil.move(source_path, destination_path)
# 	except Exception as e:
# 		print('ファイルの移動中にエラーが発生しました')

# 	files.remove(latest_file)

def move_file(paitiants_list):
	files = []
	for entry in os.listdir(download_folder):
		full_path = os.path.join(download_folder, entry)
		if os.path.isfile(full_path):
			files.append(full_path)
			print(full_path)

	for paitiant in reversed(paitiants_list):
		latest_file = max(files, key=os.path.getmtime)
		file_name = os.path.basename(latest_file)
		source_path = latest_file
		destination_path = os.path.join(destination_folder, file_name)

	try:
		shutil.move(source_path, destination_path)
	except Exception as e:
		print('ファイルの移動中にエラーが発生しました')

	files.remove(latest_file)

# latest_file = max(files, key=os.path.getmtime)

# #ファイルの名称をfile_nameに格納。あとで移動元のディレクトリの名称とjoinさせる際に使用する
# file_name = os.path.basename(latest_file)

# #移動もとのファイルのパスを変数に格納
# source_path = latest_file

# #ファイルを移動する処理をした際のパスを指定
# destination_path = os.path.join(destination_folder, file_name)

# try:
# 	shutil.move(source_path, destination_path)
# 	print(f"最も新しいファイル '{file_name}' を '{destination_folder}' に移動しました。")
# 	print(f"移動元: {source_path}")
# 	print(f"移動先: {destination_path}")
# except Exception as e:
# 	print(f"ファイルの移動中にエラーが発生しました: {e}")