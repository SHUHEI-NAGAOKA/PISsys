import os
import requests
import shutil

destination_folder = '/Users/nagaokashuuhei/Desktop/sys_practice'

#ダウンロードフォルダのパスを変数に格納
download_folder = os.path.join(os.path.expanduser('~'), 'Downloads')

if not os.path.exists(download_folder):
    print(f"エラー: ダウンロードフォルダ '{download_folder}' が見つかりません。")

files = []
for entry in os.listdir(download_folder):
	full_path = os.path.join(download_folder, entry)
	if os.path.isfile(full_path):
		files.append(full_path)

if not files:
	print('ダウンロードフォルダにファイルが見つかりませんでした')

latest_file = max(files, key=os.path.getmtime)

file_name = os.path.basename(latest_file)

