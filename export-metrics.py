import qbittorrentapi
import openpyxl
import os
import datetime
from dotenv import load_dotenv

time_format = "%d%B%Y"
intervals = (
('weeks', 604800),  # 60 * 60 * 24 * 7
('days', 86400),    # 60 * 60 * 24
('hours', 3600),    # 60 * 60
('minutes', 60),
('seconds', 1),)

def format_bytes(size):
    # 2**10 = 1024
    power = 2**10
    n = 0
    power_labels = {0 : '', 1: 'K', 2: 'M', 3: 'G', 4: 'T'}
    while size > power:
        size /= power
        n += 1
    return f"{round(size, 2)} {power_labels[n]}B"

def format_seconds(seconds, granularity=2):
    result = []
    for name, count in intervals:
        value = seconds // count
        if value:
            seconds -= value * count
            if value == 1:
                name = name.rstrip('s')
            result.append("{} {}".format(value, name))
    return ', '.join(result[:granularity])

def create_workbook():
	new_workbook = openpyxl.Workbook()
	worksheet = new_workbook.active
	worksheet.title = 'Torrents'
	header_row = (
		'Name',
		'Size',
		'Added On',
		'Seeded For',
		'Uploaded',
		'Ratio',
		'Category',
		'Hash',
	)
	worksheet.append(header_row)	
	# This makes the rows bold
	for row in worksheet.iter_rows():
		for cell in row:
			cell.font = openpyxl.styles.Font(bold = True)
	return(new_workbook)

def add_torrents_to_workbook(workbook, torrents):
	worksheet = workbook['Torrents']
	for torrent in torrents:
		if "PSA" in torrent.name:
			new_row = (
				torrent.name.replace(".mkv",""),
				format_bytes(torrent.size),
				datetime.datetime.fromtimestamp(torrent.added_on).strftime("%d %B %Y, %I:%M %p"),
				format_seconds(torrent.seeding_time),
				format_bytes(torrent.uploaded),
				round(torrent.ratio, 2),
				"None" if torrent.category == "" else torrent.category,
				torrent.hash,
			)
			worksheet.append(new_row)
	return(workbook)

def save_workbook(workbook):
	filename = f'qBittorrent_Metrics_{datetime.datetime.today().strftime(time_format)}.xlsx'
	workbook.save(filename)
	print(f'Saved {filename}.')

def main():
	BASEDIR = os.path.abspath(os.path.dirname(__name__))
	load_dotenv(os.path.join(BASEDIR,"credentials.env"))
	qbit_hostname = os.getenv("QBIT_HOSTNAME")
	qbit_port = os.getenv("QBIT_PORT")
	qbit_username = os.getenv("QBIT_USERNAME")
	qbit_password = os.getenv("QBIT_PASSWORD")
	qbt_client = qbittorrentapi.Client(host=qbit_hostname, port=qbit_port, username=qbit_username, password=qbit_password)
	try:
		qbt_client.auth_log_in()
	except qbittorrentapi.LoginFailed as e:
		print(e)
	all_torrents = qbt_client.torrents_info()
	#Create the workbook
	workbook = create_workbook()
	#Add torrents to the workbook
	workbook = add_torrents_to_workbook(workbook, all_torrents)	
	#Save the workbook
	save_workbook(workbook)

if __name__=="__main__":
    main()