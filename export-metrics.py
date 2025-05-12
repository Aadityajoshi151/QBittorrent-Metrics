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

def format_completion_time(timestamp):
    try:
        if timestamp <= 0:
            return "Not completed"
        return datetime.datetime.fromtimestamp(timestamp).strftime("%d %B %Y, %I:%M %p")
    except (OSError, ValueError, OverflowError):
        return "Not completed"

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
		'Path',
		'Hash',
		'Magnet Link',
		'Time Active',
		'Downloaded',
		'Completed On',
	)
	worksheet.append(header_row)	
	# This makes the rows bold
	for row in worksheet.iter_rows():
		for cell in row:
			cell.font = openpyxl.styles.Font(bold = True)
	# This set the column widths
	worksheet.column_dimensions['A'].width = 55
	worksheet.column_dimensions['B'].width = 9
	worksheet.column_dimensions['C'].width = 25
	worksheet.column_dimensions['D'].width = 20
	worksheet.column_dimensions['E'].width = 10
	worksheet.column_dimensions['F'].width = 5
	worksheet.column_dimensions['G'].width = 8
	worksheet.column_dimensions['H'].width = 18
	worksheet.column_dimensions['I'].width = 5
	worksheet.column_dimensions['J'].width = 11
	worksheet.column_dimensions['K'].width = 20
	worksheet.column_dimensions['L'].width = 11
	worksheet.column_dimensions['M'].width = 25
	return(new_workbook)

def add_torrents_to_workbook(workbook, torrents):
	worksheet = workbook['Torrents']
	for torrent in torrents:
		new_row = (
			torrent.name.replace(".mkv",""),
			format_bytes(torrent.size),
			datetime.datetime.fromtimestamp(torrent.added_on).strftime("%d %B %Y, %I:%M %p"),
			format_seconds(torrent.seeding_time),
			format_bytes(torrent.uploaded),
			round(torrent.ratio, 2),
			"None" if torrent.category == "" else torrent.category,
			torrent.save_path,
			torrent.hash,
			torrent.magnet_uri,
			format_seconds(torrent.time_active),
			format_bytes(torrent.downloaded),
			format_completion_time(torrent.completion_on),
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