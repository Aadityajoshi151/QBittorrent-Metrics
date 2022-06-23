# QBittorrent-Metrics

A simple script to save qbittorrent metrics in an excel file.

# Usage:

- Clone the repo.
- Install requirements using the command `pip install -r requirements.txt`.
- Create a file in the project directory named "credentials.env" and save your qbittorrent web-UI credentials in the following format:

```
QBIT_HOSTNAME = "localhost"
QBIT_PORT = "8080"
QBIT_USERNAME = "your-username-goes-here"
QBIT_PASSWORD = "your-password-goes-here"
```

If you are not running this on localhost, then enter your ip address and port number (if required) else just add username and password.

- Run `python export.py` or `python3 export.py` (on linux) to generate the excel file.

## What is exported in the excel file?

| Property   | Description                                                           |
| ---------- | --------------------------------------------------------------------- |
| Name       | Full name of torrent. Removes _.mkv_ file extension if present        |
| Size       | Torrent size in MB/GB                                                 |
| Added on   | Date and time on which torrent was added in the client                |
| Seeded For | For how much time the torrent was seeded                              |
| Uploaded   | How much data is uploaded from your end. Size in MB/GB                |
| Ratio      | Share ratio                                                           |
| Category   | If the specefic torrent is categorized. _None_ if no category present |
| Hash       | Complete hash of the torrent                                           |

These are the metrics that I need hence this is what the script exports. You can modify the script according to your own needs. Refer the [qbitorrent-api docs](https://qbittorrent-api.readthedocs.io/en/latest/) for more info.
