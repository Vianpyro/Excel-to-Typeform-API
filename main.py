# -*- coding:utf-8 -*-
import json
import subprocess
import sys
import urllib.request as urlreq

########################################
# IMPORT LIBRARIES
########################################

def install_library(library) -> None:
    print(f'Installing {library}...')
    try:
        subprocess.Popen(['pip', 'install', library])
    except:
        raise ImportError(f'Unable to find the library {library}!')
    else:
        print(f'Successfuly installed {library}.')

# Import the "selenium" library which allows Python to use a web browser
try:
    import selenium
except:
    install_library('selenium')
    import selenium

########################################
# DOWNLOAD AND UNZIP THE LATEST DRIVER
########################################
# Retrieve latest firefox geckodriver.
web_browser = 'mozilla'
repository = 'geckodriver'
base_request_link = f'https://api.github.com/repos/{web_browser}/{repository}/releases/latest'

try:
    response = urlreq.urlopen(base_request_link)
    data = json.loads(response.read())
    latest_version = data['tag_name']
    print(web_browser, repository, data['tag_name'])
except:
    raise InterruptedError('Unable to access to the internet')

# Detect user Operating System.
os_name = sys.platform
os_bits = '64' if sys.maxsize > 2**32 else '32'

# Prepare to download the driver.
base_download_link = f'https://github.com/{web_browser}/{repository}/releases/download/{latest_version}/{repository}-{latest_version}-'

if os_name == 'linux':
    download_link = base_download_link + 'linux' + os_bits + '.tar.gz'
elif os_name == 'win32':
    download_link = base_download_link + 'win' + os_bits + '.zip'
elif os_name == 'darwin':
    download_link = base_download_link + 'macos.tar.gz'
else:
    raise EnvironmentError('This program can not run on this OS:', os_name)

# Download the driver.
print(f'Downloading {web_browser} {repository} {latest_version}...')
try:
    urlreq.urlretrieve(download_link, f'{web_browser}_{repository}' + '.zip' if os_name == 'win32' else '.tar.gz')
except:
    raise InterruptedError(f'Unable to download {web_browser} {repository} {latest_version} :(.')
else:
    print(f'Successfuly downloaded {web_browser} {repository} {latest_version}!')
