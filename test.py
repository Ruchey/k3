# -*- coding: utf-8 -*-

import time
from instapy_cli import client

username = 'shkafkupebel'
password = 'A621e238'
image = r'd:\dj\skb\media\photolog\JPG\73.jpg'

image_files = [r'd:\dj\skb\media\photolog\JPG\72.jpg', r'd:\dj\skb\media\photolog\JPG\71.jpg', r'd:\dj\skb\media\photolog\JPG\70.jpg']

with client(username, password, bypass_suspicious_attempt=True) as cli:
    cli.upload(image, story=True)