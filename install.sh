#!/bin/bash

echo -e '=> Installing Updates ...\n'
sudo apt update
#sudo apt upgrade -yecho -e '=> Installing needed packages ...\n'
sudo apt -y upgrade 
sudo apt install python3 -y
sudo apt install pip

# you can use the pip3 part without sudo, just make sure the packets above are installed
pip3 install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
pip3 install --upgrade pandas gspread_formatting pygsheets pytube
