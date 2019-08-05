Title: Install Python 3.7 on RaspberryPi
Date: 2019-06-24 20:37
Author: Lulef
Category: raspberry
Slug: install-python-3-7-on-raspberrypi
Status: published

```
sudo apt-get update
sudo apt-get install -y build-essential tk-dev libncurses5-dev libncursesw5-dev libreadline6-dev libdb5.3-dev libgdbm-dev libsqlite3-dev libssl-dev libbz2-dev libexpat1-dev liblzma-dev zlib1g-dev libffi-dev
## Compile (takes a while!)
wget https://www.python.org/ftp/python/3.7.3/Python-3.7.3.tar.xz
tar xf Python-3.7.3.tar.xz
cd Python-3.7.3
./configure --prefix=/usr/local/opt/python-3.7.3
make -j 4
## Install
sudo make altinstall
## Make Python 3.7 the default version, make aliases
## sudo ln -s /usr/local/opt/python-3.7.3/bin/pydoc3.7 /usr/bin/pydoc3.7
## sudo ln -s /usr/local/opt/python-3.7.3/bin/python3.7 /usr/bin/python3.7
## sudo ln -s /usr/local/opt/python-3.7.3/bin/python3.7m /usr/bin/python3.7m
sudo ln -s /usr/local/opt/python-3.7.3/bin/pyvenv-3.7 /usr/bin/pyvenv-3.7
sudo ln -s /usr/local/opt/python-3.7.3/bin/pip3.7 /usr/bin/pip3.7
alias python='/usr/bin/python3.7'
alias python3='/usr/bin/python3.7'
ls /usr/bin/python*
cd ..
sudo rm -r Python-3.7.3
rm Python-3.7.3.tar.xz
. ~/.bashrc
## And verify:
python -V
## And if you want to revert:
update-alternatives --config python
```
