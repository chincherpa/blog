Title: pip install opencv
Date: 2018-09-26 06:04
Author: Lulef
Category: Sammlung
Slug: pip-install-opencv
Status: published

<https://www.pyimagesearch.com/2018/09/19/pip-install-opencv/>

## Install Python3

<https://gist.github.com/dschep/24aa61672a2092246eaca2824400d37f>

```
$ sudo apt-get update
$ sudo apt-get install build-essential tk-dev libncurses5-dev libncursesw5-dev libreadline6-dev libdb5.3-dev libgdbm-dev libsqlite3-dev libssl-dev libbz2-dev libexpat1-dev liblzma-dev zlib1g-dev
```
If one of the packages cannot be found, try a newer version number (e.g. `libdb5.4-dev` instead of `libdb5.3-dev`).

### Download and install Python 3.7. When downloading the source code, select the most recent release of Python 3.7, available on the [official site](https://www.python.org/downloads/source/). Adjust the file names accordingly.
```
$ wget https://www.python.org/ftp/python/3.7.X/Python-3.7.X.tar.xz
$ tar xf Python-3.7.X.tar.xz
$ cd Python-3.7.X
$ ./configure
$ make
$ sudo make altinstall
```

#### Install pip on your Raspberry Pi

The Python package manager, “pip”, can be obtained via wget:
