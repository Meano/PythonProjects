from urllib2 import urlopen

url = "http://dl.platformio.org/packages/toolchain-xtensa-windows-1.40802.0.tar.gz"

r = urlopen(url)
i = 0
for chunk in r.read(1024):
    i += 1
    print i