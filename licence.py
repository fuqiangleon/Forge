import base64
import platform
import hashlib


cert = platform.machine() + platform.node() + platform.platform() + platform.processor() + platform.release()
#print cert

yanxx = 'eDg2enpieHh6eFdpbmRvd3MtWFAtNS4xLjI2MDAtU1AzeDg2IEZhbWlseSA2IE1vZGVsIDIzIFN0ZXBwaW5nIDEwLCBHZW51aW5lSW50ZWxYUA=='

certnew = base64.decodestring(yanxx)



m = hashlib.md5(base64.encodestring(certnew))

print m.hexdigest()
f = file('./licence', 'w')
s = ''
for t in range(len(m.hexdigest()) - 1, -1, -1):
#    print m.hexdigest()[t]
    s += m.hexdigest()[t]
f.write(s)
f.close()
print s




