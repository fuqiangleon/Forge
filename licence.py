import base64
# import platform
import hashlib
import sys
import os

# node = uuid.getnode()
# mac = uuid.UUID(int = node).hex[-12:]
# print mac




mac = None    
if sys.platform == "win32":
        for line in os.popen("ipconfig /all"):
#             print line
            if line.lstrip().startswith("Physical Address"):
                mac = line.split(":")[1].strip().replace("-", ":")
                break
print mac

yanxx = 'eDg2enpieHh6eFdpbmRvd3MtWFAtNS4xLjI2MDAtU1AzeDg2IEZhbWlseSA2IE1vZGVsIDIzIFN0ZXBwaW5nIDEwLCBHZW51aW5lSW50ZWxYUA=='

certnew = base64.decodestring(mac)

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




