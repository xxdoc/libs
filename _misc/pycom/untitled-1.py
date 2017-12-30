import pefile

pth = "C:\\Documents and Settings\\david\\Desktop\\53701\\fixed_d15.dll_"
pe = pefile.PE(pth)
print "Import Hash: %s" % pe.get_imphash()