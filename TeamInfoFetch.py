import urllib.request
with urllib.request.urlopen("http://es01.usfirst.org/teams/_search?size=100000") as response:
	html = response.read()
t = "{0}".format(html)
#infostart = t.find("<h3>Info")
#print(infostart)
#t = t[infostart:-1]
#print(t)
