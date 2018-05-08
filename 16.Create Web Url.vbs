Set objshell = createobject("wscript.shell")
Set weburl = objshell.createshortcut("H:\google.url")
weburl.targetpath = "http://www.google.com"
Weburl.save