import json,re
page=open('page.html').read(); data=open('data.json').read()
out=page.replace('/*__DATA__*/[]',data)
open('scouting-lsc-2027.html','w').write(out)
# standalone for GitHub Pages
m=re.search(r'<title>.*?</title>',out); title=m.group(0)
rest=out.replace(title,'',1)
links=re.findall(r'<link[^>]+>',rest)
for l in links: rest=rest.replace(l,'',1)
head='<!doctype html>\n<html lang="es">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">\n'+title+'\n'+'\n'.join(links)+'\n<style>*,*::before,*::after{box-sizing:border-box}body{margin:0}img{max-width:100%}[hidden]{display:none!important}</style>\n'
style=re.search(r'<style>.*?</style>',rest,re.S).group(0); rest=rest.replace(style,'',1)
open('index.html','w').write(head+style+'\n</head>\n<body>\n'+rest.strip()+'\n</body>\n</html>\n')
print(len(out))
