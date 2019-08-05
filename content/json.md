Title: json
Date: 2019-03-11 21:46
Author: Lulef
Category: Sammlung
Tags: python
Slug: json
Status: published

```
import json
conf = json.load(open('jsonfile.json'))
print(type(conf))
if conf['use_dropbox']:
    print('jep ->', conf['use_dropbox'])
else:
    print('NOPE')
for k, v in conf.items():
    print(k, v)
```
